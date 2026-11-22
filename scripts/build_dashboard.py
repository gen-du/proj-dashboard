#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""项目推进看板生成器

读取《财政项目推进总台账》Excel，按「项目类别」筛选工作包、按「工作包类别」分类统计，
生成三类图表（里程碑甘特图 / 类别饼图 / 完成情况直方图）的单文件 HTML 看板，
可选导出 Word 到本地（通过本地 editor_sdk）。

用法:
  python3 build_dashboard.py <excel路径> [选项]

选项:
  --category  项目类别筛选（默认: 新系统建设）
  --outdir    输出目录（默认: /workspace）
  --title     看板标题（默认: <项目类别>项目推进看板）
  --today     数据截止日期 YYYY-MM-DD（默认: 系统当天）
  --word      额外导出 Word 文档（需本地 editor_sdk + playwright）
"""
import argparse, datetime, json, os, re, subprocess, sys, tempfile, urllib.request

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
TEMPLATE = os.path.join(SKILL_DIR, 'assets', 'dashboard_template.html')
ECHARTS_CACHE = os.path.join(SKILL_DIR, 'assets', 'echarts.min.js')
ECHARTS_URL = 'https://unpkg.com/echarts@5.5.0/dist/echarts.min.js'

WP_ORDER = ['需求与范围', '方案设计', '采购', '开发', '版本交付', '测试',
            '安全测评', '环境部署', '投产验收', '培训交付', '合规监管']

# 列名 → 内部字段（按「包含」匹配，容忍列名小差异）
COLUMNS = {
    '编号': 'id', '项目类别': 'category', '工作包类别': 'wp', '工作包名称': 'name',
    '里程碑': 'milestone', '交付成果': 'deliver', '验收标准': 'deliver',
    '主责人': 'owner', '协作': 'dep', '前置': 'dep',
    '计划完成日': 'plan', '完成日期': 'plan', '推进状态': 'status',
    '风险级别': 'risk', '风险': 'risknote', '卡点': 'risknote',
    '下一步': 'nextstep', '最近更新': 'update',
}


def find_header(ws):
    """定位含「编号」的表头行，返回 {字段: 列索引}。"""
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        cells = [str(c) if c is not None else '' for c in row]
        if '编号' in cells:
            colmap = {}
            for j, h in enumerate(cells):
                for key, field in COLUMNS.items():
                    if key and key in h:
                        colmap.setdefault(field, j)
            return colmap, cells.index('编号')
    raise ValueError('未找到表头行（含「编号」列）')


def read_tasks(excel_path, category):
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    colmap, _ = find_header(ws)
    tasks = []
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        def g(field, default=None):
            j = colmap.get(field)
            return row[j] if j is not None and j < len(row) else default
        if g('category') != category:
            continue
        tasks.append({
            'id': str(g('id') or ''),
            'wp': str(g('wp') or ''),
            'name': str(g('name') or '').strip(),
            'milestone': (g('milestone') == '是'),
            'deliver': str(g('deliver') or ''),
            'owner': str(g('owner') or ''),
            'dep': str(g('dep') or ''),
            'plan': g('plan'),
            'status': str(g('status') or ''),
            'risk': str(g('risk') or ''),
            'update': g('update'),
        })
    return tasks


def _to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()[:10]
    try:
        return datetime.date.fromisoformat(s)
    except ValueError:
        return None


def build_data(tasks, today):
    tasks = [t for t in tasks if t['id']]  # 丢弃空编号行
    for t in tasks:
        t['plan'] = _to_date(t['plan'])
        t['overdue'] = bool(t['plan'] and t['plan'] < today and t['status'] != '已完成')
        t['plan'] = t['plan'].isoformat() if t['plan'] else ''

    tasks.sort(key=lambda t: t['plan'])
    done = sum(1 for t in tasks if t['status'] == '已完成')
    prog = sum(1 for t in tasks if t['status'] == '进行中')
    notstart = sum(1 for t in tasks if t['status'] == '未开始')
    overdue = sum(1 for t in tasks if t['overdue'])

    # 工作包类别统计（按逻辑顺序）
    wp = {}
    for t in tasks:
        c = t['wp'] or '未分类'
        wp.setdefault(c, {'total': 0, 'done': 0, 'undone': 0})
        wp[c]['total'] += 1
        if t['status'] == '已完成':
            wp[c]['done'] += 1
        else:
            wp[c]['undone'] += 1
    wpList = [{'name': c, 'total': wp[c]['total'], 'done': wp[c]['done'], 'undone': wp[c]['undone']}
              for c in WP_ORDER if c in wp] + \
             [{'name': c, 'total': wp[c]['total'], 'done': wp[c]['done'], 'undone': wp[c]['undone']}
              for c in sorted(wp) if c not in WP_ORDER]

    # 仅保留 JSON 安全的字段，避免 datetime 对象进入序列化
    ms_fields = ['id', 'name', 'wp', 'plan', 'status', 'overdue', 'deliver', 'owner', 'risk']
    milestones = sorted(
        [{k: t.get(k, '') for k in ms_fields} for t in tasks if t['milestone']],
        key=lambda t: t['plan'])

    return {
        'today': today.isoformat(),
        'total': len(tasks), 'done': done, 'progress': prog,
        'notstart': notstart, 'overdue': overdue,
        'rate': round(done / len(tasks) * 100, 1) if tasks else 0,
        'wpList': wpList, 'milestones': milestones,
    }


def ensure_echarts():
    if os.path.exists(ECHARTS_CACHE) and os.path.getsize(ECHARTS_CACHE) > 100000:
        return open(ECHARTS_CACHE, encoding='utf-8').read()
    print('下载 ECharts ...', flush=True)
    req = urllib.request.Request(ECHARTS_URL, headers={'User-Agent': 'Mozilla/5.0'})
    data = urllib.request.urlopen(req, timeout=60).read().decode('utf-8')
    with open(ECHARTS_CACHE, 'w', encoding='utf-8') as f:
        f.write(data)
    return data


def render_html(data, category, title, source_name):
    tmpl = open(TEMPLATE, encoding='utf-8').read()
    echarts = ensure_echarts()
    ms = data['milestones']
    plans = [t['plan'] for t in ms if t['plan']]
    if plans:
        earliest = min(plans)
        latest = max(plans)
    else:
        earliest = data['today']
        latest = data['today']
    base_start = (datetime.date.fromisoformat(earliest) - datetime.timedelta(days=1)).isoformat()
    x_min = (datetime.date.fromisoformat(earliest) - datetime.timedelta(days=4)).isoformat()
    x_max = (datetime.date.fromisoformat(latest) + datetime.timedelta(days=6)).isoformat()
    gantt_h = max(360, len(ms) * 34 + 60)

    goal = ('项目总目标：云财一体化系统生产运行平稳，进一步压降业务风险；'
            '2027年1月1日完成关键系统投产上线运行，覆盖全量基础业务。')

    repl = {
        '__TITLE__': title, '__TAG__': f'财政项目推进总台账 · 项目类别筛选：{category}',
        '__GOAL__': goal, '__CATEGORY__': category, '__SOURCE__': source_name,
        '__TOTAL__': str(data['total']), '__MS__': str(len(ms)), '__TODAY__': data['today'],
        '__GANTT_H__': str(gantt_h), '__BASE_START__': base_start,
        '__X_MIN__': x_min, '__X_MAX__': x_max,
        '__DATA__': json.dumps(data, ensure_ascii=False),
        '__ECHARTS__': echarts,
    }
    for k, v in repl.items():
        tmpl = tmpl.replace(k, v)
    return tmpl


def export_word(data, html_path, out_path, category, title, source_name):
    """通过本地 editor_sdk 生成 Word（含图表截图）。"""
    EDS = '/root/.codebuddy/skills/workbuddy-local-office-edit/edsdk.py'
    if not os.path.exists(EDS):
        print('[word] 未找到本地 editor_sdk，跳过 Word 导出', file=sys.stderr)
        return False
    tmp = tempfile.mkdtemp(prefix='dashboard_shots_')
    os.makedirs(tmp, exist_ok=True)

    # 1) playwright 截图三张图表
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1400, 'height': 1000}, device_scale_factor=2)
        page.goto('file://' + html_path)
        page.wait_for_timeout(2500)
        for sel, name in [('#gantt', 'chart_gantt'), ('#pie', 'chart_pie'), ('#bar', 'chart_bar')]:
            el = page.query_selector(sel)
            el.scroll_into_view_if_needed()
            page.wait_for_timeout(400)
            el.screenshot(path=os.path.join(tmp, name + '.png'))
        browser.close()

    def call(tool, **kw):
        cmd = ['python3', EDS, 'call', tool, '--json', json.dumps(kw, ensure_ascii=False)]
        r = subprocess.run(cmd, capture_output=True, text=True)
        return r.stdout.strip()

    def call_json(tool, **kw):
        out = call(tool, **kw)
        for line in reversed(out.splitlines()):
            line = line.strip()
            if line.startswith('{'):
                try:
                    return json.loads(line)
                except Exception:
                    pass
        return {}

    m = re.search(r'file_id=([^\s,]+)', call('create_doc'))
    fid = m.group(1) if m else None
    if not fid:
        print('[word] create_doc 失败', file=sys.stderr)
        return False

    pos = [0]
    def step(tool, **kw):
        kw['file_id'] = fid
        kw['idx'] = pos[0]
        r = call_json(tool, **kw)
        p = r.get('position', r.get('last_edit_index'))
        if p is not None:
            pos[0] = int(p)

    d = data
    ms = d['milestones']
    wp = d['wpList']
    overdue_list = [t for t in ms if t['overdue']]

    step('doc_insert_markdown', markdown=(
        f"# {title}\n\n"
        f"**数据来源**：{source_name}（项目类别 = {category}）｜**数据截止**：{d['today']}\n\n"
        f"**项目总目标**：云财一体化系统生产运行平稳，进一步压降业务风险；2027年1月1日完成关键系统投产上线运行，覆盖全量基础业务。\n\n"
        f"## 一、核心指标\n\n"
        f"| 指标 | 数值 |\n| --- | --- |\n"
        f"| 工作包总数 | {d['total']} 项 |\n| 已完成 | {d['done']} 项 |\n| 进行中 | {d['progress']} 项 |\n"
        f"| 未开始 | {d['notstart']} 项 |\n| 已逾期 | {d['overdue']} 项 |\n| 完成率 | {d['rate']}% |\n\n"
        f"## 二、里程碑甘特图\n\n仅列出「里程碑 = 是」的工作包，共 {len(ms)} 项，按计划完成日升序排列。"
    ))
    step('doc_insert_image', image_path=os.path.join(tmp, 'chart_gantt.png'), w=600)
    step('doc_insert_paragraph')

    rows = []
    for i, t in enumerate(ms, 1):
        st = t['status'] + ('（逾期）' if t['overdue'] else '')
        rows.append(f"| {i} | {t['name'] or t['id']} | {t['wp']} | {t['plan']} | {st} |")
    step('doc_insert_markdown', markdown=(
        "| 序号 | 工作包名称 | 工作包类别 | 计划完成日 | 推进状态 |\n| --- | --- | --- | --- | --- |\n"
        + "\n".join(rows)))
    step('doc_insert_markdown', markdown="## 三、工作包类别分布\n\n按「工作包类别」统计工作包总数，饼图标示每个类别的数量与占比。")
    step('doc_insert_image', image_path=os.path.join(tmp, 'chart_pie.png'), w=430)
    step('doc_insert_paragraph')
    step('doc_insert_markdown', markdown="## 四、工作包类别完成情况\n\n每个类别一个方柱，柱高为该类别工作包总数；绿色 = 已完成，灰色 = 未完成。")
    step('doc_insert_image', image_path=os.path.join(tmp, 'chart_bar.png'), w=430)
    step('doc_insert_paragraph')

    rows2 = []
    for w in wp:
        rate = round(w['done'] / w['total'] * 100) if w['total'] else 0
        rows2.append(f"| {w['name']} | {w['total']} | {w['done']} | {w['undone']} | {rate}% |")
    step('doc_insert_markdown', markdown=(
        "| 工作包类别 | 总数 | 已完成 | 未完成 | 完成率 |\n| --- | --- | --- | --- | --- |\n"
        + "\n".join(rows2)))
    if overdue_list:
        items = "\n".join([f"- **{t['name'] or t['id']}**（{t['wp']}）：计划完成 {t['plan']}，当前「{t['status']}」" for t in overdue_list])
    else:
        items = "- 无逾期事项"
    step('doc_insert_markdown', markdown="## 五、逾期事项\n\n" + items)

    res = call('save_file', file_id=fid, file_path=out_path)
    print('[word]', res)
    return True


def main():
    ap = argparse.ArgumentParser(description='项目推进看板生成器')
    ap.add_argument('excel', help='台账 Excel 路径')
    ap.add_argument('--category', default='新系统建设', help='项目类别筛选')
    ap.add_argument('--outdir', default='/workspace', help='输出目录')
    ap.add_argument('--title', default=None, help='看板标题')
    ap.add_argument('--today', default=None, help='数据截止日期 YYYY-MM-DD')
    ap.add_argument('--word', action='store_true', help='额外导出 Word')
    args = ap.parse_args()

    if not os.path.exists(args.excel):
        print(f'❌ 文件不存在: {args.excel}', file=sys.stderr)
        sys.exit(1)

    category = args.category
    title = args.title or f'{category}项目推进看板'
    source_name = os.path.basename(args.excel)

    print(f'读取 {source_name}，筛选项目类别 = {category} ...')
    tasks = read_tasks(args.excel, category)
    print(f'  筛选出 {len(tasks)} 条记录')

    # 截止日：优先 --today；默认系统当天（「截止当日」字面语义）
    today = datetime.date.fromisoformat(args.today) if args.today else datetime.date.today()
    print(f'  截止日 = {today.isoformat()}')
    data = build_data(tasks, today)
    print(f"  已完成 {data['done']} / 进行中 {data['progress']} / 未开始 {data['notstart']} / 逾期 {data['overdue']} / 完成率 {data['rate']}%")
    print(f"  工作包类别 {len(data['wpList'])} 类，里程碑 {len(data['milestones'])} 项")

    os.makedirs(args.outdir, exist_ok=True)
    html_path = os.path.join(args.outdir, title + '.html')
    html = render_html(data, category, title, source_name)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'✅ HTML 看板: {html_path} ({len(html)} bytes)')

    if args.word:
        docx_path = os.path.join(args.outdir, title + '.docx')
        print('导出 Word ...')
        if export_word(data, html_path, docx_path, category, title, source_name):
            print(f'✅ Word 文档: {docx_path}')
        else:
            print('⚠️ Word 导出未完成（依赖缺失）')


if __name__ == '__main__':
    main()
