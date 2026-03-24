#!/usr/bin/env python3
"""导出汽车之家车型口碑为正式版 Excel。

功能：
- 抓取指定车型的最满意 / 最不满意口碑页
- 区分车主购车口碑与试驾探店口碑
- 按正式版字段导出为 4 个 sheet 的 Excel

依赖：
- agent-browser 可用
- openpyxl 已安装

输出：
- `ZJ+车型名称+最满意or最不满意_页数范围.xlsx`
"""

import argparse
import json
import subprocess
import time
from collections import Counter
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

HEADERS = [
    '口碑维度','数据类型','用户名','发表日期','口碑标题','综合口碑','车型','行驶里程','电耗',
    '裸车购买价','参考价格','购买时间','探店时间','购买地点','探店地点','评价详情','来源链接','抓取页码'
]


def run(cmd: str, cwd: Path, timeout: int = 120):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=str(cwd), timeout=timeout)


def url_for(series_id: int, dimensionid: int, page: int) -> str:
    if page == 1:
        return f'https://k.autohome.com.cn/{series_id}?dimensionid={dimensionid}&order=0&yearid=0#listcontainer'
    return f'https://k.autohome.com.cn/{series_id}/index_{page}.html?dimensionid={dimensionid}&order=0&yearid=0#listcontainer'


def get_snapshot(cwd: Path, url: str, session: str, retries: int = 2):
    last_err = ''
    for _ in range(retries + 1):
        r1 = run(f"agent-browser --session {session} open '{url}' >/dev/null", cwd)
        if r1.returncode == 0:
            r2 = run(f"agent-browser --session {session} snapshot -c", cwd)
            if r2.returncode == 0 and '条口碑在售' in r2.stdout:
                return r2.stdout.splitlines()
            last_err = r2.stderr or r2.stdout
        else:
            last_err = r1.stderr or r1.stdout
        time.sleep(1)
    raise RuntimeError(last_err)


def section(lines, start, end):
    s = next(i for i, l in enumerate(lines) if start in l)
    e = next(i for i, l in enumerate(lines[s + 1:], s + 1) if end in l)
    return lines[s + 1:e]


def split_cards(sec_lines):
    cards, cur = [], []
    for l in sec_lines:
        if l.startswith('    - listitem:'):
            if cur:
                cards.append(cur)
            cur = [l]
        elif cur:
            cur.append(l)
    if cur:
        cards.append(cur)
    return [c for c in cards if any('查看完整口碑' in x for x in c)]


def norm_user(raw: str) -> str:
    raw = raw.strip()
    if raw.endswith(' 风云X3L认证'):
        return raw[:-8].strip() + '_风云X3L认证'
    if raw.endswith(' 风云X3L'):
        return raw[:-6].strip()
    return raw


def parse_card(lines, dim_name: str, page: int):
    import re
    row = {h: '' for h in HEADERS}
    row['口碑维度'] = dim_name
    row['抓取页码'] = page
    row['数据类型'] = '车主购车口碑'

    for i, l in enumerate(lines):
        s = l.strip()
        if s.startswith('- /url: https://i.autohome.com.cn/') and i + 1 < len(lines):
            m = re.match(r'- link "([^"]+)"', lines[i + 1].strip())
            if m and not row['用户名']:
                row['用户名'] = norm_user(m.group(1))
        if '发表口碑' in s:
            m = re.search(r'(20\d{2}-\d{2}-\d{2})\s+发表口碑', s)
            if m:
                row['发表日期'] = m.group(1)
        if '综合口碑评分' in s:
            m = re.search(r'综合口碑评分\s+([0-9.]+)', s)
            if m:
                row['综合口碑'] = m.group(1)
        m = re.match(r'- link "([^"]+)" \[ref=.*\]:$', s)
        if m and i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if 'https://k.autohome.com.cn/detail/view_' in nxt and not row['口碑标题']:
                row['口碑标题'] = m.group(1)
        m = re.match(r'- link "(20\d{2}款 [^"]+)"', s)
        if m and not row['车型']:
            row['车型'] = m.group(1)
        m = re.match(r'- listitem: (.+)', s)
        if m:
            val = m.group(1).strip()
            if val.endswith('行驶里程'):
                row['行驶里程'] = val[:-4].strip()
            elif '电耗' in val:
                row['电耗'] = val
            elif val.endswith('裸车购买价'):
                row['裸车购买价'] = val[:-5].strip()
            elif val.endswith('参考价格'):
                row['参考价格'] = val[:-4].strip()
                row['数据类型'] = '试驾探店口碑'
            elif val.endswith('购买时间'):
                row['购买时间'] = val[:-4].strip()
            elif val.endswith('探店时间'):
                row['探店时间'] = val[:-4].strip()
                row['数据类型'] = '试驾探店口碑'
            elif val.endswith('购买地点'):
                row['购买地点'] = val[:-4].strip()
            elif val.endswith('探店地点'):
                row['探店地点'] = val[:-4].strip()
                row['数据类型'] = '试驾探店口碑'
        for prefix, dtype in [('- text: 满意 ', '车主购车口碑'), ('- text: 不满意 ', '车主购车口碑'), ('- text: 好评 ', '试驾探店口碑'), ('- text: 槽点 ', '试驾探店口碑')]:
            if s.startswith(prefix):
                row['评价详情'] = s[len(prefix):].strip()
                row['数据类型'] = dtype
        if 'link "查看完整口碑"' in s and i + 1 < len(lines):
            m = re.match(r'- /url: (https://k\.autohome\.com\.cn/detail/view_[^\s]+)', lines[i + 1].strip())
            if m:
                row['来源链接'] = m.group(1)

    if any([row['参考价格'], row['探店时间'], row['探店地点']]):
        row['数据类型'] = '试驾探店口碑'

    valid = all(row[k] for k in ['用户名', '综合口碑', '车型', '来源链接'])
    return row, valid


def collect_dimension(cwd: Path, series_id: int, dimensionid: int, start_page: int, end_page: int):
    dim_name = '最满意' if dimensionid == 10 else '最不满意'
    rows, bad = [], []
    for page in range(start_page, end_page + 1):
        lines = get_snapshot(cwd, url_for(series_id, dimensionid, page), f'koubei_{series_id}_{dimensionid}_{page}')
        marker = next((l.strip() for l in lines if '条口碑在售' in l), None)
        if not marker:
            raise RuntimeError('未找到口碑列表标记')
        cards = split_cards(section(lines, marker, '相关车系口碑推荐'))
        for c in cards:
            row, valid = parse_card(c, dim_name, page)
            if valid:
                rows.append(row)
            else:
                bad.append(row)
    return rows, bad


def write_xlsx(out_path: Path, groups: dict):
    wb = Workbook()
    sheet_names = ['最满意_车主购车口碑', '最满意_试驾探店口碑', '最不满意_车主购车口碑', '最不满意_试驾探店口碑']
    first = True
    for name in sheet_names:
        ws = wb.active if first else wb.create_sheet(name)
        first = False
        ws.title = name
        ws.append(HEADERS)
        for r in groups.get(name, []):
            ws.append([r.get(h, '') for h in HEADERS])
        for c in ws[1]:
            c.font = Font(bold=True)
        for col in ['A','B','C','D','F','G','H','I','J','K','L','M','N','O','R']:
            ws.column_dimensions[col].width = 16
        ws.column_dimensions['E'].width = 28
        ws.column_dimensions['P'].width = 100
        ws.column_dimensions['Q'].width = 60
        for row in ws.iter_rows(min_row=2):
            row[15].alignment = Alignment(wrap_text=True, vertical='top')
            row[16].alignment = Alignment(wrap_text=True, vertical='top')
    wb.save(out_path)


def validate_results(sat_rows, unsat_rows):
    page_counts = {
        '最满意': Counter(r['抓取页码'] for r in sat_rows),
        '最不满意': Counter(r['抓取页码'] for r in unsat_rows),
    }
    issues = []

    if len(sat_rows) != len(unsat_rows):
        issues.append({
            'type': 'total_mismatch',
            'message': f'最满意总数 {len(sat_rows)} != 最不满意总数 {len(unsat_rows)}',
            'sat_total': len(sat_rows),
            'unsat_total': len(unsat_rows),
        })

    pages = sorted(set(page_counts['最满意']) | set(page_counts['最不满意']))
    for page in pages:
        sat_count = page_counts['最满意'].get(page, 0)
        unsat_count = page_counts['最不满意'].get(page, 0)
        if sat_count != unsat_count:
            issues.append({
                'type': 'page_mismatch',
                'message': f'第 {page} 页数量不一致：最满意 {sat_count}，最不满意 {unsat_count}',
                'page': page,
                'sat_count': sat_count,
                'unsat_count': unsat_count,
            })

    return {
        'ok': not issues,
        'issues': issues,
        'page_counts': {
            k: dict(sorted(v.items())) for k, v in page_counts.items()
        },
        'sat_total': len(sat_rows),
        'unsat_total': len(unsat_rows),
    }


def write_validation_report(path: Path, report: dict):
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser(description='Export Autohome koubei to xlsx')
    parser.add_argument('--series-id', type=int, required=True)
    parser.add_argument('--start-page', type=int, default=1)
    parser.add_argument('--end-page', type=int, required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--workdir', default='.')
    parser.add_argument('--strict-validate', action='store_true', help='总数或分页数量不一致时返回非零退出码')
    args = parser.parse_args()

    cwd = Path(args.workdir).resolve()
    sat_rows, sat_bad = collect_dimension(cwd, args.series_id, 10, args.start_page, args.end_page)
    unsat_rows, unsat_bad = collect_dimension(cwd, args.series_id, 11, args.start_page, args.end_page)

    groups = {
        '最满意_车主购车口碑': [r for r in sat_rows if r['数据类型'] == '车主购车口碑'],
        '最满意_试驾探店口碑': [r for r in sat_rows if r['数据类型'] == '试驾探店口碑'],
        '最不满意_车主购车口碑': [r for r in unsat_rows if r['数据类型'] == '车主购车口碑'],
        '最不满意_试驾探店口碑': [r for r in unsat_rows if r['数据类型'] == '试驾探店口碑'],
    }

    out_path = Path(args.output).resolve()
    write_xlsx(out_path, groups)

    report = validate_results(sat_rows, unsat_rows)
    report['anomalies'] = len(sat_bad) + len(unsat_bad)
    report_path = out_path.with_suffix('.validation.json')
    write_validation_report(report_path, report)

    print(f'Wrote: {out_path}')
    print(f'Validation: {report_path}')
    print('Counts:')
    for k, v in groups.items():
        print(f'- {k}: {len(v)}')
    print(f'- anomalies: {len(sat_bad) + len(unsat_bad)}')
    if report['ok']:
        print('- validation: OK')
    else:
        print('- validation: FAILED')
        for issue in report['issues']:
            print(f"  * {issue['message']}")
        if args.strict_validate:
            raise SystemExit(2)


if __name__ == '__main__':
    main()
