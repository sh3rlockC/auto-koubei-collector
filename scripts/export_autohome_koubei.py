#!/usr/bin/env python3
"""导出汽车之家车型口碑为按来源链接对齐的 Excel。

功能：
- 抓取指定车型的最满意 / 最不满意口碑页
- 按来源链接严格对齐同一条口碑
- 统一判定为车主购车口碑 / 试驾探店口碑
- 导出为 2 个 sheet：购车口碑 / 试驾探店口碑
- 每行包含同一链接的“最满意 / 最不满意”两列

依赖：
- agent-browser 可用
- openpyxl 已安装
"""

import argparse
import hashlib
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Optional
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font


class ProgressTracker:
    def __init__(self, label: str, progress_path: Optional[Path] = None, quiet: bool = False):
        self.label = label
        self.progress_path = progress_path
        self.quiet = quiet
        self.state = {
            'label': label,
            'stage': 'init',
            'current': 0,
            'total': 0,
            'percent': 0,
            'success': 0,
            'retry': 0,
            'failed': 0,
            'records': 0,
            'message': '',
            'updated_at': '',
            'overall': {'current': 0, 'total': 0, 'percent': 0},
            'current_dimension': '',
            'dimension_progress': {},
            'output_path': '',
            'validation_path': '',
            'failed_pages': [],
        }

    def _render_bar(self, current: int, total: int, width: int = 24) -> str:
        if total <= 0:
            return '[' + '.' * width + ']'
        filled = min(width, int(width * current / total))
        return '[' + '#' * filled + '.' * (width - filled) + ']'

    def emit(self):
        self.state['updated_at'] = datetime.now().isoformat(timespec='seconds')
        if self.progress_path:
            self.progress_path.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding='utf-8')
        if self.quiet:
            return
        overall = self.state.get('overall') or {}
        bar = self._render_bar(int(overall.get('current', 0)), int(overall.get('total', 0)))
        total = overall.get('total') or '?'
        current_dimension = self.state.get('current_dimension') or '-'
        dim_progress = (self.state.get('dimension_progress') or {}).get(current_dimension) or {}
        dim_current = dim_progress.get('current', 0)
        dim_total = dim_progress.get('total', 0)
        print(
            f"{self.state['stage']} {bar} 总体 {overall.get('current', 0)}/{total} ({overall.get('percent', 0)}%) | "
            f"{current_dimension} {dim_current}/{dim_total} | ok {self.state['success']} retry {self.state['retry']} "
            f"fail {self.state['failed']} rows {self.state['records']} | {self.state['message']}",
            flush=True,
        )

    def update(self, **kwargs):
        self.state.update(kwargs)
        total = int(self.state.get('total') or 0)
        current = int(self.state.get('current') or 0)
        percent = int(current * 100 / total) if total > 0 else 0
        self.state['percent'] = percent
        self.state['overall'] = {'current': current, 'total': total, 'percent': percent}
        self.emit()

HEADERS = [
    '数据类型', '用户名', '发表日期', '口碑标题', '综合口碑', '车型', '行驶里程', '电耗',
    '裸车购买价', '参考价格', '购买时间', '探店时间', '购买地点', '探店地点',
    '最满意', '最不满意', '来源链接', '最满意抓取页码', '最不满意抓取页码'
]


def run(cmd: str, cwd: Path, timeout: int = 120):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=str(cwd), timeout=timeout)


def url_for(series_id: int, dimensionid: int, page: int) -> str:
    if page == 1:
        return f'https://k.autohome.com.cn/{series_id}?dimensionid={dimensionid}&order=0&yearid=0#listcontainer'
    return f'https://k.autohome.com.cn/{series_id}/index_{page}.html?dimensionid={dimensionid}&order=0&yearid=0#listcontainer'


def get_snapshot(cwd: Path, url: str, session: str, retries: int = 2):
    last_err = ''
    retry_count = 0
    for _ in range(retries + 1):
        r1 = run(f"agent-browser --session {session} open '{url}' >/dev/null", cwd)
        if r1.returncode == 0:
            r2 = run(f"agent-browser --session {session} snapshot -c", cwd)
            if r2.returncode == 0:
                out = r2.stdout
                if '查看完整口碑' in out and 'detail/view_' in out:
                    return out.splitlines(), retry_count
            last_err = r2.stderr or r2.stdout
        else:
            last_err = r1.stderr or r1.stdout
        retry_count += 1
        time.sleep(1)
    raise RuntimeError(last_err)


def get_page_html(cwd: Path, url: str, session: str, retries: int = 2):
    last_err = ''
    for _ in range(retries + 1):
        r1 = run(f"agent-browser --session {session} open '{url}' >/dev/null", cwd)
        if r1.returncode != 0:
            last_err = r1.stderr or r1.stdout
            time.sleep(1)
            continue
        r2 = run(f"agent-browser --session {session} eval 'document.documentElement.outerHTML'", cwd)
        if r2.returncode == 0 and r2.stdout.strip():
            return r2.stdout
        last_err = r2.stderr or r2.stdout
        time.sleep(1)
    raise RuntimeError(last_err)


def get_snapshot_any(cwd: Path, url: str, session: str, retries: int = 2):
    last_err = ''
    for _ in range(retries + 1):
        r1 = run(f"agent-browser --session {session} open '{url}' >/dev/null", cwd)
        if r1.returncode == 0:
            r2 = run(f"agent-browser --session {session} snapshot -c", cwd)
            if r2.returncode == 0 and r2.stdout.strip():
                return r2.stdout.splitlines()
            last_err = r2.stderr or r2.stdout
        else:
            last_err = r1.stderr or r1.stdout
        time.sleep(1)
    raise RuntimeError(last_err)


def detect_max_page(cwd: Path, series_id: int, dimensionid: int) -> int:
    html = get_page_html(cwd, url_for(series_id, dimensionid, 1), f'koubei_detect_{series_id}_{dimensionid}')
    html = unescape(html)

    candidates = set()

    for m in re.finditer(rf'/{series_id}/index_(\d+)\.html\?dimensionid={dimensionid}', html):
        candidates.add(int(m.group(1)))

    for m in re.finditer(r'ace-pagination__link[^>]*>(\d+)<', html):
        candidates.add(int(m.group(1)))

    for m in re.finditer(r'分页[^\n]{0,200}?共\s*(\d+)\s*页', html):
        candidates.add(int(m.group(1)))

    for m in re.finditer(r'共\s*(\d+)\s*页', html):
        candidates.add(int(m.group(1)))

    for m in re.finditer(r'尾页[^\n]{0,120}?index_(\d+)\.html', html):
        candidates.add(int(m.group(1)))

    candidates = {x for x in candidates if x >= 1}
    return max(candidates) if candidates else 1


def norm_user(raw: str) -> str:
    raw = raw.strip()
    if raw.endswith(' 风云X3L认证'):
        return raw[:-8].strip() + '_风云X3L认证'
    if raw.endswith(' 风云X3L'):
        return raw[:-6].strip()
    return raw


def extract_cards(lines):
    cards = []
    current = []
    in_target = False
    for line in lines:
        s = line.strip()
        if 'heading ' in s and '/detail/view_' not in s:
            pass
        if 'heading ' in s and current and any('查看完整口碑' in x for x in current):
            cards.append(current)
            current = []
        if current or ('heading ' in s and '[level=1]' in s):
            current.append(line)
            if '相关车系口碑推荐' in s:
                if current and any('查看完整口碑' in x for x in current):
                    cards.append(current)
                current = []
                break
    if current and any('查看完整口碑' in x for x in current):
        cards.append(current)

    detail_cards = []
    for c in cards:
        if any('https://k.autohome.com.cn/detail/view_' in x for x in c):
            detail_cards.append(c)

    # 去重：同一链接保留首次出现
    seen = set()
    uniq = []
    for c in detail_cards:
        link = ''
        for line in c:
            m = re.search(r'https://k\.autohome\.com\.cn/detail/view_[^\s]+', line)
            if m:
                link = m.group(0)
                break
        if link and link not in seen:
            seen.add(link)
            uniq.append(c)
    return uniq


def parse_card(lines, dim_name: str, page: int, cwd=None):
    row = {
        '口碑维度': dim_name,
        '抓取页码': page,
        '数据类型': '车主购车口碑',
        '用户名': '', '发表日期': '', '口碑标题': '', '综合口碑': '', '车型': '',
        '行驶里程': '', '电耗': '', '裸车购买价': '', '参考价格': '', '购买时间': '',
        '探店时间': '', '购买地点': '', '探店地点': '', '评价详情': '', '来源链接': ''
    }

    compact = ' '.join(l.strip() for l in lines)

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

        for prefix, dtype in [
            ('- text: 满意 ', '车主购车口碑'),
            ('- text: 不满意 ', '车主购车口碑'),
            ('- text: 好评 ', '试驾探店口碑'),
            ('- text: 槽点 ', '试驾探店口碑')
        ]:
            if s.startswith(prefix):
                row['评价详情'] = s[len(prefix):].strip()
                row['数据类型'] = dtype

        if s in ('- text: 满意', '- text: 不满意') and not row['评价详情']:
            m = re.search(r'- text: (?:满意|不满意)\s+(.*?)\s+- listitem:', compact)
            if m:
                candidate = m.group(1).strip()
                if candidate and not candidate.startswith('- listitem:'):
                    row['评价详情'] = candidate
                    row['数据类型'] = '车主购车口碑'

        if s in ('- text: 好评', '- text: 槽点') and not row['评价详情']:
            m = re.search(r'- text: (?:好评|槽点)\s+(.*?)\s+- listitem:', compact)
            if m:
                candidate = m.group(1).strip()
                if candidate and not candidate.startswith('- listitem:'):
                    row['评价详情'] = candidate
                    row['数据类型'] = '试驾探店口碑'

        m = re.search(r'https://k\.autohome\.com\.cn/detail/view_[^\s]+', s)
        if m and not row['来源链接']:
            row['来源链接'] = m.group(0)

    if any([row['参考价格'], row['探店时间'], row['探店地点']]):
        row['数据类型'] = '试驾探店口碑'

    if not row['评价详情'] and row['来源链接'] and cwd is not None:
        try:
            detail_session = f"kd_{hashlib.md5(row['来源链接'].encode('utf-8')).hexdigest()[:12]}"
            detail_lines = get_snapshot_any(cwd, row['来源链接'], detail_session)
            detail_text = extract_detail_text(detail_lines, row['口碑维度'])
            if detail_text:
                row['评价详情'] = detail_text
        except Exception:
            pass

    valid = all(row[k] for k in ['用户名', '综合口碑', '车型', '来源链接', '评价详情'])
    return row, valid


def extract_detail_text(lines, dim_name: str) -> str:
    target_heading = dim_name
    collecting = False
    parts = []
    for raw in lines:
        s = raw.strip()
        if s.startswith(f'- heading "{target_heading}"'):
            collecting = True
            continue
        if collecting and s.startswith('- heading "'):
            break
        if collecting and s.startswith('- paragraph:'):
            text = s[len('- paragraph:'):].strip()
            text = text.strip('"').strip()
            if text and '上述内容的版权归' not in text:
                parts.append(text)
    return ' '.join(parts).strip()


def collect_dimension(cwd: Path, series_id: int, dimensionid: int, start_page: int, end_page: int, tracker=None, offset: int = 0, total_steps: int = 0, failed_page_list=None, target_pages=None):
    dim_name = '最满意' if dimensionid == 10 else '最不满意'
    rows, bad = [], []
    page_link_counts = {}
    retry_total = 0
    success_pages = 0
    failed_pages = 0
    pages = target_pages or list(range(start_page, end_page + 1))
    page_total = len(pages)
    for index, page in enumerate(pages, start=1):
        current_step = offset + index
        if tracker:
            dim_progress = dict(tracker.state.get('dimension_progress') or {})
            dim_progress[dim_name] = {'current': index, 'total': page_total}
            tracker.update(
                stage='抓取页面',
                current=current_step,
                total=total_steps,
                success=success_pages,
                retry=retry_total,
                failed=failed_pages,
                records=len(rows),
                current_dimension=dim_name,
                dimension_progress=dim_progress,
                message=f'{dim_name} 第 {page} 页',
            )
        try:
            lines, retries_used = get_snapshot(cwd, url_for(series_id, dimensionid, page), f'koubei_{series_id}_{dimensionid}_{page}')
            retry_total += retries_used
        except RuntimeError as e:
            failed_pages += 1
            if failed_page_list is not None:
                failed_page_list.append({'dimension': dim_name, 'page': page, 'reason': str(e)})
            bad.append({'口碑维度': dim_name, '抓取页码': page, '错误': str(e)})
            if tracker:
                tracker.update(
                    stage='抓取页面',
                    current=current_step,
                    total=total_steps,
                    success=success_pages,
                    retry=retry_total,
                    failed=failed_pages,
                    records=len(rows),
                    message=f'{dim_name} 第 {page} 页失败',
                )
            continue

        cards = extract_cards(lines)
        page_link_counts[page] = len(cards)
        if not cards:
            failed_pages += 1
            if failed_page_list is not None:
                failed_page_list.append({'dimension': dim_name, 'page': page, 'reason': '未解析到口碑卡片'})
            bad.append({'口碑维度': dim_name, '抓取页码': page, '错误': '未解析到口碑卡片'})
            if tracker:
                tracker.update(
                    stage='抓取页面',
                    current=current_step,
                    total=total_steps,
                    success=success_pages,
                    retry=retry_total,
                    failed=failed_pages,
                    records=len(rows),
                    message=f'{dim_name} 第 {page} 页未解析到口碑卡片',
                )
            continue

        success_pages += 1
        for c in cards:
            row, valid = parse_card(c, dim_name, page, cwd=cwd)
            if valid:
                rows.append(row)
            else:
                bad.append(row)
        if tracker:
            tracker.update(
                stage='抓取页面',
                current=current_step,
                total=total_steps,
                success=success_pages,
                retry=retry_total,
                failed=failed_pages,
                records=len(rows),
                message=f'{dim_name} 第 {page} 页完成，新增 {len(cards)} 条卡片',
            )
    return rows, bad, page_link_counts, retry_total, success_pages, failed_pages


def merge_aligned(sat_rows, unsat_rows):
    sat_map = {r['来源链接']: r for r in sat_rows}
    unsat_map = {r['来源链接']: r for r in unsat_rows}

    sat_links = set(sat_map)
    unsat_links = set(unsat_map)
    common = sorted(sat_links & unsat_links)
    only_sat = sorted(sat_links - unsat_links)
    only_unsat = sorted(unsat_links - sat_links)

    merged = []
    anomalies = []

    for link in common:
        a = sat_map[link]
        b = unsat_map[link]
        dtype = '试驾探店口碑' if any([
            a['数据类型'] == '试驾探店口碑',
            b['数据类型'] == '试驾探店口碑',
            a['参考价格'], b['参考价格'],
            a['探店时间'], b['探店时间'],
            a['探店地点'], b['探店地点'],
        ]) else '车主购车口碑'

        row = {
            '数据类型': dtype,
            '用户名': a['用户名'] or b['用户名'],
            '发表日期': a['发表日期'] or b['发表日期'],
            '口碑标题': a['口碑标题'] or b['口碑标题'],
            '综合口碑': a['综合口碑'] or b['综合口碑'],
            '车型': a['车型'] or b['车型'],
            '行驶里程': a['行驶里程'] or b['行驶里程'],
            '电耗': a['电耗'] or b['电耗'],
            '裸车购买价': a['裸车购买价'] or b['裸车购买价'],
            '参考价格': a['参考价格'] or b['参考价格'],
            '购买时间': a['购买时间'] or b['购买时间'],
            '探店时间': a['探店时间'] or b['探店时间'],
            '购买地点': a['购买地点'] or b['购买地点'],
            '探店地点': a['探店地点'] or b['探店地点'],
            '最满意': a['评价详情'],
            '最不满意': b['评价详情'],
            '来源链接': link,
            '最满意抓取页码': a['抓取页码'],
            '最不满意抓取页码': b['抓取页码'],
        }

        if not row['最满意'] or not row['最不满意']:
            anomalies.append({'type': 'missing_dimension_text', 'link': link, 'sat': a, 'unsat': b})
            continue

        if dtype == '车主购车口碑':
            row['参考价格'] = ''
            row['探店时间'] = ''
            row['探店地点'] = ''

        merged.append(row)

    for link in only_sat:
        anomalies.append({'type': 'missing_unsat', 'link': link, 'row': sat_map[link]})
    for link in only_unsat:
        anomalies.append({'type': 'missing_sat', 'link': link, 'row': unsat_map[link]})

    groups = {
        '购车口碑': [r for r in merged if r['数据类型'] == '车主购车口碑'],
        '试驾口碑': [r for r in merged if r['数据类型'] == '试驾探店口碑'],
    }
    return groups, anomalies, {'common': len(common), 'only_sat': only_sat, 'only_unsat': only_unsat}


def apply_sheet_style(ws):
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','R','S']:
        ws.column_dimensions[col].width = 16
    ws.column_dimensions['O'].width = 90
    ws.column_dimensions['P'].width = 90
    ws.column_dimensions['Q'].width = 60
    for row in ws.iter_rows(min_row=2):
        row[14].alignment = Alignment(wrap_text=True, vertical='top')
        row[15].alignment = Alignment(wrap_text=True, vertical='top')
        row[16].alignment = Alignment(wrap_text=True, vertical='top')


def write_xlsx(out_path: Path, groups: dict):
    wb = Workbook()
    sheet_names = ['购车口碑', '试驾口碑']
    first = True
    for name in sheet_names:
        ws = wb.active if first else wb.create_sheet(name)
        first = False
        ws.title = name
        ws.append(HEADERS)
        for r in groups.get(name, []):
            ws.append([r.get(h, '') for h in HEADERS])
        apply_sheet_style(ws)
    wb.save(out_path)


def read_existing_groups(path: Path):
    wb = load_workbook(path)
    groups = {'购车口碑': [], '试驾口碑': []}
    for sheet_name in groups.keys():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header = [cell.value for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            item = {}
            for i, key in enumerate(header):
                if key:
                    item[str(key)] = '' if row[i] is None else str(row[i])
            if any(item.values()):
                groups[sheet_name].append(item)
    return groups


def merge_groups(existing_groups: dict, new_groups: dict, mode: str = 'keep-extra'):
    merged = {'购车口碑': [], '试驾口碑': []}
    for sheet_name in merged.keys():
        by_link = {}
        order = []
        if mode != 'strict':
            for row in existing_groups.get(sheet_name, []):
                key = row.get('来源链接') or ''
                if not key:
                    continue
                if key not in by_link:
                    order.append(key)
                by_link[key] = row
        for row in new_groups.get(sheet_name, []):
            key = row.get('来源链接') or ''
            if not key:
                continue
            if key not in by_link:
                order.append(key)
            by_link[key] = row
        merged[sheet_name] = [by_link[key] for key in order]
    return merged


def write_validation_report(path: Path, report: dict):
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def load_retry_pages_by_dimension(path: Path):
    data = json.loads(path.read_text(encoding='utf-8'))
    result = {'最满意': [], '最不满意': []}
    for item in data:
        page = item.get('page')
        dim = item.get('dimension')
        if dim not in result:
            continue
        if isinstance(page, int) and page > 0:
            result[dim].append(page)
        elif isinstance(page, str) and page.isdigit() and int(page) > 0:
            result[dim].append(int(page))
    for key in result:
        result[key] = sorted(set(result[key]))
    return result


def main():
    parser = argparse.ArgumentParser(description='Export aligned Autohome koubei to xlsx')
    parser.add_argument('--series-id', type=int, required=True)
    parser.add_argument('--start-page', type=int, default=1)
    parser.add_argument('--end-page', type=int)
    parser.add_argument('--auto-detect-pages', action='store_true', help='Auto detect max page when end page is omitted')
    parser.add_argument('--output', required=True)
    parser.add_argument('--workdir', default='.')
    parser.add_argument('--strict-validate', action='store_true')
    parser.add_argument('--progress-file', help='进度 JSON 输出路径；默认与 output 同目录同名 .progress.json')
    parser.add_argument('--quiet', action='store_true', help='静默模式，仅写 progress json，不打印进度行')
    parser.add_argument('--retry-failed-pages', help='从 *.failed-pages.json 读取失败页，仅补抓这些页')
    parser.add_argument('--merge-into', help='将补抓结果合并进已有 Excel，按 来源链接 覆盖旧记录')
    parser.add_argument('--merge-mode', choices=['keep-extra', 'strict'], default='keep-extra', help='merge-into 时的合并模式：keep-extra 保留旧表其他记录；strict 仅保留本轮新结果')
    args = parser.parse_args()

    cwd = Path(args.workdir).resolve()

    retry_pages_by_dimension = None
    if args.retry_failed_pages:
        retry_pages_by_dimension = load_retry_pages_by_dimension(Path(args.retry_failed_pages).resolve())
        if not retry_pages_by_dimension['最满意'] and not retry_pages_by_dimension['最不满意']:
            raise SystemExit('retry-failed-pages 中没有可用页码')
        sat_target_pages = retry_pages_by_dimension['最满意']
        unsat_target_pages = retry_pages_by_dimension['最不满意']
        all_target_pages = sat_target_pages + unsat_target_pages
        args.start_page = min(all_target_pages)
        args.end_page = max(all_target_pages)
    else:
        if args.end_page is None:
            if args.auto_detect_pages or args.start_page == 1:
                sat_max = detect_max_page(cwd, args.series_id, 10)
                unsat_max = detect_max_page(cwd, args.series_id, 11)
                args.end_page = max(sat_max, unsat_max)
                print(f'Auto-detected end page: {args.end_page} (sat={sat_max}, unsat={unsat_max})')
            else:
                raise SystemExit('--end-page is required when start-page is not 1; pass --auto-detect-pages if you want automatic detection')
        sat_target_pages = list(range(args.start_page, args.end_page + 1))
        unsat_target_pages = list(range(args.start_page, args.end_page + 1))

    if args.start_page < 1 or args.end_page < args.start_page:
        raise SystemExit('invalid page range: require start-page >= 1 and end-page >= start-page')

    out_path = Path(args.output).resolve()
    progress_path = Path(args.progress_file).resolve() if args.progress_file else out_path.with_suffix('.progress.json')
    tracker = ProgressTracker(label=f'autohome:{args.series_id}', progress_path=progress_path, quiet=args.quiet)
    total_steps = len(sat_target_pages) + len(unsat_target_pages) + 2
    failed_page_list = []
    tracker.update(
        stage='初始化',
        current=0,
        total=total_steps,
        current_dimension='最满意',
        dimension_progress={
            '最满意': {'current': 0, 'total': len(sat_target_pages)},
            '最不满意': {'current': 0, 'total': len(unsat_target_pages)},
        },
        output_path=str(out_path),
        message='准备开始抓取汽车之家口碑' + ('（失败页补抓）' if args.retry_failed_pages else ''),
        failed_pages=failed_page_list,
    )

    sat_rows, sat_bad, sat_pages, sat_retry, sat_success, sat_failed = collect_dimension(
        cwd, args.series_id, 10, args.start_page, args.end_page, tracker=tracker, offset=0, total_steps=total_steps, failed_page_list=failed_page_list, target_pages=sat_target_pages
    )
    page_span = len(sat_target_pages)
    unsat_rows, unsat_bad, unsat_pages, unsat_retry, unsat_success, unsat_failed = collect_dimension(
        cwd, args.series_id, 11, args.start_page, args.end_page, tracker=tracker, offset=page_span, total_steps=total_steps, failed_page_list=failed_page_list, target_pages=unsat_target_pages
    )

    tracker.update(
        stage='结果对齐',
        current=page_span * 2 + 1,
        total=total_steps,
        success=sat_success + unsat_success,
        retry=sat_retry + unsat_retry,
        failed=sat_failed + unsat_failed,
        records=len(sat_rows) + len(unsat_rows),
        current_dimension='最不满意',
        failed_pages=failed_page_list,
        message='合并最满意/最不满意记录',
    )
    groups, anomalies, meta = merge_aligned(sat_rows, unsat_rows)

    final_groups = groups
    merge_summary = None
    if args.merge_into:
        merge_path = Path(args.merge_into).resolve()
        existing_groups = read_existing_groups(merge_path)
        final_groups = merge_groups(existing_groups, groups, mode=args.merge_mode)
        merge_summary = {
            'merge_into': str(merge_path),
            'merge_mode': args.merge_mode,
            'existing_rows': {k: len(v) for k, v in existing_groups.items()},
            'new_rows': {k: len(v) for k, v in groups.items()},
            'merged_rows': {k: len(v) for k, v in final_groups.items()},
        }

    write_xlsx(out_path, final_groups)

    report = {
        'ok': len(anomalies) == 0,
        'input': {
            'start_page': args.start_page,
            'end_page': args.end_page,
            'retry_failed_pages': args.retry_failed_pages,
            'target_pages': {
                '最满意': sat_target_pages,
                '最不满意': unsat_target_pages,
            },
        },
        'sat_total_raw': len(sat_rows),
        'unsat_total_raw': len(unsat_rows),
        'aligned_total': sum(len(v) for v in final_groups.values()),
        'groups': {k: len(v) for k, v in final_groups.items()},
        'page_link_counts': {
            '最满意': sat_pages,
            '最不满意': unsat_pages,
        },
        'raw_parse_anomalies': len(sat_bad) + len(unsat_bad),
        'alignment_meta': meta,
        'anomalies': anomalies,
        'merge': merge_summary,
    }
    report_path = out_path.with_suffix('.validation.json')
    write_validation_report(report_path, report)
    failed_pages_path = out_path.with_suffix('.failed-pages.json')
    failed_pages_path.write_text(json.dumps(failed_page_list, ensure_ascii=False, indent=2), encoding='utf-8')

    tracker.update(
        stage='完成',
        current=total_steps,
        total=total_steps,
        success=sat_success + unsat_success,
        retry=sat_retry + unsat_retry,
        failed=sat_failed + unsat_failed + len(anomalies),
        records=sum(len(v) for v in final_groups.values()),
        validation_path=str(report_path),
        failed_pages=failed_page_list,
        message=f'导出完成: {out_path.name}',
    )

    print(f'Wrote: {out_path}')
    print(f'Validation: {report_path}')
    if args.retry_failed_pages:
        print(f'Retry pages 最满意: {sat_target_pages}')
        print(f'Retry pages 最不满意: {unsat_target_pages}')
    print('Counts:')
    for k, v in final_groups.items():
        print(f'- {k}: {len(v)}')
    if merge_summary:
        print(f"Merged into existing workbook: {merge_summary['merge_into']}")
        print(f"Existing rows: {merge_summary['existing_rows']}")
        print(f"New rows: {merge_summary['new_rows']}")
        print(f"Final rows: {merge_summary['merged_rows']}")
    print(f'- raw sat: {len(sat_rows)}')
    print(f'- raw unsat: {len(unsat_rows)}')
    print(f'- anomalies: {len(anomalies)}')
    if report['ok']:
        print('- validation: OK')
    else:
        print('- validation: FAILED')
        if args.strict_validate:
            raise SystemExit(2)


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
