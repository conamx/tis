# excel_source.py
"""사례 엑셀 읽기 → 지침 입력 형식(라벨형)으로 변환.

- 시트마다 열 순서가 달라도 머리글 이름(사건번호/범죄명/결과/결과2/요약/처분이유…)으로 찾는다.
- '이름'(실명), '범죄명 분류', '사용체크', '비고', '순' 열은 AI에 보내지 않는다.
- 이미 쓴 사례는 사용기록.csv(프로젝트별) + 엑셀 '사용체크' 표시로 판별한다.
- 원본 엑셀은 읽기만 하고 절대 수정하지 않는다.
"""
import csv
import os
from datetime import datetime

from app_config import USAGE_LOG


def _s(v):
    if v is None:
        return ''
    return str(v).replace('\r\n', '\n').replace('\r', '\n').strip()


class Case:
    def __init__(self, sheet, row_no, data):
        self.sheet = sheet
        self.row_no = row_no
        self.case_id = data.get('사건번호', '')
        self.crime = data.get('범죄명', '')
        self.category = data.get('범죄명 분류', '')
        self.result1 = data.get('결과', '')
        self.result2 = data.get('결과2', '')
        self.summary = data.get('요약', '')
        self.reason = data.get('처분이유', '')
        self.reason2 = data.get('처분이유2', '')
        self.bad = data.get('처분이유(불리한점)', '')
        self.good = data.get('처분이유(유리한점)', '')
        self.used_mark = data.get('사용체크', '')

    @property
    def key(self):
        return self.case_id or f'{self.sheet}#{self.row_no}'

    def is_usable(self):
        empty = ('', '(내용없음)')
        return bool(self.crime) and self.summary not in empty

    def to_prompt(self):
        """지침의 라벨형 입력. 6항목(불리/유리 분리) 또는 5항목(종합 처분이유)."""
        lines = []
        if self.case_id:
            lines.append(f'사건번호: {self.case_id}')
        lines.append(f'범죄명: {self.crime}')
        lines.append(f'결과1: {self.result1}')
        lines.append(f'결과2: {self.result2}')
        lines.append(f'사건 요약: {self.summary}')
        if self.bad or self.good:
            lines.append(f'불리한 점: {self.bad}')
            lines.append(f'유리한 점: {self.good}')
        else:
            reason = self.reason
            if self.reason2:
                reason = (reason + '\n' + self.reason2).strip()
            lines.append(f'종합 처분이유: {reason}')
        return '\n'.join(lines)

    def short(self):
        return self.summary.replace('\n', ' ')[:60]


def _find_header(ws, max_scan=10):
    for r in range(1, max_scan + 1):
        vals = [_s(c.value) for c in ws[r]]
        if '사건번호' in vals and '요약' in vals:
            return r, vals
    return None, None


def load_cases(path):
    """{시트명: [Case, ...]}"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    try:
        for ws in wb.worksheets:
            hr, header = _find_header(ws)
            if not hr:
                continue
            idx = {}
            for i, h in enumerate(header):
                h = h.replace(' ', '') if h.startswith('처분이유') else h
                if h and h not in idx:
                    idx[h] = i
            cases = []
            for r_no, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), start=hr + 1):
                data = {k: _s(row[i]) if i < len(row) else '' for k, i in idx.items()}
                c = Case(ws.title, r_no, data)
                if c.is_usable():
                    cases.append(c)
            out[ws.title] = cases
    finally:
        wb.close()
    return out


# ── 사용기록 ────────────────────────────────────────────────

def load_usage(project_name):
    used = set()
    if not os.path.exists(USAGE_LOG):
        return used
    with open(USAGE_LOG, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            if row.get('프로젝트') == project_name:
                used.add(row.get('사례키', ''))
    return used


def record_usage(project_name, case):
    new = not os.path.exists(USAGE_LOG)
    with open(USAGE_LOG, 'a', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=['날짜', '프로젝트', '시트', '사례키', '범죄명'])
        if new:
            w.writeheader()
        w.writerow({'날짜': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    '프로젝트': project_name, '시트': case.sheet,
                    '사례키': case.key, '범죄명': case.crime.replace('\n', ' ')})


# ── 로톡 업로드 기록 (새 엑셀, 사례 엑셀 원본은 건드리지 않음) ─────────────

UPLOAD_HEADER = ['날짜', '섹션', '사건번호', '링크', '제목']


def record_upload(path, rows):
    """rows: [{'지침','사건번호','링크','제목'}, ...] → 업로드기록.xlsx 맨 아래에 추가.
    같은 링크가 이미 있으면 건너뜀. 반환: 추가한 행 수"""
    import openpyxl
    from openpyxl.styles import Font
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '로톡 업로드'
        ws.append(UPLOAD_HEADER)
        for c in ws[1]:
            c.font = Font(bold=True)
        for col, w in zip('ABCDE', (17, 18, 18, 50, 60)):
            ws.column_dimensions[col].width = w
    have = {str(r[3]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r and r[3]}
    added = 0
    for r in rows:
        url = (r.get('링크') or '').strip()
        if not url or url in have:
            continue
        ws.append([datetime.now().strftime('%Y-%m-%d %H:%M'), r.get('섹션', ''), r.get('사건번호', ''),
                   url, r.get('제목', '')])
        cell = ws.cell(row=ws.max_row, column=4)
        cell.hyperlink = url
        cell.font = Font(color='0563C1', underline='single')
        have.add(url)
        added += 1
    wb.save(path)   # 엑셀에서 열려 있으면 PermissionError
    return added
