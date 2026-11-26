# xlsx_append.py
"""기존 엑셀(.xlsx)에 행을 '이어서' 붙인다 — 파일 안쪽 XML만 고쳐서 기존 서식·링크·열 너비를 그대로 둔다.
한셀 등에서 만든 파일처럼 openpyxl로 열면 깨지는 파일도 처리할 수 있게 직접 수정한다.

- 새 행의 칸 서식은 바로 윗줄 같은 열의 서식을 그대로 복사
- 글자는 inlineStr(공유 문자열 표를 건드리지 않음)
- url 열은 클릭되는 하이퍼링크로 (시트 rels에 관계 추가)
- 머리글 줄에 자동 필터(정렬·필터 버튼)를 켬
"""
import os
import re
import shutil
import tempfile
import zipfile
from xml.sax.saxutils import escape

_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_HL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'


def _col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _first_sheet(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rid = re.search(r'<(?:\w+:)?sheet\b[^>]*\br:id="([^"]+)"', wb).group(1)
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    tgt = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"' % re.escape(rid), rels)
    if not tgt:
        tgt = re.search(r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="%s"' % re.escape(rid), rels)
    t = tgt.group(1).lstrip('/')
    path = t if t.startswith('xl/') else 'xl/' + t
    name = re.search(r'<(?:\w+:)?sheet\b[^>]*\bname="([^"]+)"', wb).group(1)
    return path, name


def append_rows(path, rows, link_col=None):
    """rows: [[값, 값, ...], ...] (A열부터). link_col: 하이퍼링크로 만들 열 번호(1부터), 없으면 None"""
    if not rows:
        return 0
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        infos = {i.filename: i for i in z.infolist()}
        sheet, sheet_name = _first_sheet(z)
    xml = parts[sheet].decode('utf-8')
    m = re.search(r'<(\w+:)?worksheet\b', xml)
    p = m.group(1) or ''                       # 'x:' (한셀) 또는 '' (엑셀)
    rp = re.search(r'xmlns:(\w+)="%s"' % re.escape(_REL), xml)
    if rp:
        rpre = rp.group(1)
    else:                                      # r 네임스페이스 선언이 없으면 추가
        rpre = 'r'
        xml = re.sub(r'(<%sworksheet\b)' % p, r'\1 xmlns:r="%s"' % _REL, xml, count=1)

    # 마지막 행 번호와 그 행의 열별 서식
    row_tags = list(re.finditer(r'<%srow\b[^>]*\br="(\d+)"[^>]*>(.*?)</%srow>' % (p, p), xml, re.S))
    last_r = int(row_tags[-1].group(1)) if row_tags else 0
    styles = {}
    if row_tags:
        for c in re.finditer(r'<%sc\b[^>]*\br="([A-Z]+)\d+"[^>]*>' % p, row_tags[-1].group(2)):
            sm = re.search(r'\bs="(\d+)"', c.group(0))
            if sm:
                styles[c.group(1)] = sm.group(1)
    ncol = max(len(r) for r in rows)

    # 시트 rels (하이퍼링크)
    rels_path = sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    rels = parts.get(rels_path, b'').decode('utf-8') or \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    next_id = (max(used) + 1) if used else 1

    new_rows, new_links, new_rels = [], [], []
    for k, vals in enumerate(rows, 1):
        r = last_r + k
        cells = []
        for ci in range(1, ncol + 1):
            v = vals[ci - 1] if ci - 1 < len(vals) else ''
            col = _col(ci)
            s_attr = f' s="{styles[col]}"' if col in styles else ''
            if v is None or v == '':
                cells.append(f'<{p}c r="{col}{r}"{s_attr}/>')
            else:
                cells.append(f'<{p}c r="{col}{r}"{s_attr} t="inlineStr"><{p}is><{p}t xml:space="preserve">'
                             f'{escape(str(v))}</{p}t></{p}is></{p}c>')
            if link_col == ci and v:
                rid = f'rId{next_id}'; next_id += 1
                new_links.append(f'<{p}hyperlink ref="{col}{r}" {rpre}:id="{rid}"/>')
                new_rels.append(f'<Relationship Id="{rid}" Type="{_HL}" Target="{escape(str(v), {chr(34): "&quot;"})}" '
                                f'TargetMode="External"/>')
        new_rows.append(f'<{p}row r="{r}" spans="1:{ncol}">' + ''.join(cells) + f'</{p}row>')
    end_r = last_r + len(rows)

    # 행 추가
    if f'</{p}sheetData>' in xml:
        xml = xml.replace(f'</{p}sheetData>', ''.join(new_rows) + f'</{p}sheetData>', 1)
    else:
        xml = re.sub(r'<%ssheetData\s*/>' % p, f'<{p}sheetData>' + ''.join(new_rows) + f'</{p}sheetData>', xml, 1)
    # 범위 갱신
    last_col = _col(max(ncol, 5))
    xml = re.sub(r'(<%sdimension\b[^>]*\bref=")[^"]*(")' % p, r'\g<1>A1:%s%d\g<2>' % (last_col, end_r), xml, 1)
    # 자동 필터 (머리글 줄 정렬·필터 버튼)
    af = f'<{p}autoFilter ref="A1:{last_col}{end_r}"/>'
    if re.search(r'<%sautoFilter\b' % p, xml):
        xml = re.sub(r'(<%sautoFilter\b[^>]*\bref=")[^"]*(")' % p, r'\g<1>A1:%s%d\g<2>' % (last_col, end_r), xml, 1)
    else:
        anchor = None
        for tag in ('scenarios', 'protectedRanges', 'sheetProtection', 'sheetCalcPr'):
            mm = list(re.finditer(r'<%s%s\b[^>]*/>|</%s%s>' % (p, tag, p, tag), xml))
            if mm:
                anchor = mm[-1].end(); break
        if anchor is None:
            anchor = xml.index(f'</{p}sheetData>') + len(f'</{p}sheetData>')
        xml = xml[:anchor] + af + xml[anchor:]
    # 하이퍼링크
    if new_links:
        if f'</{p}hyperlinks>' in xml:
            xml = xml.replace(f'</{p}hyperlinks>', ''.join(new_links) + f'</{p}hyperlinks>', 1)
        else:
            block = f'<{p}hyperlinks>' + ''.join(new_links) + f'</{p}hyperlinks>'
            mm = re.search(r'<%s(printOptions|pageMargins|pageSetup|headerFooter|drawing|legacyDrawing|tableParts|extLst)\b' % p, xml)
            pos = mm.start() if mm else xml.rindex(f'</{p}worksheet>')
            xml = xml[:pos] + block + xml[pos:]
        rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>', 1)
        parts[rels_path] = rels.encode('utf-8')
    parts[sheet] = xml.encode('utf-8')

    # 필터용 이름 (엑셀이 자동 필터 범위를 기억하는 곳)
    wb = parts['xl/workbook.xml'].decode('utf-8')
    wp = re.search(r'<(\w+:)?workbook\b', wb).group(1) or ''
    sn = sheet_name if re.fullmatch(r'\w+', sheet_name) else "'" + sheet_name.replace("'", "''") + "'"
    dn = (f'<{wp}definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
          f'{escape(sn)}!$A$1:${last_col}${end_r}</{wp}definedName>')
    if '_xlnm._FilterDatabase' in wb:
        wb = re.sub(r'(_xlnm\._FilterDatabase"[^>]*>)[^<]*', lambda mm: mm.group(1) + escape(sn) +
                    f'!$A$1:${last_col}${end_r}', wb, 1)
    elif f'</{wp}definedNames>' in wb:
        wb = wb.replace(f'</{wp}definedNames>', dn + f'</{wp}definedNames>', 1)
    else:
        wb = wb.replace(f'</{wp}sheets>', f'</{wp}sheets><{wp}definedNames>' + dn + f'</{wp}definedNames>', 1)
    parts['xl/workbook.xml'] = wb.encode('utf-8')

    # 다시 묶기 (임시 파일 → 교체; 엑셀에서 열려 있으면 PermissionError)
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(os.path.abspath(path)))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as out:
            for name, data in parts.items():
                zi = infos.get(name) or zipfile.ZipInfo(name)
                zi.compress_type = zipfile.ZIP_DEFLATED
                out.writestr(zi, data)
        with open(path, 'ab'):
            pass                                # 열려 있는지 먼저 확인
        shutil.move(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return len(rows)
