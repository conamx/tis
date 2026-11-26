# onenote_com.py
import os, re, shutil, subprocess, tempfile, xml.etree.ElementTree as ET
from pathlib import Path

_DEFAULT_TIMEOUT = 300
_PAGE_TIMEOUT    = 300
_MAX_RETRY       = 2

# ──────────────────────────────────────────────────────────────
# 서식 설정 (마크다운형 원고용) — 복붙 시 네이버 블로그 글자크기에 맞춘 값
#   11pt ≈ 15px(기본) / 14pt ≈ 19px / 18pt ≈ 24px / 21pt ≈ 28px
#   크기를 바꾸고 싶으면 여기 숫자만 고치면 됩니다.
# ──────────────────────────────────────────────────────────────
SIZE_BODY    = 11                              # 본문
SIZE_HEADING = {1: 21, 2: 18, 3: 14, 4: 12, 5: 11, 6: 11}   # # ~ ###### 소제목
SIZE_KEYWORD = 12                              # 상단 키워드 (굵게)
SIZE_FOOTER  = 20                              # 하단 문구 (기존과 동일)
TABLE_ALIGN  = 'center'
# 로톡 대제목/소제목: OneNote '제목' 스타일(h2/h3)로 넣어서 복사할 때 제목으로 넘어가게 함
#   # ## → 대제목(h2), ### 이하 → 소제목(h3). 하단 문구는 제목이 아님(굵은 글씨만)
HEADING_STYLE = {1: 'h2', 2: 'h2', 3: 'h3', 4: 'h3', 5: 'h3', 6: 'h3'}
HEADING_FONT_PT = {'h2': 18, 'h3': 14}                        # 표(텍스트 줄) 정렬: 'center' / 'left' / 'right'

TITLE_CHARS_PER_LINE = 28    # 제목이 한 줄에 들어가는 글자 수(대략). 겹치면 줄이고, 간격이 크면 늘리세요.
TITLE_LINE_HEIGHT    = 30    # 제목 한 줄 높이(pt)
BODY_TOP_Y           = 86.4  # 제목이 한 줄일 때 본문 시작 위치(pt)


def _body_top_y(title):
    """제목이 여러 줄로 꺾이면 그만큼 본문 시작 위치를 아래로 내려 겹침을 막는다."""
    import math
    lines = max(1, math.ceil(len(title or '') / float(TITLE_CHARS_PER_LINE)))
    return BODY_TOP_Y + (TITLE_LINE_HEIGHT * (lines - 1) + 12 if lines > 1 else 0)


_ONE_NS = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
_INLINE_TOKEN = re.compile(r'\*\*|\*')


def _esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _cdata(s):
    return '<![CDATA[' + s.replace(']]>', ']]]]><![CDATA[>') + ']]>'


def _inline_runs(text, force_bold=False):
    """한 문단의 인라인 표기를 (글자, 볼드, 이탤릭, 색없음) 조각으로 분해.
    **볼드**, *이탤릭* 지원. ([BLUE] 표시는 체크용이라 건드리지 않고 글자 그대로 둔다) 짝이 안 맞는 기호는 화면에 남지 않게 버린다."""
    tokens, pos = [], 0
    for m in _INLINE_TOKEN.finditer(text):
        if m.start() > pos:
            tokens.append(('t', text[pos:m.start()]))
        v = m.group()
        tokens.append(('m', v))
        pos = m.end()
    if pos < len(text):
        tokens.append(('t', text[pos:]))
    drop_b = sum(1 for k, v in tokens if k == 'm' and v == '**') % 2 == 1
    drop_i = sum(1 for k, v in tokens if k == 'm' and v == '*') % 2 == 1
    bold = ital = blue = False
    runs = []
    for k, v in tokens:
        if k == 't':
            if v:
                runs.append((v, bold or force_bold, ital, blue))
        elif v == '**':
            if not drop_b: bold = not bold
        elif v == '*':
            if not drop_i: ital = not ital
    return runs


def _run_html(text, size, bold, ital, blue):
    st = [f'font-size:{size:.1f}pt']
    if bold: st.append('font-weight:bold')
    if ital: st.append('font-style:italic')
    return f"<span style='{';'.join(st)}'>{_esc(text)}</span>"


def _oe_para(text, size=SIZE_BODY, bold=False, align=None, qs=None):
    """문단 1개 → <one:OE>. 모든 글자에 크기/볼드를 명시해서 복붙해도 서식이 유지되게 한다.
    qs='h2'/'h3' 이면 OneNote 제목 스타일 자리표시를 단다 (페이지 만들 때 실제 번호로 바뀜)."""
    html = ''.join(_run_html(t, size, b, i, c)
                   for t, b, i, c in _inline_runs(text, force_bold=bold))
    st = [f'font-size:{size:.1f}pt']
    if bold: st.append('font-weight:bold')
    al = f' alignment="{align}"' if align else ''
    q = f' quickStyleIndex="__QS_{qs}__"' if qs else ''
    return f"<one:OE{q}{al} style='{';'.join(st)}'><one:T>{_cdata(html)}</one:T></one:OE>"


def _oe_blank(size=SIZE_BODY):
    return f"<one:OE style='font-size:{size:.1f}pt'><one:T><![CDATA[]]></one:T></one:OE>"


def _table_lines(tbl):
    """표를 진짜 OneNote 표로 만들지 않고, 한 행 = 한 줄 ("칸1 | 칸2") 텍스트로 만든다.
    TABLE_ALIGN 으로 줄 정렬, 제목 행은 볼드."""
    out = []
    for ri, row in enumerate(tbl['rows']):
        cells = [re.sub(r'\s*<br\s*/?>\s*', ' ', c, flags=re.I).strip() for c in row]
        out.append(_oe_para(' | '.join(cells), SIZE_BODY,
                            bold=(tbl['header'] and ri == 0), align=TABLE_ALIGN))
    return ''.join(out)


def build_md_body_xml(article, top_keyword='', footer_presets=None):
    """마크다운형 원고(blocks)를 OneNote <one:OE> 목록 XML 문자열로 만든다.
    표가 없는 글이면 표 부분이 그냥 생략된다."""
    parts = []
    if top_keyword and top_keyword.strip():
        parts.append(_oe_para(top_keyword.strip(), SIZE_KEYWORD, bold=True))
        parts.append(_oe_blank())
    for b in article.get('blocks') or []:
        t = b.get('type')
        if t == 'blank':
            parts.append(_oe_blank())
        elif t == 'heading':
            lv = b.get('level', 2)
            qs = HEADING_STYLE.get(lv)
            parts.append(_oe_para(b['text'], HEADING_FONT_PT.get(qs, SIZE_HEADING.get(lv, SIZE_BODY)),
                                  bold=True, qs=qs))
        elif t == 'bullet':
            parts.append(_oe_para('• ' + b['text']))
        elif t == 'table':
            parts.append(_table_lines(b))
        else:                                   # para, quote
            parts.append(_oe_para(b.get('text', '')))
    if footer_presets:
        parts.append(_oe_blank())
        for preset_lines in footer_presets:
            parts.append(_oe_blank())
            for text, is_bold in preset_lines:
                parts.append(_oe_para(text, SIZE_FOOTER, bold=is_bold))
    return ''.join(parts)


# 페이지에 OneNote 제목 스타일(h2=대제목, h3=소제목)을 등록하고 본문의 자리표시를 실제 번호로 바꾸는 PowerShell
#  - 이미 같은 이름(h2/h3) 스타일이 있으면 그 번호를 쓰되 크기·굵기·색을 우리 값으로 맞춤 (OneNote 기본 제목색 파랑 방지)
#  - 새로 만들 때는 기존 QuickStyleDef 뒤(없으면 TagDef 뒤, 그것도 없으면 맨 앞)에 끼움 (OneNote XML 순서 규칙)
QUICKSTYLE_PS = (
    "$qsDefs=@(@('h2','" + str(HEADING_FONT_PT['h2']) + ".0'),@('h3','" + str(HEADING_FONT_PT['h3']) + ".0'))\n"
    "$qsAll=@($doc.DocumentElement.SelectNodes('one:QuickStyleDef',$mgr))\n"
    "$qsMax=-1; $qsFont='Malgun Gothic'\n"
    "foreach($q in $qsAll){ $i=[int]$q.GetAttribute('index'); if($i -gt $qsMax){$qsMax=$i}; "
    "if($q.GetAttribute('name') -eq 'p' -and $q.GetAttribute('font')){$qsFont=$q.GetAttribute('font')} }\n"
    "foreach($d in $qsDefs){\n"
    "  $hit=$null; foreach($q in $qsAll){ if($q.GetAttribute('name') -eq $d[0]){$hit=$q} }\n"
    "  if(-not $hit){\n"
    "    $qsMax++; $hit=$doc.CreateElement('one:QuickStyleDef',$ns); $hit.SetAttribute('index',[string]$qsMax)\n"
    "    $hit.SetAttribute('name',$d[0])\n"
    "    $last=$doc.DocumentElement.SelectSingleNode('one:QuickStyleDef[last()]',$mgr)\n"
    "    if(-not $last){ $last=$doc.DocumentElement.SelectSingleNode('one:TagDef[last()]',$mgr) }\n"
    "    if($last){ $doc.DocumentElement.InsertAfter($hit,$last)|Out-Null } else { $doc.DocumentElement.PrependChild($hit)|Out-Null }\n"
    "    $qsAll=@($doc.DocumentElement.SelectNodes('one:QuickStyleDef',$mgr))\n"
    "  }\n"
    "  $hit.SetAttribute('fontColor','automatic'); $hit.SetAttribute('highlightColor','automatic')\n"
    "  $hit.SetAttribute('font',$qsFont); $hit.SetAttribute('fontSize',$d[1]); $hit.SetAttribute('bold','true')\n"
    "  $hit.SetAttribute('spaceBefore','0.0'); $hit.SetAttribute('spaceAfter','0.0')\n"
    "  $bodyXml=$bodyXml.Replace('__QS_'+$d[0]+'__',$hit.GetAttribute('index'))\n"
    "}\n"
)


# 제목 맨 앞의 'ㅡ' 표시(업로드 전 표시)를 지움. 제목 안의 서식 태그는 그대로 둠
TITLE_DASH_RE = r'^((?:<[^>]+>)*)\s*[ㅡ—–\-]+\s*'
TITLE_DASH_PS = (
    "$tt=$doc.SelectSingleNode('//one:Title/one:OE/one:T',$mgr)\n"
    "if($tt){\n"
    "  $tv=$tt.InnerText\n"
    "  $nv=[regex]::Replace($tv,'" + TITLE_DASH_RE + "','$1')\n"
    "  if($nv -ne $tv){ $tt.InnerXml=''; $tt.AppendChild($doc.CreateCDataSection($nv))|Out-Null }\n"
    "}\n"
)


def strip_title_dash(title):
    """TITLE_DASH_PS 와 같은 규칙 (화면 표시·테스트용)"""
    return re.sub(TITLE_DASH_RE, r'\1', title or '')


# 섹션의 페이지마다: ID, 제목, 만든 시각, 본문 텍스트 앞부분(사건번호·로톡 링크 찾기용)을 JSON으로 저장
PAGES_PS = (
    "$ErrorActionPreference='Stop'\n"
    "$sid=[System.IO.File]::ReadAllText(__SID__,[System.Text.Encoding]::UTF8).Trim()\n"
    "$on=New-Object -ComObject OneNote.Application\n"
    "[string]$h=''\n"
    "$on.GetHierarchy($sid,4,[ref]$h)\n"
    "$hd=[xml]$h\n"
    "$m=New-Object System.Xml.XmlNamespaceManager($hd.NameTable)\n"
    "$m.AddNamespace('one',$hd.DocumentElement.NamespaceURI)\n"
    "$rows=@()\n"
    "foreach($pg in $hd.SelectNodes('//one:Page',$m)){\n"
    "  [string]$x=''\n"
    "  try{ $on.GetPageContent($pg.GetAttribute('ID'),[ref]$x,0) }catch{ continue }\n"
    "  $d=[xml]$x\n"
    "  $m2=New-Object System.Xml.XmlNamespaceManager($d.NameTable)\n"
    "  $m2.AddNamespace('one',$d.DocumentElement.NamespaceURI)\n"
    "  $ts=@($d.SelectNodes('//one:Outline//one:T',$m2) | Select-Object -First 6 | ForEach-Object { $_.InnerText })\n"
    "  $rows+=[pscustomobject]@{id=$pg.GetAttribute('ID');title=$pg.GetAttribute('name');"
    "time=$pg.GetAttribute('dateTime');head=($ts -join \"`n\")}\n"
    "}\n"
    "$json=ConvertTo-Json -InputObject @($rows) -Depth 3\n"
    "[System.IO.File]::WriteAllText(__OUT__,$json,[System.Text.Encoding]::UTF8)\n"
)


def parse_page_rows(data):
    """PAGES_PS 결과 → 사건번호·로톡 링크를 뽑아 정리 (OneNote 순서 그대로)"""
    rows = []
    for r in data or []:
        head = r.get('head') or ''
        text = re.sub(r'<[^>]+>', ' ', head)
        cm = re.search(r'사건번호\s*[:：]\s*([0-9]{2,4}\s*[가-힣]{1,4}\s*[0-9]{1,7}(?:호)?)', text)
        lm = re.search(r'href=["\']([^"\']+)["\']', head) if '로톡 링크' in head else None
        if not lm and '로톡 링크' in text:
            lm = re.search(r'로톡 링크\s*[:：]\s*(https?://\S+)', text)
        rows.append({'id': r.get('id', ''), 'title': (r.get('title') or '').strip(),
                     'time': r.get('time') or '',
                     'case_no': cm.group(1).replace(' ', '') if cm else '',
                     'link': lm.group(1) if lm else ''})
    return rows   # OneNote 페이지 목록과 같은 순서 (한 번에 붙여넣기가 위에서부터 맞게 들어가도록)


class SectionInfo:
    def __init__(self, sid, name, notebook):
        self.id = sid; self.name = name; self.notebook = notebook
    def display_name(self):
        return f"{self.notebook} > {self.name}"


class OneNoteManager:
    def __init__(self):
        self.section_id = ''; self.connected = False; self._sections = []
        self.last_page_id = ''   # 마지막으로 만든 페이지 ID (다시 쓰기 때 옛 페이지 교체용)

    def connect(self):
        try:
            self._get_hierarchy_root(); self.connected = True; return True
        except Exception as e:
            print(f"[OneNote] 연결 실패: {e}"); self.connected = False; return False

    def get_all_sections(self):
        root = self._get_hierarchy_root(); sections = []
        def _walk(node, nb=''):
            for child in node:
                tag = re.sub(r'\{.*?\}', '', child.tag)
                if tag == 'Notebook':   _walk(child, child.get('name','?'))
                elif tag == 'SectionGroup': _walk(child, nb)
                elif tag == 'Section':
                    s = child.get('ID','')
                    if s: sections.append(SectionInfo(s, child.get('name','?'), nb))
        _walk(root); self._sections = sections; return sections

    def set_section(self, sid): self.section_id = sid

    def get_first_section_id(self):
        return self._find_section_id(self._get_hierarchy_root())

    def create_page(self, article, title_prefix='', add_footer=False, footer_presets=None, top_keyword='', retry=_MAX_RETRY):
        if not self.section_id:
            self.section_id = self.get_first_section_id()
        for attempt in range(1, retry + 2):
            try:
                prefix = (title_prefix.strip() + ' ') if title_prefix.strip() else ''
                title  = f"{prefix}{article['title']}"
                self._create_page_all(title, article, footer_presets=footer_presets or [], top_keyword=top_keyword)
                return True
            except Exception as e:
                print(f"[OneNote] 페이지 생성 실패 (시도 {attempt}/{retry+1}): {e}")
                if attempt > retry:
                    import traceback; traceback.print_exc(); return False
        return False

    def batch_create_pages(self, articles, title_prefix='', progress_cb=None, add_footer=False, footer_presets=None, top_keyword=''):
        result = {'success': 0, 'fail': 0, 'failed_nums': []}
        footer_presets = footer_presets or []
        for i, art in enumerate(articles, 1):
            if progress_cb: progress_cb(i, len(articles), art)
            if self.create_page(art, title_prefix=title_prefix, add_footer=add_footer, footer_presets=footer_presets, top_keyword=top_keyword):
                result['success'] += 1
            else:
                result['fail'] += 1; result['failed_nums'].append(art.get('number','?'))
        if add_footer:
            try:
                self.create_page({'number': '0', 'title': 'ㅡ' * 20,
                                   'focus_keyword': '', 'sub_keywords': '',
                                   'toc': [], 'summary': '', 'sections': [],
                                   'raw_content': ''}, title_prefix='')
            except Exception:
                pass
        return result

    # ── PowerShell ────────────────────────────────────────────

    def _run_ps_file(self, path, timeout=_DEFAULT_TIMEOUT):
        r = subprocess.run(
            ['powershell','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File', path],
            capture_output=True, timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW)
        out = self._dec(r.stdout); err = self._dec(r.stderr)
        if r.returncode != 0:
            raise RuntimeError(f"PowerShell 실행 실패 (코드 {r.returncode}): {(err or out).strip()}")
        return out

    @staticmethod
    def _dec(b):
        for enc in ('utf-8','cp949','euc-kr'):
            try: return b.decode(enc)
            except: pass
        return b.decode('utf-8', errors='replace')

    @staticmethod
    def _q(s): return "'" + s.replace("'","''") + "'"

    # ── 계층 조회 ─────────────────────────────────────────────

    def _get_hierarchy_root(self):
        tmp = tempfile.mkdtemp(prefix='on_hier_')
        try:
            hf  = os.path.join(tmp,'h.xml')
            ps1 = os.path.join(tmp,'h.ps1')
            Path(ps1).write_text(
                "$ErrorActionPreference='Stop'\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$xml=''\n"
                "$on.GetHierarchy('',4,[ref]$xml)\n"
                "[System.IO.File]::WriteAllText("+self._q(hf)+","
                "$xml,[System.Text.Encoding]::UTF8)\n", encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_DEFAULT_TIMEOUT)
            return ET.fromstring(Path(hf).read_text(encoding='utf-8-sig'))
        finally: shutil.rmtree(tmp, ignore_errors=True)

    @staticmethod
    def _find_section_id(root):
        def _s(node):
            for c in node:
                tag = re.sub(r'\{.*?\}','',c.tag)
                if tag=='Section':
                    v=c.get('ID','')
                    if v: return v
                elif tag in ('Notebook','SectionGroup'):
                    r=_s(c)
                    if r: return r
            return ''
        sid=_s(root)
        if not sid: raise RuntimeError('섹션을 찾을 수 없습니다.')
        return sid

    # ── 페이지 단계 ───────────────────────────────────────────

    def _create_page_all(self, title, article, footer_presets=None, top_keyword=''):
        """페이지 생성 + 제목 + 본문을 단일 PowerShell 세션에서 처리"""
        if article.get('blocks') is not None:      # 마크다운형 원고("# 1" 형식)
            return self._create_page_md(title, article, footer_presets=footer_presets,
                                        top_keyword=top_keyword)
        from parser import strip_markdown, to_onenote_richtext
        tmp = tempfile.mkdtemp(prefix='on_page_')
        try:
            sf   = os.path.join(tmp,'sid.txt')
            tf   = os.path.join(tmp,'title.txt')
            ps1  = os.path.join(tmp,'p.ps1')
            Path(sf).write_text(self.section_id, encoding='utf-8')
            Path(tf).write_text(title, encoding='utf-8')

            # 본문 라인 구성
            lines = []
            def add(text, h=False, size=None): lines.append((text, h, size))
            # 상단 키워드 삽입 (본문 맨 위, 굵게)
            if top_keyword and top_keyword.strip():
                add(to_onenote_richtext(top_keyword.strip()), h=True, size=12)
                add('')
            if article.get('focus_keyword'):
                add(to_onenote_richtext("Focus Keyword: " + article['focus_keyword']))
            if article.get('sub_keywords'):
                add(to_onenote_richtext("보조 키워드: " + article['sub_keywords']))
            if article.get('focus_keyword') or article.get('sub_keywords'):
                add('')
            has_toc = bool(article.get('toc'))
            has_sum = bool(article.get('summary','').strip())
            has_sec = bool(article.get('sections'))
            if has_toc:
                add('목차', h=True)
                for item in article['toc']:
                    rich = to_onenote_richtext(item)
                    add(f"  {rich}" if re.match(r'^\d+[\.\)]', item) else f"  • {rich}")
                add('')
            if has_sum:
                add('요약문', h=True)
                for p in re.split(r'\n\s*\n', article['summary'].strip()):
                    p = p.strip()
                    if p:
                        add(to_onenote_richtext(p))
                add('')
            if has_sec:
                for st, sb in article['sections']:
                    parts = sb.strip().split('\n\n')
                    # 본문 첫 줄이 단독 숫자면 섹션 제목 위에 먼저 출력
                    if parts and re.match(r'^\d+$', parts[0].strip()):
                        add(parts[0].strip()); add('')
                        parts = parts[1:]
                    add(to_onenote_richtext(strip_markdown(st)), h=True); add('')
                    for p in parts:
                        p = p.strip()
                        if not p:
                            continue
                        if re.match(r'^\d+$', p):
                            add(p); add('')
                        else:
                            add(to_onenote_richtext(p)); add('')
            if not has_toc and not has_sum and not has_sec:
                add('본문', h=True); add('')
                for ln in re.split(r'\n\s*\n', article.get('raw_content','').strip()):
                    ln = ln.strip()
                    if ln:
                        add(to_onenote_richtext(ln)); add('')
            if footer_presets:
                add('')
                for preset_lines in footer_presets:
                    add('')
                    for text, is_bold in preset_lines:
                        add(to_onenote_richtext(text), h=is_bold, size=20)

            # 라인 파일 저장
            meta = []
            for i,(t,h,sz) in enumerate(lines):
                lf = os.path.join(tmp,f'l{i}.txt')
                Path(lf).write_text(t, encoding='utf-8')
                meta.append((lf, h, sz))

            # 단일 PS 스크립트: 생성 + 제목 + 본문
            ps = (
                "$ErrorActionPreference='Stop'\n"
                "$sectionId=[System.IO.File]::ReadAllText("+self._q(sf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$newTitle=[System.IO.File]::ReadAllText("+self._q(tf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                # 페이지 생성
                "[string]$pageId=''\n"
                "$on.CreateNewPage($sectionId,[ref]$pageId,0)\n"
                # 제목 설정
                "[string]$xml=''\n"
                "$on.GetPageContent($pageId,[ref]$xml,0)\n"
                "$doc=[xml]$xml\n"
                "$ns=$doc.DocumentElement.NamespaceURI\n"
                "$mgr=New-Object System.Xml.XmlNamespaceManager($doc.NameTable)\n"
                "$mgr.AddNamespace('one',$ns)\n"
                "$tn=$doc.SelectSingleNode('//one:Title/one:OE/one:T',$mgr)\n"
                "if($tn){$tn.InnerText=$newTitle}\n"
                # Outline 준비
                "$outline=$doc.SelectSingleNode('//one:Outline',$mgr)\n"
                "if(-not $outline){\n"
                "  $outline=$doc.CreateElement('one:Outline',$ns)\n"
                "  $pos=$doc.CreateElement('one:Position',$ns)\n"
                "  $pos.SetAttribute('x','36');$pos.SetAttribute('y','"+str(_body_top_y(title))+"');$pos.SetAttribute('z','0')\n"
                "  $outline.AppendChild($pos)|Out-Null\n"
                "  $sz=$doc.CreateElement('one:Size',$ns)\n"
                "  $sz.SetAttribute('width','720');$sz.SetAttribute('height','100')\n"
                "  $outline.AppendChild($sz)|Out-Null\n"
                "  $doc.DocumentElement.AppendChild($outline)|Out-Null\n"
                "}\n"
                "$oec=$outline.SelectSingleNode('one:OEChildren',$mgr)\n"
                "if(-not $oec){\n"
                "  $oec=$doc.CreateElement('one:OEChildren',$ns)\n"
                "  $outline.AppendChild($oec)|Out-Null\n"
                "}\n"
            )
            for lf, is_h, sz in meta:
                ps += (
                    "$lineText=[System.IO.File]::ReadAllText("+self._q(lf)+","
                    "[System.Text.Encoding]::UTF8)\n"
                    "$oe=$doc.CreateElement('one:OE',$ns)\n"
                    "$tNode=$doc.CreateElement('one:T',$ns)\n"
                    # 줄 텍스트는 이미 to_onenote_richtext()에서 &/</> 이스케이프 + <b>/<i> 태그로 변환되어 있음.
                    # CDATA로 감싸 InnerXml에 넣어야 OneNote가 <b>/<i>를 서식으로 인식한다
                    # (InnerText를 쓰면 태그가 그대로 이스케이프되어 화면에 문자로 노출됨)
                    "$tNode.InnerXml=('<![CDATA[' + $lineText + ']]>')\n"
                )
                style_parts = []
                if is_h:
                    style_parts.append('font-weight:bold')
                if sz:
                    style_parts.append(f'font-size:{sz}.0pt')
                if style_parts:
                    ps += f"$oe.SetAttribute('style','{';'.join(style_parts)}')\n"
                ps += "$oe.AppendChild($tNode)|Out-Null\n$oec.AppendChild($oe)|Out-Null\n"
            ps += "$on.UpdatePageContent($doc.OuterXml,[DateTime]::MinValue)\n"

            Path(ps1).write_text(ps, encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=300)
        finally: shutil.rmtree(tmp, ignore_errors=True)

    def _create_page_md(self, title, article, footer_presets=None, top_keyword=''):
        """마크다운형 원고 전용: 본문 XML을 파이썬에서 완성해 한 번에 넣는다.
        (표/볼드/글자크기/색을 모두 XML로 만들고, OneNote 호출 전에 XML 문법을 먼저 검증)"""
        body_xml = build_md_body_xml(article, top_keyword=top_keyword,
                                     footer_presets=footer_presets)
        # OneNote를 건드리기 전에 XML이 올바른지 검증 → 문제 있으면 빈 페이지가 생기기 전에 실패
        ET.fromstring(f'<root xmlns:one="{_ONE_NS}">{body_xml}</root>')

        tmp = tempfile.mkdtemp(prefix='on_md_')
        try:
            sf  = os.path.join(tmp, 'sid.txt')
            tf  = os.path.join(tmp, 'title.txt')
            bf  = os.path.join(tmp, 'body.xml')
            pf  = os.path.join(tmp, 'pid.txt')
            ps1 = os.path.join(tmp, 'p.ps1')
            Path(sf).write_text(self.section_id, encoding='utf-8')
            Path(tf).write_text(title, encoding='utf-8')
            Path(bf).write_text(body_xml, encoding='utf-8')

            ps = (
                "$ErrorActionPreference='Stop'\n"
                "$sectionId=[System.IO.File]::ReadAllText("+self._q(sf)+",[System.Text.Encoding]::UTF8).Trim()\n"
                "$newTitle=[System.IO.File]::ReadAllText("+self._q(tf)+",[System.Text.Encoding]::UTF8).Trim()\n"
                "$bodyXml=[System.IO.File]::ReadAllText("+self._q(bf)+",[System.Text.Encoding]::UTF8)\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$pageId=''\n"
                "$on.CreateNewPage($sectionId,[ref]$pageId,0)\n"
                "[System.IO.File]::WriteAllText("+self._q(pf)+",$pageId,[System.Text.Encoding]::UTF8)\n"
                "[string]$xml=''\n"
                "$on.GetPageContent($pageId,[ref]$xml,0)\n"
                "$doc=[xml]$xml\n"
                "$ns=$doc.DocumentElement.NamespaceURI\n"
                "$mgr=New-Object System.Xml.XmlNamespaceManager($doc.NameTable)\n"
                "$mgr.AddNamespace('one',$ns)\n"
                "$tn=$doc.SelectSingleNode('//one:Title/one:OE/one:T',$mgr)\n"
                "if($tn){$tn.InnerText=$newTitle}\n"
                "$outline=$doc.SelectSingleNode('//one:Outline',$mgr)\n"
                "if(-not $outline){\n"
                "  $outline=$doc.CreateElement('one:Outline',$ns)\n"
                "  $pos=$doc.CreateElement('one:Position',$ns)\n"
                "  $pos.SetAttribute('x','36');$pos.SetAttribute('y','"+str(_body_top_y(title))+"');$pos.SetAttribute('z','0')\n"
                "  $outline.AppendChild($pos)|Out-Null\n"
                "  $sz=$doc.CreateElement('one:Size',$ns)\n"
                "  $sz.SetAttribute('width','720');$sz.SetAttribute('height','100')\n"
                "  $outline.AppendChild($sz)|Out-Null\n"
                "  $doc.DocumentElement.AppendChild($outline)|Out-Null\n"
                "}else{\n"
                "  $op=$outline.SelectSingleNode('one:Position',$mgr)\n"
                "  if($op){$op.SetAttribute('y','"+str(_body_top_y(title))+"')}\n"
                "}\n"
                "$oec=$outline.SelectSingleNode('one:OEChildren',$mgr)\n"
                "if(-not $oec){\n"
                "  $oec=$doc.CreateElement('one:OEChildren',$ns)\n"
                "  $outline.AppendChild($oec)|Out-Null\n"
                "}\n"
                + QUICKSTYLE_PS +
                # 본문 XML을 별도 문서로 읽은 뒤 페이지 문서로 옮겨 붙임 (네임스페이스 안전)
                "$frag=New-Object System.Xml.XmlDocument\n"
                "$frag.LoadXml('<root xmlns:one=\"'+$ns+'\">'+$bodyXml+'</root>')\n"
                "foreach($n in @($frag.DocumentElement.ChildNodes)){\n"
                "  $imp=$doc.ImportNode($n,$true)\n"
                "  $oec.AppendChild($imp)|Out-Null\n"
                "}\n"
                "$on.UpdatePageContent($doc.OuterXml,[DateTime]::MinValue)\n"
            )
            Path(ps1).write_text(ps, encoding='utf-8-sig')
            self.last_page_id = ''
            self._run_ps_file(ps1, timeout=300)
            if os.path.exists(pf):
                self.last_page_id = Path(pf).read_text(encoding='utf-8-sig').strip()
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def list_pages(self, section_id):
        """섹션의 페이지 목록 + 각 페이지의 사건번호·로톡 링크를 OneNote에서 직접 읽는다.
        매크로를 껐다 켜도, 예전에 만든 페이지여도 언제든 쓸 수 있게 하기 위함.
        반환: [{'id','title','time','case_no','link'}] (OneNote 페이지 목록과 같은 순서)"""
        if not section_id:
            section_id = self.section_id or self.get_first_section_id()
        tmp = tempfile.mkdtemp(prefix='on_pages_')
        try:
            sf = os.path.join(tmp, 'sid.txt'); of = os.path.join(tmp, 'out.json'); ps1 = os.path.join(tmp, 'l.ps1')
            Path(sf).write_text(section_id, encoding='utf-8')
            Path(ps1).write_text(PAGES_PS.replace('__SID__', self._q(sf)).replace('__OUT__', self._q(of)),
                                 encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_DEFAULT_TIMEOUT)
            import json
            raw = Path(of).read_text(encoding='utf-8-sig').strip()
            data = json.loads(raw) if raw else []
            if isinstance(data, dict):
                data = [data]
            return parse_page_rows(data)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def add_link(self, page_id, url, label='로톡 링크', strip_dash=False, new_title=None):
        """페이지 본문 맨 위에 '로톡 링크: 주소' 한 줄을 넣는다 (누르면 열리는 링크).
        strip_dash=True 면 제목 맨 앞의 'ㅡ'(업로드 전 표시)도 지운다.
        new_title 을 주면 제목을 그 글자로 바꾼다 (날짜 없는 글에 '사용 날짜 로톡' 붙일 때)."""
        if not page_id or not url:
            return False
        tmp = tempfile.mkdtemp(prefix='on_link_')
        try:
            pgf = os.path.join(tmp, 'pg.txt'); lf = os.path.join(tmp, 'l.txt'); ps1 = os.path.join(tmp, 'k.ps1')
            tf = os.path.join(tmp, 't.txt')
            Path(pgf).write_text(page_id, encoding='utf-8')
            Path(tf).write_text(_esc(new_title or ''), encoding='utf-8')
            u = _esc(url).replace('"', '&quot;')
            Path(lf).write_text(f"<span style='font-size:{SIZE_BODY:.1f}pt;font-weight:bold'>{_esc(label)}: </span>"
                                f"<a href=\"{u}\"><span style='font-size:{SIZE_BODY:.1f}pt'>{_esc(url)}</span></a>",
                                encoding='utf-8')
            Path(ps1).write_text(
                "$ErrorActionPreference='Stop'\n"
                "$pageId=[System.IO.File]::ReadAllText("+self._q(pgf)+",[System.Text.Encoding]::UTF8).Trim()\n"
                "$line=[System.IO.File]::ReadAllText("+self._q(lf)+",[System.Text.Encoding]::UTF8)\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$xml=''\n"
                "$on.GetPageContent($pageId,[ref]$xml,0)\n"
                "$doc=[xml]$xml\n"
                "$ns=$doc.DocumentElement.NamespaceURI\n"
                "$mgr=New-Object System.Xml.XmlNamespaceManager($doc.NameTable)\n"
                "$mgr.AddNamespace('one',$ns)\n"
                "$oec=$doc.SelectSingleNode('//one:Outline/one:OEChildren',$mgr)\n"
                "if(-not $oec){ throw 'no outline' }\n"
                "$oe=$doc.CreateElement('one:OE',$ns)\n"
                "$t=$doc.CreateElement('one:T',$ns)\n"
                "$t.AppendChild($doc.CreateCDataSection($line))|Out-Null\n"
                "$oe.AppendChild($t)|Out-Null\n"
                "$oec.PrependChild($oe)|Out-Null\n"
                + (("$nt=[System.IO.File]::ReadAllText(" + self._q(tf) + ",[System.Text.Encoding]::UTF8)\n"
                    "$tt=$doc.SelectSingleNode('//one:Title/one:OE/one:T',$mgr)\n"
                    "if($tt -and $nt){ $tt.InnerXml=''; $tt.AppendChild($doc.CreateCDataSection($nt))|Out-Null }\n")
                   if new_title else (TITLE_DASH_PS if strip_dash else '')) +
                "$on.UpdatePageContent($doc.OuterXml,[DateTime]::MinValue)\n", encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_PAGE_TIMEOUT)
            return True
        except Exception as e:
            print(f"[OneNote] 링크 넣기 실패: {e}")
            return False
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def delete_page(self, page_id):
        """페이지 삭제 (OneNote '삭제된 페이지'로 이동 — 필요하면 복구 가능)"""
        if not page_id:
            return False
        tmp = tempfile.mkdtemp(prefix='on_del_')
        try:
            pgf = os.path.join(tmp, 'pg.txt')
            ps1 = os.path.join(tmp, 'd.ps1')
            Path(pgf).write_text(page_id, encoding='utf-8')
            Path(ps1).write_text(
                "$ErrorActionPreference='Stop'\n"
                "$pageId=[System.IO.File]::ReadAllText("+self._q(pgf)+",[System.Text.Encoding]::UTF8).Trim()\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "$on.DeleteHierarchy($pageId,[DateTime]::MinValue,$false)\n", encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_PAGE_TIMEOUT)
            return True
        except Exception as e:
            print(f"[OneNote] 페이지 삭제 실패: {e}")
            return False
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def _create_empty_page(self):
        tmp = tempfile.mkdtemp(prefix='on_new_')
        try:
            sf   = os.path.join(tmp,'sid.txt')
            pgf  = os.path.join(tmp,'pgid.txt')
            ps1  = os.path.join(tmp,'c.ps1')
            Path(sf).write_text(self.section_id, encoding='utf-8')
            Path(ps1).write_text(
                "$ErrorActionPreference='Stop'\n"
                "$sectionId=[System.IO.File]::ReadAllText("+self._q(sf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$newPageId=''\n"
                "$on.CreateNewPage($sectionId,[ref]$newPageId,0)\n"
                "[System.IO.File]::WriteAllText("+self._q(pgf)+","
                "$newPageId,[System.Text.Encoding]::UTF8)\n", encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_PAGE_TIMEOUT)
            return Path(pgf).read_text(encoding='utf-8').strip()
        finally: shutil.rmtree(tmp, ignore_errors=True)

    def _set_page_title(self, page_id, title):
        tmp = tempfile.mkdtemp(prefix='on_title_')
        try:
            pgf  = os.path.join(tmp,'pg.txt')
            tf   = os.path.join(tmp,'t.txt')
            ps1  = os.path.join(tmp,'t.ps1')
            Path(pgf).write_text(page_id, encoding='utf-8')
            Path(tf).write_text(title, encoding='utf-8')
            Path(ps1).write_text(
                "$ErrorActionPreference='Stop'\n"
                "$pageId=[System.IO.File]::ReadAllText("+self._q(pgf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$newTitle=[System.IO.File]::ReadAllText("+self._q(tf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$xml=''\n"
                "$on.GetPageContent($pageId,[ref]$xml,0)\n"
                "$doc=[xml]$xml\n"
                "$ns=$doc.DocumentElement.NamespaceURI\n"
                "$mgr=New-Object System.Xml.XmlNamespaceManager($doc.NameTable)\n"
                "$mgr.AddNamespace('one',$ns)\n"
                "$tn=$doc.SelectSingleNode('//one:Title/one:OE/one:T',$mgr)\n"
                "if($tn){$tn.InnerText=$newTitle;$on.UpdatePageContent($doc.OuterXml,[DateTime]::MinValue)}\n",
                encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=_PAGE_TIMEOUT)
        finally: shutil.rmtree(tmp, ignore_errors=True)

    def _add_page_content(self, page_id, article):
        tmp = tempfile.mkdtemp(prefix='on_body_')
        try:
            pgf = os.path.join(tmp,'pg.txt')
            ps1 = os.path.join(tmp,'b.ps1')
            Path(pgf).write_text(page_id, encoding='utf-8')

            lines = []
            def add(text, h=False): lines.append((text, h))

            from parser import strip_markdown, to_onenote_richtext
            if article.get('focus_keyword'):
                add(to_onenote_richtext("Focus Keyword: " + article['focus_keyword']))
            if article.get('sub_keywords'):
                add(to_onenote_richtext("보조 키워드: " + article['sub_keywords']))
            if article.get('focus_keyword') or article.get('sub_keywords'):
                add('')

            has_toc  = bool(article.get('toc'))
            has_sum  = bool(article.get('summary','').strip())
            has_sec  = bool(article.get('sections'))

            if has_toc:
                add('목차', h=True)
                for item in article['toc']:
                    rich = to_onenote_richtext(item)
                    add(f"  {rich}" if re.match(r'^\d+[\.\)]', item) else f"  • {rich}")
                add('')
            if has_sum:
                add('요약문', h=True)
                for p in re.split(r'\n\s*\n', article['summary'].strip()):
                    p = p.strip()
                    if p:
                        add(to_onenote_richtext(p))
                add('')
            if has_sec:
                for st, sb in article['sections']:
                    parts = sb.strip().split('\n\n')
                    if parts and re.match(r'^\d+$', parts[0].strip()):
                        add(parts[0].strip()); add('')
                        parts = parts[1:]
                    add(to_onenote_richtext(strip_markdown(st)), h=True); add('')
                    for p in parts:
                        p = p.strip()
                        if not p:
                            continue
                        if re.match(r'^\d+$', p):
                            add(p); add('')
                        else:
                            add(to_onenote_richtext(p)); add('')
            if not has_toc and not has_sum and not has_sec:
                add('본문', h=True); add('')
                for ln in re.split(r'\n\s*\n', article.get('raw_content','').strip()):
                    ln = ln.strip()
                    if ln:
                        add(to_onenote_richtext(ln)); add('')

            # 줄 파일 저장
            meta = []
            for i,(t,h) in enumerate(lines):
                lf = os.path.join(tmp,f'l{i}.txt')
                Path(lf).write_text(t, encoding='utf-8')
                meta.append((lf, h))

            ps = (
                "$ErrorActionPreference='Stop'\n"
                "$pageId=[System.IO.File]::ReadAllText("+self._q(pgf)+","
                "[System.Text.Encoding]::UTF8).Trim()\n"
                "$on=New-Object -ComObject OneNote.Application\n"
                "[string]$xml=''\n"
                "$on.GetPageContent($pageId,[ref]$xml,0)\n"
                "$doc=[xml]$xml\n"
                "$ns=$doc.DocumentElement.NamespaceURI\n"
                "$mgr=New-Object System.Xml.XmlNamespaceManager($doc.NameTable)\n"
                "$mgr.AddNamespace('one',$ns)\n"
                "$outline=$doc.SelectSingleNode('//one:Outline',$mgr)\n"
                "if(-not $outline){\n"
                "  $outline=$doc.CreateElement('one:Outline',$ns)\n"
                "  $pos=$doc.CreateElement('one:Position',$ns)\n"
                "  $pos.SetAttribute('x','36');$pos.SetAttribute('y','86.4');$pos.SetAttribute('z','0')\n"
                "  $outline.AppendChild($pos)|Out-Null\n"
                "  $sz=$doc.CreateElement('one:Size',$ns)\n"
                "  $sz.SetAttribute('width','720');$sz.SetAttribute('height','100')\n"
                "  $outline.AppendChild($sz)|Out-Null\n"
                "  $doc.DocumentElement.AppendChild($outline)|Out-Null\n"
                "}\n"
                "$oec=$outline.SelectSingleNode('one:OEChildren',$mgr)\n"
                "if(-not $oec){\n"
                "  $oec=$doc.CreateElement('one:OEChildren',$ns)\n"
                "  $outline.AppendChild($oec)|Out-Null\n"
                "}\n"
            )
            for lf, is_h in meta:
                ps += (
                    "$lineText=[System.IO.File]::ReadAllText("+self._q(lf)+","
                    "[System.Text.Encoding]::UTF8)\n"
                    "$oe=$doc.CreateElement('one:OE',$ns)\n"
                    "$tNode=$doc.CreateElement('one:T',$ns)\n"
                    "$tNode.InnerXml=('<![CDATA[' + $lineText + ']]>')\n"
                )
                if is_h:
                    ps += "$oe.SetAttribute('style','font-weight:bold')\n"
                ps += "$oe.AppendChild($tNode)|Out-Null\n$oec.AppendChild($oe)|Out-Null\n"
            ps += "$on.UpdatePageContent($doc.OuterXml,[DateTime]::MinValue)\n"

            Path(ps1).write_text(ps, encoding='utf-8-sig')
            self._run_ps_file(ps1, timeout=300)
        finally: shutil.rmtree(tmp, ignore_errors=True)
