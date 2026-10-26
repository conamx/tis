# onenote_com.py
import os, re, shutil, subprocess, tempfile, xml.etree.ElementTree as ET
from pathlib import Path

_DEFAULT_TIMEOUT = 300
_PAGE_TIMEOUT    = 300
_MAX_RETRY       = 2


class SectionInfo:
    def __init__(self, sid, name, notebook):
        self.id = sid; self.name = name; self.notebook = notebook
    def display_name(self):
        return f"{self.notebook} > {self.name}"


class OneNoteManager:
    def __init__(self):
        self.section_id = ''; self.connected = False; self._sections = []

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
