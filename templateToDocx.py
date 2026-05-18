# -*- coding: utf-8 -*-
import io
import json
import os
import re
import argparse
import html as _html
from copy import deepcopy
from pathlib import Path
from html.parser import HTMLParser
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

_HTML_RE = re.compile(
    r'<(?:br|/?b|/?u|/?i|/?p|/?span|/?strong|/?em|/?div|/?ul|/?ol|/?li|/?table|/?tr|/?td|/?th|/?tbody|img)\b[^>]*/?>', re.I)
_PLACEHOLDER_RE = re.compile(r'{{\s*([^}]+?)\s*}}')

# Colour name → OOXML highlight value (closest match)
_HIGHLIGHT_MAP = {
    '#ffff00': 'yellow', '#ff0000': 'red', '#00ff00': 'green',
    '#0000ff': 'blue',   '#00ffff': 'cyan', '#ff00ff': 'magenta',
    '#000000': 'black',  '#ffffff': 'white',
    '#808080': 'darkGray', '#c0c0c0': 'lightGray',
    '#800000': 'darkRed', '#808000': 'darkYellow',
    '#008000': 'darkGreen', '#008080': 'darkCyan',
    '#000080': 'darkBlue', '#800080': 'darkMagenta',
}

def _closest_highlight(hex_color):
    """Map an arbitrary hex colour to the nearest OOXML highlight value,
    or return None to use shading instead."""
    if not hex_color:
        return None
    hc = hex_color.lower().strip()
    if hc in _HIGHLIGHT_MAP:
        return _HIGHLIGHT_MAP[hc]
    # Not a standard highlight colour – caller should use shading
    return None

def _int_to_roman(n, lower=True):
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    res = ""
    i = 0
    while n > 0:
        for _ in range(n // val[i]):
            res += syb[i]
            n -= val[i]
        i += 1
    return res if lower else res.upper()

def _int_to_alpha(n, lower=True):
    res = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        res = chr((ord('a') if lower else ord('A')) + r) + res
    return res

def _format_list_counter(n, style_type):
    st = (style_type or '').lower().strip()
    if st in ('lower-roman', 'roman'): return _int_to_roman(n, lower=True)
    if st == 'upper-roman': return _int_to_roman(n, lower=False)
    if st in ('lower-alpha', 'lower-latin'): return _int_to_alpha(n, lower=True)
    if st in ('upper-alpha', 'upper-latin'): return _int_to_alpha(n, lower=False)
    if st == 'lower-greek':
        alphabet = 'αβγδεζηθικλμνξοπρστυφχψω'
        return alphabet[n-1] if 1 <= n <= len(alphabet) else str(n)
    return str(n)

def _is_html(text):
    return bool(_HTML_RE.search(text))

def _extract_placeholder_keys(text):
    return [m.group(1).strip() for m in _PLACEHOLDER_RE.finditer(text or '')]

def _has_static_text_outside_placeholders(text):
    """True when paragraph includes literal text besides placeholders/separators."""
    if not text:
        return False
    stripped = _PLACEHOLDER_RE.sub('', text)
    # Keep only meaningful chars; ignore common separators around placeholders.
    stripped = re.sub(r'[\s,:;|(){}\[\]\-–—]+', '', stripped)
    return bool(stripped)

def _trim_to_placeholder_line(text):
    """For 'header\\n{{item...}}' templates, keep only the item line."""
    if not text:
        return text
    m = _PLACEHOLDER_RE.search(text)
    if not m:
        return text
    first_ph = m.start()
    nl = text.rfind('\n', 0, first_ph)
    if nl >= 0:
        return text[nl + 1:]
    return text[first_ph:]

def _render_template_text_for_item(tpl_text, item):
    repl = {}
    for k, v in (item or {}).items():
        sv = '' if v is None else str(v)
        repl[k] = sv
        repl[k.lower()] = sv
    return _replace_placeholders_in_text(tpl_text, repl)

def _render_static_prefix_with_inline_items(tpl_text, items):
    """Render 'prefix {{a}} ... {{b}}' as one line with comma-separated items."""
    m = _PLACEHOLDER_RE.search(tpl_text or '')
    if not m:
        return tpl_text or ''
    prefix = (tpl_text or '')[:m.start()]
    item_tpl = (tpl_text or '')[m.start():]
    parts = []
    for item in items:
        t = _render_template_text_for_item(item_tpl, item).strip()
        if t:
            parts.append(t)
    return prefix + ', '.join(parts)

def _render_inline_collection_in_paragraph(paragraph, items):
    """Render static+placeholders paragraph inline, preserving run styles.

    Uses the template run styles for each literal/placeholder segment, and
    repeats the placeholder-bearing tail per item with ', ' between items.
    """
    runs = list(paragraph.runs)
    full_text = ''.join((r.text or '') for r in runs)
    matches = list(_PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return

    spans = _run_spans_from_runs(runs)
    first_ph_start = matches[0].start()

    def add_literal_segments(start, end, out):
        if end > start:
            out.extend(_slice_text_by_run_style(full_text, spans, start, end))

    # Prefix before first placeholder stays once.
    segments = []
    add_literal_segments(0, first_ph_start, segments)

    # Build template tokens from first placeholder to end.
    tail_tokens = []
    cursor = first_ph_start
    for m in _PLACEHOLDER_RE.finditer(full_text, pos=first_ph_start):
        s, e = m.span()
        key = m.group(1).strip()
        if s > cursor:
            tail_tokens.extend(_slice_text_by_run_style(full_text, spans, cursor, s))
        tail_tokens.append((('placeholder', key), _rpr_for_index(spans, s)))
        cursor = e
    if cursor < len(full_text):
        tail_tokens.extend(_slice_text_by_run_style(full_text, spans, cursor, len(full_text)))

    # Expand tail once per item, preserving placeholder-specific style.
    for idx, item in enumerate(items):
        item = item if isinstance(item, dict) else {}
        for token, rpr in tail_tokens:
            if isinstance(token, tuple) and token and token[0] == 'placeholder':
                key = token[1]
                val = item.get(key, item.get(key.lower(), ''))
                segments.append(('' if val is None else str(val), rpr))
            else:
                segments.append((token, rpr))
        if idx < len(items) - 1:
            sep_rpr = tail_tokens[-1][1] if tail_tokens else _rpr_for_index(spans, first_ph_start)
            segments.append((', ', sep_rpr))

    # Remove existing text nodes and insert reconstructed styled segments.
    had_text_runs = []
    for r in runs:
        r_elem = r._r
        removed = False
        for child in list(r_elem):
            if child.tag == qn('w:t') or child.tag == qn('w:br'):
                r_elem.remove(child)
                removed = True
        if removed:
            had_text_runs.append(r_elem)
    if not had_text_runs:
        paragraph.text = ''.join(s for s, _ in segments)
        return
    first = had_text_runs[0]
    _insert_plain_segments(first, segments)

def _render_header_plus_inline_items(tpl_text, items):
    """Render 'header\\n{{item...}}' as 'header item1, item2, ...'."""
    first_ph = _PLACEHOLDER_RE.search(tpl_text or '')
    if not first_ph:
        return tpl_text or ''
    first_ph_idx = first_ph.start()
    nl = (tpl_text or '').rfind('\n', 0, first_ph_idx)
    if nl < 0:
        line_tpl = tpl_text or ''
        header = ''
    else:
        header = (tpl_text or '')[:nl].strip()
        line_tpl = (tpl_text or '')[nl + 1:]
    rendered = []
    for item in items:
        t = _render_template_text_for_item(line_tpl, item).strip()
        if t:
            rendered.append(t)
    joined = ', '.join(rendered)
    if header and joined:
        return f"{header} {joined}"
    return header or joined

def _parse_style(style_str):
    """Extract relevant CSS properties from an inline style string."""
    props = {}
    if not style_str:
        return props
    for part in style_str.split(';'):
        if ':' not in part:
            continue
        k, v = part.split(':', 1)
        props[k.strip().lower()] = v.strip()
    return props

def _html_segments(markup):
    """Parse HTML into list of segment dicts with rich formatting info."""
    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.segs = []
            # Formatting state stacks
            self.bold = False
            self.underline = False
            self.italic = False
            self.highlight = None   # OOXML highlight name
            self.shading = None     # raw hex (e.g. 'E67E23') for w:shd
            self.font = None        # font family override
            self.indent = 0         # nesting depth for list items
            self._stack = []        # (attr_name, old_value) restore stack
            # List counters: stack of (list_type, counter)
            self._list_stack = []
            # Track whether we're inside a table
            self._in_table = 0
            self._td_count = 0
            # Span boundary markers — each span open pushes a marker
            self._span_depth = 0

        def _fmt(self):
            return dict(bold=self.bold, underline=self.underline,
                        italic=self.italic, highlight=self.highlight,
                        shading=self.shading, font=self.font,
                        indent=self.indent)

        def _push(self, attr, new_val):
            self._stack.append((attr, getattr(self, attr)))
            setattr(self, attr, new_val)

        def _pop_one(self, tags):
            """Pop most-recent stack entry whose attr name is in *tags*."""
            for i in range(len(self._stack) - 1, -1, -1):
                if self._stack[i][0] in tags:
                    attr, old_val = self._stack.pop(i)
                    setattr(self, attr, old_val)
                    return

        def _pop_to_marker(self, marker):
            """Pop all stack entries down to (and including) *marker*."""
            while self._stack:
                attr, old_val = self._stack.pop()
                if attr == marker:
                    break
                setattr(self, attr, old_val)

        def _ensure_newline(self):
            if self.segs and self.segs[-1].get('text') != '\n':
                self.segs.append(dict(text='\n', **self._fmt()))

        def handle_starttag(self, tag, attrs):
            t = tag.lower()
            attr_d = dict(attrs) if attrs else {}
            style = _parse_style(attr_d.get('style', ''))
            tag_map = {'b': 'bold', 'strong': 'bold', 'u': 'underline', 'i': 'italic', 'em': 'italic'}
            if t in tag_map:
                self._push(tag_map[t], True)
            elif t == 'span':
                self._span_depth += 1
                self._stack.append((f'__span_{self._span_depth}', None))
                bg, td, ff = style.get('background-color'), style.get('text-decoration'), style.get('font-family')
                if bg: self._push('shading', bg.lstrip('#').upper() if not _closest_highlight(bg) else None)
                if bg and _closest_highlight(bg): self._push('highlight', _closest_highlight(bg))
                if td and 'underline' in td: self._push('underline', True)
                if ff: self._push('font', [f.strip().strip("'\"") for f in ff.split(',')][0])
            elif t == 'br':
                self.segs.append(dict(text='\n', **self._fmt()))
            elif t in ('p', 'div'):
                self._ensure_newline()
            elif t in ('ul', 'ol'):
                self._list_stack.append([t, 0, style.get('list-style-type')])
                self.indent = len(self._list_stack)

            elif t == 'li':
                self._ensure_newline()
                prefix = ''
                if self._list_stack:
                    lt, counter, ls = self._list_stack[-1]
                    if lt == 'ul':
                        bullets = ['•', '◦', '▪', '▸']
                        bullet = bullets[min(len(self._list_stack) - 1, len(bullets) - 1)]
                        prefix = '  ' * (self.indent - 1) + bullet + ' '
                    else:
                        counter += 1
                        self._list_stack[-1][1] = counter
                        formatted = _format_list_counter(counter, ls)
                        prefix = '  ' * (self.indent - 1) + f'{formatted}. '
                if prefix:
                    self.segs.append(dict(text=prefix, **self._fmt()))

            elif t == 'table':
                self._in_table += 1
                self._ensure_newline()

            elif t == 'tr':
                self._td_count = 0

            elif t in ('td', 'th'):
                if self._td_count > 0:
                    self.segs.append(dict(text=' | ', **self._fmt()))
                self._td_count += 1
                if t == 'th':
                    self._push('bold', True)

            elif t == 'img':
                alt = attr_d.get('alt', '')
                self.segs.append(dict(text=f'[image: {alt}]' if alt else '[image]', **self._fmt()))

        def handle_endtag(self, tag):
            t = tag.lower()
            if t in ('b', 'strong'):
                self._pop_one({'bold'})
            elif t == 'u':
                self._pop_one({'underline'})
            elif t in ('i', 'em'):
                self._pop_one({'italic'})
            elif t == 'span':
                # Pop all attributes this span pushed, back to its marker
                if self._span_depth > 0:
                    marker = f'__span_{self._span_depth}'
                    self._pop_to_marker(marker)
                    self._span_depth -= 1
            elif t in ('p', 'div'):
                self._ensure_newline()
            elif t in ('ul', 'ol'):
                if self._list_stack:
                    self._list_stack.pop()
                self.indent = len(self._list_stack)
                self._ensure_newline()
            elif t == 'li':
                self._ensure_newline()
            elif t == 'tr':
                self._ensure_newline()
            elif t == 'table':
                self._in_table = max(0, self._in_table - 1)
                self._ensure_newline()
            elif t in ('td', 'th'):
                if t == 'th':
                    self._pop_one({'bold'})

        def handle_data(self, data):
            if data:
                self.segs.append(dict(text=data, **self._fmt()))

        def handle_entityref(self, name):
            ch = _html.unescape(f'&{name};')
            self.segs.append(dict(text=ch, **self._fmt()))

        def handle_charref(self, name):
            ch = _html.unescape(f'&#{name};')
            self.segs.append(dict(text=ch, **self._fmt()))

    p = P()
    p.feed(markup)
    # # Trim trailing newlines to avoid redundant breaks at the end of content
    # while p.segs and p.segs[-1].get('text') == '\n':
    #     p.segs.pop()
    return p.segs

def _get_num_id(doc, num_fmt='bullet'):
    """Find existing numId in the document's numbering for a given format (e.g., 'bullet', 'decimal')."""
    try:
        numbering = doc.part.numbering_part.element
        for abstract in numbering.xpath('//w:abstractNum'):
            lvls = abstract.xpath('w:lvl')
            if lvls and lvls[0].xpath(f'w:numFmt[@w:val="{num_fmt}"]'):
                abs_id = abstract.get(qn('w:abstractNumId'))
                for num in numbering.xpath(f'//w:num[w:abstractNumId[@w:val="{abs_id}"]]'):
                    return num.get(qn('w:numId'))
    except Exception:
        pass
    return '3' if num_fmt == 'bullet' else '1'

def _apply_run_formatting(rpr, seg):
    """Applies Bold, Italic, Underline, Highlight, Shading, and Font from a segment to a w:rPr element."""
    if seg.get('bold'):
        if rpr.find(qn('w:b')) is None:
            etree.SubElement(rpr, qn('w:b'))
            etree.SubElement(rpr, qn('w:bCs'))
    if seg.get('underline'):
        if rpr.find(qn('w:u')) is None:
            etree.SubElement(rpr, qn('w:u')).set(qn('w:val'), 'single')
    if seg.get('italic'):
        if rpr.find(qn('w:i')) is None:
            etree.SubElement(rpr, qn('w:i'))
            etree.SubElement(rpr, qn('w:iCs'))
    if seg.get('highlight'):
        if rpr.find(qn('w:highlight')) is None:
            etree.SubElement(rpr, qn('w:highlight')).set(qn('w:val'), seg['highlight'])
    if seg.get('shading'):
        if rpr.find(qn('w:shd')) is None:
            shd = etree.SubElement(rpr, qn('w:shd'))
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), seg['shading'])
    if seg.get('font'):
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = etree.SubElement(rpr, qn('w:rFonts'))
        fname = seg['font']
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
            rfonts.set(qn(attr), fname)

def _insert_html_runs(first_run, html_text, doc=None):
    """Insert HTML-formatted content. Uses proper Word lists for <ul>/<ol>."""
    if '<ul' not in html_text.lower() and '<ol' not in html_text.lower():
        return _insert_html_runs_simple(first_run, html_text)
    return _insert_html_with_proper_lists(first_run, html_text, doc=doc)

def _insert_html_runs_simple(first_run, html_text):
    """Simple insertion for non-list HTML content."""
    base_rpr = first_run.find(qn('w:rPr'))
    prev = first_run
    for seg in _html_segments(html_text):
        text = seg['text']
        run = etree.Element(qn('w:r'))
        if text == '\n':
            etree.SubElement(run, qn('w:br'))
        else:
            rpr = deepcopy(base_rpr) if base_rpr is not None else etree.Element(qn('w:rPr'))
            _apply_run_formatting(rpr, seg)
            if len(rpr):
                run.insert(0, rpr)
            t_el = etree.SubElement(run, qn('w:t'))
            t_el.set(XML_SPACE, 'preserve')
            t_el.text = text
        prev.addnext(run)
        prev = run

def _make_list_pPr(orig_pPr, level, is_numbered, doc=None):
    """Build a w:pPr element with proper w:numPr for list paragraphs."""
    if orig_pPr is not None:
        new_pPr = deepcopy(orig_pPr)
    else:
        new_pPr = etree.Element(qn('w:pPr'))

    # Remove any old numPr
    old = new_pPr.find(qn('w:numPr'))
    if old is not None:
        new_pPr.remove(old)

    # Add w:numPr with ilvl and numId
    numPr = etree.SubElement(new_pPr, qn('w:numPr'))
    ilvl = etree.SubElement(numPr, qn('w:ilvl'))
    ilvl.set(qn('w:val'), str(max(0, level - 1)))
    numId_el = etree.SubElement(numPr, qn('w:numId'))

    # Determine numId dynamically from doc or use defaults
    numId_el.set(qn('w:val'), _get_num_id(doc, 'decimal' if is_numbered else 'bullet'))

    # Ensure pStyle is ListParagraph for proper continuation
    pStyle = new_pPr.find(qn('w:pStyle'))
    if pStyle is None:
        pStyle = etree.SubElement(new_pPr, qn('w:pStyle'))
    pStyle.set(qn('w:val'), 'ListParagraph')

    return new_pPr

def _add_formatted_run(p_elem, seg, base_rpr=None):
    """Add a formatted run to a paragraph element."""
    run = etree.SubElement(p_elem, qn('w:r'))
    rpr = deepcopy(base_rpr) if base_rpr is not None else etree.Element(qn('w:rPr'))
    _apply_run_formatting(rpr, seg)
    if len(rpr):
        run.insert(0, rpr)
    return run

def _insert_html_with_proper_lists(first_run, html_text, doc=None):
    """Create real Word list paragraphs for <ul>/<ol> so Enter continues the list."""
    p_elem = first_run.getparent()  # w:r parent is w:p directly
    if p_elem.tag != qn('w:p'):
        # Fallback: try grandparent
        p_elem = first_run.getparent()
        while p_elem is not None and p_elem.tag != qn('w:p'):
            p_elem = p_elem.getparent()
        if p_elem is None:
            return _insert_html_runs_simple(first_run, html_text)

    parent = p_elem.getparent()  # w:body or w:tc
    if parent is None:
        return _insert_html_runs_simple(first_run, html_text)

    # Save original pPr and base rPr for copying
    orig_pPr = p_elem.find(qn('w:pPr'))
    base_rpr = first_run.find(qn('w:rPr'))

    # Clear all runs from the original paragraph
    for child in list(p_elem):
        if child.tag == qn('w:r'):
            p_elem.remove(child)

    # Parse HTML into segments
    segs = _html_segments(html_text)

    # Group segments into logical paragraphs:
    # - newline segments cause paragraph breaks
    # - list item segments (with indent > 0 and bullet prefix) become list paragraphs
    para_groups = []  # list of (is_list, list_type, level, segments)
    current_segs = []
    current_is_list = False
    current_list_type = None
    current_level = 0

    for seg in segs:
        text = seg['text']
        indent = seg.get('indent', 0)

        if text == '\n':
            # Flush current group
            if current_segs:
                para_groups.append((current_is_list, current_list_type, current_level, current_segs))
                current_segs = []
                current_is_list = False
                current_list_type = None
                current_level = 0
            continue

        # Detect bullet/number prefix from _html_segments
        is_bullet = any(text.startswith(b + ' ') for b in ['•', '◦', '▪', '▸'])
        # Check for indented bullet prefix like "  • "
        stripped = text.lstrip()
        if not is_bullet:
            is_bullet = any(stripped.startswith(b + ' ') for b in ['•', '◦', '▪', '▸'])

        is_numbered = False
        if not is_bullet and indent > 0:
            # Check for numbered prefix like "1. " or "a. " or "i. " etc
            import re as _re
            is_numbered = bool(_re.match(r'^\s*(?:\w+[\.\)])\s', text))

        if is_bullet and indent > 0:
            # Start of a bullet list item
            if current_segs:
                para_groups.append((current_is_list, current_list_type, current_level, current_segs))
            # Strip the bullet prefix
            clean = stripped
            for b in ['•', '◦', '▪', '▸']:
                if clean.startswith(b + ' '):
                    clean = clean[len(b)+1:]
                    break
            current_segs = [dict(seg, text=clean)] if clean.strip() else []
            current_is_list = True
            current_list_type = 'ul'
            current_level = indent
        elif is_numbered and indent > 0:
            if current_segs:
                para_groups.append((current_is_list, current_list_type, current_level, current_segs))
            # Strip the number prefix
            clean = _re.sub(r'^\s*\w+[\.\)]\s*', '', text)
            current_segs = [dict(seg, text=clean)] if clean.strip() else []
            current_is_list = True
            current_list_type = 'ol'
            current_level = indent
        else:
            # Continuation text or normal text
            if current_is_list and indent > 0:
                # Continue adding to current list item
                current_segs.append(seg)
            else:
                if text.strip():
                    current_segs.append(seg)

    # Flush last group
    if current_segs:
        para_groups.append((current_is_list, current_list_type, current_level, current_segs))

    # Now create Word paragraphs
    insert_after = p_elem
    first_group = True
    for is_list, list_type, level, group_segs in para_groups:
        if not group_segs:
            continue

        if first_group:
            # Use the original paragraph for the first group
            target_p = p_elem
            first_group = False
        else:
            # Create a new paragraph
            target_p = etree.Element(qn('w:p'))
            insert_after.addnext(target_p)

        if is_list:
            # Set list paragraph properties
            pPr = _make_list_pPr(orig_pPr, level, is_numbered=(list_type == 'ol'), doc=doc)
            # Remove any existing pPr
            old_pPr = target_p.find(qn('w:pPr'))
            if old_pPr is not None:
                target_p.remove(old_pPr)
            target_p.insert(0, pPr)
        else:
            # Normal paragraph - copy original pPr if new paragraph
            if target_p is not p_elem and orig_pPr is not None:
                target_p.insert(0, deepcopy(orig_pPr))

        # Add runs with text
        for seg in group_segs:
            text = seg['text']
            if not text or not text.strip():
                continue
            run = _add_formatted_run(target_p, seg, base_rpr)
            t_el = etree.SubElement(run, qn('w:t'))
            t_el.set(XML_SPACE, 'preserve')
            t_el.text = text

        insert_after = target_p

def _replace_placeholders_in_text(text, replacements):
    def repl(match):
        key = match.group(1).strip()
        return replacements.get(key, replacements.get(key.lower(), match.group(0)))
    return _PLACEHOLDER_RE.sub(repl, text)

def _copy_rpr(run_elem):
    rpr = run_elem.find(qn('w:rPr'))
    return deepcopy(rpr) if rpr is not None else None

def _run_spans_from_runs(runs):
    spans = []
    pos = 0
    for r in runs:
        txt = r.text or ''
        ln = len(txt)
        if ln:
            spans.append((pos, pos + ln, r._r))
            pos += ln
    return spans

def _rpr_for_index(spans, idx):
    for start, end, r_elem in spans:
        if start <= idx < end:
            return _copy_rpr(r_elem)
    if spans:
        return _copy_rpr(spans[-1][2])
    return None

def _slice_text_by_run_style(text, spans, start, end):
    out = []
    i = start
    while i < end:
        found = None
        for s, e, r_elem in spans:
            if s <= i < e:
                found = (s, e, r_elem)
                break
        if found is None:
            out.append((text[i:end], None))
            break
        s, e, r_elem = found
        j = min(end, e)
        out.append((text[i:j], _copy_rpr(r_elem)))
        i = j
    return out

def _build_plain_segments_with_styles(full_text, runs, replacements):
    spans = _run_spans_from_runs(runs)
    segments = []
    cursor = 0
    for m in _PLACEHOLDER_RE.finditer(full_text):
        start, end = m.span()
        key = m.group(1).strip()
        rep = replacements.get(key, replacements.get(key.lower(), m.group(0)))
        if start > cursor:
            segments.extend(_slice_text_by_run_style(full_text, spans, cursor, start))
        segments.append((rep, _rpr_for_index(spans, start)))
        cursor = end
    if cursor < len(full_text):
        segments.extend(_slice_text_by_run_style(full_text, spans, cursor, len(full_text)))
    return segments

def _insert_plain_segments(first_run, segments):
    prev = first_run
    for text, rpr in segments:
        if text is None:
            text = ''
        run = etree.Element(qn('w:r'))
        if rpr is not None:
            run.append(deepcopy(rpr))
        parts = str(text).split('\n')
        t = etree.SubElement(run, qn('w:t'))
        t.set(XML_SPACE, 'preserve')
        t.text = parts[0]
        for part in parts[1:]:
            etree.SubElement(run, qn('w:br'))
            t = etree.SubElement(run, qn('w:t'))
            t.set(XML_SPACE, 'preserve')
            t.text = part
        prev.addnext(run)
        prev = run

def _replace_in_paragraph(paragraph, data, doc=None):
    def fmt(v):
        if isinstance(v, list): return '\n'.join(_item_to_text(item) for item in v)
        return _item_to_text(v) if isinstance(v, (dict, list)) else ('' if v is None else str(v))

    # Build full text from runs (used for detection + fallback path)
    runs = list(paragraph.runs)
    full_text = ''.join((r.text or '') for r in runs)

    # Quick check: if no placeholders present, skip
    para_keys = _extract_placeholder_keys(full_text)
    if not para_keys:
        return

    replacements = {}
    for k, v in data.items():
        sv = fmt(v)
        replacements[k] = sv
        replacements[k.lower()] = sv
    # Also try to match placeholder keys that differ only by extra chars
    # e.g. template has {{meeting_cc_role}} but data has meeting_CCs_role
    lower_keys = {k.lower(): fmt(v) for k, v in data.items()}
    for pk in para_keys:
        pkl = pk.lower()
        if pkl not in replacements:
            # Try fuzzy: remove double letters, strip s
            for dk, dv in lower_keys.items():
                # Normalize both keys: lowercase, remove consecutive duplicates
                norm_pk = re.sub(r'(.)\1+', r'\1', pkl.replace('_', ''))
                norm_dk = re.sub(r'(.)\1+', r'\1', dk.replace('_', ''))
                if norm_pk == norm_dk:
                    replacements[pkl] = dv
                    break
                # Also try stripping trailing 's' from data key segments
                dk_no_s = re.sub(r's_', '_', dk)
                if pkl == dk_no_s:
                    replacements[pkl] = dv
                    break
                # Aggressive prefix stripping: if pkl starts with common template prefixes
                # e.g. "meeting_summary_html" matches data key "summary_html"
                for prefix in ['meeting_', 'task_', 'arr_']:
                    if pkl.startswith(prefix) and pkl[len(prefix):] == dk:
                        replacements[pkl] = dv
                        break
                if pkl in replacements: break
    html_keys = {k.lower() for k, v in replacements.items() if _is_html(v)}
    has_html_placeholder = any(k.lower() in html_keys for k in para_keys)

    # Preferred path: replace inside each run's text to preserve template styling.
    # Fallback to paragraph-wide XML rewrite when placeholders span runs or contain HTML.
    used_run_level = False
    if not has_html_placeholder:
        updated_any = False
        for r in runs:
            t = r.text or ''
            nt = _replace_placeholders_in_text(t, replacements)
            if nt != t:
                r.text = nt
                updated_any = True
        if updated_any:
            new_full = ''.join((r.text or '') for r in paragraph.runs)
            unresolved_keys = _extract_placeholder_keys(new_full)
            unresolved = any(k.lower() in replacements for k in unresolved_keys)
            if not unresolved:
                used_run_level = True

    if used_run_level:
        return

    new_text = _replace_placeholders_in_text(full_text, replacements)
    plain_segments = None
    if not _is_html(new_text):
        # Capture per-placeholder run styles BEFORE text nodes are removed.
        plain_segments = _build_plain_segments_with_styles(full_text, runs, replacements)

    # Remove only text (w:t) and break (w:br) elements in runs, keep drawings
    had_text_runs = []
    for r in runs:
        r_elem = r._r
        removed = False
        for child in list(r_elem):
            if child.tag == qn('w:t') or child.tag == qn('w:br'):
                r_elem.remove(child)
                removed = True
        if removed:
            had_text_runs.append(r_elem)

    if not had_text_runs:
        # fallback
        paragraph.text = new_text
        return

    first = had_text_runs[0]
    if _is_html(new_text):
        _insert_html_runs(first, new_text, doc=doc)
    else:
        # Rebuild text as styled segments so placeholders on the same line can
        # keep distinct template styles (per run / per tag).
        _insert_plain_segments(first, plain_segments or [(new_text, None)])

def _item_to_text(item, sep=' | '):
    if isinstance(item, dict):
        vals = [str(x) for x in item.values() if x is not None and str(x) != '']
        return sep.join(vals)
    if isinstance(item, list):
        return '\n'.join(_item_to_text(v, sep=sep) for v in item)
    if item is None:
        return ''
    return str(item)

def _flatten_dict(d, parent_key='', sep='_'):
    """Recursively flatten a nested dictionary."""
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict) and not (isinstance(v.get('data'), list) and 'type' in v):
            items.extend(_flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)

def _format_collection_value(coll):
    ctype = str(coll.get('type', '')).strip().lower()
    items = coll.get('data', [])
    if not isinstance(items, list):
        return ''
    if ctype == 'bullets':
        lines = []
        for item in items:
            txt = _item_to_text(item, sep=' ')
            if txt:
                lines.append(f'• {txt}')
        return '\n'.join(lines)
    if ctype == 'html':
        return '\n'.join(_item_to_text(i, sep=' ') for i in items if _item_to_text(i, sep=' '))
    return '\n'.join(_item_to_text(i) for i in items if _item_to_text(i))

def _extract_collections(data):
    collections = {}
    for name, v in data.items():
        ctype = ''
        items = None
        if isinstance(v, dict) and isinstance(v.get('data'), list):
            ctype = str(v.get('type', '')).strip().lower()
            items = v.get('data', [])
        elif isinstance(v, list):
            # Also support bare arrays of dicts at top-level.
            items = v
        if items is None:
            continue
        rows = [it for it in items if isinstance(it, dict)]
        if not rows:
            continue
        keys = {k.lower() for row in rows for k in row.keys()}
        collections[name] = {
            'name': name,
            'type': ctype,
            'items': rows,
            'item_keys': keys,
        }
    return collections

def _flatten_collection_keys(collections):
    """Expose inner item keys as global placeholders.

    Example: for {"meeting_CCs": {"data": [{"meeting_CC": "A"}, {"meeting_CC": "B"}]}}
    expose "meeting_CC" -> "A\\nB" so templates can use {{meeting_CC}}
    even if {{meeting_CCs}} is absent.
    """
    out = {}
    for coll in collections.values():
        per_key = {}
        for item in coll['items']:
            for k, v in item.items():
                if v and str(v).strip():
                    per_key.setdefault(k, []).append(str(v))
        for k, vals in per_key.items():
            combined = '\n'.join(vals)
            out[k] = f"{out[k]}\n{combined}" if k in out else combined
    return out

def _expand_collection_paragraphs_in_block(block, collections):
    paras = list(block.paragraphs)
    for coll in collections.values():
        item_keys = coll['item_keys']
        if not item_keys:
            continue
        idxs = []
        inline_static_idxs = []
        for i, p in enumerate(paras):
            txt = ''.join((r.text or '') for r in p.runs)
            keys = {k.lower() for k in _extract_placeholder_keys(txt)}
            coll_keys_in_para = keys & item_keys
            if coll_keys_in_para and _has_static_text_outside_placeholders(txt):
                inline_static_idxs.append(i)
                continue
            # Repeat only "template-like" lines; avoid duplicating section titles
            # such as "משתתפים: {{meeting_participant}}".
            # Exception: if paragraph contains multiple placeholders from the same
            # collection item, repeat it even when it has static text so values
            # stay aligned row-by-row (e.g. name + email on same line).
            if coll_keys_in_para and (
                not _has_static_text_outside_placeholders(txt) or len(coll_keys_in_para) >= 2
            ):
                idxs.append(i)

        # Render static-text paragraphs inline from collection items, once only.
        for i in inline_static_idxs:
            p = paras[i]
            txt = ''.join((r.text or '') for r in p.runs)
            if '\n' in txt:
                # Keep legacy newline-header behavior in one paragraph.
                p.text = _render_header_plus_inline_items(txt, coll['items'])
            else:
                _render_inline_collection_in_paragraph(p, coll['items'])
        if not idxs:
            continue

        groups = []
        start = prev = idxs[0]
        for i in idxs[1:]:
            if i == prev + 1:
                prev = i
                continue
            groups.append((start, prev))
            start = prev = i
        groups.append((start, prev))

        for start, end in reversed(groups):
            group_paras = paras[start:end + 1]
            first_elem = group_paras[0]._p
            parent = first_elem.getparent()
            insert_idx = parent.index(first_elem)
            template_elems = [deepcopy(p._p) for p in group_paras]
            template_texts = [''.join((r.text or '') for r in p.runs) for p in group_paras]

            for p in group_paras:
                p._p.getparent().remove(p._p)

            for item_idx, item in enumerate(coll['items']):
                for template_elem, tpl_txt in zip(template_elems, template_texts):
                    keys = {k.lower() for k in _extract_placeholder_keys(tpl_txt)}
                    coll_keys_in_para = keys & item_keys
                    header_inline_mode = (
                        len(coll_keys_in_para) >= 2
                        and _has_static_text_outside_placeholders(tpl_txt)
                        and '\n' in tpl_txt
                    )
                    if header_inline_mode and item_idx > 0:
                        continue
                    new_elem = deepcopy(template_elem)
                    parent.insert(insert_idx, new_elem)
                    insert_idx += 1
                    p_new = Paragraph(new_elem, block)
                    if header_inline_mode:
                        p_new.text = _render_header_plus_inline_items(tpl_txt, coll['items'])
                    else:
                        # If template has a static header + newline + item placeholders
                        # in one paragraph, keep header only for the first rendered item.
                        if (
                            item_idx > 0
                            and len(coll_keys_in_para) >= 2
                            and _has_static_text_outside_placeholders(tpl_txt)
                            and '\n' in tpl_txt
                        ):
                            p_new.text = _trim_to_placeholder_line(tpl_txt)
                        _replace_in_paragraph(p_new, item)

def _expand_table_rows_in_doc(doc, collections):
    for table in doc.tables:
        r_idx = 0
        while r_idx < len(table.rows):
            row = table.rows[r_idx]
            row_text = '\n'.join(
                ''.join((r.text or '') for r in p.runs)
                for c in row.cells for p in c.paragraphs
            )
            row_keys = {k.lower() for k in _extract_placeholder_keys(row_text)}
            if not row_keys:
                r_idx += 1
                continue

            matches = [c for c in collections.values() if row_keys & c['item_keys']]
            if len(matches) != 1:
                r_idx += 1
                continue

            coll = matches[0]
            tbl = table._tbl
            template_tr = deepcopy(row._tr)
            insert_idx = tbl.index(row._tr)
            tbl.remove(row._tr)

            if not coll['items']:
                continue

            row_pos = r_idx
            for item in coll['items']:
                new_tr = deepcopy(template_tr)
                tbl.insert(insert_idx, new_tr)
                new_row = table.rows[row_pos]
                for p in new_row.cells[0].paragraphs: # or iterate all cells
                     pass # Wait, need to check how table rows are expanded
                # Actually, the original code had:
                for cell in new_row.cells:
                    for p in cell.paragraphs:
                        _replace_in_paragraph(p, item, doc=doc)
                insert_idx += 1
                row_pos += 1

            r_idx = row_pos

def _process_block(block, data, doc=None, collections=None):
    collections = collections or {}
    _expand_collection_paragraphs_in_block(block, collections)
    for p in block.paragraphs:
        _replace_in_paragraph(p, data, doc=doc)
    for table in getattr(block, 'tables', []):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, data, doc=doc)

def fill_template_from_flat(template_path, data, output=None):
    """Minimal template filler: replaces {{key}} with values from flat `data` dict.

    Keeps implementation tiny; updates main doc, tables, and header/footer blocks.
    """
    doc = Document(template_path)
    collections = _extract_collections(data)

    # Replace collection wrappers ({"type","data"}) with plain text fallback
    # so inline placeholders like {{meeting_participants}} resolve naturally.
    normalized_data = {}
    for name, v in data.items():
        val = v
        if isinstance(v, dict) and isinstance(v.get('data'), list):
            val = _format_collection_value(v)
        if isinstance(val, str) and _is_html(val):
            val = val.replace('\n', '')
        normalized_data[name] = val
    # Flatten all nested objects (non-collections)
    flattened_input = _flatten_dict(data)
    for k, v in flattened_input.items():
        if k not in normalized_data:
            normalized_data[k] = v

    # Make inner collection keys addressable directly in template placeholders.
    # Do not override explicit top-level keys provided by input data.
    flat_keys = _flatten_collection_keys(collections)
    for k, v in flat_keys.items():
        normalized_data.setdefault(k, v)
        normalized_data.setdefault(k.lower(), v)
        # Also add fuzzy variants: e.g. meeting_CCs_role -> meeting_cc_role
        fuzzy = re.sub(r'([A-Z])s_', lambda m: m.group(1).lower() + '_', k)
        normalized_data.setdefault(fuzzy, v)
        normalized_data.setdefault(fuzzy.lower(), v)
        # Strip doubled chars variant
        simple = re.sub(r'(.)\1+', r'\1', k.lower())
        normalized_data.setdefault(simple, v)

    _expand_table_rows_in_doc(doc, collections)
    _process_block(doc, normalized_data, doc=doc, collections=collections)

    def _replace_in_header_footer(hf):
        # iterate all <w:p> elements so we include paragraphs nested in tables/sdts
        for p_elem in hf._element.iter(qn('w:p')):
            p = Paragraph(p_elem, hf)
            _replace_in_paragraph(p, normalized_data, doc=doc)

    for section in doc.sections:
        _replace_in_header_footer(section.header)
        _replace_in_header_footer(section.footer)

    if output is None:
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf
    elif isinstance(output, str):
        doc.save(output)
    else:
        doc.save(output)
        output.seek(0)

def run_from_paths(template_path, json_path, output_path=None):
    """Load JSON from disk and fill a DOCX template."""
    template_path = Path(template_path)
    json_path = Path(json_path)

    with json_path.open('r', encoding='utf-8') as f:
        data = json.load(f)

    if output_path is None:
        output_path = template_path.with_name(f"{template_path.stem}_output.docx")
    else:
        output_path = Path(output_path)

    fill_template_from_flat(str(template_path), data, str(output_path))
    return output_path

def _build_arg_parser():
    parser = argparse.ArgumentParser(
        description='Fill a DOCX template using a JSON file.'
    )
    parser.add_argument('template_path', help='Path to template .docx file')
    parser.add_argument('json_path', help='Path to input .json data file')
    parser.add_argument(
        '-o',
        '--output',
        dest='output_path',
        default=None,
        help='Optional output .docx path (default: <template>_output.docx)',
    )
    return parser

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def fill_template_from_data(template_path, data, output=None):
    """Fill template from flat JSON dict (API entry point)."""
    return fill_template_from_flat(template_path, data, output)


if __name__ == '__main__':
    args = _build_arg_parser().parse_args()
    out = run_from_paths(args.template_path, args.json_path, args.output_path)
    print(out)
