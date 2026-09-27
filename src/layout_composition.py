"""Explicit, revision-bound layout transforms. No rewriting or renderer calls."""
from collections import Counter
from copy import deepcopy
import hashlib
from io import BytesIO
import math
import re
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from .layout_structure import map_document, unsupported_features


SPEC_KEYS = {'schema_version', 'expected_workspace_sha256', 'expected_inspection_sha256',
             'orders', 'typography', 'paragraph_rules', 'highlights', 'compact', 'whitespace',
             'linearize_tables', 'theme'}
RULE_KEYS = {'block', 'size_pt', 'bold', 'space_before_pt', 'space_after_pt',
             'line_spacing', 'keep_with_next', 'keep_together', 'page_break_before',
             'widow_control', 'indent_inches', 'first_line_indent_inches', 'alignment',
             'color_role', 'shading_role', 'borders', 'tab_stops'}
THEME_ROLES = {'ink', 'teal', 'accent', 'surface', 'highlight', 'text', 'muted', 'paper'}
ALIGNMENTS = {'left', 'center', 'right', 'justify'}
TAB_ALIGNMENTS = {'left', 'center', 'right', 'decimal'}
TAB_LEADERS = {'none', 'dots', 'dashes', 'lines', 'heavy', 'middle_dot'}
BORDER_EDGES = {'left', 'right', 'top', 'bottom'}
BORDER_STYLES = {'single', 'double', 'dotted', 'dashed'}


def luminance(color):
    values = [int(color[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    values = [value / 12.92 if value <= .04045 else ((value + .055) / 1.055) ** 2.4
              for value in values]
    return .2126 * values[0] + .7152 * values[1] + .0722 * values[2]


def contrast_ratio(first, second):
    low, high = sorted((luminance(first), luminance(second)))
    return (high + .05) / (low + .05)


def theme_contrast(theme):
    pairs = [
        ('text', 'paper', 4.5, 'normal text'),
        ('muted', 'paper', 4.5, 'secondary text'),
        ('ink', 'paper', 4.5, 'primary text'),
        ('teal', 'paper', 4.5, 'emphasized text'),
        ('accent', 'paper', 3.0, 'visual boundary'),
        ('ink', 'surface', 4.5, 'section heading'),
        ('teal', 'highlight', 4.5, 'phrase highlight'),
    ]
    evidence = []
    for foreground, background, minimum, purpose in pairs:
        ratio = contrast_ratio(theme[foreground], theme[background])
        evidence.append({'foreground_role': foreground, 'background_role': background,
                         'ratio': round(ratio, 2), 'minimum': minimum, 'purpose': purpose})
        if ratio < minimum:
            raise ValueError('Unsafe theme contrast for %s on %s' % (foreground, background))
    return evidence


def validate_role(value, theme):
    if not isinstance(value, str) or theme is None or value not in theme:
        raise ValueError('Unknown theme role: ' + str(value))


def keys(value, allowed, required=()):
    if not isinstance(value, dict) or set(value) - allowed or set(required) - set(value):
        raise ValueError('Invalid composition keys; allowed: ' + ', '.join(sorted(allowed)))


def number(value, low, high):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or not low <= value <= high:
        raise ValueError('Invalid composition numeric value')


def validate_composition(params):
    keys(params, {'specification'}, {'specification'})
    spec = params['specification']
    keys(spec, SPEC_KEYS, {'schema_version', 'expected_workspace_sha256', 'expected_inspection_sha256'})
    if type(spec['schema_version']) is not int or spec['schema_version'] != 1:
        raise ValueError('Unsupported composition schema version')
    for key in ('expected_workspace_sha256', 'expected_inspection_sha256'):
        value = spec[key]
        if not isinstance(value, str) or len(value) != 64 or any(c not in '0123456789abcdef' for c in value):
            raise ValueError('Invalid composition checksum')
    for name in ('orders', 'paragraph_rules', 'highlights', 'compact', 'whitespace', 'linearize_tables'):
        if not isinstance(spec.get(name, []), list):
            raise ValueError(name + ' must be a list')
    theme = spec.get('theme')
    if theme is not None:
        if not isinstance(theme, dict) or set(theme) != THEME_ROLES:
            raise ValueError('Theme must define exactly: ' + ', '.join(sorted(THEME_ROLES)))
        for role, color in theme.items():
            if not isinstance(color, str) or not re.fullmatch(r'#[0-9A-Fa-f]{6}', color):
                raise ValueError('Theme color for %s must be #RRGGBB' % role)
        theme_contrast(theme)
    typography = spec.get('typography', {})
    keys(typography, {'font', 'body_size_pt', 'margin_inches', 'color'})
    if 'color' in typography and (not isinstance(typography['color'], str) or not re.fullmatch(r'#[0-9A-Fa-f]{6}', typography['color'])):
        raise ValueError('Color must be #RRGGBB')
    if 'font' in typography and (not isinstance(typography['font'], str) or not typography['font'].strip()):
        raise ValueError('Font must be nonempty')
    for key, low, high in [('body_size_pt', 6, 30), ('margin_inches', .2, 2)]:
        if key in typography:
            number(typography[key], low, high)
    for order in spec.get('orders', []):
        keys(order, {'container', 'blocks'}, {'container', 'blocks'})
        if not isinstance(order['container'], str) or not isinstance(order['blocks'], list) or any(not isinstance(x, str) for x in order['blocks']):
            raise ValueError('Invalid block order')
    for item in spec.get('linearize_tables', []):
        keys(item, {'table', 'paragraph_order'}, {'table', 'paragraph_order'})
        if not isinstance(item['table'], str) or not isinstance(item['paragraph_order'], list) or any(not isinstance(x, str) for x in item['paragraph_order']):
            raise ValueError('Invalid table linearization rule')
    for rule in spec.get('paragraph_rules', []):
        keys(rule, RULE_KEYS, {'block'})
        if not isinstance(rule['block'], str):
            raise ValueError('Invalid paragraph target')
        for key in ('bold', 'keep_with_next', 'keep_together', 'page_break_before', 'widow_control'):
            if key in rule and not isinstance(rule[key], bool):
                raise ValueError(key + ' must be boolean')
        for key, low, high in [('size_pt', 6, 30), ('space_before_pt', 0, 72),
                              ('space_after_pt', 0, 72), ('line_spacing', .8, 3),
                              ('indent_inches', 0, 2), ('first_line_indent_inches', -2, 2)]:
            if key in rule:
                number(rule[key], low, high)
        if 'alignment' in rule and rule['alignment'] not in ALIGNMENTS:
            raise ValueError('Unsupported paragraph alignment')
        for key in ('color_role', 'shading_role'):
            if key in rule:
                validate_role(rule[key], theme)
        if 'borders' in rule:
            borders = rule['borders']
            if not isinstance(borders, dict) or not borders or set(borders) - BORDER_EDGES:
                raise ValueError('Invalid paragraph borders')
            for edge, border in borders.items():
                keys(border, {'color_role', 'style', 'width_pt', 'space_pt'},
                     {'color_role', 'style', 'width_pt', 'space_pt'})
                validate_role(border['color_role'], theme)
                if border['style'] not in BORDER_STYLES:
                    raise ValueError('Unsupported border style')
                number(border['width_pt'], .25, 6)
                number(border['space_pt'], 0, 31)
                background = rule.get('shading_role', 'paper')
                if contrast_ratio(theme[border['color_role']], theme[background]) < 3:
                    raise ValueError('Unsafe paragraph border contrast')
        if 'tab_stops' in rule:
            tabs = rule['tab_stops']
            if not isinstance(tabs, list) or not tabs:
                raise ValueError('tab_stops must be a nonempty list')
            positions = []
            for tab in tabs:
                keys(tab, {'position_inches', 'alignment', 'leader'},
                     {'position_inches', 'alignment'})
                number(tab['position_inches'], .1, 10)
                if tab['alignment'] not in TAB_ALIGNMENTS:
                    raise ValueError('Unsupported tab alignment')
                if tab.get('leader', 'none') not in TAB_LEADERS:
                    raise ValueError('Unsupported tab leader')
                positions.append(tab['position_inches'])
            if positions != sorted(set(positions)):
                raise ValueError('Tab stops must be unique and increasing')
        if theme is not None and ('color_role' in rule or 'shading_role' in rule):
            foreground = rule.get('color_role', 'text')
            background = rule.get('shading_role', 'paper')
            large = rule.get('size_pt', 0) >= 18 or (rule.get('size_pt', 0) >= 14 and rule.get('bold') is True)
            if contrast_ratio(theme[foreground], theme[background]) < (3 if large else 4.5):
                raise ValueError('Unsafe paragraph foreground/background contrast')
    for item in spec.get('highlights', []):
        keys(item, {'block', 'text', 'expected_matches', 'color_role', 'shading_role', 'bold'},
             {'block', 'text', 'expected_matches'})
        if not isinstance(item['block'], str) or not isinstance(item['text'], str) or not item['text'] or type(item['expected_matches']) is not int or item['expected_matches'] < 1:
            raise ValueError('Invalid highlight selector')
        if 'bold' in item and not isinstance(item['bold'], bool):
            raise ValueError('Highlight bold must be boolean')
        for key in ('color_role', 'shading_role'):
            if key in item:
                validate_role(item[key], theme)
        if theme is not None and ('color_role' in item or 'shading_role' in item):
            foreground = item.get('color_role', 'text')
            background = item.get('shading_role', 'paper')
            if contrast_ratio(theme[foreground], theme[background]) < 4.5:
                raise ValueError('Unsafe highlight foreground/background contrast')
    for item in spec.get('compact', []):
        keys(item, {'blocks', 'separator'}, {'blocks', 'separator'})
        if not isinstance(item['blocks'], list) or len(item['blocks']) < 2 or any(not isinstance(x, str) for x in item['blocks']):
            raise ValueError('Compact needs at least two paragraph IDs')
        if item['separator'] not in (' ', ' | ', '\n', '\t'):
            raise ValueError('Unsupported compact separator')
    paragraph_rules = {rule['block']: rule for rule in spec.get('paragraph_rules', [])}
    for item in spec.get('compact', []):
        if item['separator'] == '\t' and not paragraph_rules.get(item['blocks'][0], {}).get('tab_stops'):
            raise ValueError('Tab compaction requires a tab stop on the surviving paragraph')
    for item in spec.get('whitespace', []):
        keys(item, {'block', 'action'}, {'block', 'action'})
        if not isinstance(item['block'], str) or item['action'] != 'remove_empty':
            raise ValueError('Only explicit remove_empty whitespace rules are supported')


def block_children(container):
    return [node for node in container if node.tag in (qn('w:p'), qn('w:tbl'))]


def text_of(node):
    return ''.join(n.text or '' for n in node.iter(qn('w:t')))


def canonical(node):
    return etree.tostring(node, method='c14n') if node is not None else b''


def bookmark_inventory(node):
    position, values = 0, []
    for child in node.iter():
        if child.tag == qn('w:t'):
            position += len(child.text or '')
        elif child.tag in (qn('w:bookmarkStart'), qn('w:bookmarkEnd')):
            values.append((child.tag, tuple(sorted(child.attrib.items())), position))
    return values


def validate_bookmarks(node):
    paragraphs = [node] if node.tag == qn('w:p') else list(node.iter(qn('w:p')))
    for para in paragraphs:
        starts, ends = {}, {}
        for item in para.iter():
            if item.tag in (qn('w:bookmarkStart'), qn('w:bookmarkEnd')):
                selected = starts if item.tag == qn('w:bookmarkStart') else ends
                identifier = item.get(qn('w:id'))
                if not identifier or identifier in selected:
                    raise ValueError('Unsupported duplicate bookmark range')
                selected[identifier] = item
        if set(starts) != set(ends):
            raise ValueError('Bookmark spans paragraphs; cannot preserve it during composition')
        ordered = list(para.iter())
        if any(ordered.index(starts[b]) >= ordered.index(ends[b]) for b in starts):
            raise ValueError('Invalid bookmark range order')


def safe_node(node):
    forbidden = ('fldChar', 'instrText', 'fldSimple', 'ins', 'del', 'moveFrom', 'moveTo',
                 'drawing', 'pict', 'txbxContent', 'sdt', 'sectPr',
                 'commentRangeStart', 'commentRangeEnd', 'commentReference')
    if any(next(node.iter(qn('w:' + name)), None) is not None for name in forbidden):
        raise ValueError('Affected block has unsupported fields, section boundaries, annotations, or drawings')
    if unsupported_features(node):
        raise ValueError('Affected block has unsupported structure: ' + ', '.join(unsupported_features(node)))
    validate_bookmarks(node)
    if node.tag == qn('w:tbl'):
        if len(list(node.iter(qn('w:tbl')))) > 1 or any(next(node.iter(qn('w:' + name)), None) is not None for name in ('gridSpan', 'vMerge')):
            raise ValueError('Affected table has nested or merged cells')


def paragraph_target(nodes, block):
    node = nodes.get(block)
    if node is None or node.tag != qn('w:p'):
        raise ValueError('Unknown or non-paragraph target: ' + block)
    safe_node(node)
    return node


def effective_numbering(doc, node):
    direct = node.find('./' + qn('w:pPr') + '/' + qn('w:numPr'))
    if direct is not None:
        return direct
    style = Paragraph(node, doc).style
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        value = style.element.find('./' + qn('w:pPr') + '/' + qn('w:numPr'))
        if value is not None:
            return value
        style = style.base_style
    return None


def compatible_run_defaults(paragraphs):
    defaults = [p.find('./' + qn('w:pPr') + '/' + qn('w:rPr')) for p in paragraphs]
    if len({canonical(p) for p in defaults}) == 1:
        return True
    properties = set(c.tag for p in defaults if p is not None for c in p)
    for tag in properties:
        values = [p.find(tag) if p is not None else None for p in defaults]
        if len({canonical(value) for value in values}) == 1:
            continue
        # Differing paragraph defaults are safe only when every moved run explicitly
        # overrides the complete property. Otherwise the surviving paragraph could
        # change another paragraph's inherited appearance.
        attributes = set(k for value in values if value is not None for k in value.attrib)
        for para in paragraphs:
            for run in para.iter(qn('w:r')):
                direct = run.find('./' + qn('w:rPr') + '/' + tag)
                if direct is None or not attributes.issubset(set(direct.attrib)):
                    return False
    return True


def inventory(node):
    """Character and control tokens retain their hyperlink context across run splits."""
    tokens = []
    for n in node.iter():
        if n.tag in (qn('w:t'), qn('w:tab'), qn('w:br'), qn('w:cr'), qn('w:lastRenderedPageBreak')):
            if not any(parent.tag == qn('w:r') for parent in n.iterancestors()):
                continue
            link = next((p for p in n.iterancestors() if p.tag == qn('w:hyperlink')), None)
            context = tuple(sorted(link.attrib.items())) if link is not None else ()
            if n.tag == qn('w:t'):
                tokens.extend(('text', char, context) for char in n.text or '')
            else:
                tokens.append((n.tag, tuple(sorted(n.attrib.items())), context))
    return tokens


def has_content_break_or_tab(node):
    return any(item.tag in (qn('w:br'), qn('w:cr'), qn('w:tab'))
               and any(parent.tag == qn('w:r') for parent in item.iterancestors())
               for item in node.iter())


PARAGRAPH_ALIGNMENT_VALUES = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}
TAB_ALIGNMENT_VALUES = {
    'left': WD_TAB_ALIGNMENT.LEFT,
    'center': WD_TAB_ALIGNMENT.CENTER,
    'right': WD_TAB_ALIGNMENT.RIGHT,
    'decimal': WD_TAB_ALIGNMENT.DECIMAL,
}
TAB_LEADER_VALUES = {
    'none': WD_TAB_LEADER.SPACES,
    'dots': WD_TAB_LEADER.DOTS,
    'dashes': WD_TAB_LEADER.DASHES,
    'lines': WD_TAB_LEADER.LINES,
    'heavy': WD_TAB_LEADER.HEAVY,
    'middle_dot': WD_TAB_LEADER.MIDDLE_DOT,
}


def replace_shading(properties, color):
    for node in properties.findall(qn('w:shd')):
        properties.remove(node)
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), color[1:].upper())
    properties.append(shading)


def apply_run_visual(run_node, theme, color_role=None, shading_role=None, bold=None):
    run = Run(run_node, None)
    if color_role is not None:
        run.font.color.rgb = RGBColor.from_string(theme[color_role][1:])
    if shading_role is not None:
        replace_shading(run_node.get_or_add_rPr(), theme[shading_role])
    if bold is not None:
        run.bold = bold


def apply_borders(node, borders, theme):
    properties = node.get_or_add_pPr()
    existing = properties.find(qn('w:pBdr'))
    if existing is not None:
        properties.remove(existing)
    group = OxmlElement('w:pBdr')
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge not in borders:
            continue
        settings = borders[edge]
        border = OxmlElement('w:' + edge)
        border.set(qn('w:val'), settings['style'])
        border.set(qn('w:sz'), str(round(settings['width_pt'] * 8)))
        border.set(qn('w:space'), str(round(settings['space_pt'])))
        border.set(qn('w:color'), theme[settings['color_role']][1:].upper())
        group.append(border)
    properties.append(group)


def validate_tab_widths(doc, rules):
    widths = [(section.page_width - section.left_margin - section.right_margin) / 914400
              for section in doc.sections]
    maximum = min(widths)
    for rule in rules:
        for tab in rule.get('tab_stops', []):
            if tab['position_inches'] > maximum:
                raise ValueError('Tab stop falls outside printable page width')


def apply_rule(node, rule, theme=None):
    para = Paragraph(node, None)
    fmt = para.paragraph_format
    for key, prop in [('space_before_pt', 'space_before'), ('space_after_pt', 'space_after')]:
        if key in rule:
            setattr(fmt, prop, Pt(rule[key]))
    for key, prop in [('indent_inches', 'left_indent'), ('first_line_indent_inches', 'first_line_indent')]:
        if key in rule:
            setattr(fmt, prop, Inches(rule[key]))
            ind = node.find('./' + qn('w:pPr') + '/' + qn('w:ind'))
            for attribute in (('leftChars',) if key == 'indent_inches' else ('firstLineChars', 'hangingChars')):
                if ind is not None:
                    ind.attrib.pop(qn('w:' + attribute), None)
    for key in ('line_spacing', 'keep_with_next', 'keep_together', 'page_break_before', 'widow_control'):
        if key in rule:
            setattr(fmt, key, rule[key])
    if 'alignment' in rule:
        fmt.alignment = PARAGRAPH_ALIGNMENT_VALUES[rule['alignment']]
    if 'tab_stops' in rule:
        fmt.tab_stops.clear_all()
        for tab in rule['tab_stops']:
            fmt.tab_stops.add_tab_stop(Inches(tab['position_inches']),
                                       TAB_ALIGNMENT_VALUES[tab['alignment']],
                                       TAB_LEADER_VALUES[tab.get('leader', 'none')])
    if 'shading_role' in rule:
        replace_shading(node.get_or_add_pPr(), theme[rule['shading_role']])
    if 'borders' in rule:
        apply_borders(node, rule['borders'], theme)
    for element in node.iter(qn('w:r')):
        run = Run(element, None)
        if 'size_pt' in rule:
            run.font.size = Pt(rule['size_pt'])
        if 'bold' in rule:
            run.bold = rule['bold']
        if 'color_role' in rule:
            run.font.color.rgb = RGBColor.from_string(theme[rule['color_role']][1:])


def highlight(node, phrase, count, theme=None, color_role=None, shading_role=None, bold=True):
    # Word's lastRenderedPageBreak is cached, nonprinting pagination metadata.
    # Isolate such markers so text-run splitting keeps each marker exactly once.
    for run in list(node.iter(qn('w:r'))):
        if run.find(qn('w:lastRenderedPageBreak')) is not None:
            if any(c.tag not in (qn('w:rPr'), qn('w:t'), qn('w:lastRenderedPageBreak')) for c in run):
                raise ValueError('Highlight target contains non-text runs')
            parent, index = run.getparent(), run.getparent().index(run)
            for child in run:
                if child.tag == qn('w:rPr'):
                    continue
                part = deepcopy(run)
                for value in list(part):
                    if value.tag != qn('w:rPr'):
                        part.remove(value)
                part.append(deepcopy(child))
                parent.insert(index, part)
                index += 1
            parent.remove(run)
    runs = list(node.iter(qn('w:r')))
    # Restrict to text-only runs. Splitting control/field runs would be destructive.
    for run in runs:
        if any(c.tag not in (qn('w:rPr'), qn('w:t'), qn('w:lastRenderedPageBreak')) for c in run):
            raise ValueError('Highlight target contains non-text runs')
    text = ''.join(text_of(run) for run in runs)
    matches, start = [], 0
    while True:
        pos = text.find(phrase, start)
        if pos < 0:
            break
        matches.append((pos, pos + len(phrase)))
        start = pos + len(phrase)
    if len(matches) != count:
        raise ValueError('Highlight expected %d matches, found %d: %s' % (count, len(matches), phrase))
    offset = 0
    for run in runs:
        content = text_of(run)
        end = offset + len(content)
        if not content:
            continue
        cuts = {offset, end}
        for low, high in matches:
            if offset < low < end:
                cuts.add(low)
            if offset < high < end:
                cuts.add(high)
        parent, index = run.getparent(), run.getparent().index(run)
        boundaries = sorted(cuts)
        if not any(low < end and high > offset for low, high in matches):
            offset = end
            continue
        for low, high in zip(boundaries, boundaries[1:]):
            new = deepcopy(run)
            for child in list(new):
                if child.tag != qn('w:rPr'):
                    new.remove(child)
            value = OxmlElement('w:t')
            value.text = content[low - offset:high - offset]
            value.set(qn('xml:space'), 'preserve')
            new.append(value)
            if any(a <= low and high <= b for a, b in matches):
                apply_run_visual(new, theme, color_role, shading_role, bold)
            parent.insert(index, new)
            index += 1
        parent.remove(run)
        offset = end
    return len(matches)


def linearize_tables(nodes, containers, rules):
    """Replace supported body layout tables with their declared leaf paragraphs."""
    removed, results = set(), []
    for rule in rules:
        table_id = rule['table']
        table = nodes.get(table_id)
        if table_id in removed or table is None or table.tag != qn('w:tbl') or table.getparent().tag != qn('w:body'):
            raise ValueError('Linearization requires a distinct body-level table')
        features = set(unsupported_features(table))
        if features - {'merged_cells', 'nested_tables'}:
            raise ValueError('Unsupported content in linearized table: ' + ', '.join(sorted(features)))
        if any(next(table.iter(qn('w:' + name)), None) is not None for name in ('vMerge', 'hMerge', 'sectPr', 'commentRangeStart', 'commentRangeEnd')):
            raise ValueError('Linearization supports horizontal gridSpan merges only; no section/annotation boundaries')
        validate_bookmarks(table)
        structural_allowed = {
            qn('w:tbl'): {qn('w:tblPr'), qn('w:tblGrid'), qn('w:tr')},
            qn('w:tr'): {qn('w:trPr'), qn('w:tc')},
            qn('w:tc'): {qn('w:tcPr'), qn('w:p'), qn('w:tbl')},
        }
        for node in table.iter():
            if node.tag in structural_allowed and any(c.tag not in structural_allowed[node.tag] for c in node):
                raise ValueError('Table contains unsupported wrapped structure')
        leaves = set(table.iter(qn('w:p')))
        ids = rule['paragraph_order']
        if len(ids) != len(set(ids)) or any(b not in nodes or nodes[b].tag != qn('w:p') for b in ids):
            raise ValueError('Unknown or duplicate linearization paragraph')
        if set(nodes[b] for b in ids) != leaves:
            raise ValueError('Linearization must include every table paragraph exactly once')
        shells = [b for b, node in nodes.items() if node.tag == qn('w:tbl') and (node is table or table in node.iterancestors())]
        shell_properties = [{'block': b, 'xml_sha256': hashlib.sha256(canonical(nodes[b])).hexdigest()}
                            for b in shells]
        parent, index = table.getparent(), table.getparent().index(table)
        for offset, b in enumerate(ids):
            parent.insert(index + offset, nodes[b])
        parent.remove(table)
        removed.update(shells)
        results.append({'table': table_id, 'paragraph_order': ids,
                        'body_positions': {b: index + i for i, b in enumerate(ids)},
                        'removed_table_shells': shell_properties})
    return removed, results


def compose_layout(data, params):
    validate_composition(params)
    spec = params['specification']
    if hashlib.sha256(data).hexdigest() != spec['expected_workspace_sha256']:
        raise ValueError('Stale composition workspace checksum')
    doc = Document(BytesIO(data))
    nodes, containers = map_document(doc)
    before = {key: inventory(node) for key, node in nodes.items() if node.tag == qn('w:p')}
    numbering = {key: canonical(node.find('./' + qn('w:pPr') + '/' + qn('w:numPr')))
                 for key, node in nodes.items() if node.tag == qn('w:p')}
    bookmarks = {key: bookmark_inventory(node) for key, node in nodes.items() if node.tag == qn('w:p')}
    # Snapshot raw package elements. Only changed document/header/footer XML is emitted.
    parts = {str(part.partname).lstrip('/'): (part, canonical(part._element))
             for part in doc.part.package.parts if hasattr(part, '_element')}
    removed_shells, linearized = linearize_tables(nodes, containers, spec.get('linearize_tables', []))
    affected_containers = set()
    for order in spec.get('orders', []):
        name = order['container']
        if name not in containers or name in affected_containers:
            raise ValueError('Unknown or duplicate order container')
        affected_containers.add(name)
        container = containers[name]
        if container.tag not in (qn('w:body'), qn('w:tc')):
            raise ValueError('Only body and table-cell ordering is supported')
        current = block_children(container)
        requested = order['blocks']
        if len(requested) != len(set(requested)) or any(b not in nodes for b in requested):
            raise ValueError('Unknown or duplicate ordered block')
        if set(nodes[b] for b in requested) != set(current):
            raise ValueError('Order must contain every direct container block exactly once')
        if any(n.tag not in (qn('w:p'), qn('w:tbl'), qn('w:sectPr'), qn('w:tcPr')) for n in container):
            raise ValueError('Container has unsupported wrapped blocks')
        for n in current:
            safe_node(n)
        slots = [container.index(n) for n in current]
        for n in current:
            container.remove(n)
        for slot, block in zip(slots, requested):
            container.insert(slot, nodes[block])
    typography = spec.get('typography', {})
    if typography:
        for node in nodes.values():
            if node.tag != qn('w:p'):
                continue
            safe_node(node)
            for element in node.iter(qn('w:r')):
                run = Run(element, None)
                if 'font' in typography:
                    run.font.name = typography['font']
                    fonts = element.find('./' + qn('w:rPr') + '/' + qn('w:rFonts'))
                    for key in ('asciiTheme', 'hAnsiTheme', 'csTheme', 'cstheme'):
                        fonts.attrib.pop(qn('w:' + key), None)
                    fonts.set(qn('w:cs'), typography['font'])
                if 'body_size_pt' in typography:
                    run.font.size = Pt(typography['body_size_pt'])
                if 'color' in typography:
                    run.font.color.rgb = RGBColor.from_string(typography['color'][1:])
        if 'margin_inches' in typography:
            for section in doc.sections:
                section.top_margin = section.bottom_margin = Inches(typography['margin_inches'])
                section.left_margin = section.right_margin = Inches(typography['margin_inches'])
    validate_tab_widths(doc, spec.get('paragraph_rules', []))
    targeted = set()
    for rule in spec.get('paragraph_rules', []):
        if rule['block'] in targeted:
            raise ValueError('Duplicate paragraph rule')
        targeted.add(rule['block'])
        apply_rule(paragraph_target(nodes, rule['block']), rule, spec.get('theme'))
    highlights = []
    for item in spec.get('highlights', []):
        matched = highlight(paragraph_target(nodes, item['block']), item['text'], item['expected_matches'],
                            spec.get('theme'), item.get('color_role'), item.get('shading_role'),
                            item.get('bold', True))
        highlights.append(dict(item, matched=matched))
    removed, whitespace, compact_targets = set(removed_shells), [], set()
    for item in spec.get('compact', []):
        ids = item['blocks']
        if len(ids) != len(set(ids)) or any(b in compact_targets for b in ids):
            raise ValueError('Overlapping compact groups')
        compact_targets.update(ids)
        ps = [paragraph_target(nodes, b) for b in ids]
        parent = ps[0].getparent()
        if any(p.getparent() is not parent for p in ps):
            raise ValueError('Compact paragraphs must share a container')
        positions = [parent.index(p) for p in ps]
        if positions != list(range(positions[0], positions[0] + len(ps))):
            raise ValueError('Compact paragraphs must be adjacent in declared order')
        if any(effective_numbering(doc, p) is not None for p in ps):
            raise ValueError('Numbered paragraphs cannot be compacted')
        if len({Paragraph(p, doc).style.style_id for p in ps}) != 1:
            raise ValueError('Compact paragraphs must have compatible inherited styles')
        if not compatible_run_defaults(ps):
            raise ValueError('Compact paragraphs must have compatible paragraph run defaults')
        if any(has_content_break_or_tab(p)
               or Paragraph(p, None).paragraph_format.page_break_before for p in ps):
            raise ValueError('Compact paragraphs cannot contain explicit breaks or tabs')
        # Only simple text runs and hyperlinks are reconstructed; all link nodes move intact.
        if any(c.tag not in (qn('w:pPr'), qn('w:r'), qn('w:hyperlink')) for p in ps for c in p):
            raise ValueError('Unsupported compact paragraph structure')
        expected = list(before[ids[0]])
        first = ps[0]
        for b, p in zip(ids[1:], ps[1:]):
            run = OxmlElement('w:r')
            if item['separator'] == '\n':
                run.append(OxmlElement('w:br'))
                expected.append((qn('w:br'), (), ()))
            elif item['separator'] == '\t':
                run.append(OxmlElement('w:tab'))
                expected.append((qn('w:tab'), (), ()))
            else:
                t = OxmlElement('w:t')
                t.text = item['separator']
                t.set(qn('xml:space'), 'preserve')
                run.append(t)
                expected.extend(('text', ch, ()) for ch in item['separator'])
            first.append(run)
            for child in list(p):
                if child.tag != qn('w:pPr'):
                    first.append(child)
            expected.extend(before[b])
            parent.remove(p)
            removed.add(b)
        before[ids[0]] = expected
        result = {'compact_blocks': ids, 'separator': item['separator']}
        if item['separator'] == '\t':
            result['inserted_token'] = 'tab'
        whitespace.append(result)
    for item in spec.get('whitespace', []):
        b = item['block']
        if b in removed or b in compact_targets:
            raise ValueError('Overlapping whitespace targets')
        node = paragraph_target(nodes, b)
        if before[b] or effective_numbering(doc, node) is not None or any(child.tag not in (qn('w:pPr'), qn('w:r')) for child in node) or any(child.tag != qn('w:rPr') for r in node.findall(qn('w:r')) for child in r):
            raise ValueError('remove_empty requires a truly empty paragraph')
        if node.getparent().tag == qn('w:tc') and (node is node.getparent()[-1] or len(node.getparent().findall(qn('w:p'))) == 1):
            raise ValueError('Cannot remove a cell\'s only or trailing paragraph')
        node.getparent().remove(node)
        removed.add(b)
        whitespace.append(item)
    for b, expected in before.items():
        if b in removed:
            continue
        if inventory(nodes[b]) != expected or canonical(nodes[b].find('./' + qn('w:pPr') + '/' + qn('w:numPr'))) != numbering[b]:
            raise ValueError('Composition failed text/link/numbering preservation: ' + b)
        if bookmark_inventory(nodes[b]) != bookmarks[b]:
            raise ValueError('Composition failed bookmark preservation: ' + b)
    # Block movement must be a permutation, minus explicitly compacted/empty nodes.
    after_nodes, _ = map_document(doc)
    if Counter(id(n) for n in after_nodes.values()) != Counter(id(n) for b, n in nodes.items() if b not in removed):
        raise ValueError('Composition lost or duplicated document blocks')
    replacements = {}
    for name, (part, original) in parts.items():
        if canonical(part._element) != original:
            if name != 'word/document.xml' and not (name.startswith('word/header') or name.startswith('word/footer')):
                raise ValueError('Composition changed an unsupported package part')
            replacements[name] = etree.tostring(part._element, xml_declaration=True, encoding='UTF-8', standalone=True)
    output = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(output, 'w') as target:
        for entry in source.infolist():
            target.writestr(entry, replacements.get(entry.filename, source.read(entry.filename)))
    return output.getvalue(), {'specification': spec, 'input_sha256': hashlib.sha256(data).hexdigest(),
                              'output_sha256': hashlib.sha256(output.getvalue()).hexdigest(),
                              'theme_contrast': theme_contrast(spec['theme']) if 'theme' in spec else [],
                              'visual_rules': [{key: value for key, value in rule.items()
                                                if key in {'block', 'alignment', 'color_role', 'shading_role',
                                                           'borders', 'tab_stops'}}
                                               for rule in spec.get('paragraph_rules', [])
                                               if set(rule) & {'alignment', 'color_role', 'shading_role',
                                                               'borders', 'tab_stops'}],
                              'highlights': highlights, 'whitespace': whitespace,
                              'linearized_tables': linearized,
                              'preservation': {'text': True, 'hyperlinks': True, 'numbering': True, 'bookmarks': True,
                                               'block_inventory': True, 'untouched_package_parts': True}}
