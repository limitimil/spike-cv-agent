"""Read-only, revision-scoped inventory of document blocks and their XML order.

Identifiers describe one inspected revision, not a semantic section or a stable
identity across edits.  Walking actual cells avoids the aliases python-docx uses
for merged cells.  Header/footer references are read without creating new parts.
"""

from hashlib import sha256
from io import BytesIO

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def map_document(doc):
    """Return ordered dictionaries of block IDs and container IDs to XML nodes."""
    nodes, containers = {}, {}

    def walk(container, identifier):
        containers[identifier] = container
        block_index = 0
        for child in container:
            if child.tag not in (qn('w:p'), qn('w:tbl')):
                continue
            block_id = f'{identifier}/b{block_index}'
            nodes[block_id] = child
            block_index += 1
            if child.tag == qn('w:tbl'):
                for row_index, row in enumerate(child.findall(qn('w:tr'))):
                    for cell_index, cell in enumerate(row.findall(qn('w:tc'))):
                        walk(cell, f'{block_id}/r{row_index}/c{cell_index}')

    walk(doc.element.body, 'body')
    seen = set()
    counts = {'header': 0, 'footer': 0}
    for section in doc.sections:
        for reference in section._sectPr:
            if reference.tag not in (qn('w:headerReference'), qn('w:footerReference')):
                continue
            part = doc.part.related_parts[reference.get(qn('r:id'))]
            identity = str(part.partname)
            if identity in seen:
                continue
            seen.add(identity)
            kind = 'header' if reference.tag == qn('w:headerReference') else 'footer'
            identifier = f'{kind}{counts[kind]}'
            counts[kind] += 1
            walk(part.element, identifier)
    return nodes, containers


def _nearest_paragraph(element):
    for ancestor in element.iterancestors():
        if ancestor.tag == qn('w:p'):
            return ancestor
    return None


def _run_text(run):
    # Textbox paragraphs are separate content; do not duplicate their text in
    # the host paragraph. Their presence is an unsupported transformation marker.
    result = []
    for child in run:
        if child.tag in (qn('w:t'), qn('w:delText')):
            result.append(child.text or '')
        elif child.tag == qn('w:tab'):
            result.append('\t')
        elif child.tag in (qn('w:br'), qn('w:cr')):
            result.append('\n')
        elif child.tag == qn('w:noBreakHyphen'):
            result.append('\u2011')
        elif child.tag == qn('w:softHyphen'):
            result.append('\u00ad')
    return ''.join(result)


def paragraph_runs(element):
    """Yield actual runs, including hyperlink runs, in XML document order."""
    for run in element.iter(qn('w:r')):
        if _nearest_paragraph(run) is element:
            yield run


def paragraph_text(element):
    return ''.join(_run_text(run) for run in paragraph_runs(element))


def unsupported_features(element):
    """Name structures whose affected transformations must be rejected."""
    features = set()
    tags = {child.tag for child in element.iter()}
    if tags & {qn('w:fldChar'), qn('w:fldSimple'), qn('w:instrText')}:
        features.add('fields')
    if tags & {qn('w:ins'), qn('w:del'), qn('w:moveFrom'), qn('w:moveTo'),
               qn('w:pPrChange'), qn('w:rPrChange'), qn('w:tblPrChange'),
               qn('w:tcPrChange'), qn('w:trPrChange'), qn('w:sectPrChange')}:
        features.add('tracked_changes')
    if tags & {qn('w:drawing'), qn('w:pict'), qn('w:object')}:
        features.add('drawings')
    if qn('w:txbxContent') in tags:
        features.add('textboxes')
    if tags & {qn('w:sdt'), qn('w:altChunk')}:
        if qn('w:sdt') in tags:
            features.add('content_controls')
        if qn('w:altChunk') in tags:
            features.add('external_content')
    if tags & {qn('w:footnoteReference'), qn('w:endnoteReference')}:
        features.add('note_references')
    tables = list(element.iter(qn('w:tbl')))
    ancestors = [ancestor for ancestor in element.iterancestors() if ancestor.tag == qn('w:tbl')]
    if (element.tag == qn('w:tbl') and len(tables) > 1) or len(ancestors) > 1 or (tables and ancestors):
        features.add('nested_tables')
    if any(table.findall('.//' + qn('w:gridSpan')) or table.findall('.//' + qn('w:vMerge'))
           or table.findall('.//' + qn('w:hMerge')) for table in tables + ancestors):
        features.add('merged_cells')
    if any(ancestor.tag in (qn('w:ins'), qn('w:del'), qn('w:moveFrom'), qn('w:moveTo'), qn('w:sdt'))
           for ancestor in element.iterancestors()):
        features.add('unsupported_container')
    return sorted(features)


def _attributes(element):
    return {key.split('}')[-1]: value for key, value in element.attrib.items()}


def _xml_properties(element):
    if element is None:
        return []
    return [{'name': child.tag.split('}')[-1], 'attributes': _attributes(child),
             'children': _xml_properties(child)} for child in element]


def _part_for_element(doc, element):
    root = element.getroottree().getroot()
    if root is doc.element:
        return doc.part
    for part in doc.part.related_parts.values():
        if getattr(part, 'element', None) is root:
            return part
    raise ValueError('Cannot resolve document part for inspected element')


def _run_report(doc, paragraph, element, index):
    run = Run(element, paragraph)
    report = {'index': index, 'text': _run_text(element), 'bold': run.bold,
              'italic': run.italic, 'underline': None if run.underline is None else int(run.underline),
              'font': run.font.name, 'size_pt': None if run.font.size is None else run.font.size.pt,
              'properties': _xml_properties(element.find(qn('w:rPr')))}
    report['breaks'] = [_attributes(child) for child in element if child.tag == qn('w:br')]
    for parent in element.iterancestors():
        if parent.tag == qn('w:p'):
            break
        if parent.tag == qn('w:hyperlink'):
            rid = parent.get(qn('r:id'))
            link = {'relationship_id': rid, 'anchor': parent.get(qn('w:anchor')),
                    'target': None, 'external': False}
            if rid:
                rel = _part_for_element(doc, element).rels.get(rid)
                if rel is None:
                    link['unresolved'] = True
                else:
                    link['target'] = rel.target_ref
                    link['external'] = rel.is_external
            report['hyperlink'] = link
            break
    return report


def _container_report(identifier, element, blocks):
    allowed = {qn('w:p'), qn('w:tbl'), qn('w:sectPr'), qn('w:tcPr')}
    unexpected = [child for child in element if child.tag not in allowed]
    unsupported = sorted({feature for child in unexpected for feature in unsupported_features(child)})
    if unexpected:
        unsupported = sorted(set(unsupported) | {'unaddressable_blocks'})
    return {'id': identifier,
            'block_ids': [block['id'] for block in blocks if block['container'] == identifier],
            'unsupported': unsupported,
            'unaddressable_elements': [child.tag.split('}')[-1] for child in unexpected],
            'properties': _xml_properties(element.find(qn('w:tcPr')))}


def inspect_layout(data, params):
    """Operation executor: inventory workspace bytes without modifying them."""
    if params:
        raise ValueError('layout.inspect takes no parameters')
    doc = Document(BytesIO(data))
    nodes, containers = map_document(doc)
    blocks = []
    for identifier, element in nodes.items():
        container = identifier.rsplit('/b', 1)[0]
        item = {'id': identifier, 'container': container,
                'type': 'paragraph' if element.tag == qn('w:p') else 'table',
                'unsupported': unsupported_features(element)}
        if item['type'] == 'paragraph':
            paragraph = Paragraph(element, doc)
            item['text'] = paragraph_text(element)
            item['style'] = paragraph.style.name
            item['runs'] = [_run_report(doc, paragraph, run, i)
                            for i, run in enumerate(paragraph_runs(element))]
            item['properties'] = _xml_properties(element.find(qn('w:pPr')))
        else:
            item['text'] = '\n'.join(paragraph_text(p) for p in element.iter(qn('w:p')))
            item['properties'] = _xml_properties(element.find(qn('w:tblPr')))
            item['rows'] = len(element.findall(qn('w:tr')))
            item['grid'] = _xml_properties(element.find(qn('w:tblGrid')))
            item['row_properties'] = [_xml_properties(row.find(qn('w:trPr')))
                                      for row in element.findall(qn('w:tr'))]
        blocks.append(item)
    report = {'schema_version': 1, 'workspace_sha256': sha256(data).hexdigest(),
              'blocks': blocks,
              'containers': [_container_report(identifier, element, blocks)
                             for identifier, element in containers.items()],
              'sections': [{'index': i, 'properties': _xml_properties(section._sectPr)}
                           for i, section in enumerate(doc.sections)]}
    return None, report
