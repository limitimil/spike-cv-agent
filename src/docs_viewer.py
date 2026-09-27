"""Read-only Markdown pages, confined to the workspace's real docs directory."""
from html import escape
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
from urllib.parse import parse_qs, quote, unquote, urlsplit

from markdown_it import MarkdownIt
from mdit_py_plugins.tasklists import tasklists_plugin


class DocumentUnavailable(ValueError):
    pass


def document_path(root, relative):
    parts = relative.split('/')
    if (not relative or any(p in ('', '.', '..') for p in parts)
            or any(c in relative for c in '\\:\x00')
            or any(p.endswith((' ', '.')) for p in parts)
            or PurePosixPath(relative).suffix.lower() != '.md'):
        raise DocumentUnavailable()
    base = Path(root).resolve() / 'docs'
    current = base
    # Reject reparse aliases, including an aliased docs root, before following them.
    for part in [None] + parts:
        if part is not None:
            current = current / part
        if current.is_symlink() or (hasattr(current, 'is_junction') and current.is_junction()):
            raise DocumentUnavailable()
    resolved = current.resolve()
    if not resolved.is_relative_to(base) or not resolved.is_file():
        raise DocumentUnavailable()
    return resolved


def documents(root):
    base = Path(root).resolve() / 'docs'
    if base.is_symlink() or (hasattr(base, 'is_junction') and base.is_junction()):
        return []
    found = []
    for directory, dirs, files in os.walk(base, followlinks=False):
        dirs[:] = [d for d in dirs if not (Path(directory) / d).is_symlink()
                   and not (hasattr(Path(directory) / d, 'is_junction')
                            and (Path(directory) / d).is_junction())]
        for name in files:
            relative = (Path(directory) / name).relative_to(base).as_posix()
            try:
                document_path(root, relative)
                found.append(relative)
            except (DocumentUnavailable, OSError, ValueError):
                pass
    return sorted(found, key=str.casefold)


def safe_link(root, relative, href):
    try:
        link = urlsplit(href)
        if link.scheme:
            return href if link.scheme.lower() in ('https', 'http', 'mailto') else None
        if link.netloc or link.query:
            return None
        if not link.path:
            return '#' + quote(unquote(link.fragment), safe='-._~')
        path = unquote(link.path, errors='strict')
        if path.startswith('/docs/'):
            target = posixpath.normpath(path[6:])
        elif path.startswith('/'):
            return None
        else:
            target = posixpath.normpath(posixpath.join(posixpath.dirname(relative), path))
        document_path(root, target)
        return '/docs/' + quote(target, safe='/') + (
            '#' + quote(unquote(link.fragment), safe='-._~') if link.fragment else '')
    except (OSError, ValueError, UnicodeError):
        return None


def render_markdown(root, relative, source):
    md = MarkdownIt('js-default').use(tasklists_plugin)
    # Images become alt text; rendering a document never fetches embedded resources.
    def image_text(tokens, idx, options, env):
        return '<span class="image-alt">' + escape(tokens[idx].content) + '</span>'
    md.renderer.rules['image'] = image_text
    tokens = md.parse(source)
    headings, used = [], set()
    for i, token in enumerate(tokens):
        if token.type == 'heading_open':
            label = ''.join(t.content for t in (tokens[i + 1].children or [])
                            if t.type in ('text', 'code_inline', 'image'))
            slug = re.sub(r'[^\w\- ]', '', label.lower()).replace(' ', '-') or 'section'
            anchor, suffix = slug, 1
            while anchor in used:
                anchor = slug + '-' + str(suffix)
                suffix += 1
            used.add(anchor)
            token.attrSet('id', anchor)
            headings.append((anchor, label))
        for child in token.children or []:
            if child.type == 'link_open':
                href = safe_link(root, relative, child.attrGet('href') or '')
                if href is None:
                    child.attrs.pop('href', None)
                    child.attrSet('title', 'Link outside available documents')
                else:
                    child.attrSet('href', href)
                    child.attrSet('rel', 'noreferrer noopener')
    return md.renderer.render(tokens, md.options, {}), headings


def page(root, request_target):
    """Return (status, HTML bytes); never initialize or recover the resume store."""
    request = urlsplit(request_target)
    relative = unquote(request.path.removeprefix('/docs/'), errors='strict')
    if request.path in ('/docs', '/docs/'):
        relative = ''
    query = parse_qs(request.query).get('q', [''])[0][:200]
    names = documents(root)
    status, headings = 200, []
    title = relative or 'Plans & docs'
    if relative:
        try:
            source = document_path(root, relative).read_text(encoding='utf-8-sig')
            article, headings = render_markdown(root, relative, source)
            if not article.strip():
                article = '<p class="muted">This document is empty.</p>'
        except (OSError, ValueError, UnicodeError):
            status, title = 404, 'Document unavailable'
            article = '<h1>Document unavailable</h1><p>Choose a Markdown document from the list.</p>'
    else:
        article = '<h1>Plans & docs</h1><p>Choose a document to read.</p><p class="muted">Reading a plan does not approve it. Decisions remain in the conversation.</p>'
    visible = [name for name in names if query.casefold() in name.casefold()]
    links = ''.join('<li><a ' + ('aria-current="page" ' if name == relative else '')
                    + 'href="/docs/' + quote(name, safe='/') + '">' + escape(name) + '</a></li>'
                    for name in visible)
    if not links:
        links = '<li class="muted">' + ('No matching documents.' if names else 'No Markdown documents in docs/.') + '</li>'
    toc = ''.join('<li><a href="#' + escape(anchor, quote=True) + '">' + escape(label) + '</a></li>'
                  for anchor, label in headings)
    template = Path(__file__).with_name('docs_viewer.html').read_text(encoding='utf-8')
    values = {'title': escape(title), 'document': escape(relative or 'Documentation'),
              'query': escape(query, quote=True), 'links': links, 'article': article,
              'toc': ('<nav class="desktop-toc" aria-label="On this page"><h2>On this page</h2><ul>'
                      + toc + '</ul></nav><details class="mobile-toc"><summary>On this page</summary>'
                      + '<nav aria-label="On this page"><ul>' + toc + '</ul></nav></details>') if toc else ''}
    # Single substitution pass: document text must never become a template directive.
    result = re.sub(r'@@(\w+)@@', lambda m: values[m[1]], template)
    return status, result.encode('utf-8')
