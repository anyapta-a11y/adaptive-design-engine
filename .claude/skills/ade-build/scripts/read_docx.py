"""DOCX → Markdown-подобный текст с сохранением порядка абзацев и таблиц.
Нужен skill'у, чтобы видеть структуру документа (заголовки, списки, таблицы) дословно."""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit('}', 1)[-1]
        if tag == 'p':
            yield Paragraph(child, doc)
        elif tag == 'tbl':
            yield Table(child, doc)

def render(path):
    out = []
    for b in iter_blocks(Document(path)):
        if isinstance(b, Paragraph):
            text = b.text.strip()
            if not text:
                continue
            style = (b.style.name or '').lower()
            if style in ('title', 'название'):
                out.append('# ' + text)
            elif style.startswith('heading') or style.startswith('заголовок'):
                level = ''.join(ch for ch in style if ch.isdigit()) or '1'
                out.append('#' * int(level) + ' ' + text)
            elif 'list' in style or 'спис' in style or b._p.pPr is not None and b._p.pPr.numPr is not None:
                out.append('- ' + text)
            else:
                out.append(text)
        else:
            rows = [[c.text.strip() for c in r.cells] for r in b.rows]
            if rows:
                out.append('| ' + ' | '.join(rows[0]) + ' |')
                out.append('|' + '---|' * len(rows[0]))
                out += ['| ' + ' | '.join(r) + ' |' for r in rows[1:]]
        out.append('')
    return '\n'.join(out)

if __name__ == '__main__':
    print(render(sys.argv[1]))
