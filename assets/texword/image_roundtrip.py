"""Lossless embedded-picture import shared by block and section synchronization."""
from __future__ import annotations
import copy, os, re
from pathlib import Path
from decimal import Decimal
from lxml import etree as E
from texword import ConversionError
from roundtrip_model import W, NS, sha, field_parts

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{' + NS['r'] + '}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
CAPTION = re.compile(r'^\s*(?:图\s*S?\d+|Figure\s+S?\d+)[\s\u3000.:：、]*', re.I)
MIME = {'image/png': '.png', 'image/jpeg': '.jpeg', 'image/gif': '.gif',
        'image/bmp': '.bmp', 'image/x-ms-bmp': '.bmp', 'image/tiff': '.tiff'}


def text(element):
    return ''.join(n.text or '' for n in element.iter(W + 't'))


def drawings(element):
    return list(element.iter(W + 'drawing'))


def caption(element):
    if element.tag != W + 'p':
        return False
    style = element.find('w:pPr/w:pStyle', NS)
    return bool(CAPTION.match(text(element)) or
                (style is not None and style.get(W + 'val') == 'ImageCaption' and
                 (text(element).strip() or any(kind == 'field' and re.match(r'^SEQ\s+Figure\b', value)
                                              for kind, value, _ in field_parts(element)))))


def empty(element):
    return element.tag == W + 'p' and all(kind == 'text' and not value.strip()
                                         for kind, value, _ in field_parts(element))


def picture_only(element):
    return element.tag == W + 'p' and bool(drawings(element)) and all(
        kind == 'drawing' or (kind == 'text' and not value.strip())
        for kind, value, _ in field_parts(element))


def expand(elements):
    """Unwrap figure layout cells; never flatten an ordinary data table."""
    result = []
    for element in elements:
        if element.tag == W + 'tbl':
            rows = element.findall('w:tr', NS)
            cells = element.findall('w:tr/w:tc', NS)
            if len(rows) == len(cells) == 1:
                children = [e for e in cells[0] if e.tag != W + 'tcPr']
                if (drawings(element) or any(caption(e) for e in children)) and all(e.tag == W + 'p' and (picture_only(e) or caption(e) or empty(e)) for e in children):
                    result.extend(expand(children))
                    continue
        # Word can anchor the following picture in the preceding caption. Keep
        # the preceding image/caption pair together before emitting that anchor.
        anchored = (element.tag == W + 'p' and drawings(element) and
                    all(d.find(WP + 'anchor') is not None for d in drawings(element)))
        if element.tag == W + 'p' and drawings(element) and (caption(element) or anchored):
            prose = copy.deepcopy(element)
            for drawing in drawings(prose):
                drawing.getparent().remove(drawing)
            pictures = []
            for drawing in drawings(element):
                p = E.Element(W + 'p'); r = E.SubElement(p, W + 'r')
                r.append(copy.deepcopy(drawing)); pictures.append(p)
            preceding = next((e for e in reversed(result) if not empty(e)), None)
            result.extend(([prose] + pictures) if not caption(element) or (preceding is not None and picture_only(preceding))
                          else (pictures + [prose]))
        else:
            result.append(element)
    return result


def figure_at(elements, index):
    """Pair standalone pictures with an adjacent caption, ignoring empty carriers."""
    first = elements[index]
    if not (picture_only(first) or caption(first)):
        return None
    j = index + 1
    while j < len(elements) and empty(elements[j]):
        j += 1
    if j < len(elements) and ((picture_only(first) and caption(elements[j])) or
                             (caption(first) and picture_only(elements[j]))):
        return elements[index:j + 1], j + 1 - index
    if picture_only(first):
        return [first], 1
    return None


class ImageImport:
    def __init__(self, doc, source, manifest, destination=None):
        self.doc = doc
        source_root = Path(source).resolve().parent
        self.root = Path(destination).resolve().parent if destination else source_root
        self.manifest = manifest
        self.pending = {}
        self.used = set()
        self.existing = {}
        self.figure_sources = []
        self.reserved = set(re.findall(r'\\label\{([^}]+)\}', manifest['source']))
        for match in re.finditer(r'\\includegraphics(?:\[[^]]*\])?\{([^}]+)\}', manifest['source']):
            path = source_root / match.group(1)
            if path.is_file():
                self.existing.setdefault(sha(path.read_bytes()), Path(os.path.relpath(path, self.root)).as_posix())
        for record in manifest['spans']:
            for unit in record.get('units', [record]):
                if '\\begin{figure}' in unit['latex']:
                    paths = re.findall(r'\\includegraphics(?:\[[^]]*\])?\{([^}]+)\}', unit['latex'])
                    labels = re.findall(r'\\label\{([^}]+)\}', unit['latex'])
                    if paths and len(labels) == 1:
                        path = source_root / paths[0]
                        self.figure_sources.append((labels[0], sha(path.read_bytes()) if path.is_file() else None,
                                                    unit['latex']))

    def asset(self, drawing):
        blips = list(drawing.iter(A + 'blip'))
        if len(blips) != 1 or not blips[0].get(R + 'embed') or blips[0].get(R + 'link'):
            raise ConversionError('Picture must contain one embedded image; linked or compound drawings cannot be imported')
        part = self.doc.part.related_parts.get(blips[0].get(R + 'embed'))
        if part is None or part.content_type not in MIME:
            raise ConversionError('Unsupported embedded picture format; original Word was preserved')
        rects = list(drawing.iter(A + 'srcRect'))
        if len(rects) > 1:
            raise ConversionError('Multiple picture crop rectangles are unsupported')
        crop = [int(rects[0].get(k, '0')) for k in ('l', 't', 'r', 'b')] if rects else [0]*4
        if crop[0] + crop[2] >= 100000 or crop[1] + crop[3] >= 100000:
            raise ConversionError('Picture crop leaves no visible image')
        if any(any(int(n.get(k, '0')) for k in ('rot', 'flipH', 'flipV')) for n in drawing.iter(A + 'xfrm')):
            raise ConversionError('Rotated or flipped pictures require explicit transform support; Word was preserved')
        extent = next(drawing.iter(WP + 'extent'), None)
        if extent is None or any(int(extent.get(k, '0')) <= 0 for k in ('cx', 'cy')):
            raise ConversionError('Picture has no valid displayed dimensions')
        raw = part.blob; digest = sha(raw)
        path = self.existing.get(digest)
        if path is None:
            path = 'figures/word_' + digest + MIME[part.content_type]
            self.pending[path] = raw
        # Decimal conversion retains Word's EMU size well below one EMU.
        def cm(key):
            return format(Decimal(extent.get(key)) / Decimal(360000), '.12f').rstrip('0').rstrip('.')
        return {'path': path, 'sha256': digest, 'width_cm': cm('cx'), 'height_cm': cm('cy'), 'crop': crop}

    def inline(self, drawing):
        image = self.asset(drawing)
        prefix = ('\\twimagecrop{' + ','.join(map(str, image['crop'])) + '}') if any(image['crop']) else ''
        return (prefix + '\\includegraphics[width=' + image['width_cm'] + 'cm,height=' + image['height_cm'] +
                'cm]{' + image['path'] + '}')

    def label(self, parts, images, old):
        labels = set()
        for element in parts:
            for bookmark in element.iter(W + 'bookmarkStart'):
                info = self.manifest.get('references', {}).get(bookmark.get(W + 'name'))
                if info and info['kind'] == 'figure':
                    labels.add(info['key'])
        if len(labels) > 1:
            raise ConversionError('A picture group contains multiple figure labels; separate the figures before syncing')
        if not labels:
            matches = [label for label, digest, latex in self.figure_sources
                       if digest == images[0]['sha256'] and label not in self.used and
                       any(latex == unit.get('latex') for unit in old)]
            if len(matches) == 1:
                labels.add(matches[0])
            else:
                candidates = [label for label, _, latex in self.figure_sources
                              if any(latex == unit.get('latex') for unit in old) and label not in self.used]
                if len(candidates) == 1:
                    labels.add(candidates[0])
        if labels:
            label = labels.pop()
            if label in self.used:
                raise ConversionError('Copied figure bookmark is duplicated; remove the duplicate bookmark before syncing')
        else:
            base = 'fig:word_' + images[0]['sha256'][:16]; label = base; count = 2
            while label in self.reserved or label in self.used:
                label = base + '_' + str(count); count += 1
        self.used.add(label)
        return label

    def figure(self, parts, old, reviews):
        from roundtrip import block_to_latex
        for unit in old:
            if '\\includegraphics' not in unit.get('latex', ''):
                continue
            for match in re.finditer('%', unit['latex']):
                prefix = unit['latex'][:match.start()]
                if (len(prefix) - len(prefix.rstrip('\\'))) % 2 == 0:
                    raise ConversionError('Changed source figure contains LaTeX comments; preserve them in LaTeX before syncing')
        pics = [drawing for element in parts for drawing in drawings(element)]
        images = [self.asset(drawing) for drawing in pics]
        caps = [element for element in parts if caption(element)]
        if len(caps) > 1:
            raise ConversionError('Multiple captions attached to one image group')
        # An uncaptioned newly pasted picture stays unnumbered. Existing figures
        # retain their label and an empty caption when only their caption is removed.
        matched = any(digest == images[0]['sha256'] and any(latex == u.get('latex') for u in old)
                      for _, digest, latex in self.figure_sources)
        if not caps and not matched and not any('\\begin{figure}' in u.get('latex', '') for u in old):
            return '\n'.join(self.inline(drawing) for drawing in pics)
        label = self.label(parts, images, old)
        value = ''
        if caps:
            cap = copy.deepcopy(caps[0])
            # Remove template numbering in its own Word representation, before
            # converting formatted caption prose/citations/comments.
            for node in list(cap.iter(W + 'bookmarkStart')) + list(cap.iter(W + 'bookmarkEnd')):
                node.getparent().remove(node)
            record = {'id': 'figure-caption', 'kind': 'figure_caption', 'latex': '', 'math_sources': [], 'math_baselines': []}
            value = block_to_latex(self.doc, [cap], record, self.manifest, reviews)
        return ('\\begin{figure}[htbp]\n\\centering\n' + '\n'.join(self.inline(d) for d in pics) +
                '\n\\caption{' + value + '}\\label{' + label + '}\n\\end{figure}')

    def inventory(self):
        return [{'path': path, 'sha256': sha(raw), 'bytes': len(raw)} for path, raw in self.pending.items()]

    def validate(self, updated):
        remaining = set(re.findall(r'\\label\{([^}]+)\}', updated))
        referenced = set(re.findall(r'\\(?:eqref|ref)\{([^}]+)\}', updated))
        deleted = {label for label, _, _ in self.figure_sources} - remaining
        dangling = deleted & referenced
        if dangling:
            raise ConversionError('Deleted figures are still referenced: ' + ', '.join(sorted(dangling)) +
                                  '; remove or update those references in Word or LaTeX before syncing')

    def publish(self, destination, temporary):
        """Commit source after lossless assets; roll back only files created here."""
        created = []
        try:
            for relative, raw in self.pending.items():
                path = self.root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                if path.exists():
                    if path.read_bytes() != raw:
                        raise ConversionError('Imported image path is occupied by different bytes: ' + str(path))
                    continue
                try:
                    with path.open('xb') as stream:
                        created.append(path)
                        stream.write(raw)
                except FileExistsError:
                    if path.read_bytes() != raw:
                        raise ConversionError('Imported image changed concurrently: ' + str(path))
            os.replace(temporary, destination)
        except Exception:
            for path in created:
                path.unlink()
            raise
