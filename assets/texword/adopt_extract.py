"""Extract editable semantic blocks from an existing Word document via Pandoc.

The source DOCX is read-only.  A short-lived copy receives paragraph markers
and placeholders for native Word reference fields so a single Pandoc pass can
preserve each original body block and its citation identities.
"""
from __future__ import annotations

import copy
import hashlib
import json
import re
import tempfile
import uuid
import zipfile
from pathlib import Path

from lxml import etree as E

try:  # Package import (``tools.texword.adopt_extract``).
    from .roundtrip_model import NS, W, field_parts
    from .texword import ConversionError, pandoc
except ImportError:  # Script-path import used by the adoption CLI.
    from roundtrip_model import NS, W, field_parts
    from texword import ConversionError, pandoc


M = "{" + NS["m"] + "}"
XML = "http://www.w3.org/XML/1998/namespace"
_FIELD_BEGIN = W + "fldChar"
_INSTR = W + "instrText"
_TEXT = W + "t"
_REF_NAME = re.compile(r"^BIB_[A-Za-z0-9_]+$")
_CAPTION = re.compile(r"^\s*表\s*(\d+)\s*[\u3000 \t]*(.*)$", re.S)
_CHINESE_HEADING = re.compile(r"^[〇零一二三四五六七八九十百]+[、.．]\s*")
_HEADING_NUMBERS = {
    1: re.compile(r"^\d+(?:\.\d+){0}\s*"),
    2: re.compile(r"^\d+(?:\.\d+){1}\s*"),
    3: re.compile(r"^\d+(?:\.\d+){2}\s*"),
}


def _text(element):
    return "".join(element.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS))


def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _style_table(styles):
    result = {}
    for style in styles.findall("w:style", NS):
        sid = style.get(W + "styleId")
        if not sid:
            continue
        name = style.find("w:name", NS)
        parent = style.find("w:basedOn", NS)
        outline = style.find("w:pPr/w:outlineLvl", NS)
        result[sid] = {
            "name": name.get(W + "val", "") if name is not None else "",
            "based_on": parent.get(W + "val") if parent is not None else None,
            "numbered": style.find("w:pPr/w:numPr", NS) is not None,
            "outline": int(outline.get(W + "val")) if outline is not None and outline.get(W + "val", "").isdigit() else None,
        }
    return result


def _heading_has_number(paragraph, level, styles):
    value = _text(paragraph).strip()
    if _HEADING_NUMBERS[level].match(value) or (level == 1 and _CHINESE_HEADING.match(value)):
        return True
    if paragraph.find("w:pPr/w:numPr", NS) is not None:
        return True
    sid = _paragraph_style_id(paragraph)
    seen = set()
    while sid and sid in styles and sid not in seen:
        seen.add(sid)
        if styles[sid].get("numbered"):
            return True
        sid = styles[sid].get("based_on")
    return False


def _paragraph_style_id(paragraph):
    style = paragraph.find("w:pPr/w:pStyle", NS)
    return style.get(W + "val") if style is not None else None


def _style_names(style_id, styles):
    names = []
    seen = set()
    while style_id and style_id in styles and style_id not in seen:
        seen.add(style_id)
        style = styles[style_id]
        names.append(style["name"])
        style_id = style["based_on"]
    return names


def _heading_level(paragraph, styles):
    sid = _paragraph_style_id(paragraph)
    names = " ".join(_style_names(sid, styles))
    # Specific names precede inherited outline levels (af7 is based on af4 in
    # a source template, while its displayed name correctly says level 3).
    for level, terms in (
        (3, ("三级标题", "heading 3", "heading3")),
        (2, ("二级标题", "heading 2", "heading2")),
        (1, ("一级标题", "heading 1", "heading1")),
    ):
        if any(term in names.lower() for term in terms):
            return level
    value = _text(paragraph).strip()
    if _CHINESE_HEADING.match(value):
        return 1
    if re.match(r"^\d+(?:\.\d+){2}(?:\s|\S)", value):
        return 3
    if re.match(r"^\d+\.\d+(?:\s|\S)", value):
        return 2
    style = styles.get(sid, {})
    if style.get("outline") is not None and style["outline"] < 3:
        return style["outline"] + 1
    return None


def _is_bibliography_heading(paragraph):
    value = re.sub(r"\s+", "", _text(paragraph)).casefold()
    return value in ("参考文献", "references", "bibliography")


def _is_payload_paragraph(paragraph):
    if _text(paragraph).strip():
        return True
    return bool(paragraph.xpath(".//w:tab | .//w:br | .//w:drawing | .//w:object | .//m:oMath | .//m:oMathPara", namespaces=NS))


def _top_level_paragraphs(element):
    if element.tag == W + "p":
        return [element]
    if element.tag == W + "tbl":
        return element.xpath(".//w:p", namespaces=NS)
    return []


def _normalized_instruction(value):
    return " ".join((value or "").split())


def _field_spans(paragraph):
    """Return field (child-start, child-end, instruction, cached text) tuples.

    The current source stores each complex-field component in its own run.
    A field sharing a run with unrelated visible text is rejected so placeholder
    replacement cannot silently delete that text.
    """
    children = list(paragraph)
    active = None
    spans = []
    for index, child in enumerate(children):
        if child.tag == W + "fldSimple":
            instruction = child.get(W + "instr", "")
            cached = "".join(child.xpath(".//w:t/text()", namespaces=NS))
            spans.append((index, index, instruction, cached, True))
            continue
        if child.tag != W + "r":
            continue
        for part in child:
            local = E.QName(part).localname
            if part.tag == _FIELD_BEGIN:
                kind = part.get(W + "fldCharType")
                if kind == "begin":
                    if active is not None:
                        raise ConversionError("Nested Word fields cannot be extracted safely.")
                    active = {"start": index, "instruction": [], "cached": [], "phase": "instruction"}
                elif kind == "separate" and active is not None:
                    active["phase"] = "cached"
                elif kind == "end":
                    if active is None:
                        raise ConversionError("Malformed Word field: end without begin.")
                    spans.append((active["start"], index, "".join(active["instruction"]), "".join(active["cached"]), False))
                    active = None
            elif local == "instrText" and active is not None:
                active["instruction"].append(part.text or "")
            elif local == "t" and active is not None and active["phase"] == "cached":
                active["cached"].append(part.text or "")
    if active is not None:
        raise ConversionError("Malformed Word field: missing end.")

    for start, end, instruction, cached, simple in spans:
        for index in range(start, end + 1):
            run = children[index]
            if simple or run.tag != W + "r":
                continue
            # All visible text in the field run is either field instruction or
            # cached result. Runs may retain formatting properties only.
            for part in run:
                local = E.QName(part).localname
                if local not in ("rPr", "fldChar", "instrText", "t"):
                    raise ConversionError(f"Unsupported node inside Word field: {local}.")
    return spans


def _field_parts(paragraph, source_index):
    try:
        parts = field_parts(paragraph)
        parsed = [value for kind, value, _ in parts if kind == "field"]
        math_count = sum(kind == "math" for kind, _, _ in parts)
        spans = _field_spans(paragraph)
    except Exception as error:
        if isinstance(error, ConversionError):
            raise
        raise ConversionError(f"Could not parse Word fields in body index {source_index}: {error}") from error
    if [_normalized_instruction(item) for item in parsed] != [_normalized_instruction(item[2]) for item in spans]:
        raise ConversionError(f"Word field parser disagreement at body index {source_index}.")
    return spans, math_count


def _set_run_text(run, value):
    node = E.SubElement(run, _TEXT)
    node.set("{" + XML + "}space", "preserve")
    node.text = value


def _replace_field(paragraph, span, value):
    start, end = span[0], span[1]
    children = list(paragraph)
    replacement = E.Element(W + "r")
    _set_run_text(replacement, value)
    for child in children[start : end + 1]:
        paragraph.remove(child)
    paragraph.insert(start, replacement)


def _strip_field_brackets(paragraph, replacements):
    """Remove the visible [number] wrappers surrounding REF/SEQ fields.

    The Word citation number itself is cached inside the field.  Removing the
    adjacent bracket characters before Pandoc import prevents them from being
    escaped as literal punctuation around the later ``\\cite`` command.
    """
    children = list(paragraph)
    spans = sorted((item[0] for item in replacements), key=lambda span: span[0])
    if not spans:
        return
    by_start = {span[0]: span for span in spans}
    events = []
    index = 0
    while index < len(children):
        span = by_start.get(index)
        if span is not None:
            events.append(("field", span, None))
            index = span[1] + 1
            continue
        child = children[index]
        for text_node in child.xpath(".//w:t", namespaces=NS):
            events.append(("text", text_node, text_node.text or ""))
        index += 1

    offsets = {}
    for event_index, event in enumerate(events):
        if event[0] != "field":
            continue
        previous = next((events[j] for j in range(event_index - 1, -1, -1) if events[j][0] == "text" and events[j][2].strip()), None)
        following = next((events[j] for j in range(event_index + 1, len(events)) if events[j][0] == "text" and events[j][2].strip()), None)
        if previous is not None:
            value = previous[2]
            position = len(value.rstrip()) - 1
            if position >= 0 and value[position] == "[":
                offsets.setdefault(previous[1], set()).add(position)
        if following is not None:
            value = following[2]
            position = len(value) - len(value.lstrip())
            if position < len(value) and value[position] == "]":
                offsets.setdefault(following[1], set()).add(position)
    for text_node, positions in offsets.items():
        value = text_node.text or ""
        for position in sorted(positions, reverse=True):
            value = value[:position] + value[position + 1 :]
        text_node.text = value


def _bookmark_names(paragraph):
    return [node.get(W + "name", "") for node in paragraph.findall("w:bookmarkStart", NS)]


def _marker_paragraph(token):
    paragraph = E.Element(W + "p")
    run = E.SubElement(paragraph, W + "r")
    _set_run_text(run, token)
    return paragraph


def _replace_xml_part(source, target, document_xml):
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            payload = document_xml if info.filename == "word/document.xml" else zin.read(info.filename)
            zout.writestr(info, payload)


def _inline_text(inlines):
    out = []
    if isinstance(inlines, list):
        for item in inlines:
            if isinstance(item, dict):
                kind = item.get("t")
                value = item.get("c")
                if kind in ("Str", "Code") and isinstance(value, str):
                    out.append(value)
                elif kind == "Space":
                    out.append(" ")
                elif kind in ("SoftBreak", "LineBreak"):
                    out.append("\n")
                else:
                    out.append(_inline_text(value))
            elif isinstance(item, list):
                out.append(_inline_text(item))
    return "".join(out)


def _block_text(block):
    if not isinstance(block, dict):
        return ""
    return _inline_text(block.get("c"))


def _find_marker_blocks(blocks, marker_tokens):
    positions = {}
    for index, block in enumerate(blocks):
        for token in marker_tokens:
            if token in _block_text(block):
                if token in positions:
                    raise ConversionError(f"Pandoc duplicated extraction marker {token}.")
                positions[token] = index
    missing = sorted(set(marker_tokens) - set(positions))
    if missing:
        raise ConversionError("Pandoc omitted extraction markers: " + ", ".join(missing[:8]))
    return positions


def _leaf_strs(value):
    if isinstance(value, dict):
        if value.get("t") == "Str" and isinstance(value.get("c"), str):
            yield value
            return
        yield from _leaf_strs(value.get("c"))
    elif isinstance(value, list):
        for item in value:
            yield from _leaf_strs(item)


def _strip_prefix_inlines(inlines, pattern):
    """Remove one known leading numbering prefix while retaining inline style."""
    for leaf in _leaf_strs(inlines):
        new, count = pattern.subn("", leaf["c"], count=1)
        if count:
            leaf["c"] = new
            break
        if leaf["c"].strip():
            break
    _drop_leading_spaces(inlines)


def _drop_leading_spaces(inlines):
    if not isinstance(inlines, list):
        return
    while inlines and isinstance(inlines[0], dict) and inlines[0].get("t") in ("Space", "SoftBreak"):
        inlines.pop(0)
    if inlines and isinstance(inlines[0], dict) and inlines[0].get("t") in ("Strong", "Emph", "Underline", "SmallCaps"):
        _drop_leading_spaces(inlines[0].get("c"))


def _heading_ast(segment, level, unnumbered=False):
    paragraph = next((item for item in segment if item.get("t") in ("Para", "Plain", "Header")), None)
    if paragraph is None:
        raise ConversionError("A recognized heading did not produce a Pandoc paragraph.")
    # Pandoc recognizes built-in Word heading styles as Header, whereas
    # visually styled headings may arrive as ordinary paragraphs.
    inlines = copy.deepcopy(paragraph["c"][2] if paragraph.get("t") == "Header" else paragraph.get("c", []))
    if not unnumbered:
        pattern = _CHINESE_HEADING if level == 1 else _HEADING_NUMBERS[level]
        _strip_prefix_inlines(inlines, pattern)
    attr = ["", ["unnumbered"], []] if unnumbered else ["", [], []]
    return [{"t": "Header", "c": [level, attr, inlines]}]


def _clean_str(value, cite_tokens):
    """Split a Pandoc Str around citation placeholders and emit raw \cite nodes."""
    token_pattern = re.compile("|".join(re.escape(token) for token in sorted(cite_tokens, key=len, reverse=True)))
    if not cite_tokens or not token_pattern.search(value):
        return [{"t": "Str", "c": value}]

    def parse_group(inner):
        keys = []
        position = 0
        separator = re.compile(r"[\s,;，、/\-–—]*")
        while position < len(inner):
            while position < len(inner) and inner[position].isspace():
                position += 1
            match = token_pattern.match(inner, position)
            if match is None:
                return None
            keys.append(cite_tokens[match.group(0)])
            position = match.end()
            if position == len(inner):
                break
            gap = separator.match(inner, position)
            if gap is None or gap.end() == position:
                return None
            position = gap.end()
        return keys or None

    result = []
    position = 0
    cursor = 0
    while cursor < len(value):
        if value[cursor] == "[":
            close = value.find("]", cursor + 1)
            if close >= 0:
                keys = parse_group(value[cursor + 1 : close])
                if keys:
                    if cursor > position:
                        result.append({"t": "Str", "c": value[position:cursor]})
                    result.append({"t": "RawInline", "c": ["latex", r"\cite{" + ",".join(keys) + "}"]})
                    cursor = close + 1
                    position = cursor
                    continue
        match = token_pattern.search(value, cursor)
        if match is None:
            break
        if match.start() > position:
            result.append({"t": "Str", "c": value[position : match.start()]})
        key = cite_tokens[match.group(0)]
        result.append({"t": "RawInline", "c": ["latex", r"\cite{" + key + "}"]})
        cursor = match.end()
        position = cursor
    if position < len(value):
        result.append({"t": "Str", "c": value[position:]})
    return [item for item in result if item.get("c") not in ("", [])]


def _transform_citations(value, cite_tokens):
    if isinstance(value, list):
        result = []
        for item in value:
            if isinstance(item, dict) and item.get("t") == "Str" and isinstance(item.get("c"), str):
                result.extend(_clean_str(item["c"], cite_tokens))
            else:
                result.append(_transform_citations(item, cite_tokens))
        return result
    if isinstance(value, dict):
        return {key: _transform_citations(item, cite_tokens) for key, item in value.items()}
    return value


def _remove_sequence_token(value, token):
    pattern = re.compile(r"\[\s*" + re.escape(token) + r"\s*\]\s*")
    found = False

    def visit(node):
        nonlocal found
        if isinstance(node, dict):
            if node.get("t") == "Str" and isinstance(node.get("c"), str) and token in node["c"]:
                if found:
                    raise ConversionError(f"Pandoc duplicated bibliography field token {token}.")
                new, n = pattern.subn("", node["c"], count=1)
                if n != 1:
                    new = node["c"].replace(token, "", 1)
                node["c"] = new
                found = True
            else:
                for child in node.values():
                    visit(child)
        elif isinstance(node, list):
            for child in node:
                visit(child)

    visit(value)
    if not found:
        raise ConversionError(f"Pandoc omitted bibliography field token {token}.")


def _set_table_caption(segment, caption_inlines, label):
    tables = []
    for block in segment:
        if isinstance(block, dict) and block.get("t") == "Table":
            tables.append(block)
    if len(tables) != 1:
        raise ConversionError(f"Captioned source unit produced {len(tables)} Pandoc tables; expected one.")
    table = tables[0]
    content = copy.deepcopy(caption_inlines)
    content.append({"t": "RawInline", "c": ["latex", r"\label{" + label + "}"]})
    if len(table.get("c", [])) < 2:
        raise ConversionError("Unsupported Pandoc table AST without a caption slot.")
    # Pandoc 2.12's Table caption is [short-caption, long-caption-blocks].
    table["c"][1] = [None, [{"t": "Plain", "c": content}]]
    return [block for block in segment if not (isinstance(block, dict) and block.get("t") in ("Para", "Plain") and _block_text(block).lstrip().startswith("表"))]


def _token_occurrences(value, tokens):
    counts = {token: 0 for token in tokens}
    if isinstance(value, dict):
        if value.get("t") == "Str" and isinstance(value.get("c"), str):
            for token in tokens:
                counts[token] += value["c"].count(token)
        else:
            for child in value.values():
                nested = _token_occurrences(child, tokens)
                for token in tokens:
                    counts[token] += nested[token]
    elif isinstance(value, list):
        for child in value:
            nested = _token_occurrences(child, tokens)
            for token in tokens:
                counts[token] += nested[token]
    return counts


def _replace_segments(blocks, descriptors, cite_tokens, seq_tokens):
    marker_tokens = [token for descriptor in descriptors for token in (descriptor["begin"], descriptor["end"])]
    positions = _find_marker_blocks(blocks, marker_tokens)
    ordered = sorted(descriptors, key=lambda item: positions[item["begin"]])
    result = []
    cursor = 0
    segment_map = {}
    for descriptor in ordered:
        first = positions[descriptor["begin"]]
        last = positions[descriptor["end"]]
        if first >= last or first < cursor:
            raise ConversionError("Extraction markers are nested or out of source order.")
        result.extend(blocks[cursor : first + 1])
        segment = copy.deepcopy(blocks[first + 1 : last])
        if not segment:
            raise ConversionError(f"Pandoc returned empty content for original index {descriptor['indices'][0]}.")
        counts = _token_occurrences(segment, cite_tokens)
        for token, count in counts.items():
            if count and count != 1:
                raise ConversionError(f"Pandoc duplicated citation field token {token}.")
        segment = _transform_citations(segment, cite_tokens)

        if descriptor["kind"] == "heading":
            segment = _heading_ast(segment, descriptor["level"], descriptor.get("unnumbered", False))
        elif descriptor["kind"] == "captioned_table":
            caption_node = next((item for item in segment if item.get("t") in ("Para", "Plain") and _CAPTION.match(_block_text(item))), None)
            if caption_node is None:
                raise ConversionError(f"Pandoc omitted the caption at body index {descriptor['indices'][0]}.")
            caption = copy.deepcopy(caption_node.get("c", []))
            _strip_prefix_inlines(caption, re.compile(r"^表\s*\d+\s*[\u3000 \t]*"))
            segment = _set_table_caption(segment, caption, descriptor["label"])
        elif descriptor["kind"] == "bibliography_item":
            _remove_sequence_token(segment, descriptor["sequence_token"])

        result.extend(segment)
        result.append(blocks[last])
        cursor = last + 1
        segment_map[descriptor["id"]] = (descriptor, segment)
    result.extend(blocks[cursor:])

    # Every placeholder must survive DOCX parsing exactly once before it is
    # rewritten into a citation or removed from its bibliography label.
    counts = _token_occurrences(blocks, list(cite_tokens) + list(seq_tokens))
    for token in list(cite_tokens) + list(seq_tokens):
        if counts[token] != 1:
            raise ConversionError(f"Pandoc preserved {counts[token]} copies of field placeholder {token}; expected one.")
    return result, segment_map


def _render_segments(rendered, descriptors):
    result = {}
    for descriptor in descriptors:
        begin = descriptor["begin"]
        end = descriptor["end"]
        first = rendered.find(begin)
        last = rendered.find(end, first + len(begin)) if first >= 0 else -1
        if first < 0 or last < 0:
            raise ConversionError(f"Pandoc LaTeX output lost block markers for original index {descriptor['indices'][0]}.")
        if rendered.find(begin, first + len(begin)) >= 0:
            raise ConversionError(f"Pandoc duplicated LaTeX marker {begin}.")
        result[descriptor["id"]] = rendered[first + len(begin) : last].strip()
    return result


def extract_blocks(docx, body_start, body_end, workdir):
    """Return semantic LaTeX blocks for the half-open top-level DOCX range.

    ``body_start`` and ``body_end`` are indices into ``w:body`` (the end is
    exclusive). The result is a registry with ``blocks``, a BIB-keyed
    ``bibliography`` map, and the SHA-256 of the untouched source DOCX.
    """
    source = Path(docx).resolve()
    workspace = Path(workdir).resolve()
    if not source.is_file():
        raise ConversionError(f"Source DOCX does not exist: {source}")
    if not workspace.is_dir():
        raise ConversionError(f"Extraction workdir must already exist: {workspace}")
    if not isinstance(body_start, int) or not isinstance(body_end, int) or body_start < 0 or body_start >= body_end:
        raise ConversionError("DOCX body range must be a non-empty half-open integer range.")

    source_digest = _sha256(source)
    with zipfile.ZipFile(source, "r") as archive:
        if "word/document.xml" not in archive.namelist():
            raise ConversionError("Source DOCX has no word/document.xml part.")
        document = E.fromstring(archive.read("word/document.xml"))
        styles = E.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else E.Element("styles")
        original_children = list(document.find("w:body", NS))
        style_table = _style_table(styles)
    if body_end > len(original_children):
        raise ConversionError(f"Body range ends at {body_end}, but the DOCX body has {len(original_children)} children.")

    body = document.find("w:body", NS)
    seed = uuid.uuid4().hex[:12]
    serial = 0
    cite_tokens = {}
    sequence_tokens = {}
    ref_items = []
    bibliography = {}
    field_replacements = {}
    top_text = {}
    heading_info = {}
    native_math_nodes = 0
    for index in range(body_start, body_end):
        element = original_children[index]
        if element.tag not in (W + "p", W + "tbl"):
            raise ConversionError(f"Unsupported top-level Word node at body index {index}: {E.QName(element).localname}.")
        top_text[index] = _text(element)
        paragraphs = _top_level_paragraphs(element)
        for paragraph in paragraphs:
            try:
                spans, paragraph_math_count = _field_parts(paragraph, index)
                native_math_nodes += paragraph_math_count
            except ConversionError as error:
                raise ConversionError(f"Body index {index}: {error}") from error
            for span in spans:
                instruction = _normalized_instruction(span[2])
                words = instruction.split()
                if len(words) >= 2 and words[0].upper() == "REF" and _REF_NAME.fullmatch(words[1]):
                    serial += 1
                    token = f"TxCiteA{seed}{serial:06d}Q"
                    cite_tokens[token] = words[1]
                    field_replacements.setdefault(paragraph, []).append((span, token, "cite", words[1]))
                elif len(words) >= 2 and words[0].upper() == "SEQ" and words[1] == "AblBibRef":
                    # Sequence fields belong only to bibliography entries.
                    serial += 1
                    token = f"TxSeqA{seed}{serial:06d}Q"
                    sequence_tokens[token] = span[3]
                    field_replacements.setdefault(paragraph, []).append((span, token, "sequence", None))
                else:
                    raise ConversionError(f"Unsupported Word field at body index {index}: {instruction!r}.")

        if element.tag == W + "p":
            level = _heading_level(element, style_table)
            if level is not None:
                heading_info[index] = level
            sequence_fields = [item for item in _field_spans(element) if _normalized_instruction(item[2]).split()[:2] == ["SEQ", "AblBibRef"]]
            if sequence_fields:
                keys = [key for key in _bookmark_names(element) if _REF_NAME.fullmatch(key)]
                if len(keys) != 1:
                    raise ConversionError(f"Bibliography item at body index {index} must have exactly one BIB_ bookmark; found {keys}.")
                key = keys[0]
                number_text = sequence_fields[0][3].strip()
                if not number_text.isdigit():
                    raise ConversionError(f"Bibliography SEQ field at body index {index} has non-numeric cached value {number_text!r}.")
                match = re.match(r"^\s*\[\s*(\d+)\s*\]\s*(.*)$", _text(element), re.S)
                if match is None:
                    raise ConversionError(f"Bibliography item at body index {index} has no cached [number] prefix.")
                if int(match.group(1)) != int(number_text):
                    raise ConversionError(f"Bibliography number mismatch at body index {index}: field={number_text}, text={match.group(1)}.")
                if key in bibliography:
                    raise ConversionError(f"Duplicate bibliography bookmark: {key}.")
                item_text = match.group(2)
                bibliography[key] = {"number": int(number_text), "text": item_text, "original_index": index}
                ref_items.append({"index": index, "key": key, "number": int(number_text), "text": item_text})

    # Body citations must resolve to the document's bibliography registry.
    unknown_keys = sorted(set(cite_tokens.values()) - set(bibliography))
    if unknown_keys:
        raise ConversionError("Unresolved body citation keys: " + ", ".join(unknown_keys[:12]))
    if ref_items:
        numbers = [item["number"] for item in ref_items]
        if numbers != list(range(1, len(numbers) + 1)):
            raise ConversionError("Bibliography SEQ cached numbers are not one contiguous sequence starting at 1.")
        indices = [item["index"] for item in ref_items]
        if indices != list(range(indices[0], indices[-1] + 1)):
            raise ConversionError("Bibliography entries contain non-reference nodes between items.")

    for paragraph, replacements in field_replacements.items():
        _strip_field_brackets(paragraph, replacements)
        for span, token, kind, key in sorted(replacements, key=lambda item: item[0][0], reverse=True):
            _replace_field(paragraph, span, token)

    # Build ordered extraction units. Caption+table pairs and the complete
    # bibliography each stay a single main block with their original indices.
    units = []
    bib_index_set = {item["index"] for item in ref_items}
    bib_group = None
    if ref_items:
        bib_group = {
            "kind": "bibliography",
            "indices": [item["index"] for item in ref_items],
            "items": ref_items,
            "id": "bibliography",
        }

    index = body_start
    while index < body_end:
        element = original_children[index]
        if index in bib_index_set:
            if index == ref_items[0]["index"]:
                units.append(bib_group)
            index += 1
            continue
        if element.tag == W + "p" and _is_bibliography_heading(element):
            unit = {"kind": "heading", "indices": [index], "level": 1, "unnumbered": True, "id": f"block{len(units):05d}"}
            units.append(unit)
            index += 1
            continue
        if element.tag == W + "p":
            caption = _CAPTION.match(_text(element))
            if caption and index + 1 < body_end and original_children[index + 1].tag == W + "tbl":
                units.append({
                    "kind": "captioned_table",
                    "indices": [index, index + 1],
                    "label": f"tab:{caption.group(1)}",
                    "table_number": int(caption.group(1)),
                    "id": f"block{len(units):05d}",
                })
                index += 2
                continue
            if not _is_payload_paragraph(element):
                units.append({"kind": "layout", "indices": [index], "latex": "", "source_text": "", "reason": "blank paragraph or section break", "id": f"block{len(units):05d}"})
            elif index in heading_info:
                level = heading_info[index]
                units.append({"kind": "heading", "indices": [index], "level": level,
                              "unnumbered": not _heading_has_number(element, level, styles),
                              "id": f"block{len(units):05d}"})
            else:
                units.append({"kind": "paragraph", "indices": [index], "id": f"block{len(units):05d}"})
            index += 1
            continue
        if element.tag == W + "tbl":
            units.append({"kind": "table", "indices": [index], "id": f"block{len(units):05d}"})
            index += 1
            continue
        raise ConversionError(f"Unsupported Word node at body index {index}: {E.QName(element).localname}.")

    # Add one begin/end marker pair per editable source unit. The bibliography
    # has internal pairs per reference, then returns as one semantic block.
    marked = []
    for unit in units:
        if unit["kind"] == "layout":
            continue
        if unit["kind"] == "bibliography":
            item_descriptors = []
            for item in unit["items"]:
                serial += 1
                item_id = f"bib{serial:06d}"
                begin = f"TxBeginA{seed}{serial:06d}Q"
                end = f"TxEndA{seed}{serial:06d}Q"
                seq_token = next(token for token, number in sequence_tokens.items() if int(number) == item["number"] and token not in {x["sequence_token"] for x in item_descriptors})
                descriptor = {"kind": "bibliography_item", "indices": [item["index"]], "id": item_id, "begin": begin, "end": end, "sequence_token": seq_token, "key": item["key"], "number": item["number"]}
                item_descriptors.append(descriptor)
                marked.append(descriptor)
            unit["item_descriptors"] = item_descriptors
            continue
        serial += 1
        unit["begin"] = f"TxBeginA{seed}{serial:06d}Q"
        unit["end"] = f"TxEndA{seed}{serial:06d}Q"
        marked.append(unit)

    starts = {}
    ends = {}
    for unit in marked:
        first_index = unit["indices"][0]
        last_index = unit["indices"][-1]
        starts.setdefault(first_index, []).append(unit["begin"])
        ends.setdefault(last_index, []).append(unit["end"])
    rewritten_children = []
    for child_index, child in enumerate(original_children):
        for token in starts.get(child_index, []):
            rewritten_children.append(_marker_paragraph(token))
        rewritten_children.append(child)
        for token in ends.get(child_index, []):
            rewritten_children.append(_marker_paragraph(token))
    for child in list(body):
        body.remove(child)
    for child in rewritten_children:
        body.append(child)
    xml = E.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)

    with tempfile.TemporaryDirectory(prefix=".adopt-extract-", dir=str(workspace)) as temporary:
        temp_docx = Path(temporary) / "source-marked.docx"
        _replace_xml_part(source, temp_docx, xml)
        raw_json, _ = pandoc(["-f", "docx", "-t", "json", temp_docx])
        ast = json.loads(raw_json)
        descriptors = marked
        ast_blocks, segment_ast = _replace_segments(ast["blocks"], descriptors, cite_tokens, sequence_tokens)
        ast["blocks"] = ast_blocks
        rendered, _ = pandoc(["-f", "json", "-t", "latex", "--wrap=none"], input=json.dumps(ast, ensure_ascii=False))
        rendered_segments = _render_segments(rendered, descriptors)

    blocks = []
    by_id = {descriptor["id"]: descriptor for descriptor in marked}
    for unit in units:
        if unit["kind"] == "layout":
            block = {
                "kind": "layout",
                "latex": "",
                "original_indices": list(unit["indices"]),
                "source_text": unit.get("source_text", ""),
                "reason": unit.get("reason", "blank paragraph or section break"),
            }
            blocks.append(block)
            continue
        if unit["kind"] == "bibliography":
            entries = []
            for item_descriptor in unit["item_descriptors"]:
                item = next(x for x in unit["items"] if x["index"] == item_descriptor["indices"][0])
                latex_text = rendered_segments[item_descriptor["id"]]
                if not latex_text:
                    raise ConversionError(f"Pandoc produced empty bibliography text at body index {item['index']}.")
                entries.append(f"\\bibitem{{{item['key']}}}\n{latex_text}")
            latex = "\\begin{thebibliography}{999}\n" + "\n\n".join(entries) + "\n\\end{thebibliography}"
            blocks.append({
                "kind": "bibliography",
                "latex": latex,
                "original_indices": list(unit["indices"]),
                "reference_keys": [item["key"] for item in unit["items"]],
                "reference_count": len(unit["items"]),
                "source_text": "\n".join(top_text[index] for index in unit["indices"]),
            })
            continue
        latex = rendered_segments[unit["id"]]
        if not latex.strip():
            raise ConversionError(f"Pandoc produced empty LaTeX for source indices {unit['indices']}.")
        source_text = "\n".join(top_text[index] for index in unit["indices"])
        if unit["kind"] == "captioned_table":
            if not re.search(r"\\begin\{(?:longtable|tabular|table)\}", latex) or not re.search(r"\\label\{" + re.escape(unit["label"]) + r"\}", latex):
                raise ConversionError(f"Captioned table {unit['label']} did not retain one LaTeX table and its label.")
            block = {
                "kind": "table",
                "latex": latex,
                "original_indices": list(unit["indices"]),
                "table_number": unit["table_number"],
                "label": unit["label"],
                "source_text": source_text,
            }
        else:
            public_kind = "header" if unit["kind"] == "heading" else unit["kind"]
            block = {
                "kind": public_kind,
                "latex": latex,
                "original_indices": list(unit["indices"]),
                "source_text": source_text,
            }
            if unit["kind"] == "heading":
                block["level"] = unit["level"]
                if unit.get("unnumbered"):
                    block["unnumbered"] = True
        blocks.append(block)

    flattened = [index for block in blocks for index in block["original_indices"]]
    expected = list(range(body_start, body_end))
    if sorted(flattened) != expected or len(flattened) != len(set(flattened)):
        raise ConversionError("Extraction blocks do not partition the requested source body range exactly once.")
    if any("\\begin{thebibliography}{999}" in block.get("latex", "") for block in blocks if block["kind"] != "bibliography"):
        raise ConversionError("Bibliography environment escaped its grouped bibliography block.")

    return {
        "blocks": blocks,
        "bibliography": bibliography,
        "source_sha256": source_digest,
        "body_start": body_start,
        "body_end": body_end,
        "source_path": str(source),
        "counts": {
            "source_top_level_nodes": body_end - body_start,
            "blocks": len(blocks),
            "layout_blocks": sum(block["kind"] == "layout" for block in blocks),
            "tables": sum(block["kind"] == "table" for block in blocks),
            "references": len(bibliography),
            "native_citations": len(cite_tokens),
            "native_math_nodes": native_math_nodes,
        },
    }
