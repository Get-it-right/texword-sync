"""Lexical handling for review annotations embedded in LaTeX source.

Review commands are converted to stable text markers around their source
anchor. The annotation itself is returned as plain data for the caller to
render separately, leaving all unrelated source bytes untouched.
"""
from __future__ import annotations

import hashlib
import re
from typing import Iterator

try:  # Works both as a top-level module and as tools.texword.review_syntax.
    from source_spans import (
        _VERBATIM_ENVS,
        _command_at,
        _find_verbatim_close,
        _read_environment,
        _skip_comment,
        _skip_verb,
    )
except ImportError:  # pragma: no cover - exercised by package-style imports.
    from .source_spans import (
        _VERBATIM_ENVS,
        _command_at,
        _find_verbatim_close,
        _read_environment,
        _skip_comment,
        _skip_verb,
    )


_REVIEW_COMMANDS = {"twcomment": ("comment", 3), "twsuggest": ("suggestion", 4), "twrevision": ("revision", 4)}
_SECTION_COMMANDS = {
    "part", "chapter", "section", "subsection", "subsubsection",
    "paragraph", "subparagraph",
}
_KEY_RE = re.compile(r"[A-Za-z][A-Za-z0-9:_-]*\Z")
_DEFAULT_AUTHOR = "Codex"


def _iter_commands(source: str, start: int, limit: int) -> Iterator[object]:
    """Yield control sequences while skipping comments and literal regions."""
    pos = start
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if char != "\\":
            pos += 1
            continue

        command = _command_at(source, pos, limit)
        if command is None:
            pos += 1
            continue
        if command.name == "verb":
            pos = _skip_verb(source, command, limit)
            continue
        if command.name == "begin":
            environment = _read_environment(source, command, limit)
            if environment is not None and environment.name in _VERBATIM_ENVS:
                close = _find_verbatim_close(
                    source, environment.after_name, environment.name, limit
                )
                pos = limit if close is None else close[1]
                continue
        yield command
        pos = command.end


def _document_body(source: str) -> tuple[int, int]:
    """Find a complete, top-level document body or fall back to all input."""
    pos = 0
    limit = len(source)
    brace_depth = 0
    body_start: int | None = None

    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if char == "\\":
            command = _command_at(source, pos, limit)
            if command is None:
                pos += 1
                continue
            if command.name == "verb":
                pos = _skip_verb(source, command, limit)
                continue
            if command.name in {"begin", "end"}:
                environment = _read_environment(source, command, limit)
                if environment is not None:
                    if environment.name in _VERBATIM_ENVS:
                        close = _find_verbatim_close(
                            source, environment.after_name, environment.name, limit
                        )
                        pos = limit if close is None else close[1]
                        continue
                    if brace_depth == 0 and environment.name == "document":
                        if command.name == "begin" and body_start is None:
                            body_start = environment.after_name
                        elif command.name == "end" and body_start is not None:
                            return body_start, command.start
                    pos = environment.after_name
                    continue
            pos = command.end
            continue
        if char == "{":
            brace_depth += 1
        elif char == "}" and brace_depth:
            brace_depth -= 1
        pos += 1

    return 0, limit


def _skip_space_comments(source: str, pos: int, limit: int) -> int:
    while pos < limit:
        if source[pos].isspace():
            pos += 1
        elif source[pos] == "%":
            pos = _skip_comment(source, pos + 1, limit)
        else:
            break
    return pos


def _read_group(
    source: str, pos: int, limit: int, opening: str, closing: str
) -> tuple[int, int, int] | None:
    """Read a balanced argument; return inner start, inner end, group end."""
    pos = _skip_space_comments(source, pos, limit)
    if pos >= limit or source[pos] != opening:
        return None

    inner_start = pos + 1
    pos += 1
    depth = 1
    brace_depth = 0
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if char == "\\":
            command = _command_at(source, pos, limit)
            if command is not None:
                if command.name == "verb":
                    pos = _skip_verb(source, command, limit)
                else:
                    pos = command.end
                continue
        if opening == "[":
            if char == "{":
                brace_depth += 1
            elif char == "}" and brace_depth:
                brace_depth -= 1
            elif brace_depth == 0 and char == opening:
                depth += 1
            elif brace_depth == 0 and char == closing:
                depth -= 1
                if depth == 0:
                    return inner_start, pos, pos + 1
        elif char == opening:
            depth += 1
        elif char == closing:
            depth -= 1
            if depth == 0:
                return inner_start, pos, pos + 1
        pos += 1
    return None


def _without_comments(value: str) -> str:
    pieces: list[str] = []
    cursor = 0
    limit = len(value)
    while cursor < limit:
        if value[cursor] == "\\":
            command = _command_at(value, cursor, limit)
            if command is not None:
                pieces.append(value[cursor:command.end])
                cursor = command.end
                continue
        if value[cursor] == "%":
            cursor = _skip_comment(value, cursor + 1, limit)
            continue
        pieces.append(value[cursor])
        cursor += 1
    return "".join(pieces)


def _decode_plain(value: str) -> str:
    """Decode plain review text, including the small set of TeX escapes."""
    out: list[str] = []
    pos = 0
    limit = len(value)
    escaped_symbols = {"{": "{", "}": "}", "%": "%", "&": "&",
                       "_": "_", "#": "#", "$": "$"}
    named_escapes = {
        "textbackslash": "\\",
        "textasciicircum": "^",
        "textasciitilde": "~",
    }

    while pos < limit:
        char = value[pos]
        if char == "%":
            pos = _skip_comment(value, pos + 1, limit)
            continue
        if char == "\\":
            command = _command_at(value, pos, limit)
            if command is None:
                pos += 1
                continue
            if command.symbol == "\\":
                out.append("\n")
                pos = command.end
                continue
            if command.symbol in escaped_symbols:
                out.append(escaped_symbols[command.symbol])
                pos = command.end
                continue
            if command.name in named_escapes:
                out.append(named_escapes[command.name])
                pos = command.end
                # encode_plain emits these commands with an empty group.
                if value.startswith("{}", pos):
                    pos += 2
                continue
            raise ValueError(f"plain review text contains unsupported TeX command \\{command.name}")
        if char in "{}":
            # Unescaped braces group TeX input but do not render as text.
            pos += 1
            continue
        out.append(char)
        pos += 1
    return "".join(out)


def _decode_revision_proposal(value: str) -> tuple[str, list[dict[str, object]], list[str]]:
    """Decode a revision proposal with bounded inline citations kept as tokens."""
    parts: list[dict[str, object]] = []
    citations: list[str] = []
    cursor = 0
    for command in _iter_commands(value, 0, len(value)):
        if command.name != "cite":
            continue
        parts.append({"text": _decode_plain(value[cursor:command.start])})
        group = _read_group(value, command.end, len(value), "{", "}")
        if group is None:
            raise ValueError("revision proposal contains malformed \\cite; expected one braced key list")
        keys = [key.strip() for key in value[group[0]:group[1]].split(",")]
        if not keys or any(not _KEY_RE.fullmatch(key) for key in keys):
            raise ValueError("revision proposal contains an invalid \\cite key")
        if len(keys) > 3:
            raise ValueError("Each citation group supports at most three references.")
        parts.append({"citations": keys})
        citations.extend(keys)
        cursor = group[2]
    parts.append({"text": _decode_plain(value[cursor:])})
    display = "".join(
        str(part.get("text", ""))
        if "text" in part
        else "".join(f"[{key}]" for key in part["citations"])
        for part in parts
    )
    return display, parts, citations


def _marker(key: str, edge: str) -> str:
    digest = hashlib.sha256(f"texword-review:{edge}:{key}".encode("utf-8")).hexdigest().upper()
    return f"TWR{edge.upper()}{digest}"


def _validate_anchor(anchor: str, key: str, *, allow_empty: bool = False) -> None:
    visible = _without_comments(anchor)
    if not visible.strip() and not allow_empty:
        raise ValueError(f"review {key!r} has an empty anchor")
    if re.search(r"(?:\r\n|\r|\n)[ \t\f\v]*(?:\r\n|\r|\n)", visible):
        raise ValueError(f"review {key!r} anchor contains a blank paragraph")

    for command in _iter_commands(anchor, 0, len(anchor)):
        if command.name in _REVIEW_COMMANDS:
            raise ValueError(f"review {key!r} contains a nested review macro")
        if command.name in _SECTION_COMMANDS or command.name in {"begin", "end"}:
            raise ValueError(
                f"review {key!r} anchor contains a section or environment command"
            )


def _parse_review(source: str, command: object, limit: int) -> dict[str, object]:
    kind, argument_count = _REVIEW_COMMANDS[command.name]
    cursor = command.end
    probe = _skip_space_comments(source, cursor, limit)
    author = _DEFAULT_AUTHOR
    if probe < limit and source[probe] == "[":
        group = _read_group(source, probe, limit, "[", "]")
        if group is None:
            raise ValueError(f"malformed \\{command.name}: unclosed author argument")
        author = _decode_plain(source[group[0]:group[1]])
        cursor = group[2]

    arguments: list[str] = []
    for index in range(argument_count):
        group = _read_group(source, cursor, limit, "{", "}")
        if group is None:
            raise ValueError(f"malformed \\{command.name}: expected {argument_count} braced arguments")
        arguments.append(source[group[0]:group[1]])
        cursor = group[2]

    key = arguments[0]
    if not _KEY_RE.fullmatch(key):
        raise ValueError(f"malformed \\{command.name}: invalid stable id {key!r}")

    anchor = arguments[1]
    _validate_anchor(anchor, key, allow_empty=kind == "revision")
    for value in [author, *arguments]:
        for nested in _iter_commands(value, 0, len(value)):
            if nested.name in _REVIEW_COMMANDS:
                raise ValueError(f"review {key!r} contains a nested review macro")

    result: dict[str, object] = {
        "key": key,
        "kind": kind,
        "author": author,
        "anchor": anchor,
        "raw": source[command.start:cursor],
        "start": command.start,
        "end": cursor,
        "start_marker": _marker(key, "start"),
        "end_marker": _marker(key, "end"),
    }
    if kind == "comment":
        result["note"] = _decode_plain(arguments[2])
    elif kind == "revision":
        proposal, proposal_parts, proposal_citations = _decode_revision_proposal(arguments[2])
        result["proposal"] = proposal
        result["proposal_parts"] = proposal_parts
        result["proposal_citations"] = proposal_citations
        result["anchor_empty"] = not _without_comments(anchor).strip()
        result["proposal_empty"] = not proposal.strip()
        result["proposal_source"] = arguments[2]
        result["reason"] = _decode_plain(arguments[3])
    else:
        result["proposal"] = _decode_plain(arguments[2])
        result["reason"] = _decode_plain(arguments[3])
    return result


def prepare_reviews(text: str) -> tuple[str, list[dict[str, object]]]:
    """Replace review macros with deterministic markers surrounding anchors.

    Offsets and ``raw`` refer to the original input string. A complete
    ``document`` environment limits discovery to its body; without one, all
    input is considered body text.
    """
    if not isinstance(text, str):
        raise TypeError("text must be a string")

    body_start, body_end = _document_body(text)
    reviews: list[dict[str, object]] = []
    seen: set[str] = set()
    for command in _iter_commands(text, body_start, body_end):
        spec = _REVIEW_COMMANDS.get(command.name)
        if spec is None:
            continue
        review = _parse_review(text, command, body_end)
        key = review["key"]
        if key in seen:
            raise ValueError(f"duplicate review id {key!r}")
        seen.add(key)
        reviews.append(review)

    if not reviews:
        return text, reviews

    output: list[str] = []
    cursor = 0
    for review in reviews:
        start = review["start"]
        end = review["end"]
        output.append(text[cursor:start])
        output.append(review["start_marker"])
        output.append(review["anchor"])
        output.append(review["end_marker"])
        cursor = end
    output.append(text[cursor:])
    return "".join(output), reviews


def format_review(review: dict[str, object]) -> str:
    """Return the plain text displayed for a parsed review annotation."""
    if review.get("kind") in {"suggestion", "revision"}:
        return (
            "待确认的修改建议\n"
            f"建议改为：{review.get('proposal', '')}\n"
            f"理由：{review.get('reason', '')}"
        )
    if review.get("kind") == "comment":
        return str(review.get("note", ""))
    raise ValueError(f"unknown review kind {review.get('kind')!r}")


def encode_plain(text: str) -> str:
    """Escape plain text for safe inclusion as one TeX braced argument."""
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    escapes = {
        "\\": r"\textbackslash{}",
        "{": r"\{",
        "}": r"\}",
        "%": r"\%",
        "&": r"\&",
        "_": r"\_",
        "#": r"\#",
        "$": r"\$",
        "^": r"\textasciicircum{}",
        "~": r"\textasciitilde{}",
    }
    return "".join(escapes.get(char, char) for char in text)
