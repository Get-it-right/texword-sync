"""Deterministic, lossless block spans for the body of a LaTeX source file.

The scanner is deliberately lexical rather than a LaTeX parser.  It recognizes
the block forms needed by the Word roundtrip path and returns exact slices of
the caller's original string.  Unrecognized environments remain intact as one
span so a downstream parser can make the final coverage decision.
"""
from __future__ import annotations

from typing import NamedTuple


_HEADER_COMMANDS = {"section", "subsection", "subsubsection"}
_EQUATION_ENVS = {
    "equation", "equation*", "align", "align*", "gather", "gather*",
}
_KNOWN_ENV_KINDS = {
    "table": "table", "table*": "table", "longtable": "table",
    "figure": "figure", "figure*": "figure",
}
_VERBATIM_ENVS = {"verbatim", "verbatim*", "lstlisting", "minted"}
_HORIZONTAL_SPACE = " \t\f\v"


class _Command(NamedTuple):
    start: int
    end: int
    name: str
    symbol: str | None


class _Environment(NamedTuple):
    name: str
    after_name: int


class _Event(NamedTuple):
    start: int
    end: int
    kind: str


def _newline_end(source: str, pos: int, limit: int) -> int | None:
    if pos >= limit:
        return None
    if source[pos] == "\r":
        return pos + 2 if pos + 1 < limit and source[pos + 1] == "\n" else pos + 1
    if source[pos] == "\n":
        return pos + 1
    return None


def _blank_line_end(source: str, pos: int, limit: int) -> int | None:
    """Return the end of a blank-line separator beginning at a line ending."""
    first_end = _newline_end(source, pos, limit)
    if first_end is None:
        return None
    second = first_end
    while second < limit and source[second] in _HORIZONTAL_SPACE:
        second += 1
    return _newline_end(source, second, limit)


def _skip_comment(source: str, pos: int, limit: int) -> int:
    while pos < limit and source[pos] not in "\r\n":
        pos += 1
    # Keep the end-of-line visible: a following empty line is still a
    # paragraph separator, while the comment text itself is never tokenized.
    return pos


def _skip_space_comments(source: str, pos: int, limit: int) -> int:
    while pos < limit:
        if source[pos].isspace():
            pos += 1
        elif source[pos] == "%":
            pos = _skip_comment(source, pos + 1, limit)
        else:
            break
    return pos


def _command_at(source: str, pos: int, limit: int) -> _Command | None:
    """Read one TeX control sequence at a backslash position."""
    if pos >= limit or source[pos] != "\\" or pos + 1 >= limit:
        return None
    end = pos + 1
    if source[end].isalpha():
        end += 1
        while end < limit and source[end].isalpha():
            end += 1
        return _Command(pos, end, source[pos + 1:end], None)
    return _Command(pos, end + 1, "", source[end])


def _next_command(source: str, pos: int, limit: int) -> _Command | None:
    """Find the next unescaped control sequence, ignoring comments."""
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
        elif char == "\\":
            command = _command_at(source, pos, limit)
            if command is not None:
                return command
            pos += 1
        else:
            pos += 1
    return None


def _read_group(source: str, pos: int, limit: int, opening: str, closing: str) -> tuple[int, int] | None:
    """Read a balanced TeX brace/bracket group; return inner start and end."""
    pos = _skip_space_comments(source, pos, limit)
    if pos >= limit or source[pos] != opening:
        return None
    inner_start = pos + 1
    depth = 1
    brace_depth = 0
    pos += 1
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if char == "\\":
            command = _command_at(source, pos, limit)
            if command is not None:
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
                    return inner_start, pos + 1
        elif char == opening:
            depth += 1
        elif char == closing:
            depth -= 1
            if depth == 0:
                return inner_start, pos + 1
        pos += 1
    return None


def _read_environment(source: str, command: _Command, limit: int) -> _Environment | None:
    group = _read_group(source, command.end, limit, "{", "}")
    if group is None:
        return None
    inner_start, after_name = group
    name = source[inner_start:after_name - 1].strip()
    if not name or any(c.isspace() for c in name):
        return None
    return _Environment(name, after_name)


def _skip_verb(source: str, command: _Command, limit: int) -> int:
    """Skip the literal payload of \verb or \verb*."""
    pos = command.end
    if pos < limit and source[pos] == "*":
        pos += 1
    if pos >= limit or source[pos] in "\r\n":
        return pos
    delimiter = source[pos]
    end = source.find(delimiter, pos + 1, limit)
    return limit if end < 0 else end + 1


def _find_verbatim_close(source: str, pos: int, name: str, limit: int) -> tuple[int, int] | None:
    """Find a literal close token inside a verbatim-like environment."""
    marker = r"\end{" + name + "}"
    found = source.find(marker, pos, limit)
    if found < 0:
        return None
    # Verbatim environments end at the literal marker. Other commands,
    # comments, and escaped delimiters inside the payload have no syntax.
    return found, found + len(marker)


def _find_environment_close(source: str, opening: _Environment, pos: int,
                            limit: int) -> tuple[int, int]:
    """Find the matching environment close, respecting nested environments."""
    stack = [opening.name]
    while pos < limit:
        if stack[-1] in _VERBATIM_ENVS:
            close = _find_verbatim_close(source, pos, stack[-1], limit)
            if close is None:
                return limit, limit
            close_start, pos = close
            stack.pop()
            if not stack:
                return close_start, pos
            continue

        command = _next_command(source, pos, limit)
        if command is None:
            return limit, limit
        pos = command.end
        if command.name == "verb":
            pos = _skip_verb(source, command, limit)
            continue
        if command.name not in {"begin", "end"}:
            continue
        env = _read_environment(source, command, limit)
        if env is None:
            continue
        pos = env.after_name
        if command.name == "begin":
            stack.append(env.name)
            continue
        if env.name in stack:
            # Well-formed TeX closes the top entry.  Popping through a
            # mismatched close lets this lexical pass recover its outer edge;
            # downstream parsing remains responsible for rejecting bad TeX.
            matching = len(stack) - 1 - stack[::-1].index(env.name)
            del stack[matching:]
            if not stack:
                return command.start, pos
    return limit, limit


def _find_environment_end(source: str, opening: _Environment, pos: int, limit: int) -> int:
    return _find_environment_close(source, opening, pos, limit)[1]


def _read_header(source: str, command: _Command, limit: int) -> int | None:
    """Read a section command and an immediately following label, if present."""
    pos = command.end
    if pos < limit and source[pos] == "*":
        pos += 1
    pos = _skip_space_comments(source, pos, limit)
    if pos < limit and source[pos] == "[":
        optional = _read_group(source, pos, limit, "[", "]")
        if optional is None:
            return None
        _, pos = optional
    title = _read_group(source, pos, limit, "{", "}")
    if title is None:
        return None
    _, after_title = title

    # A label may follow with spaces or comments between it and the title.
    label_pos = _skip_space_comments(source, after_title, limit)
    label_command = _command_at(source, label_pos, limit) if label_pos < limit else None
    if label_command is not None and label_command.name == "label":
        label = _read_group(source, label_command.end, limit, "{", "}")
        if label is not None:
            return label[1]
    return after_title


def _find_display_math_end(source: str, pos: int, limit: int, close_symbol: str) -> int:
    while pos < limit:
        command = _next_command(source, pos, limit)
        if command is None:
            return limit
        if command.symbol == close_symbol:
            return command.end
        if command.name == "verb":
            pos = _skip_verb(source, command, limit)
        else:
            pos = command.end
    return limit


def _find_dollar_math_end(source: str, pos: int, limit: int, delimiter: str) -> int:
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if char == "\\":
            command = _command_at(source, pos, limit)
            if command is not None:
                pos = command.end
                continue
        if source.startswith(delimiter, pos, limit):
            return pos + len(delimiter)
        pos += 1
    return limit


def _next_block_event(source: str, start: int, limit: int) -> _Event | None:
    """Find the next block opener or paragraph separator after ``start``."""
    pos = start
    brace_depth = 0
    while pos < limit:
        char = source[pos]
        if char == "%":
            pos = _skip_comment(source, pos + 1, limit)
            continue
        if brace_depth == 0 and char in "\r\n":
            separator_end = _blank_line_end(source, pos, limit)
            if separator_end is not None:
                return _Event(pos, separator_end, "separator")
        if char == "\\":
            command = _command_at(source, pos, limit)
            if command is None:
                pos += 1
                continue
            if command.name == "verb":
                pos = _skip_verb(source, command, limit)
                continue
            if brace_depth == 0:
                if command.name in _HEADER_COMMANDS:
                    header_end = _read_header(source, command, limit)
                    if header_end is not None:
                        return _Event(command.start, header_end, "header")
                if command.name == "begin":
                    environment = _read_environment(source, command, limit)
                    if environment is not None:
                        kind = ("equation" if environment.name in _EQUATION_ENVS else
                                _KNOWN_ENV_KINDS.get(environment.name, "environment"))
                        end = _find_environment_end(source, environment, environment.after_name, limit)
                        return _Event(command.start, end, kind)
                if command.symbol == "[":
                    return _Event(command.start,
                                  _find_display_math_end(source, command.end, limit, "]"),
                                  "equation")
                if command.symbol == "(":
                    pos = _find_display_math_end(source, command.end, limit, ")")
                    continue
            pos = command.end
            continue
        if char == "{" and brace_depth >= 0:
            brace_depth += 1
            pos += 1
            continue
        if char == "}" and brace_depth:
            brace_depth -= 1
            pos += 1
            continue
        if brace_depth == 0 and char == "$":
            delimiter = "$$" if source.startswith("$$", pos, limit) else "$"
            close = _find_dollar_math_end(source, pos + len(delimiter), limit, delimiter)
            if delimiter == "$$":
                return _Event(pos, close, "equation")
            pos = close
            continue
        pos += 1
    return None


def _document_body(source: str) -> tuple[int, int]:
    """Return the document body range, or the full input if markers are absent."""
    pos = 0
    limit = len(source)
    while True:
        command = _next_command(source, pos, limit)
        if command is None:
            return 0, limit
        pos = command.end
        if command.name not in {"begin", "end"}:
            continue
        environment = _read_environment(source, command, limit)
        if environment is None or environment.name != "document":
            continue
        if command.name == "begin":
            close_start, _ = _find_environment_close(
                source, environment, environment.after_name, limit)
            if close_start < limit:
                return environment.after_name, close_start
            return 0, limit
        else:
            pos = environment.after_name


def split_source(source_text: str) -> list[dict[str, object]]:
    """Split source into exact, non-overlapping body spans.

    Offsets are Python character offsets into ``source_text`` and ``latex`` is
    always exactly ``source_text[start:end]``. Whitespace between spans stays
    outside them. If a complete document environment is absent, the full input
    is scanned.
    """
    body_start, body_end = _document_body(source_text)
    spans: list[dict[str, object]] = []
    cursor = body_start

    def add(start: int, end: int, kind: str) -> None:
        while start < end and source_text[start].isspace():
            start += 1
        while end > start and source_text[end - 1].isspace():
            end -= 1
        if start >= end:
            return
        spans.append({
            "id": f"s{len(spans) + 1:04d}",
            "start": start,
            "end": end,
            "kind": kind,
            "latex": source_text[start:end],
        })

    while cursor < body_end:
        event = _next_block_event(source_text, cursor, body_end)
        if event is None:
            add(cursor, body_end, "paragraph")
            break
        if event.kind == "separator":
            add(cursor, event.start, "paragraph")
            cursor = event.end
            continue
        add(cursor, event.start, "paragraph")
        add(event.start, event.end, event.kind)
        cursor = event.end

    return spans


__all__ = ["split_source"]
