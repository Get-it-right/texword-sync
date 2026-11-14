"""Content-preserving normalization for TeX layout wrappers.

Normalize only the Pandoc input; the source-map owner retains exact LaTeX.
"""
import re
from source_spans import _next_command, _read_group


def unwrap_layout(source):
    """Expose resizebox content that Pandoc otherwise silently omits."""
    # Pandoc places an abstract in document metadata, outside editable body
    # controls. A quote has the same body content but stays in the block AST.
    # The exact original abstract environment remains in the source map.
    source = source.replace(r'\begin{abstract}', r'\begin{quote}')
    source = source.replace(r'\end{abstract}', r'\end{quote}')
    # tabularx's width and column decorators affect layout, not cell content.
    cursor = 0
    while (match := re.search(r'\\begin\{tabularx\}', source[cursor:])) is not None:
        begin = cursor + match.start()
        position = cursor + match.end()
        width = _read_group(source, position, len(source), '{', '}')
        columns = _read_group(source, width[1], len(source), '{', '}') if width else None
        if columns is None:
            raise ValueError('Malformed tabularx: expected width and column specification')
        start, end = columns
        spec = source[start:end - 1]
        spec = re.sub(r'[>@<]\{[^{}]*\}', '', spec).replace('X', 'l')
        replacement = r'\begin{tabular}{' + spec + '}'
        source = source[:begin] + replacement + source[end:]
        cursor = begin + len(replacement)
    source = source.replace(r'\end{tabularx}', r'\end{tabular}')
    cursor = 0
    while (command := _next_command(source, cursor, len(source))) is not None:
        cursor = command.end
        if command.name == 'addlinespace':
            optional = _read_group(source, cursor, len(source), '[', ']')
            end = optional[1] if optional else cursor
            source = source[:command.start] + source[end:]
            cursor = command.start
            continue
        if command.name == 'multicolumn':
            span = _read_group(source, cursor, len(source), '{', '}')
            columns = _read_group(source, span[1], len(source), '{', '}') if span else None
            if columns is None:
                raise ValueError('Malformed multicolumn')
            start, end = columns
            spec = re.sub(r'@\{[^{}]*\}', '', source[start:end - 1])
            source = source[:start] + spec + source[end - 1:]
            cursor = start + len(spec) + 1
            continue
        if command.name != 'resizebox':
            continue
        position = command.end
        if source[position:position + 1] == '*':
            position += 1
        for _ in range(3):
            group = _read_group(source, position, len(source), '{', '}')
            if group is None:
                raise ValueError('Malformed resizebox: expected dimensions and content')
            start, position = group
        source = source[:command.start] + source[start:position - 1] + source[position:]
        cursor = command.start
    return source
