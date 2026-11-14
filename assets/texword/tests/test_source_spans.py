from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from source_spans import split_source


class SourceSpanTests(unittest.TestCase):
    def test_document_body_spans_keep_exact_slices_and_stable_offsets(self) -> None:
        source = r"""\documentclass{article}
% The preamble must not be included: \begin{equation}
\begin{document}
\section* [Short {] title}] {Header {detail}}% label comment
\label{sec:header}

Paragraph with inline $x^2$ and \verb|\begin{figure}| remains intact.
% \begin{table} and \end{table} here are comments.
Next line has escaped delimiters: \\begin{equation} and \% percent.

\begin{align*}
a &= b + c \\
\begin{aligned}d &= e\end{aligned}
% \end{align*} is ignored.
\end{align*}

Prose \[x+y=z\] continues after display math.

\begin{table*}[ht]
\begin{tabular}{c}value\end{tabular}
\end{table*}

\begin{longtable}{c}generated table\end{longtable}

\begin{figure}
\begin{minipage}{1cm}nested content\end{minipage}
\end{figure}

\begin{vendorblock}
outer
\begin{quote}nested\end{quote}
inner
\end{vendorblock}
\end{document}
"""

        spans = split_source(source)
        kinds = [span["kind"] for span in spans]
        self.assertEqual(
            kinds,
            ["header", "paragraph", "equation", "paragraph", "equation",
             "paragraph", "table", "table", "figure", "environment"],
        )
        self.assertEqual([span["id"] for span in spans],
                         [f"s{i:04d}" for i in range(1, len(spans) + 1)])

        previous_end = 0
        for span in spans:
            start, end = span["start"], span["end"]
            self.assertIsInstance(start, int)
            self.assertIsInstance(end, int)
            self.assertLess(start, end)
            self.assertGreaterEqual(start, previous_end)
            self.assertEqual(span["latex"], source[start:end])
            previous_end = end

        self.assertTrue(all(char.isspace() for char in source[
            source.index(r"\begin{document}") + len(r"\begin{document}"):spans[0]["start"]
        ]))
        self.assertTrue(all(char.isspace() for char in source[
            spans[-1]["end"]:source.index(r"\end{document}")
        ]))
        self.assertNotIn(r"\documentclass", "".join(span["latex"] for span in spans))
        self.assertNotIn(r"\begin{document}", "".join(span["latex"] for span in spans))
        self.assertNotIn(r"\end{document}", "".join(span["latex"] for span in spans))

        header = spans[0]["latex"]
        self.assertIn(r"\section* [Short {] title}] {Header {detail}}", header)
        self.assertIn(r"\label{sec:header}", header)
        self.assertIn("% label comment", header)
        self.assertIn(r"\verb|\begin{figure}|", spans[1]["latex"])
        self.assertIn(r"\\begin{equation}", spans[1]["latex"])
        self.assertIn(r"\begin{aligned}", spans[2]["latex"])
        self.assertIn("Prose", spans[3]["latex"])
        self.assertEqual(spans[4]["latex"], r"\[x+y=z\]")
        self.assertIn("continues after display math", spans[5]["latex"])
        self.assertTrue(spans[6]["latex"].startswith(r"\begin{table*}"))
        self.assertTrue(spans[7]["latex"].startswith(r"\begin{longtable}"))

    def test_adjacent_prose_display_math_and_paragraphs_are_separate(self) -> None:
        source = "before \\[a+b\\] after\n\nnext paragraph with $$c=d$$."
        spans = split_source(source)
        self.assertEqual([span["kind"] for span in spans],
                         ["paragraph", "equation", "paragraph", "paragraph",
                          "equation", "paragraph"])
        self.assertEqual([span["latex"] for span in spans],
                         ["before", r"\[a+b\]", "after", "next paragraph with",
                          "$$c=d$$", "."])
        self.assertEqual(source[spans[0]["end"]:spans[1]["start"]], " ")
        self.assertEqual(source[spans[1]["end"]:spans[2]["start"]], " ")
        self.assertEqual(source[spans[2]["end"]:spans[3]["start"]], "\n\n")
        self.assertEqual(source[spans[3]["end"]:spans[4]["start"]], " ")
        self.assertEqual(source[spans[4]["end"]:spans[5]["start"]], "")

    def test_missing_document_markers_scans_full_input(self) -> None:
        source = "plain text with $x$\n\n\\begin{equation*}x=1\\end{equation*}"
        spans = split_source(source)
        self.assertEqual([span["kind"] for span in spans], ["paragraph", "equation"])
        self.assertEqual(spans[0]["latex"], "plain text with $x$")
        self.assertEqual(spans[1]["latex"], r"\begin{equation*}x=1\end{equation*}")
        self.assertEqual(spans[0]["start"], 0)

    def test_heading_levels_and_equation_environment_variants(self) -> None:
        blocks = [
            r"\subsection[short]{Second level}",
            r"\subsubsection*{Third level}\label{sub:third}",
            *[fr"\begin{{{env}}}x=1\end{{{env}}}"
              for env in ("equation", "equation*", "align", "align*", "gather", "gather*")],
        ]
        spans = split_source("\n\n".join(blocks))
        self.assertEqual([span["kind"] for span in spans],
                         ["header", "header", "equation", "equation", "equation",
                          "equation", "equation", "equation"])
        self.assertIn(r"\label{sub:third}", spans[1]["latex"])

    def test_comments_and_escaped_document_markers_do_not_define_body(self) -> None:
        source = (
            r"% \begin{document}" + "\n"
            + r"\% \\begin{document}" + "\n"
            + "body"
        )
        spans = split_source(source)
        self.assertEqual(len(spans), 1)
        self.assertEqual(spans[0]["start"], 0)
        self.assertEqual(spans[0]["latex"], source)

    def test_comments_and_verbatim_text_do_not_close_nested_environment(self) -> None:
        source = r"""\begin{figure}
% \end{figure}
\begin{verbatim}
\end{figure} is verbatim content
\end{verbatim}
\begin{minipage}{1cm}x\end{minipage}
\end{figure}
after"""
        spans = split_source(source)
        self.assertEqual([span["kind"] for span in spans], ["figure", "paragraph"])
        self.assertIn(r"\end{figure} is verbatim content", spans[0]["latex"])
        self.assertEqual(spans[1]["latex"], "after")


if __name__ == "__main__":
    unittest.main(verbosity=2)
