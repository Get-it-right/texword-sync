from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from review_syntax import encode_plain, format_review, prepare_reviews


class ReviewSyntaxTests(unittest.TestCase):
    def test_nested_braces_escaped_delimiters_and_author_are_parsed(self) -> None:
        source = (
            r"before \twcomment[Ana \& Li]{note:1}"
            r"{\textbf{Anchor {detail}} and $x_{i}$}"
            r"{Keep \{braces\}, 50\%, A\&B, x\_y, \#tag, and \$3} after"
        )

        prepared, reviews = prepare_reviews(source)

        self.assertEqual(len(reviews), 1)
        review = reviews[0]
        self.assertEqual(review["key"], "note:1")
        self.assertEqual(review["kind"], "comment")
        self.assertEqual(review["author"], "Ana & Li")
        self.assertEqual(review["anchor"], r"\textbf{Anchor {detail}} and $x_{i}$")
        self.assertEqual(review["note"], "Keep {braces}, 50%, A&B, x_y, #tag, and $3")
        self.assertEqual(review["raw"], source[source.index(r"\twcomment"):source.index(" after")])
        self.assertEqual(review["start"], source.index(r"\twcomment"))
        self.assertEqual(review["end"], review["start"] + len(review["raw"]))
        self.assertIn(review["start_marker"] + review["anchor"] + review["end_marker"], prepared)
        self.assertNotIn("Keep", prepared)
        self.assertEqual(format_review(review), review["note"])
        self.assertTrue(review["start_marker"].isalnum())
        self.assertTrue(review["end_marker"].isalnum())

    def test_suggestion_format_and_markers_are_deterministic_by_key(self) -> None:
        first = r"\twsuggest{stable_2}{\emph{old}}{new \& improved}{clearer wording}"
        second = r"\twcomment{other}{anchor}{note}"
        _, one = prepare_reviews(first)
        _, two = prepare_reviews(second + "\n" + first)

        suggestion_one = one[0]
        suggestion_two = two[1]
        self.assertEqual(suggestion_one["start_marker"], suggestion_two["start_marker"])
        self.assertEqual(suggestion_one["end_marker"], suggestion_two["end_marker"])
        self.assertEqual(suggestion_one["author"], "Codex")
        self.assertEqual(suggestion_one["proposal"], "new & improved")
        self.assertEqual(
            format_review(suggestion_one),
            "待确认的修改建议\n建议改为：new & improved\n理由：clearer wording",
        )

    def test_document_preamble_comments_and_verbatim_are_ignored(self) -> None:
        source = (
            r"\newcommand{\twcomment}[3]{definition body}" + "\n"
            + r"% \twcomment{commented}{a}{b}" + "\n"
            + r"\begin{document}" + "\n"
            + r"\verb|\twcomment{verb}{a}{b}|" + "\n"
            + r"\begin{verbatim}" + "\n"
            + r"\twcomment{literal}{a}{b}" + "\n"
            + r"\end{verbatim}" + "\n"
            + r"\twcomment{real}{anchor}{visible note}" + "\n"
            + r"\end{document}"
        )

        prepared, reviews = prepare_reviews(source)

        self.assertEqual([review["key"] for review in reviews], ["real"])
        self.assertIn(r"\newcommand{\twcomment}[3]{definition body}", prepared)
        self.assertIn(r"\twcomment{literal}{a}{b}", prepared)
        self.assertIn(reviews[0]["start_marker"] + "anchor" + reviews[0]["end_marker"], prepared)

    def test_malformed_arguments_invalid_keys_empty_anchors_and_duplicate_ids_reject(self) -> None:
        malformed = [
            r"\twcomment{missing}{anchor}",
            r"\twcomment{1bad}{anchor}{note}",
            r"\twcomment{empty}{   }{note}",
            r"\twcomment{unclosed}{anchor{note}",
            r"\twcomment{dup}{a}{one} \twsuggest{dup}{b}{two}{why}",
        ]
        for source in malformed:
            with self.subTest(source=source):
                with self.assertRaises(ValueError):
                    prepare_reviews(source)

    def test_nested_review_and_structural_anchor_commands_reject(self) -> None:
        bad_anchors = [
            r"\textbf{outer \twsuggest{inner}{x}{y}{why}}",
            r"\section{title}",
            r"\begin{quote}text\end{quote}",
            "first line\n\nsecond line",
        ]
        for anchor in bad_anchors:
            source = r"\twcomment{outer}{" + anchor + r"}{note}"
            with self.subTest(anchor=anchor):
                with self.assertRaises(ValueError):
                    prepare_reviews(source)

    def test_revision_supports_bounded_inline_citations_and_one_empty_side(self) -> None:
        source = (
            r"\twrevision{replace}{原文\cite{old_key}}{建议文\cite{new_key,other}}{理由}"
            "\n"
            r"\twrevision{delete}{待删除段落}{}{移除理由}"
            "\n"
            r"\twrevision{insert}{}{新增段落\cite{new_key}}{补充理由}"
        )

        prepared, reviews = prepare_reviews(source)

        self.assertEqual(len(reviews), 3)
        replacement, deletion, insertion = reviews
        self.assertEqual(replacement["proposal_citations"], ["new_key", "other"])
        self.assertEqual(replacement["proposal"], "建议文[new_key][other]")
        self.assertFalse(replacement["anchor_empty"])
        self.assertFalse(replacement["proposal_empty"])
        self.assertTrue(deletion["proposal_empty"])
        self.assertFalse(deletion["anchor_empty"])
        self.assertTrue(insertion["anchor_empty"])
        self.assertFalse(insertion["proposal_empty"])
        self.assertEqual(insertion["proposal_citations"], ["new_key"])
        self.assertIn(insertion["start_marker"] + insertion["end_marker"], prepared)

    def test_empty_sides_remain_revision_only_and_citation_groups_are_bounded(self) -> None:
        for source in (
            r"\twcomment{empty}{}{note}",
            r"\twsuggest{empty}{}{proposal}{reason}",
            r"\twrevision{too-many}{anchor}{\cite{a,b,c,d}}{reason}",
        ):
            with self.subTest(source=source), self.assertRaises(ValueError):
                prepare_reviews(source)

    def test_encode_plain_escapes_tex_and_preserves_newlines(self) -> None:
        plain = "A {brace} 50% & x_y #tag $2 ^ ~ \\path\nsecond line"
        encoded = encode_plain(plain)
        self.assertIn(r"\{brace\}", encoded)
        self.assertIn(r"50\% \& x\_y \#tag \$2", encoded)
        self.assertIn(r"\textasciicircum{} \textasciitilde{}", encoded)
        self.assertIn(r"\textbackslash{}path" + "\nsecond line", encoded)

        _, reviews = prepare_reviews(r"\twcomment{key}{anchor}{" + encoded + "}")
        self.assertEqual(reviews[0]["note"], plain)


if __name__ == "__main__":
    unittest.main(verbosity=2)
