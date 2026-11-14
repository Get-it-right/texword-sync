import json, os, sys, unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import texword
from latex_compat import unwrap_layout


class JournalLayoutTests(unittest.TestCase):
    def test_resizebox_preserves_nested_text_and_escaped_percent(self):
        original=r'\resizebox{\textwidth}{!}{A {nested} 10\% \resizebox{2cm}{!}{B}}'
        self.assertEqual(unwrap_layout(original),r'A {nested} 10\% B')
        self.assertIn('resizebox',original)

    def test_tabularx_spanning_stage_headings_and_math_survive(self):
        source=r'''\begin{table}
\caption{Control}\label{control}
\begin{tabularx}{\textwidth}{@{}ll>{\raggedright\arraybackslash}X@{}}
Name & Value & Meaning\\
\multicolumn{3}{@{}l}{\textit{General settings}}\\
$P_0$ & 40 W & Initial power\\
\addlinespace[2pt]
\multicolumn{3}{@{}l}{\textit{Stage 1}}\\
$G$ & $-0.05$ & Slope\\
\end{tabularx}\end{table}
See Table~\ref{control}.'''
        ast=json.loads(texword.pandoc(['-f','latex','-t','json'],input=source)[0])
        table=ast['blocks'][0]['c'][1][0]
        self.assertEqual(table['t'],'Table')
        rows=[row for body in table['c'][4] for row in body[3]]
        text=json.dumps(table,ensure_ascii=False)
        for word in ['General','settings','Stage','Initial','power']:
            self.assertIn(word,text)
        self.assertEqual(sum(c[3]==3 for r in rows for c in r[1]),2)
        transform=texword.AstTransform({'heading1_chinese':False})
        transform.prepare(ast['blocks'])
        transform.finish(ast)
        self.assertIn('control',transform.targets)

    def test_abstract_stays_in_body_with_source_order(self):
        ast=json.loads(texword.pandoc(['-f','latex','-t','json'],
            input=r'Before.\begin{abstract}Purpose: body.\end{abstract}After.')[0])
        self.assertNotIn('abstract',ast['meta'])
        self.assertEqual(ast['blocks'][1]['t'],'BlockQuote')
        self.assertIn('Purpose:',json.dumps(ast['blocks'][1]))


if __name__=='__main__':
    unittest.main()
