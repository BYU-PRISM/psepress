"""Regression coverage for manuscript conversion fidelity (stdlib unittest)."""
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from latex_to_word import (LatexParser, InlineLatexConverter, DocxTemplateConverter,
                           BibEntry, compute_column_widths, decode_latex_text,
                           format_bibtex_names, root_namespace_declarations, RunSpec, NS, qn)

ROOT = Path(__file__).resolve().parent


class ConversionTests(unittest.TestCase):
    def test_macros_preserve_backslashes_and_expand_arguments(self):
        parser = LatexParser.__new__(LatexParser)
        parser.active_text = r'\newcommand{\flag}{$\bullet$}\newcommand{\wrap}[1]{\textbf{#1}}'
        self.assertEqual(parser._resolve_macros(r'\wrap{A {nested} value} \flag'),
                         r'\textbf{A {nested} value} $\bullet$')

    def test_nested_macros_and_recursive_definition(self):
        parser = LatexParser.__new__(LatexParser)
        parser.active_text = r'\newcommand{\wrap}[1]{\textbf{#1}}'
        self.assertEqual(parser._resolve_macros(r'\wrap{A \wrap{B}}'), r'\textbf{A \textbf{B}}')
        parser.active_text = r'\newcommand{\loop}{\loop}'
        with self.assertRaisesRegex(ValueError, 'Recursive'):
            parser._resolve_macros(r'\loop')

    def test_figure_without_width(self):
        parser = LatexParser.__new__(LatexParser)
        figure, _ = parser._parse_figure([
            r'\begin{figure}', r'\includegraphics{example.png}',
            r'\caption{Example}', r'\end{figure}'], 0)
        self.assertIsNone(figure.width_hint)

    def test_accents_and_others(self):
        self.assertEqual(decode_latex_text(r'Pablo-Garc{\'\i}a'), 'Pablo-García')
        self.assertEqual(decode_latex_text(r'{\"O}rs and Gr{\"u}nendahl'), 'Örs and Grünendahl')
        self.assertEqual(format_bibtex_names('Smith, Anne and others'), 'A. Smith, et al.')

    def test_ref_and_autoref_have_distinct_labels(self):
        inline = InlineLatexConverter({}, [], {'tab:a': 'Table 1'})
        self.assertEqual(inline.to_plain(r'Table~\ref{tab:a}; \autoref{tab:a}'), 'Table 1; Table 1')

    def test_rotation_preserves_content(self):
        self.assertEqual(InlineLatexConverter({}, []).to_plain(r'\rotatebox[origin=l]{90}{\small Upd}').strip(), 'Upd')

    def test_widths_are_positive_and_ignore_merged_cells(self):
        rows = [['Ref.', 'System', 'Lvl', 'Upd', 'UQ', 'Hum', 'Hyb', 'Summary'],
                ['1', 'Long system name', '0', '', '', '', '', 'A substantially longer summary'],
                ['__MULTICOLUMN:8__' + 'category ' * 50]]
        widths = compute_column_widths(rows, 8, 9638)
        self.assertEqual(sum(widths), 9638)
        self.assertTrue(all(width > 0 for width in widths))
        self.assertGreater(widths[-1], widths[3])
        self.assertEqual(widths, compute_column_widths(rows[:-1], 8, 9638))
        self.assertTrue(all(w > 0 for w in compute_column_widths(rows, 8)))
        compact = compute_column_widths(rows, 8, 9638, compact=True)
        self.assertEqual(sum(compact), 9638)
        self.assertLess(compact[3], widths[3])
        self.assertGreater(compact[-1], widths[-1])

    def test_bibliography_preserves_metadata(self):
        converter = DocxTemplateConverter.__new__(DocxTemplateConverter)
        result = converter._format_bib_entry(BibEntry('inproceedings', 'key', {
            'author': 'Smith, Anne', 'title': 'A paper?', 'booktitle': 'Conference',
            'publisher': 'Publisher', 'pages': '1--8', 'year': '2026', 'doi': '10.1/example'}))
        for value in ('Conference', 'Publisher', '1–8', '10.1/example'):
            self.assertIn(value, result)
        self.assertNotIn('Ed: .', result)
        self.assertNotIn('?.', result)
        result = converter._format_bib_entry(BibEntry('misc', 'key', {
            'title': 'Preprint', 'eprint': '2506.13068', 'archiveprefix': 'arXiv',
            'primaryclass': 'cs.MA', 'note': 'Advance online publication'}))
        self.assertIn('arXiv: 2506.13068 [cs.MA]', result)
        self.assertIn('Advance online publication', result)
        online_fields = {'title': 'Workshop', 'organization': 'University', 'year': '2026',
                         'url': 'https://example.org/workshop', 'note': 'Workshop materials'}
        result = converter._format_bib_entry(BibEntry('online', 'key', online_fields))
        for value in ('University', '2026', 'https://example.org/workshop', 'Workshop materials'):
            self.assertIn(value, result)
        self.assertNotIn('[accessed', result)
        online_fields['urldate'] = '2026-11-16'
        result = converter._format_bib_entry(BibEntry('online', 'key', online_fields))
        self.assertIn('[accessed 2026-11-16]', result)

    def test_doi_display_and_hyperlink_target(self):
        converter = DocxTemplateConverter.__new__(DocxTemplateConverter)
        converter.rel_root = ET.Element("Relationships")
        converter.hyperlinks = {}
        url = "https://doi.org/10.1016/j.compchemeng.2026.109650"
        paragraph = converter._paragraph("PSECitation", [RunSpec("An article. "), RunSpec(url, hyperlink=url)])
        self.assertIsNone(paragraph.find("w:r/w:br", NS))
        self.assertIn(url, "".join(paragraph.itertext()).replace("\u200b", ""))
        self.assertIn("2026.\u200b109650", "".join(paragraph.itertext()))
        self.assertEqual(list(converter.rel_root)[0].get("Target"), url)
        other = converter._paragraph("PSEText", [RunSpec(url, hyperlink=url)])
        self.assertEqual("".join(other.itertext()), url)
        self.assertIsNone(other.find("w:r/w:br", NS))

    def test_author_terminal_period(self):
        converter = DocxTemplateConverter.__new__(DocxTemplateConverter)
        entry = BibEntry("article", "key", {"author": "Smith, Anne and others", "title": "A study", "year": "2026"})
        self.assertIn("A. Smith, et al. A study.", converter._format_bib_entry(entry))
        self.assertNotIn("et al..", converter._format_bib_entry(entry))

    def test_copyright_stays_with_logo(self):
        converter = DocxTemplateConverter.__new__(DocxTemplateConverter)
        paragraph = converter._paragraph("PSECopyright", [RunSpec("License text")])
        self.assertIsNotNone(paragraph.find("w:pPr/w:keepLines", NS))
        self.assertIsNotNone(paragraph.find("w:pPr/w:keepNext", NS))

    def test_docx_wide_table_and_inline_math(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = directory / 'main.tex'
            source.write_text(r'''
\PSESelectConference{focapo-cpc-2027}
\newcommand{\yes}{$\bullet$}
\begin{document}
\section{Introduction}
\label{sec:intro}
Table~\ref{tab:a} compares $x$.

\begin{table*}[t]
\footnotesize
\caption{A table with $x$ and following prose.}
\label{tab:a}
\begin{tabular}{lll}
A & B & C \\
\addlinespace
\multicolumn{2}{l}{Group} & End \\
1 & \yes & 3 \\
\end{tabular}
\end{table*}
\pseprintreferences
\end{document}
''')
            (directory / 'refs.bib').write_text('')
            output = directory / 'result.docx'
            template = directory / 'template.docx'
            with zipfile.ZipFile(ROOT / 'template.docx') as original, zipfile.ZipFile(template, 'w') as custom:
                for name in original.namelist():
                    data = original.read(name)
                    if name == 'word/header1.xml':
                        data = data.replace(b'[Remaining Page s]', b'[Remaining Page Headers]')
                    custom.writestr(name, data)
            DocxTemplateConverter(template, source, output).convert()
            with zipfile.ZipFile(output) as package:
                root = ET.fromstring(package.read('word/document.xml'))
                self.assertNotIn(b'[Remaining Page Headers]', package.read('word/header1.xml'))
                for name in ('word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/fontTable.xml'):
                    xml = package.read(name)
                    prefixes = ET.fromstring(xml).get(qn('mc', 'Ignorable'), '').split()
                    self.assertTrue(set(prefixes).issubset(root_namespace_declarations(xml)), name)
            text = ''.join(root.itertext())
            self.assertIn('and following prose.', text)
            self.assertNotIn('sec:intro', text)
            self.assertNotIn('Table Table', text)
            table = root.find('.//w:tbl', NS)
            self.assertEqual(len(table.findall('w:tr', NS)), 3)
            cells = table.findall('w:tr', NS)[1].findall('w:tc', NS)
            self.assertEqual(len(cells), 2)
            self.assertIn('End', ''.join(cells[-1].itertext()))
            self.assertIsNotNone(table.find('w:tr/w:trPr/w:tblHeader', NS))
            spacing = table.find('w:tr/w:tc/w:p/w:pPr/w:spacing', NS)
            self.assertEqual(spacing.get(qn('w', 'line')), '240')
            self.assertEqual(spacing.get(qn('w', 'lineRule')), 'auto')
            self.assertTrue(root.findall('.//m:oMath', NS))
            columns = [x.get(qn('w', 'num')) for x in root.findall('.//w:sectPr/w:cols', NS)]
            self.assertEqual(columns[-3:], ['2', '1', '2'])
            sections = root.findall('.//w:sectPr', NS)
            self.assertIsNotNone(sections[0].find('w:titlePg', NS))
            self.assertTrue(all(section.find('w:titlePg', NS) is None for section in sections[1:]))
            for p in root.findall('.//w:p', NS):
                style = p.find('w:pPr/w:pStyle', NS)
                if style is not None and style.get(qn('w', 'val')) == 'PSETableCaption':
                    self.assertIsNotNone(p.find('w:pPr/w:keepLines', NS))
                    self.assertIsNotNone(p.find('w:pPr/w:keepNext', NS))


if __name__ == '__main__':
    unittest.main()
