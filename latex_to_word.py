from __future__ import annotations

import argparse
import copy
import re
import struct
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NS = {
    "w": W_NS,
    "r": R_NS,
    "rel": REL_NS,
    "ct": CT_NS,
    "wp": WP_NS,
    "a": A_NS,
    "pic": PIC_NS,
    "w14": W14_NS,
    "m": M_NS,
}

for prefix, uri in (
    ("w", W_NS),
    ("r", R_NS),
    ("wp", WP_NS),
    ("a", A_NS),
    ("pic", PIC_NS),
    ("w14", W14_NS),
    ("m", M_NS),
):
    ET.register_namespace(prefix, uri)


def qn(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def split_tex_lines(value: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    i = 0
    while i < len(value):
        if value.startswith("\\\\", i) and depth == 0:
            parts.append("".join(current).strip())
            current = []
            i += 2
            continue
        ch = value[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        current.append(ch)
        i += 1
    final = "".join(current).strip()
    if final or not parts:
        parts.append(final)
    return [part for part in parts if part]


def parse_braced(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        raise ValueError(f"Expected '{{' at position {start}")
    depth = 0
    buf: list[str] = []
    i = start
    while i < len(text):
        ch = text[i]
        if ch == "{" and (i == 0 or text[i - 1] != "\\"):
            depth += 1
            if depth > 1:
                buf.append(ch)
        elif ch == "}" and (i == 0 or text[i - 1] != "\\"):
            depth -= 1
            if depth == 0:
                return "".join(buf), i + 1
            buf.append(ch)
        else:
            buf.append(ch)
        i += 1
    raise ValueError("Unbalanced braces in LaTeX input.")


def extract_command_body(text: str, command_name: str) -> str | None:
    for marker in (
        f"\\renewcommand{{\\{command_name}}}",
        f"\\newcommand{{\\{command_name}}}",
    ):
        idx = text.find(marker)
        if idx == -1:
            continue
        pos = idx + len(marker)
        while pos < len(text) and text[pos].isspace():
            pos += 1
        if pos < len(text) and text[pos] == "[":
            end = text.find("]", pos)
            pos = end + 1 if end != -1 else pos
            while pos < len(text) and text[pos].isspace():
                pos += 1
        if pos < len(text) and text[pos] == "{":
            body, _ = parse_braced(text, pos)
            return body
    return None


def strip_comments(text: str) -> str:
    out: list[str] = []
    i = 0
    while i < len(text):
        if text[i] == "%" and (i == 0 or text[i - 1] != "\\"):
            while i < len(text) and text[i] != "\n":
                i += 1
            continue
        out.append(text[i])
        i += 1
    return "".join(out)


def clean_bib_value(value: str) -> str:
    value = value.strip()
    if value.startswith("{") and value.endswith("}"):
        value = value[1:-1]
    value = value.replace("\n", " ")
    value = value.replace("~", " ")
    value = value.replace("\\&", "&")
    value = value.replace("\\_", "_")
    value = re.sub(r"[{}]", "", value)
    return normalize_space(value)


def load_latex_source(tex_path: Path, seen: set[Path] | None = None) -> str:
    tex_path = tex_path.resolve()
    if seen is None:
        seen = set()
    if tex_path in seen:
        raise ValueError(f"Recursive LaTeX input detected: {tex_path}")
    seen = set(seen)
    seen.add(tex_path)
    text = tex_path.read_text(encoding="utf-8")

    def replace(match: re.Match[str]) -> str:
        raw_target = match.group(1).strip()
        if not raw_target:
            return match.group(0)
        target = Path(raw_target)
        if target.suffix == "":
            target = target.with_suffix(".tex")
        target_path = (tex_path.parent / target).resolve()
        if not target_path.is_file():
            return match.group(0)
        return load_latex_source(target_path, seen)

    return re.sub(r"\\(?:input|include)\{([^}]+)\}", replace, text)


@dataclass
class RunSpec:
    text: str
    superscript: bool = False
    hyperlink: str | None = None
    bold: bool = False
    italic: bool = False


@dataclass
class HeadingBlock:
    level: int
    title: str


@dataclass
class ParagraphBlock:
    text: str


@dataclass
class FigureBlock:
    path: str
    caption: str
    wide: bool = False
    width_hint: str | None = None


@dataclass
class TableBlock:
    caption: str
    rows: list[list[str]]


@dataclass
class EquationBlock:
    body: str


@dataclass
class MathTextNode:
    text: str
    literal: bool = False


@dataclass
class MathSequenceNode:
    items: list["MathNode"]


@dataclass
class MathFractionNode:
    numerator: "MathNode"
    denominator: "MathNode"


@dataclass
class MathScriptNode:
    base: "MathNode"
    sub: "MathNode | None" = None
    sup: "MathNode | None" = None


@dataclass
class MathNaryNode:
    operator: str
    body: "MathNode"
    sub: "MathNode | None" = None
    sup: "MathNode | None" = None


@dataclass
class MathMatrixNode:
    rows: list[list["MathNode"]]


MathNode = MathTextNode | MathSequenceNode | MathFractionNode | MathScriptNode | MathNaryNode | MathMatrixNode


@dataclass
class MathToken:
    kind: str
    value: str


@dataclass
class CodeBlock:
    lines: list[str]


@dataclass
class DefinitionBlock:
    title: str
    body: str


@dataclass
class CommandBlock:
    name: str


Block = HeadingBlock | ParagraphBlock | FigureBlock | TableBlock | EquationBlock | CodeBlock | DefinitionBlock | CommandBlock


@dataclass
class ParsedDocument:
    metadata: dict[str, str]
    blocks: list[Block]
    bib_path: Path
    nocite_all: bool = False
    cited_keys: list[str] = field(default_factory=list)


@dataclass
class BibEntry:
    entry_type: str
    key: str
    fields: dict[str, str]


class LatexMathParser:
    GREEK_MAP = {
        "\\alpha": "α",
        "\\beta": "β",
        "\\gamma": "γ",
        "\\delta": "δ",
        "\\epsilon": "ε",
        "\\mu": "μ",
        "\\pi": "π",
        "\\sigma": "σ",
        "\\theta": "θ",
        "\\omega": "ω",
        "\\infty": "∞",
    }
    SPACING_COMMANDS = {"\\,", "\\;", "\\!", "\\quad", "\\qquad"}

    def __init__(self, text: str):
        self.tokens = self._tokenize(text)
        self.index = 0

    def parse(self) -> MathNode:
        node = self._parse_sequence()
        return self._collapse(node)

    def _tokenize(self, text: str) -> list[MathToken]:
        tokens: list[MathToken] = []
        i = 0
        while i < len(text):
            if text[i].isspace():
                i += 1
                continue
            if text.startswith("\\begin{matrix}", i):
                tokens.append(MathToken("BEGIN_MATRIX", "matrix"))
                i += len("\\begin{matrix}")
                continue
            if text.startswith("\\end{matrix}", i):
                tokens.append(MathToken("END_MATRIX", "matrix"))
                i += len("\\end{matrix}")
                continue
            if text.startswith("\\\\", i):
                tokens.append(MathToken("NEWROW", "\\\\"))
                i += 2
                continue
            if text[i] == "\\":
                if i + 1 < len(text) and text[i + 1] in ",;!":
                    i += 2
                    continue
                if i + 1 < len(text) and text[i + 1] in "{}_^&":
                    tokens.append(MathToken("TEXT", text[i + 1]))
                    i += 2
                    continue
                j = i + 1
                while j < len(text) and text[j].isalpha():
                    j += 1
                command = text[i:j]
                if command in self.SPACING_COMMANDS:
                    i = j
                    continue
                if command == "\\left":
                    tokens.append(MathToken("LEFT", command))
                    i = j
                    continue
                if command == "\\right":
                    tokens.append(MathToken("RIGHT", command))
                    i = j
                    continue
                tokens.append(MathToken("COMMAND", command))
                i = j if j > i + 1 else i + 1
                continue
            if text[i] == "{":
                tokens.append(MathToken("LBRACE", "{"))
                i += 1
                continue
            if text[i] == "}":
                tokens.append(MathToken("RBRACE", "}"))
                i += 1
                continue
            if text[i] == "_":
                tokens.append(MathToken("SUB", "_"))
                i += 1
                continue
            if text[i] == "^":
                tokens.append(MathToken("SUP", "^"))
                i += 1
                continue
            if text[i] == "&":
                tokens.append(MathToken("ALIGN", "&"))
                i += 1
                continue
            if text[i] in "=+-()[]":
                tokens.append(MathToken("TEXT", text[i]))
                i += 1
                continue
            j = i
            while j < len(text) and not text[j].isspace() and text[j] not in "\\{}_^&=+-()[]":
                j += 1
            tokens.append(MathToken("TEXT", text[i:j]))
            i = j
        tokens.append(MathToken("EOF", ""))
        return tokens

    def _current(self) -> MathToken:
        return self.tokens[self.index]

    def _advance(self) -> MathToken:
        token = self.tokens[self.index]
        self.index += 1
        return token

    def _match(self, kind: str, value: str | None = None) -> bool:
        token = self._current()
        if token.kind != kind:
            return False
        if value is not None and token.value != value:
            return False
        self.index += 1
        return True

    def _expect(self, kind: str, value: str | None = None) -> MathToken:
        token = self._current()
        if token.kind != kind or (value is not None and token.value != value):
            raise ValueError(f"Expected {kind} {value or ''} in LaTeX math, found {token.kind} {token.value!r}")
        self.index += 1
        return token

    def _parse_sequence(self, stop_kinds: set[str] | None = None, stop_text: set[str] | None = None) -> MathNode:
        items: list[MathNode] = []
        stop_kinds = stop_kinds or set()
        stop_text = stop_text or set()
        while True:
            token = self._current()
            if token.kind == "EOF" or token.kind in stop_kinds:
                break
            if token.kind == "TEXT" and token.value in stop_text:
                break
            items.append(self._parse_item())
        return self._collapse(items)

    def _parse_item(self) -> MathNode:
        token = self._current()
        if token.kind == "COMMAND" and token.value == "\\int":
            return self._parse_integral()
        base = self._parse_base()
        sub: MathNode | None = None
        sup: MathNode | None = None
        while self._current().kind in {"SUB", "SUP"}:
            if self._match("SUB"):
                sub = self._parse_script_argument()
                continue
            if self._match("SUP"):
                sup = self._parse_script_argument()
                continue
        if sub is not None or sup is not None:
            return MathScriptNode(base=base, sub=sub, sup=sup)
        return base

    def _parse_base(self) -> MathNode:
        token = self._current()
        if token.kind == "LBRACE":
            self._advance()
            node = self._parse_sequence(stop_kinds={"RBRACE"})
            self._expect("RBRACE")
            return node
        if token.kind == "LEFT":
            self._advance()
            left = self._parse_delimiter()
            inner = self._parse_sequence(stop_kinds={"RIGHT"})
            self._expect("RIGHT")
            right = self._parse_delimiter()
            items = [MathTextNode(left, literal=True), *self._flatten(inner), MathTextNode(right, literal=True)]
            return self._collapse(items)
        if token.kind == "BEGIN_MATRIX":
            return self._parse_matrix()
        if token.kind == "COMMAND":
            self._advance()
            if token.value == "\\frac":
                numerator = self._parse_group_or_item()
                denominator = self._parse_group_or_item()
                return MathFractionNode(numerator=numerator, denominator=denominator)
            if token.value in self.GREEK_MAP:
                return MathTextNode(self.GREEK_MAP[token.value])
            return MathTextNode(token.value.lstrip("\\"))
        if token.kind == "TEXT":
            self._advance()
            return MathTextNode(token.value, literal=token.value in "{}")
        raise ValueError(f"Unsupported LaTeX math token: {token.kind} {token.value!r}")

    def _parse_group_or_item(self) -> MathNode:
        if self._current().kind == "LBRACE":
            self._advance()
            node = self._parse_sequence(stop_kinds={"RBRACE"})
            self._expect("RBRACE")
            return node
        return self._parse_item()

    def _parse_script_argument(self) -> MathNode:
        if self._current().kind == "LBRACE":
            self._advance()
            node = self._parse_sequence(stop_kinds={"RBRACE"})
            self._expect("RBRACE")
            return node
        return self._parse_base()

    def _parse_delimiter(self) -> str:
        token = self._current()
        if token.kind == "TEXT":
            self._advance()
            return token.value
        if token.kind == "COMMAND" and token.value in self.GREEK_MAP:
            self._advance()
            return self.GREEK_MAP[token.value]
        raise ValueError(f"Unsupported delimiter in LaTeX math: {token.kind} {token.value!r}")

    def _parse_matrix(self) -> MathNode:
        self._expect("BEGIN_MATRIX")
        rows: list[list[MathNode]] = []
        current_row: list[MathNode] = []
        while True:
            cell = self._parse_sequence(stop_kinds={"ALIGN", "NEWROW", "END_MATRIX"})
            current_row.append(cell)
            if self._match("ALIGN"):
                continue
            if self._match("NEWROW"):
                rows.append(current_row)
                current_row = []
                continue
            self._expect("END_MATRIX")
            rows.append(current_row)
            break
        return MathMatrixNode(rows=rows)

    def _parse_integral(self) -> MathNode:
        self._expect("COMMAND", "\\int")
        sub: MathNode | None = None
        sup: MathNode | None = None
        while self._current().kind in {"SUB", "SUP"}:
            if self._match("SUB"):
                sub = self._parse_script_argument()
                continue
            if self._match("SUP"):
                sup = self._parse_script_argument()
                continue
        body = self._parse_sequence(stop_kinds={"RBRACE", "RIGHT", "ALIGN", "NEWROW", "END_MATRIX"}, stop_text={"+", "-"})
        return MathNaryNode(operator="∫", body=body, sub=sub, sup=sup)

    def _collapse(self, node_or_items: MathNode | list[MathNode]) -> MathNode:
        if isinstance(node_or_items, list):
            items = [item for item in node_or_items if not self._is_empty(item)]
        else:
            return node_or_items
        if not items:
            return MathTextNode("")
        if len(items) == 1:
            return items[0]
        return MathSequenceNode(items=items)

    def _flatten(self, node: MathNode) -> list[MathNode]:
        if isinstance(node, MathSequenceNode):
            return node.items
        return [node]

    def _is_empty(self, node: MathNode) -> bool:
        return isinstance(node, MathTextNode) and node.text == ""


class BibParser:
    @staticmethod
    def parse(path: Path) -> list[BibEntry]:
        text = path.read_text(encoding="utf-8")
        entries: list[BibEntry] = []
        i = 0
        while i < len(text):
            at = text.find("@", i)
            if at == -1:
                break
            j = at + 1
            while j < len(text) and (text[j].isalnum() or text[j] in "-_"):
                j += 1
            entry_type = text[at + 1 : j].strip().lower()
            while j < len(text) and text[j].isspace():
                j += 1
            if j >= len(text) or text[j] != "{":
                i = j + 1
                continue
            body, end = parse_braced(text, j)
            key, fields_blob = BibParser._split_key_and_fields(body)
            fields = BibParser._parse_fields(fields_blob)
            entries.append(BibEntry(entry_type=entry_type, key=key.strip(), fields=fields))
            i = end
        return entries

    @staticmethod
    def _split_key_and_fields(body: str) -> tuple[str, str]:
        depth = 0
        for i, ch in enumerate(body):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth = max(0, depth - 1)
            elif ch == "," and depth == 0:
                return body[:i], body[i + 1 :]
        return body, ""

    @staticmethod
    def _parse_fields(blob: str) -> dict[str, str]:
        fields: dict[str, str] = {}
        i = 0
        while i < len(blob):
            while i < len(blob) and blob[i] in " \t\r\n,":
                i += 1
            if i >= len(blob):
                break
            start = i
            while i < len(blob) and (blob[i].isalnum() or blob[i] in "-_"):
                i += 1
            name = blob[start:i].strip().lower()
            while i < len(blob) and blob[i].isspace():
                i += 1
            if i < len(blob) and blob[i] == "=":
                i += 1
            while i < len(blob) and blob[i].isspace():
                i += 1
            if i >= len(blob):
                fields[name] = ""
                break
            if blob[i] == "{":
                value, i = parse_braced(blob, i)
            elif blob[i] == '"':
                i += 1
                start = i
                while i < len(blob):
                    if blob[i] == '"' and blob[i - 1] != "\\":
                        break
                    i += 1
                value = blob[start:i]
                i += 1
            else:
                start = i
                while i < len(blob) and blob[i] not in ",\n":
                    i += 1
                value = blob[start:i]
            fields[name] = value.strip()
        return fields


class LatexParser:
    METADATA_KEYS = [
        "PaperTitle",
        "PaperAuthors",
        "PaperAffiliations",
        "PaperAbstract",
        "PaperKeywords",
        "PaperCopyrightText",
        "HeaderArticleType",
        "HeaderReviewType",
        "HeaderConference",
        "HeaderLocation",
        "FooterJournalInfo",
        "FirstPageFooterID",
        "BodyPageFooterID",
    ]

    def __init__(self, tex_path: Path):
        self.tex_path = tex_path
        self.text = load_latex_source(tex_path)
        self.metadata = {key: extract_command_body(self.text, key) or "" for key in self.METADATA_KEYS}
        self.document_body = self._extract_document_body()
        self.bib_path = self._extract_bib_path()
        self.nocite_all = "\\nocite{*}" in self.document_body

    def _extract_document_body(self) -> str:
        start = self.text.find("\\begin{document}")
        end = self.text.find("\\end{document}")
        if start == -1 or end == -1:
            raise ValueError("Could not find a complete document body in the LaTeX input.")
        return self.text[start + len("\\begin{document}") : end]

    def _extract_bib_path(self) -> Path:
        match = re.search(r"\\addbibresource\{([^}]+)\}", self.text)
        if match:
            return (self.tex_path.parent / match.group(1)).resolve()
        return (self.tex_path.parent / "refs.bib").resolve()

    def parse(self) -> ParsedDocument:
        body = strip_comments(self.document_body)
        lines = body.splitlines()
        blocks: list[Block] = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
            if line in {"\\psemaketitle", "\\balance"} or line.startswith("\\nocite"):
                i += 1
                continue
            if line == "\\pseprintreferences":
                blocks.append(CommandBlock("references"))
                i += 1
                continue
            if line == "\\pseprintcopyright":
                blocks.append(CommandBlock("copyright"))
                i += 1
                continue
            heading = self._parse_heading(line)
            if heading is not None:
                blocks.append(heading)
                i += 1
                continue
            if line.startswith("\\begin{figure"):
                block, i = self._parse_figure(lines, i)
                blocks.append(block)
                continue
            if line.startswith("\\begin{table"):
                block, i = self._parse_table(lines, i)
                blocks.append(block)
                continue
            if line.startswith("\\begin{equation"):
                block, i = self._parse_equation(lines, i)
                blocks.append(block)
                continue
            if line.startswith("\\begin{lstlisting"):
                block, i = self._parse_code(lines, i)
                blocks.append(block)
                continue
            if line.startswith("\\begin{psedefinition"):
                block, i = self._parse_definition(lines, i)
                blocks.append(block)
                continue

            paragraph_lines: list[str] = []
            while i < len(lines):
                current = lines[i].strip()
                if not current:
                    break
                if current.startswith("\\section{") or current.startswith("\\subsection{") or current.startswith("\\subsubsection{") or current.startswith("\\paragraph{"):
                    break
                if current.startswith("\\begin{") or current in {"\\pseprintreferences", "\\pseprintcopyright", "\\psemaketitle", "\\balance"} or current.startswith("\\nocite"):
                    break
                paragraph_lines.append(current)
                i += 1
            if paragraph_lines:
                blocks.append(ParagraphBlock(text=" ".join(paragraph_lines)))
            else:
                i += 1
        return ParsedDocument(metadata=self.metadata, blocks=blocks, bib_path=self.bib_path, nocite_all=self.nocite_all)

    def _parse_heading(self, line: str) -> HeadingBlock | None:
        for command, level in (("\\section", 1), ("\\subsection", 2), ("\\subsubsection", 3), ("\\paragraph", 4)):
            if line.startswith(f"{command}{{"):
                title, _ = parse_braced(line, len(command))
                return HeadingBlock(level=level, title=title)
        return None

    def _parse_figure(self, lines: list[str], start: int) -> tuple[FigureBlock, int]:
        wide = lines[start].strip().startswith("\\begin{figure*")
        image_path = ""
        caption = ""
        width_hint: str | None = None
        i = start + 1
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith("\\includegraphics"):
                path_match = re.search(r"\{([^}]+)\}", line)
                if path_match:
                    image_path = path_match.group(1).strip()
                opt = re.search(r"\[([^\]]+)\]", line)
                if opt:
                    width_hint = opt.group(1)
            elif line.startswith("\\caption"):
                caption, _ = parse_braced(line, len("\\caption"))
            elif line.startswith("\\end{figure"):
                return FigureBlock(path=image_path, caption=caption, wide=wide, width_hint=width_hint), i + 1
            i += 1
        raise ValueError("Unclosed figure environment.")

    def _parse_table(self, lines: list[str], start: int) -> tuple[TableBlock, int]:
        caption = ""
        rows: list[list[str]] = []
        in_tabular = False
        pending = ""
        i = start + 1
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith("\\caption"):
                caption, _ = parse_braced(line, len("\\caption"))
            elif line.startswith("\\begin{tabular"):
                in_tabular = True
            elif line.startswith("\\end{tabular"):
                in_tabular = False
            elif line.startswith("\\end{table"):
                return TableBlock(caption=caption, rows=rows), i + 1
            elif in_tabular:
                if line in {"\\toprule", "\\midrule", "\\bottomrule"} or not line:
                    i += 1
                    continue
                pending = f"{pending} {line}".strip()
                if pending.endswith("\\\\"):
                    row_text = pending[:-2].strip()
                    rows.append([cell.strip() for cell in self._split_table_row(row_text)])
                    pending = ""
            i += 1
        raise ValueError("Unclosed table environment.")

    def _split_table_row(self, row: str) -> list[str]:
        parts: list[str] = []
        current: list[str] = []
        depth = 0
        for ch in row:
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth = max(0, depth - 1)
            elif ch == "&" and depth == 0:
                parts.append("".join(current))
                current = []
                continue
            current.append(ch)
        parts.append("".join(current))
        return parts

    def _parse_equation(self, lines: list[str], start: int) -> tuple[EquationBlock, int]:
        body_lines: list[str] = []
        i = start + 1
        while i < len(lines):
            line = lines[i].rstrip()
            if line.strip().startswith("\\end{equation"):
                return EquationBlock(body=" ".join(part.strip() for part in body_lines if part.strip())), i + 1
            body_lines.append(line)
            i += 1
        raise ValueError("Unclosed equation environment.")

    def _parse_code(self, lines: list[str], start: int) -> tuple[CodeBlock, int]:
        code_lines: list[str] = []
        i = start + 1
        while i < len(lines):
            line = lines[i].rstrip("\r")
            if line.strip().startswith("\\end{lstlisting"):
                return CodeBlock(lines=code_lines), i + 1
            code_lines.append(line)
            i += 1
        raise ValueError("Unclosed lstlisting environment.")

    def _parse_definition(self, lines: list[str], start: int) -> tuple[DefinitionBlock, int]:
        title, _ = parse_braced(lines[start].strip(), len("\\begin{psedefinition}"))
        body_lines: list[str] = []
        i = start + 1
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith("\\end{psedefinition"):
                return DefinitionBlock(title=title, body=" ".join(body_lines).strip()), i + 1
            if line:
                body_lines.append(line)
            i += 1
        raise ValueError("Unclosed psedefinition environment.")


class InlineLatexConverter:
    SIMPLE_ESCAPES = {
        "\\&": "&",
        "\\%": "%",
        "\\_": "_",
        "\\#": "#",
        "\\$": "$",
        "\\{": "{",
        "\\}": "}",
        "\\,": " ",
        "\\;": " ",
    }

    def __init__(self, citation_numbers: dict[str, int], cited_keys: list[str]):
        self.citation_numbers = citation_numbers
        self.cited_keys = cited_keys

    def to_runs(self, text: str) -> list[RunSpec]:
        runs: list[RunSpec] = []
        i = 0
        text = text.replace("\r", "")
        while i < len(text):
            if text[i] == "%" and (i == 0 or text[i - 1] != "\\"):
                while i < len(text) and text[i] != "\n":
                    i += 1
                continue
            if text.startswith("\\\\", i):
                self._append_run(runs, " ")
                i += 2
                continue
            if text.startswith("\\href", i):
                url, pos = parse_braced(text, i + len("\\href"))
                label, pos = parse_braced(text, pos)
                self._append_run(runs, self.to_plain(label), hyperlink=self.to_plain(url))
                i = pos
                continue
            if text.startswith("\\url", i):
                url, pos = parse_braced(text, i + len("\\url"))
                url_text = self.to_plain(url)
                self._append_run(runs, url_text, hyperlink=url_text)
                i = pos
                continue
            if text.startswith("\\textsuperscript", i):
                inner, pos = parse_braced(text, i + len("\\textsuperscript"))
                self._append_run(runs, self.to_plain(inner), superscript=True)
                i = pos
                continue
            if text.startswith("\\texttt", i):
                inner, pos = parse_braced(text, i + len("\\texttt"))
                self._append_run(runs, self.to_plain(inner))
                i = pos
                continue
            if text.startswith("\\cite", i):
                keys, pos = parse_braced(text, i + len("\\cite"))
                numbers: list[int] = []
                for key in [item.strip() for item in keys.split(",") if item.strip()]:
                    if key in self.citation_numbers:
                        numbers.append(self.citation_numbers[key])
                    if key not in self.cited_keys:
                        self.cited_keys.append(key)
                numbers = sorted(dict.fromkeys(numbers))
                self._append_run(runs, "[" + ",".join(str(number) for number in numbers) + "]")
                i = pos
                continue
            if text.startswith("\\textbf", i):
                inner, pos = parse_braced(text, i + len("\\textbf"))
                self._append_run(runs, self.to_plain(inner), bold=True)
                i = pos
                continue
            if text.startswith("\\textit", i) or text.startswith("\\emph", i):
                command = "\\textit" if text.startswith("\\textit", i) else "\\emph"
                inner, pos = parse_braced(text, i + len(command))
                self._append_run(runs, self.to_plain(inner), italic=True)
                i = pos
                continue
            if text.startswith("\\string\\", i):
                j = i + len("\\string\\")
                start = j
                while j < len(text) and (text[j].isalpha() or text[j] in "@*"):
                    j += 1
                self._append_run(runs, "\\" + text[start:j])
                i = j
                continue
            if text.startswith("\\textcopyright", i):
                self._append_run(runs, "©")
                i += len("\\textcopyright")
                continue
            replaced = False
            for escape, replacement in self.SIMPLE_ESCAPES.items():
                if text.startswith(escape, i):
                    self._append_run(runs, replacement)
                    i += len(escape)
                    replaced = True
                    break
            if replaced:
                continue
            if text[i] == "\\":
                j = i + 1
                while j < len(text) and (text[j].isalpha() or text[j] == "*"):
                    j += 1
                if j < len(text) and text[j] == "{":
                    inner, pos = parse_braced(text, j)
                    self._append_run(runs, self.to_plain(inner))
                    i = pos
                    continue
                i = j if j > i + 1 else i + 1
                continue
            if text[i] in "{}":
                i += 1
                continue
            if text[i] == "~" or text[i] == "\n":
                self._append_run(runs, " ")
                i += 1
                continue
            start = i
            while i < len(text) and text[i] not in "\\{}~%\n":
                i += 1
            self._append_run(runs, text[start:i])
        return self._normalize_runs(runs)

    def to_plain(self, text: str) -> str:
        return normalize_space("".join(run.text for run in self.to_runs(text)))

    def _append_run(
        self,
        runs: list[RunSpec],
        text: str,
        *,
        superscript: bool = False,
        hyperlink: str | None = None,
        bold: bool = False,
        italic: bool = False,
    ) -> None:
        if not text:
            return
        if runs and (runs[-1].superscript, runs[-1].hyperlink, runs[-1].bold, runs[-1].italic) == (superscript, hyperlink, bold, italic):
            runs[-1].text += text
        else:
            runs.append(RunSpec(text=text, superscript=superscript, hyperlink=hyperlink, bold=bold, italic=italic))

    def _normalize_runs(self, runs: list[RunSpec]) -> list[RunSpec]:
        normalized: list[RunSpec] = []
        for run in runs:
            text = re.sub(r"\s+", " ", run.text)
            if not text:
                continue
            if normalized and normalized[-1].text.endswith(" ") and text.startswith(" "):
                text = text.lstrip()
            if normalized and (normalized[-1].superscript, normalized[-1].hyperlink, normalized[-1].bold, normalized[-1].italic) == (
                run.superscript,
                run.hyperlink,
                run.bold,
                run.italic,
            ):
                normalized[-1].text += text
            else:
                normalized.append(RunSpec(text=text, superscript=run.superscript, hyperlink=run.hyperlink, bold=run.bold, italic=run.italic))
        return normalized


def next_rel_id(rel_root: ET.Element) -> str:
    max_id = 0
    for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
        match = re.fullmatch(r"rId(\d+)", rel.attrib.get("Id", ""))
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def add_content_type(content_root: ET.Element, ext: str) -> None:
    ext = ext.lstrip(".").lower()
    existing = {elem.attrib.get("Extension", "").lower() for elem in content_root.findall(f"{{{CT_NS}}}Default")}
    if ext in existing:
        return
    content_type = {
        "png": "image/png",
        "gif": "image/gif",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "tif": "image/tiff",
        "tiff": "image/tiff",
    }.get(ext)
    if not content_type:
        raise ValueError(f"Unsupported image type for DOCX packaging: .{ext}")
    default = ET.SubElement(content_root, f"{{{CT_NS}}}Default")
    default.set("Extension", ext)
    default.set("ContentType", content_type)


def unique_media_name_in_set(existing_names: set[str], stem: str, suffix: str) -> str:
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "-", stem) or "image"
    candidate = f"{safe_stem}{suffix}"
    counter = 1
    while candidate in existing_names:
        candidate = f"{safe_stem}-{counter}{suffix}"
        counter += 1
    return candidate


def resolve_image_path(base_dir: Path, image_ref: str) -> Path:
    raw = image_ref.strip()
    candidate = Path(raw)
    search_paths: list[Path] = []
    if candidate.is_absolute():
        search_paths.append(candidate)
    else:
        search_paths.extend(
            [
                base_dir / raw,
                base_dir / "media" / raw,
                base_dir / "latex_assets" / "media" / raw,
                base_dir.parent / raw,
                base_dir.parent / "media" / raw,
            ]
        )
    if candidate.suffix:
        for path in search_paths:
            if path.exists():
                return path.resolve()
    else:
        for base in search_paths:
            for ext in (".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff"):
                path = base.with_suffix(ext)
                if path.exists():
                    return path.resolve()
    raise FileNotFoundError(f"Could not resolve image referenced from LaTeX: {image_ref}")


def compute_column_widths(rows: list[list[str]], col_count: int) -> list[int]:
    total_width = 4820
    scores = [1] * col_count
    for idx in range(col_count):
        scores[idx] = max(1, max((len(normalize_space(row[idx])) for row in rows if idx < len(row)), default=1))
    widths = [max(800, int(total_width * score / sum(scores))) for score in scores]
    widths[-1] += total_width - sum(widths)
    return widths


def looks_numeric(value: str) -> bool:
    candidate = normalize_space(value).lower()
    return bool(candidate) and all(ch in "0123456789.,-+/%() kmolhr" for ch in candidate)


def parse_png_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("Invalid PNG file.")
    width = struct.unpack(">I", data[16:20])[0]
    height = struct.unpack(">I", data[20:24])[0]
    return width, height


def parse_gif_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    width, height = struct.unpack("<HH", data[6:10])
    return width, height


def parse_jpeg_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if data[:2] != b"\xff\xd8":
        raise ValueError("Invalid JPEG file.")
    i = 2
    while i < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
            height = struct.unpack(">H", data[i + 5 : i + 7])[0]
            width = struct.unpack(">H", data[i + 7 : i + 9])[0]
            return width, height
        length = struct.unpack(">H", data[i + 2 : i + 4])[0]
        i += 2 + length
    raise ValueError("Could not determine JPEG dimensions.")


def image_size(path: Path) -> tuple[int, int]:
    suffix = path.suffix.lower()
    if suffix == ".png":
        return parse_png_size(path)
    if suffix == ".gif":
        return parse_gif_size(path)
    if suffix in {".jpg", ".jpeg"}:
        return parse_jpeg_size(path)
    raise ValueError(f"Unsupported image format for sizing: {path.suffix}")


def parse_width_hint(width_hint: str | None, wide: bool) -> int:
    max_cm = 17.0 if wide else 8.5
    if not width_hint:
        return int(max_cm * 360000)
    match = re.search(r"width\s*=\s*([0-9.]+)\s*(cm|mm|in|pt)", width_hint)
    if match:
        value = float(match.group(1))
        unit = match.group(2)
        if unit == "cm":
            cm = value
        elif unit == "mm":
            cm = value / 10.0
        elif unit == "in":
            cm = value * 2.54
        else:
            cm = value * 0.0352778
        return int(cm * 360000)
    if "\\columnwidth" in width_hint:
        return int(8.5 * 360000)
    if "\\textwidth" in width_hint:
        return int((17.0 if wide else 17.5) * 360000)
    return int(max_cm * 360000)


def compute_image_extent(path: Path, width_hint: str | None, wide: bool) -> tuple[int, int]:
    width_px, height_px = image_size(path)
    width_emu = parse_width_hint(width_hint, wide)
    height_emu = max(1, int(width_emu * height_px / max(1, width_px)))
    return width_emu, height_emu


class MediaManager:
    IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    def __init__(self, package_entries: dict[str, bytes], rel_root: ET.Element, content_root: ET.Element):
        self.package_entries = package_entries
        self.rel_root = rel_root
        self.content_root = content_root
        self.existing_names = {Path(name).name for name in package_entries if name.startswith("word/media/")}
        self.cache: dict[Path, str] = {}

    def add_image(self, source_path: Path) -> str:
        source_path = source_path.resolve()
        if source_path in self.cache:
            return self.cache[source_path]
        target_name = unique_media_name_in_set(self.existing_names, source_path.stem, source_path.suffix.lower())
        self.package_entries[f"word/media/{target_name}"] = source_path.read_bytes()
        self.existing_names.add(target_name)
        add_content_type(self.content_root, source_path.suffix)
        rid = next_rel_id(self.rel_root)
        rel = ET.SubElement(self.rel_root, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", self.IMAGE_REL_TYPE)
        rel.set("Target", f"media/{target_name}")
        self.cache[source_path] = rid
        return rid


class DocxTemplateConverter:
    HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

    def __init__(self, template_path: Path, tex_path: Path, output_path: Path):
        self.template_path = template_path
        self.tex_path = tex_path
        self.output_path = output_path
        self.parsed = LatexParser(tex_path).parse()
        self.bib_entries = BibParser.parse(self.parsed.bib_path)
        self.citation_numbers = {entry.key: idx + 1 for idx, entry in enumerate(self.bib_entries)}
        self.inline = InlineLatexConverter(self.citation_numbers, self.parsed.cited_keys)
        self.max_docpr_id = 100
        self.body_section_break: ET.Element | None = None
        self.final_section: ET.Element | None = None
        self.figure_paragraph_template: ET.Element | None = None
        self.cc_logo_paragraph_template: ET.Element | None = None
        self.equation_ppr_template: ET.Element | None = None
        self.sample_table_pr: ET.Element | None = None
        self.rel_root: ET.Element | None = None
        self.hyperlinks: dict[str, str] = {}

    def convert(self) -> None:
        with zipfile.ZipFile(self.template_path) as archive:
            package_entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}

        self._load_templates(package_entries)
        doc_tree = ET.ElementTree(ET.fromstring(package_entries["word/document.xml"]))
        rel_tree = ET.ElementTree(ET.fromstring(package_entries["word/_rels/document.xml.rels"]))
        content_tree = ET.ElementTree(ET.fromstring(package_entries["[Content_Types].xml"]))

        self.rel_root = rel_tree.getroot()
        self.max_docpr_id = self._compute_max_docpr_id(doc_tree.getroot())
        media = MediaManager(package_entries, self.rel_root, content_tree.getroot())

        self._build_document(doc_tree.getroot(), media)
        self._update_headers_and_footers(package_entries)

        package_entries["word/document.xml"] = ET.tostring(doc_tree.getroot(), encoding="utf-8", xml_declaration=True)
        package_entries["word/_rels/document.xml.rels"] = ET.tostring(rel_tree.getroot(), encoding="utf-8", xml_declaration=True)
        package_entries["[Content_Types].xml"] = ET.tostring(content_tree.getroot(), encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(self.output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in sorted(package_entries):
                archive.writestr(name, package_entries[name])

    def _load_templates(self, package_entries: dict[str, bytes]) -> None:
        document = ET.fromstring(package_entries["word/document.xml"])
        body = document.find("w:body", NS)
        if body is None:
            raise ValueError("Template DOCX is missing the document body.")
        table = body.find("w:tbl", NS)
        if table is not None:
            self.sample_table_pr = copy.deepcopy(table.find("w:tblPr", NS))
        paragraphs = body.findall("w:p", NS)
        self.figure_paragraph_template = copy.deepcopy(paragraphs[19])
        self.cc_logo_paragraph_template = copy.deepcopy(paragraphs[70])
        for paragraph in paragraphs:
            style = paragraph.find("w:pPr/w:pStyle", NS)
            if style is not None and style.attrib.get(qn("w", "val")) == "PSEEquation":
                ppr = paragraph.find("w:pPr", NS)
                if ppr is not None:
                    self.equation_ppr_template = copy.deepcopy(ppr)
                break
        sect_p = paragraphs[8].find("w:pPr/w:sectPr", NS)
        final_sect = body.find("w:sectPr", NS)
        if sect_p is None or final_sect is None:
            raise ValueError("Could not find the expected section properties in template.docx.")
        self.body_section_break = copy.deepcopy(sect_p)
        self.final_section = copy.deepcopy(final_sect)

    def _compute_max_docpr_id(self, root: ET.Element) -> int:
        max_id = 100
        for node in root.findall(".//wp:docPr", NS):
            try:
                max_id = max(max_id, int(node.attrib.get("id", "0")))
            except ValueError:
                pass
        return max_id

    def _build_document(self, document: ET.Element, media: MediaManager) -> None:
        body = document.find("w:body", NS)
        if body is None:
            raise ValueError("Template DOCX body missing.")
        for child in list(body):
            body.remove(child)

        metadata = self.parsed.metadata
        body.append(self._paragraph("PSETitle", self.inline.to_runs(metadata["PaperTitle"])))
        body.append(self._paragraph("PSEAuthorList", self.inline.to_runs(metadata["PaperAuthors"])))
        for line in split_tex_lines(metadata["PaperAffiliations"]):
            body.append(self._paragraph("PSEAuthorAffiliation", self.inline.to_runs(line)))
        body.append(self._paragraph("PSEAbstractHead", self.inline.to_runs("Abstract")))
        body.append(self._paragraph("PSEAbstractText", self.inline.to_runs(metadata["PaperAbstract"])))
        body.append(self._paragraph("PSEKeywords", self.inline.to_runs(f"Keywords: {metadata['PaperKeywords']}")))
        body.append(self._section_break_paragraph())

        figure_index = 0
        table_index = 0
        equation_index = 0
        for block in self.parsed.blocks:
            if isinstance(block, HeadingBlock):
                style = {1: "PSEHead1", 2: "PSEHead2", 3: "PSEHead3", 4: "PSEHead4"}[block.level]
                body.append(self._paragraph(style, self.inline.to_runs(block.title)))
            elif isinstance(block, ParagraphBlock):
                body.append(self._paragraph("PSEText", self.inline.to_runs(block.text)))
            elif isinstance(block, FigureBlock):
                figure_index += 1
                image_path = resolve_image_path(self.tex_path.parent, block.path)
                rel_id = media.add_image(image_path)
                body.append(self._figure_paragraph(rel_id, image_path, block.width_hint, block.wide))
                caption = f"Figure {figure_index}. {self.inline.to_plain(block.caption)}"
                body.append(self._paragraph("PSEFigureAndCaption", self.inline.to_runs(caption)))
            elif isinstance(block, TableBlock):
                table_index += 1
                caption = f"Table {table_index}: {self.inline.to_plain(block.caption)}"
                body.append(self._paragraph("PSETableCaption", self.inline.to_runs(caption)))
                body.append(self._table_element(block.rows))
            elif isinstance(block, EquationBlock):
                equation_index += 1
                body.append(self._equation_paragraph(block.body, equation_index))
            elif isinstance(block, CodeBlock):
                for line in block.lines:
                    body.append(self._paragraph("PSECode", [RunSpec(line)] if line else []))
            elif isinstance(block, DefinitionBlock):
                body.append(self._paragraph("PSETheoremTitle", self.inline.to_runs(block.title)))
                body.append(self._paragraph("PSEDefinition", self.inline.to_runs(block.body)))
            elif isinstance(block, CommandBlock) and block.name == "references":
                body.append(self._paragraph("PSEHead1", self.inline.to_runs("References")))
                for entry in self._selected_bib_entries():
                    body.append(self._paragraph("PSECitation", self.inline.to_runs(self._format_bib_entry(entry))))
            elif isinstance(block, CommandBlock) and block.name == "copyright":
                body.append(self._paragraph("PSECopyright", self.inline.to_runs(metadata["PaperCopyrightText"])))
                body.append(self._copyright_logo_paragraph())

        body.append(self.final_section if self.final_section is not None else ET.Element(qn("w", "sectPr")))

    def _selected_bib_entries(self) -> list[BibEntry]:
        if self.parsed.nocite_all or not self.parsed.cited_keys:
            return self.bib_entries
        cited = set(self.parsed.cited_keys)
        return [entry for entry in self.bib_entries if entry.key in cited]

    def _format_bib_entry(self, entry: BibEntry) -> str:
        fields = entry.fields
        authors = clean_bib_value(fields.get("author", "")).replace(" and ", ", ")
        if entry.entry_type == "article":
            journal = clean_bib_value(fields.get("journaltitle", fields.get("journal", "")))
            volume = clean_bib_value(fields.get("volume", ""))
            pages = clean_bib_value(fields.get("pages", ""))
            year = clean_bib_value(fields.get("date", fields.get("year", "")))
            doi = clean_bib_value(fields.get("doi", ""))
            base = f"{authors}. {clean_bib_value(fields.get('title', ''))}. {journal} {volume}:{pages} ({year})"
            return f"{base} \\url{{https://doi.org/{doi}}}" if doi else base
        if entry.entry_type == "patent":
            return (
                f"{authors}. {clean_bib_value(fields.get('title', ''))}. "
                f"{clean_bib_value(fields.get('type', 'Patent'))} {clean_bib_value(fields.get('number', ''))}. "
                f"({clean_bib_value(fields.get('date', fields.get('year', '')))})"
            )
        if entry.entry_type == "book":
            base = (
                f"{authors}. {clean_bib_value(fields.get('title', ''))}. "
                f"{clean_bib_value(fields.get('publisher', ''))} "
                f"({clean_bib_value(fields.get('date', fields.get('year', '')))})"
            )
            isbn = clean_bib_value(fields.get("isbn", ""))
            return f"{base}. ISBN {isbn}" if isbn else base
        if entry.entry_type in {"incollection", "inbook"}:
            editors = clean_bib_value(fields.get("editor", "")).replace(" and ", ", ")
            return (
                f"{authors}. {clean_bib_value(fields.get('title', ''))}. In: "
                f"{clean_bib_value(fields.get('booktitle', ''))}. Ed: {editors}. "
                f"{clean_bib_value(fields.get('publisher', ''))} "
                f"({clean_bib_value(fields.get('date', fields.get('year', '')))})"
            )
        if entry.entry_type == "online":
            author = authors or clean_bib_value(fields.get("title", ""))
            return (
                f"{author}. \\url{{{clean_bib_value(fields.get('url', ''))}}} "
                f"[accessed {clean_bib_value(fields.get('urldate', 'Date'))}]"
            )
        return normalize_space(f"{authors}. {clean_bib_value(fields.get('title', ''))}")

    def _equation_paragraph(self, value: str, equation_index: int) -> ET.Element:
        paragraph = ET.Element(qn("w", "p"))
        if self.equation_ppr_template is not None:
            paragraph.append(copy.deepcopy(self.equation_ppr_template))
        else:
            pPr = ET.SubElement(paragraph, qn("w", "pPr"))
            pStyle = ET.SubElement(pPr, qn("w", "pStyle"))
            pStyle.set(qn("w", "val"), "PSEEquation")
        try:
            equation_root = ET.SubElement(paragraph, qn("m", "oMath"))
            self._append_math_node(equation_root, LatexMathParser(value).parse())
        except Exception:
            paragraph.append(self._run(RunSpec(self.inline.to_plain(value))))
        paragraph.append(self._tab_run())
        paragraph.append(self._tab_run())
        paragraph.append(self._tab_run())
        paragraph.append(self._tab_text_run(f"({equation_index})"))
        return paragraph

    def _append_math_node(self, parent: ET.Element, node: MathNode) -> None:
        if isinstance(node, MathSequenceNode):
            for item in node.items:
                self._append_math_node(parent, item)
            return
        if isinstance(node, MathTextNode):
            if node.text:
                parent.append(self._math_run(node.text, literal=node.literal))
            return
        if isinstance(node, MathFractionNode):
            frac = ET.SubElement(parent, qn("m", "f"))
            frac_pr = ET.SubElement(frac, qn("m", "fPr"))
            self._append_math_ctrl_pr(frac_pr)
            num = ET.SubElement(frac, qn("m", "num"))
            den = ET.SubElement(frac, qn("m", "den"))
            self._append_math_node(num, node.numerator)
            self._append_math_node(den, node.denominator)
            return
        if isinstance(node, MathScriptNode):
            if node.sub is not None and node.sup is not None:
                script = ET.SubElement(parent, qn("m", "sSubSup"))
                script_pr = ET.SubElement(script, qn("m", "sSubSupPr"))
                base = ET.SubElement(script, qn("m", "e"))
                sub = ET.SubElement(script, qn("m", "sub"))
                sup = ET.SubElement(script, qn("m", "sup"))
                self._append_math_ctrl_pr(script_pr)
                self._append_math_node(base, node.base)
                self._append_math_node(sub, node.sub)
                self._append_math_node(sup, node.sup)
                return
            if node.sub is not None:
                script = ET.SubElement(parent, qn("m", "sSub"))
                script_pr = ET.SubElement(script, qn("m", "sSubPr"))
                base = ET.SubElement(script, qn("m", "e"))
                sub = ET.SubElement(script, qn("m", "sub"))
                self._append_math_ctrl_pr(script_pr)
                self._append_math_node(base, node.base)
                self._append_math_node(sub, node.sub)
                return
            if node.sup is not None:
                script = ET.SubElement(parent, qn("m", "sSup"))
                script_pr = ET.SubElement(script, qn("m", "sSupPr"))
                base = ET.SubElement(script, qn("m", "e"))
                sup = ET.SubElement(script, qn("m", "sup"))
                self._append_math_ctrl_pr(script_pr)
                self._append_math_node(base, node.base)
                self._append_math_node(sup, node.sup)
                return
            self._append_math_node(parent, node.base)
            return
        if isinstance(node, MathNaryNode):
            nary = ET.SubElement(parent, qn("m", "nary"))
            nary_pr = ET.SubElement(nary, qn("m", "naryPr"))
            symbol = ET.SubElement(nary_pr, qn("m", "chr"))
            symbol.set(qn("m", "val"), node.operator)
            lim_loc = ET.SubElement(nary_pr, qn("m", "limLoc"))
            lim_loc.set(qn("m", "val"), "subSup")
            self._append_math_ctrl_pr(nary_pr)
            if node.sub is not None:
                sub = ET.SubElement(nary, qn("m", "sub"))
                self._append_math_node(sub, node.sub)
            if node.sup is not None:
                sup = ET.SubElement(nary, qn("m", "sup"))
                self._append_math_node(sup, node.sup)
            expr = ET.SubElement(nary, qn("m", "e"))
            self._append_math_node(expr, node.body)
            return
        if isinstance(node, MathMatrixNode):
            matrix = ET.SubElement(parent, qn("m", "m"))
            matrix_pr = ET.SubElement(matrix, qn("m", "mPr"))
            mcs = ET.SubElement(matrix_pr, qn("m", "mcs"))
            mc = ET.SubElement(mcs, qn("m", "mc"))
            mc_pr = ET.SubElement(mc, qn("m", "mcPr"))
            column_count = max((len(row) for row in node.rows), default=1)
            count = ET.SubElement(mc_pr, qn("m", "count"))
            count.set(qn("m", "val"), str(column_count))
            column_align = ET.SubElement(mc_pr, qn("m", "mcJc"))
            column_align.set(qn("m", "val"), "center")
            self._append_math_ctrl_pr(matrix_pr)
            for row in node.rows:
                matrix_row = ET.SubElement(matrix, qn("m", "mr"))
                for cell in row:
                    cell_expr = ET.SubElement(matrix_row, qn("m", "e"))
                    self._append_math_node(cell_expr, cell)
            return
        raise ValueError(f"Unsupported math node: {type(node)!r}")

    def _append_math_ctrl_pr(self, parent: ET.Element) -> None:
        ctrl_pr = ET.SubElement(parent, qn("m", "ctrlPr"))
        ctrl_pr.append(self._math_word_rpr())

    def _math_word_rpr(self) -> ET.Element:
        rPr = ET.Element(qn("w", "rPr"))
        fonts = ET.SubElement(rPr, qn("w", "rFonts"))
        fonts.set(qn("w", "ascii"), "Cambria Math")
        fonts.set(qn("w", "hAnsi"), "Cambria Math")
        return rPr

    def _math_run(self, text: str, *, literal: bool = False) -> ET.Element:
        run = ET.Element(qn("m", "r"))
        if literal or all(ch in "=+-/()[]" for ch in text):
            math_rpr = ET.SubElement(run, qn("m", "rPr"))
            if literal:
                ET.SubElement(math_rpr, qn("m", "lit"))
            if all(ch in "=+-/()[]" for ch in text):
                style = ET.SubElement(math_rpr, qn("m", "sty"))
                style.set(qn("m", "val"), "p")
        run.append(self._math_word_rpr())
        math_text = ET.SubElement(run, qn("m", "t"))
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            math_text.set(f"{{{XML_NS}}}space", "preserve")
        math_text.text = text
        return run

    def _tab_run(self) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        ET.SubElement(run, qn("w", "tab"))
        return run

    def _tab_text_run(self, text: str) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        ET.SubElement(run, qn("w", "tab"))
        t = ET.SubElement(run, qn("w", "t"))
        t.text = text
        return run

    def _paragraph(self, style_id: str, runs: list[RunSpec]) -> ET.Element:
        p = ET.Element(qn("w", "p"))
        pPr = ET.SubElement(p, qn("w", "pPr"))
        pStyle = ET.SubElement(pPr, qn("w", "pStyle"))
        pStyle.set(qn("w", "val"), style_id)
        for run in runs:
            if run.hyperlink:
                hyperlink = ET.SubElement(p, qn("w", "hyperlink"))
                hyperlink.set(qn("r", "id"), self._hyperlink_rel(run.hyperlink))
                hyperlink.append(self._run(run))
            else:
                p.append(self._run(run))
        return p

    def _run(self, spec: RunSpec) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        if spec.bold or spec.italic or spec.superscript:
            rPr = ET.SubElement(run, qn("w", "rPr"))
            if spec.bold:
                ET.SubElement(rPr, qn("w", "b"))
                ET.SubElement(rPr, qn("w", "bCs"))
            if spec.italic:
                ET.SubElement(rPr, qn("w", "i"))
                ET.SubElement(rPr, qn("w", "iCs"))
            if spec.superscript:
                vert = ET.SubElement(rPr, qn("w", "vertAlign"))
                vert.set(qn("w", "val"), "superscript")
        text = ET.SubElement(run, qn("w", "t"))
        if spec.text.startswith(" ") or spec.text.endswith(" ") or "  " in spec.text:
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = spec.text
        return run

    def _hyperlink_rel(self, url: str) -> str:
        if self.rel_root is None:
            raise RuntimeError("Relationship root is not initialized.")
        if url in self.hyperlinks:
            return self.hyperlinks[url]
        rid = next_rel_id(self.rel_root)
        rel = ET.SubElement(self.rel_root, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", self.HYPERLINK_REL_TYPE)
        rel.set("Target", url)
        rel.set("TargetMode", "External")
        self.hyperlinks[url] = rid
        return rid

    def _section_break_paragraph(self) -> ET.Element:
        p = ET.Element(qn("w", "p"))
        pPr = ET.SubElement(p, qn("w", "pPr"))
        pStyle = ET.SubElement(pPr, qn("w", "pStyle"))
        pStyle.set(qn("w", "val"), "PSEKeywords")
        if self.body_section_break is not None:
            pPr.append(copy.deepcopy(self.body_section_break))
        return p

    def _figure_paragraph(self, rel_id: str, image_path: Path, width_hint: str | None, wide: bool) -> ET.Element:
        if self.figure_paragraph_template is None:
            raise ValueError("Figure template paragraph is unavailable.")
        paragraph = copy.deepcopy(self.figure_paragraph_template)
        self.max_docpr_id += 1
        width_emu, height_emu = compute_image_extent(image_path, width_hint, wide)
        for extent in paragraph.findall(".//wp:extent", NS):
            extent.set("cx", str(width_emu))
            extent.set("cy", str(height_emu))
        for ext in paragraph.findall(".//a:ext", NS):
            ext.set("cx", str(width_emu))
            ext.set("cy", str(height_emu))
        for blip in paragraph.findall(".//a:blip", NS):
            blip.set(qn("r", "embed"), rel_id)
        for docpr in paragraph.findall(".//wp:docPr", NS):
            docpr.set("id", str(self.max_docpr_id))
            docpr.set("name", f"Picture {self.max_docpr_id}")
            docpr.set("descr", image_path.name)
        for cnvpr in paragraph.findall(".//pic:cNvPr", NS):
            cnvpr.set("id", str(self.max_docpr_id))
            cnvpr.set("name", image_path.name)
            cnvpr.set("descr", image_path.name)
        return paragraph

    def _copyright_logo_paragraph(self) -> ET.Element:
        if self.cc_logo_paragraph_template is None:
            raise ValueError("Copyright logo paragraph template is unavailable.")
        return copy.deepcopy(self.cc_logo_paragraph_template)

    def _table_element(self, rows: list[list[str]]) -> ET.Element:
        tbl = ET.Element(qn("w", "tbl"))
        tblPr = copy.deepcopy(self.sample_table_pr) if self.sample_table_pr is not None else ET.Element(qn("w", "tblPr"))
        tbl.append(tblPr)
        column_count = max((len(row) for row in rows), default=1)
        widths = compute_column_widths(rows, column_count)
        grid = ET.SubElement(tbl, qn("w", "tblGrid"))
        for width in widths:
            col = ET.SubElement(grid, qn("w", "gridCol"))
            col.set(qn("w", "w"), str(width))
        for row_index, row in enumerate(rows):
            tr = ET.SubElement(tbl, qn("w", "tr"))
            if row_index == 0:
                trPr = ET.SubElement(tr, qn("w", "trPr"))
                cnf = ET.SubElement(trPr, qn("w", "cnfStyle"))
                for key, value in {
                    "val": "100000000000",
                    "firstRow": "1",
                    "lastRow": "0",
                    "firstColumn": "0",
                    "lastColumn": "0",
                    "oddVBand": "0",
                    "evenVBand": "0",
                    "oddHBand": "0",
                    "evenHBand": "0",
                    "firstRowFirstColumn": "0",
                    "firstRowLastColumn": "0",
                    "lastRowFirstColumn": "0",
                    "lastRowLastColumn": "0",
                }.items():
                    cnf.set(qn("w", key), value)
            padded = row + [""] * (column_count - len(row))
            for col_index, cell_text in enumerate(padded):
                tc = ET.SubElement(tr, qn("w", "tc"))
                tcPr = ET.SubElement(tc, qn("w", "tcPr"))
                tcW = ET.SubElement(tcPr, qn("w", "tcW"))
                tcW.set(qn("w", "w"), str(widths[col_index]))
                tcW.set(qn("w", "type"), "dxa")
                p = ET.SubElement(tc, qn("w", "p"))
                pPr = ET.SubElement(p, qn("w", "pPr"))
                pStyle = ET.SubElement(pPr, qn("w", "pStyle"))
                pStyle.set(qn("w", "val"), "PSETableContent")
                if col_index == column_count - 1 and looks_numeric(cell_text):
                    jc = ET.SubElement(pPr, qn("w", "jc"))
                    jc.set(qn("w", "val"), "right")
                if row_index == 0:
                    rPr = ET.SubElement(pPr, qn("w", "rPr"))
                    ET.SubElement(rPr, qn("w", "b"))
                    ET.SubElement(rPr, qn("w", "bCs"))
                runs = self.inline.to_runs(cell_text) or [RunSpec("")]
                for run in runs:
                    current = copy.deepcopy(run)
                    if row_index == 0:
                        current.bold = True
                    p.append(self._run(current))
        return tbl

    def _update_headers_and_footers(self, package_entries: dict[str, bytes]) -> None:
        metadata = self.parsed.metadata
        if "word/header2.xml" in package_entries:
            tree = ET.ElementTree(ET.fromstring(package_entries["word/header2.xml"]))
            header = tree.getroot()
            paragraphs = header.findall("w:p", NS)
            replacements = [
                "",
                metadata.get("HeaderArticleType", ""),
                metadata.get("HeaderReviewType", ""),
                metadata.get("HeaderConference", ""),
                metadata.get("HeaderLocation", ""),
            ]
            for idx, text in enumerate(replacements):
                if idx >= len(paragraphs):
                    break
                self._replace_paragraph_text(paragraphs[idx], self.inline.to_plain(text), preserve_drawings=idx in {0, 1})
            package_entries["word/header2.xml"] = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)

        self._rewrite_footer(package_entries, "word/footer2.xml", self.inline.to_plain(metadata.get("FirstPageFooterID", "")), self.inline.to_plain(metadata.get("FooterJournalInfo", "")))
        self._rewrite_footer(package_entries, "word/footer1.xml", self.inline.to_plain(metadata.get("BodyPageFooterID", "")), self.inline.to_plain(metadata.get("FooterJournalInfo", "")))

    def _replace_paragraph_text(self, paragraph: ET.Element, text: str, *, preserve_drawings: bool = False) -> None:
        pPr = copy.deepcopy(paragraph.find("w:pPr", NS))
        preserved: list[ET.Element] = []
        if preserve_drawings:
            for child in list(paragraph):
                if child.tag == qn("w", "r") and child.find("w:drawing", NS) is not None:
                    preserved.append(copy.deepcopy(child))
        for child in list(paragraph):
            paragraph.remove(child)
        if pPr is not None:
            paragraph.append(pPr)
        for run in preserved:
            paragraph.append(run)
        if text:
            paragraph.append(self._run(RunSpec(text)))

    def _rewrite_footer(self, package_entries: dict[str, bytes], package_path: str, prefix: str, journal_info: str) -> None:
        if package_path not in package_entries:
            return
        tree = ET.ElementTree(ET.fromstring(package_entries[package_path]))
        root = tree.getroot()
        paragraph = root.find("w:p", NS)
        if paragraph is None:
            package_entries[package_path] = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)
            return
        pPr = copy.deepcopy(paragraph.find("w:pPr", NS))
        for child in list(paragraph):
            paragraph.remove(child)
        if pPr is not None:
            paragraph.append(pPr)
        paragraph.append(self._footer_text_run(f"{prefix} {journal_info} "))
        paragraph.append(self._footer_field_run("begin"))
        instr = ET.SubElement(ET.Element(qn("w", "r")), qn("w", "instrText"))
        instr.set(f"{{{XML_NS}}}space", "preserve")
        instr.text = " PAGE  \\* Arabic  \\* MERGEFORMAT "
        paragraph.append(instr.getparent() if hasattr(instr, "getparent") else self._wrap_instr(instr))
        paragraph.append(self._footer_field_run("separate"))
        paragraph.append(self._footer_text_run("1"))
        paragraph.append(self._footer_field_run("end"))
        package_entries[package_path] = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)

    def _wrap_instr(self, instr_text: ET.Element) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        run.append(instr_text)
        return run

    def _footer_text_run(self, text: str) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        t = ET.SubElement(run, qn("w", "t"))
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        return run

    def _footer_field_run(self, field_type: str) -> ET.Element:
        run = ET.Element(qn("w", "r"))
        field = ET.SubElement(run, qn("w", "fldChar"))
        field.set(qn("w", "fldCharType"), field_type)
        return run


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert the LaTeX manuscript back into a DOCX that reuses template.docx styling."
    )
    parser.add_argument("--input", default="main.tex", help="Path to the LaTeX file. Defaults to main.tex.")
    parser.add_argument("--template", default="template.docx", help="Path to the Word template DOCX.")
    parser.add_argument("--output", default="main-from-latex.docx", help="Path to the generated DOCX.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    tex_path = Path(args.input).resolve()
    template_path = Path(args.template).resolve()
    output_path = Path(args.output).resolve()
    if not tex_path.is_file():
        raise FileNotFoundError(f"LaTeX input not found: {tex_path}")
    if not template_path.is_file():
        raise FileNotFoundError(f"Template DOCX not found: {template_path}")
    DocxTemplateConverter(template_path=template_path, tex_path=tex_path, output_path=output_path).convert()
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
