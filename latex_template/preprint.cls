\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{preprint}[2026/03/25 Local single-column preprint wrapper class]

\DeclareOption*{\PassOptionsToClass{\CurrentOption}{article}}
\ProcessOptions\relax
\LoadClass{article}
