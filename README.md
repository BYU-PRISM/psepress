# Streamlit LaTeX-to-Word Converter

This folder is a self-contained Streamlit app for generating conference-specific Word and LaTeX templates and converting a zipped LaTeX manuscript into a Word document. A single ESCAPE-based `template.docx` supplies the layout; the conference header is selected or customized for each generated template and conversion.

## Hosted app

The app is also available online through Streamlit Community Cloud:

[psepress.streamlit.app](https://psepress.streamlit.app)

Anyone can use the hosted app to create a conference template or upload a zipped LaTeX archive and convert it into a Word document without running the project locally.

## Files

- `app.py`: Streamlit UI
- `conference_templates.py`: shared conference presets, validation, date formatting, and Word/LaTeX template generators
- `latex_to_word.py`: shared converter logic used by both the Streamlit app and the offline batch tools
- `batch_convert_archives.py`: offline batch converter for folders of submission zip archives
- `convert_submissions.ps1`: Windows PowerShell wrapper for offline batch conversion
- `convert_submissions.sh`: Linux bash wrapper for offline batch conversion
- `template.docx`: ESCAPE-based master Word template reused for styling and layout
- `latex_template/latex_template.zip`: complete downloadable LaTeX project used by the template generator
- `requirements.txt`: Python dependencies for local runs or Streamlit Community Cloud

## Local run

```powershell
cd .\psepress
python -m pip install -r requirements.txt
streamlit run app.py
```

## Conference template generator

The generator is at the top of the Streamlit app. Enter the conference name, city, optional state or region, country, start date, and end date. One submission creates two in-memory downloads:

- a Word `.docx` template with the conference header updated
- a complete LaTeX project `.zip` with the same conference information in `main.tex`

The generator keeps the article type, review type, body content, logos, styles, and footers from the master templates.

## Conference selection

The web converter and command-line tools provide these built-in conference keys:

- `escape-37-2027` (default)
- `pse-2027`
- `focapo-cpc-2027`
- `latex` to read conference information from the manuscript
- `custom` with an explicit conference name and location/date line

The LaTeX template supports the same presets:

```tex
\PSESelectConference{escape-37-2027}
```

For another conference, replace the selector with:

```tex
\PSESetConference{Conference Name}{City, Country, 6-9 June 2027}
```

Legacy `\HeaderConference` and `\HeaderLocation` definitions remain supported by the converter.

## Direct conversion

```powershell
python .\latex_to_word.py --input .\latex_template\main.tex --output .\main-from-latex.docx --conference pse-2027
```

For custom header text:

```powershell
python .\latex_to_word.py --input .\latex_template\main.tex --output .\main-from-latex.docx --conference custom --conference-name "Example Conference 2028" --conference-location "Denver, Colorado, USA, 3-6 May 2028"
```

## Offline batch conversion

This directory is self-contained for both hosted and offline use. The batch tools here use the same local `latex_to_word.py` and `template.docx` as the web app.

Python entry point:

```powershell
cd .\psepress
python .\batch_convert_archives.py --input-dir .\submissions\input --output-dir .\submissions\output --conference escape-37-2027
```

Windows PowerShell wrapper:

```powershell
cd .\psepress
powershell -ExecutionPolicy Bypass -File .\convert_submissions.ps1 -InputDir .\submissions\input -OutputDir .\submissions\output -Conference escape-37-2027
```

Linux bash wrapper:

```bash
cd ./psepress
./convert_submissions.sh ./submissions/input ./submissions/output --conference escape-37-2027
```

The batch run writes one `.docx` per input archive and a `conversion-report.csv` summary in the output folder.

## Deploy on Streamlit Community Cloud

1. Put the contents of this folder into a GitHub repository, or make this folder the root of a new repo.
2. In Streamlit Community Cloud, create a new app from that repo.
3. Set the main file path to `app.py`.

## Upload format

Upload a `.zip` archive that includes:

- your manuscript entry file, typically `main.tex`
- `refs.bib` if the manuscript uses it
- figures and any other referenced files
- any additional `.tex` files included with `\input{...}`

Keep the same relative paths your manuscript expects.

## Notes

- The app does not compile LaTeX; it parses the manuscript source and rebuilds a `.docx` using `template.docx`.
- The converter is template-aware for this project rather than a general LaTeX-to-Word engine.
- For best results, upload the same project structure you use locally.
