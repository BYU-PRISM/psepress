# Streamlit LaTeX-to-Word Converter

This folder is a self-contained Streamlit app for converting a zipped LaTeX manuscript into a Word document using the bundled `template.docx` and the existing template-aware converter.

## Files

- `app.py`: Streamlit UI
- `latex_to_word.py`: converter logic copied from the main project
- `template.docx`: Word template reused for styling and layout
- `requirements.txt`: Python dependencies for local runs or Streamlit Community Cloud

## Local run

```powershell
cd .\streamlit_converter_app
python -m pip install -r requirements.txt
streamlit run app.py
```

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
