# CBSE Class 10 Science Question Sorter — Streamlit

A Streamlit app that takes up to **30 question-paper files at once**, extracts **all English-language questions** (not just MCQs), and sorts them chapter-wise for CBSE Class 10 Science.

## Important language rule

The app **does not** use “contains English letters” as its English detector.

A Hindi/Devanagari question is ignored even if it contains a few English characters because of:
- variables such as `x`, `y`, `A`, `B`
- formulas and chemical symbols
- units such as `V`, `A`, `cm`, `kg`
- option labels such as A/B/C/D
- mathematical symbols
- short English fragments

Meaningful English prose is required. OCR uses **English + Hindi** recognition for scanned material so the language filter can make a better decision.

## Supported uploads

- PDF
- DOCX
- ZIP containing supported files
- TXT
- MD
- PNG
- JPG / JPEG
- WEBP
- TIFF
- BMP

The Streamlit uploader accepts **up to 30 files at once**. A ZIP counts as one uploaded file and may contain multiple supported papers.

## Output

The app creates one DOCX containing:

- all retained English questions
- MCQs
- Assertion–Reason questions
- short/long answer questions
- case/source-based questions
- sub-parts
- chapter-wise organization
- a Review / Unclassified section for questions whose chapter match is uncertain

Source filename lines are **off by default**.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload all files in this folder to the repository root.
3. On Streamlit Community Cloud, choose the repository and set the main file to `app.py`.
4. Deploy.

`requirements.txt` installs the Python packages and `packages.txt` installs Tesseract OCR on Streamlit's Linux environment.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- The app is designed for CBSE Class 10 Science chapter-wise question sorting.
- Duplicate questions are kept because repeated questions across papers can be useful in a question bank.
- Diagrams are not reconstructed in the DOCX; the extracted question text is retained.
- OCR quality depends on scan quality.
