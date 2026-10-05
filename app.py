from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from sorter import (
    CHAPTERS,
    build_questions,
    create_docx,
)

MAX_FILES = 30
ALLOWED_EXTENSIONS = {
    "pdf", "docx", "txt", "md",
    "zip",
    "png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp",
}

st.set_page_config(
    page_title="CBSE Class 10 Science Question Sorter",
    page_icon="🧪",
    layout="wide",
)

st.title("🧪 CBSE Class 10 Science — Question Sorter")
st.caption("English-language questions only • All question types • Chapter-wise")

with st.sidebar:
    st.header("Settings")
    use_ocr = st.checkbox(
        "Use OCR for scanned PDFs/images",
        value=True,
        help="Recommended for scanned papers. OCR uses English + Hindi recognition so Hindi questions are not mistaken for English.",
    )
    english_threshold = st.slider(
        "English-language strictness",
        min_value=0.50,
        max_value=0.95,
        value=0.72,
        step=0.01,
        help="Higher values are stricter. The detector does NOT treat isolated English letters, formulas, units or option labels as English prose.",
    )
    chapter_confidence = st.slider(
        "Chapter matching strictness",
        min_value=0.20,
        max_value=0.80,
        value=0.42,
        step=0.01,
        help="Higher values send more uncertain questions to Review / Unclassified instead of forcing a chapter.",
    )
    include_source = st.checkbox(
        "Include source filename under each question",
        value=False,
    )

st.info(
    "Upload up to 30 question-paper files at once. You can upload individual PDFs/DOCX/images, "
    "or a ZIP containing them. Hindi questions are excluded even when they contain variables, formulas, units, symbols or option labels in English."
)

uploaded = st.file_uploader(
    "Upload question papers",
    type=sorted(ALLOWED_EXTENSIONS),
    accept_multiple_files=True,
    help="Supported: PDF, DOCX, TXT, MD, ZIP, PNG, JPG/JPEG, WEBP, TIFF and BMP.",
)

if uploaded:
    if len(uploaded) > MAX_FILES:
        st.error(f"Maximum {MAX_FILES} uploaded files at once. Please remove {len(uploaded) - MAX_FILES} file(s) and try again.")
        st.stop()

    total_mb = sum(len(f.getvalue()) for f in uploaded) / (1024 * 1024)
    st.write(f"**{len(uploaded)} file(s)** selected • **{total_mb:.1f} MB** total")

    if st.button("🚀 Sort all questions", type="primary", use_container_width=True):
        with st.spinner("Reading papers, detecting English questions, and sorting by chapter…"):
            with tempfile.TemporaryDirectory(prefix="cbse_streamlit_") as td:
                work = Path(td)
                zip_path = work / "uploaded_papers.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i, f in enumerate(uploaded, 1):
                        # Prefix avoids filename collisions between uploads.
                        safe = Path(f.name).name
                        zf.writestr(f"{i:02d}_{safe}", f.getvalue())

                try:
                    questions, stats, errors = build_questions(
                        zip_path,
                        use_ocr=use_ocr,
                        english_threshold=english_threshold,
                        min_chapter_conf=chapter_confidence,
                        max_files=MAX_FILES,
                    )
                    output = work / "CBSE_Class10_Science_English_Questions_Chapterwise.docx"
                    create_docx(
                        questions,
                        output,
                        title="CBSE Class 10 Science — English Questions, Chapter-wise",
                        include_source=include_source,
                        include_review=True,
                    )
                    output_bytes = output.read_bytes()
                except Exception as exc:
                    st.error(f"Processing failed: {exc}")
                    st.stop()

        st.success("Done! Your chapter-wise question bank is ready.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Files processed", stats.get("files", 0))
        c2.metric("Candidate questions", stats.get("candidate_questions", 0))
        c3.metric("English questions kept", stats.get("sorted", 0) + stats.get("review", 0))
        c4.metric("Hindi/non-English ignored", stats.get("ignored_non_english", 0))

        st.download_button(
            "⬇️ Download chapter-wise DOCX",
            data=output_bytes,
            file_name="CBSE_Class10_Science_English_Questions_Chapterwise.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.subheader("Chapter summary")
        rows = []
        for ch in CHAPTERS:
            count = sum(1 for q in questions if q.chapter_no == ch["no"])
            rows.append({"Chapter": f'{ch["no"]}. {ch["name"]}', "Questions": count})
        review_count = sum(1 for q in questions if q.chapter_no is None)
        rows.append({"Chapter": "Review / Unclassified", "Questions": review_count})
        st.dataframe(rows, use_container_width=True, hide_index=True)

        if errors:
            with st.expander(f"⚠️ {len(errors)} file(s) had processing errors"):
                for err in errors:
                    st.write(f"- {err}")

        with st.expander("Language-filter rule used"):
            st.write(
                "The sorter does not decide that a question is English merely because it contains Latin/English letters. "
                "Meaningful English prose is required. Devanagari/Hindi questions are rejected when the English material "
                "is only variables, formulas, units, symbols, option labels, or short fragments."
            )
else:
    st.markdown("### Supported formats")
    st.markdown("**PDF · DOCX · ZIP · TXT · MD · PNG · JPG/JPEG · WEBP · TIFF · BMP**")
    st.markdown("### Output")
    st.markdown("One clean **DOCX** containing all retained English questions, arranged chapter-wise, with a **Review / Unclassified** section for uncertain chapter matches.")
