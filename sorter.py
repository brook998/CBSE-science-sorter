#!/usr/bin/env python3
"""
CBSE Class 10 Science – English Question Sorter

Input: ZIP containing question papers (PDF/DOCX/TXT; images optionally via OCR)
Output: One DOCX with English-language questions grouped chapter-wise.

Key design rule:
English detection NEVER means "contains English letters".
A question is accepted as English only when there is meaningful English prose
and it is not predominantly Devanagari/Hindi. Single letters, variables,
formulas, units, option labels (A/B/C/D), symbols, and common science notation
are deliberately ignored as language evidence.
"""

from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

# Optional imports are checked at runtime so the script can show a friendly
# message instead of crashing during import.
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION
    from docx.shared import Inches, Pt, RGBColor
except ImportError:
    Document = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import pytesseract
except ImportError:
    pytesseract = None


CHAPTERS = [
    {
        "no": 1,
        "name": "Chemical Reactions and Equations",
        "keywords": {
            "chemical reaction": 8, "chemical equation": 8, "balanced equation": 8,
            "balancing": 6, "combination reaction": 8, "decomposition": 7,
            "displacement reaction": 8, "double displacement": 8, "precipitation": 6,
            "oxidation": 6, "reduction": 6, "redox": 8, "exothermic": 5,
            "endothermic": 5, "corrosion": 7, "rust": 7, "rusting": 8, "rancidity": 6, "oxidising": 5,
            "reducing agent": 5, "oxidising agent": 5, "reaction type": 5,
        },
    },
    {
        "no": 2,
        "name": "Acids, Bases and Salts",
        "keywords": {
            "acid": 4, "base": 4, "salts": 4, "pH": 7, "indicator": 6,
            "universal indicator": 7, "litmus": 6, "neutralisation": 7,
            "hydrogen ion": 6, "hydroxide ion": 6, "chlor-alkali": 8,
            "bleaching powder": 8, "baking soda": 8, "washing soda": 8,
            "plaster of paris": 8, "water of crystallisation": 8,
            "gypsum": 7, "sodium hydroxide": 7, "lime water": 5,
            "olfactory indicator": 8, "common salt": 6, "electrolyte": 7, "conducts electricity": 7, "ions in solution": 7,
        },
    },
    {
        "no": 3,
        "name": "Metals and Non-metals",
        "keywords": {
            "metal": 4, "non-metal": 5, "nonmetal": 5, "reactivity series": 9,
            "reactivity": 5, "ore": 5, "metallurgy": 7, "extraction": 5,
            "roasting": 6, "calcination": 6, "thermite": 7, "ionic compound": 8,
            "lustre": 5, "malleable": 6, "ductile": 6, "conductivity": 3,
            "amphoteric": 7, "displacement": 4, "alloy": 6, "electrolysis": 4,
        },
    },
    {
        "no": 4,
        "name": "Carbon and its Compounds",
        "keywords": {
            "carbon compound": 8, "covalent bond": 8, "hydrocarbon": 7,
            "saturated": 5, "unsaturated": 5, "homologous series": 8,
            "functional group": 8, "isomer": 8, "isomers": 8, "ethanol": 8,
            "ethanoic acid": 8, "acetic acid": 7, "esterification": 9,
            "saponification": 8, "soap": 6, "detergent": 6, "methane": 6,
            "ethene": 6, "ethyne": 6, "alkane": 6, "alkene": 6, "alkyne": 6,
            "carbon chain": 6, "electron dot structure": 7,
        },
    },
    {
        "no": 5,
        "name": "Life Processes",
        "keywords": {
            "life processes": 9, "nutrition": 7, "photosynthesis": 8, "respiration": 7,
            "aerobic": 5, "anaerobic": 5, "digestion": 7, "digestive": 6,
            "alimentary canal": 7, "stomach": 5, "small intestine": 7, "villi": 7,
            "transportation": 7, "blood": 5, "heart": 6, "circulation": 7,
            "artery": 5, "vein": 5, "haemoglobin": 6, "xylem": 7, "phloem": 7,
            "transpiration": 7, "excretion": 7, "nephron": 8, "kidney": 7,
            "urine": 5, "stomata": 7, "gas exchange": 5, "autotrophic": 6,
            "heterotrophic": 6,
        },
    },
    {
        "no": 6,
        "name": "Control and Coordination",
        "keywords": {
            "control and coordination": 10, "nervous system": 8, "neuron": 8,
            "brain": 6, "spinal cord": 8, "reflex action": 9, "reflex arc": 9,
            "sensory neuron": 7, "motor neuron": 7, "synapse": 7, "hormone": 7,
            "endocrine": 7, "adrenaline": 8, "thyroxine": 7, "insulin": 6,
            "plant hormone": 7, "auxin": 8, "tropism": 8, "phototropism": 8,
            "geotropism": 8, "coordination": 5, "receptor": 6,
        },
    },
    {
        "no": 7,
        "name": "How do Organisms Reproduce?",
        "keywords": {
            "reproduction": 7, "asexual reproduction": 9, "sexual reproduction": 9,
            "binary fission": 8, "budding": 7, "spore formation": 7,
            "vegetative propagation": 8, "reproductive organ": 7, "gamete": 8,
            "fertilisation": 9, "fertilization": 9, "pollination": 8,
            "self-pollination": 8, "cross-pollination": 8, "zygote": 7,
            "embryo": 7, "implantation": 8, "menstruation": 8, "puberty": 6,
            "contraception": 9, "vasectomy": 8, "tubectomy": 8, "ovulation": 7,
            "testis": 6, "ovary": 6, "sperm": 6, "ovum": 6,
        },
    },
    {
        "no": 8,
        "name": "Heredity",
        "keywords": {
            "heredity": 9, "mendel": 9, "inheritance": 7, "trait": 6,
            "allele": 8, "gene": 6, "genotype": 8, "phenotype": 8,
            "dominant": 7, "recessive": 7, "chromosome": 7, "variation": 6,
            "monohybrid": 8, "dihybrid": 8, "f1 generation": 7, "f2 generation": 7,
            "sex determination": 8, "speciation": 8, "evolution": 6,
        },
    },
    {
        "no": 9,
        "name": "Light - Reflection and Refraction",
        "keywords": {
            "reflection of light": 9, "refraction of light": 9, "reflection": 5,
            "refraction": 7, "spherical mirror": 8, "concave mirror": 8,
            "convex mirror": 8, "mirror formula": 8, "focal length": 7,
            "magnification": 6, "principal focus": 7, "lens": 5, "convex lens": 8,
            "concave lens": 8, "lens formula": 8, "power of lens": 8,
            "refractive index": 8, "optical centre": 8, "ray diagram": 6,
            "real image": 5, "virtual image": 5,
        },
    },
    {
        "no": 10,
        "name": "The Human Eye and the Colourful World",
        "keywords": {
            "human eye": 10, "eye": 4, "retina": 7, "iris": 7, "pupil": 6,
            "accommodation": 7, "myopia": 8, "hypermetropia": 8, "presbyopia": 8,
            "cataract": 7, "near point": 6, "far point": 6, "power of accommodation": 8,
            "prism": 6, "dispersion": 8, "spectrum": 6, "atmospheric refraction": 9,
            "twinkling": 8, "scattering": 8, "tindall": 6, "rainbow": 7,
            "blue sky": 7, "advanced sunrise": 8, "delayed sunset": 8,
        },
    },
    {
        "no": 11,
        "name": "Electricity",
        "keywords": {
            "electric current": 8, "current": 4, "potential difference": 8,
            "voltage": 5, "resistance": 8, "ohm's law": 9, "ohm law": 9,
            "resistor": 5, "series combination": 5, "parallel combination": 5,
            "electric power": 8, "electrical energy": 8, "heating effect": 8,
            "joule's law": 9, "electric circuit": 6, "ammeter": 6, "voltmeter": 6,
            "kwh": 8, "watt": 5, "fuse": 5, "resistivity": 7,
            "household circuit": 8,
        },
    },
    {
        "no": 12,
        "name": "Magnetic Effects of Electric Current",
        "keywords": {
            "magnetic effect": 8, "magnetic field": 8, "magnetic field lines": 9,
            "magnetic field due to current": 10, "compass": 5, "solenoid": 8,
            "electromagnet": 8, "fleming's left hand rule": 10,
            "fleming left hand": 9, "right hand thumb rule": 9,
            "electric motor": 9, "generator": 8, "electromagnetic induction": 10,
            "induced current": 8, "commutator": 7, "split ring": 8,
            "armature": 6, "magnetic force": 7, "current carrying conductor": 8, "motor": 7, "mechanical energy": 7, "electrical energy into mechanical energy": 10,
        },
    },
    {
        "no": 13,
        "name": "Our Environment",
        "keywords": {
            "our environment": 10, "ecosystem": 8, "food chain": 9,
            "food web": 8, "trophic level": 9, "producer": 6, "consumer": 6,
            "decomposer": 8, "energy flow": 8, "10% law": 9, "biodegradable": 9,
            "non-biodegradable": 9, "ozone": 8, "ozone layer": 9, "waste": 5,
            "biomagnification": 10, "biological magnification": 10,
            "environmental pollution": 7, "ecosystem": 8,
        },
    },
]

# Words that prove the text is English prose. Single variables such as A, B,
# x, y; units; formulae; element symbols; and option labels don't count.
ENGLISH_FUNCTION_WORDS = {
    "what", "which", "who", "whom", "whose", "where", "when", "why", "how",
    "is", "are", "was", "were", "does", "do", "did", "has", "have", "had",
    "can", "could", "will", "would", "should", "shall", "may", "might",
    "the", "a", "an", "of", "in", "on", "for", "from", "to", "with",
    "by", "at", "as", "than", "into", "over", "under", "between", "during",
    "following", "given", "below", "above", "correct", "incorrect", "select",
    "choose", "find", "calculate", "determine", "state", "define", "explain",
    "describe", "write", "give", "mention", "identify", "name", "compare",
    "differentiate", "justify", "reason", "answer", "study", "observe",
    "consider", "using", "used", "called", "known", "because", "therefore",
    "if", "then", "also", "both", "each", "any", "following", "respectively",
}

COMMON_ENGLISH_CONTENT = {
    "light", "water", "carbon", "oxygen", "hydrogen", "metal", "acid", "base",
    "salt", "reaction", "energy", "current", "electric", "magnetic", "field",
    "force", "lens", "mirror", "image", "atom", "compound", "plant", "animal",
    "cell", "body", "heart", "blood", "brain", "eye", "gene", "trait", "food",
    "organism", "environment", "ecosystem", "question", "statement", "option",
    "solution", "temperature", "pressure", "distance", "time", "speed", "mass",
    "volume", "number", "process", "formation", "formation", "property", "change",
}

STRUCTURAL_ONLY = {
    "section", "part", "case", "study", "question", "questions", "mcq", "assertion",
    "reason", "option", "options", "or", "attempt", "either", "marks", "mark",
    "note", "instructions", "visually", "impaired", "students", "answer",
}

QUESTION_START_RE = re.compile(
    r"^\s*(?:Q(?:uestion)?\s*)?(\d{1,3})\s*[\.)]?\s*(.*?)\s*$",
    flags=re.IGNORECASE,
)

SECTION_RE = re.compile(r"^\s*(?:section|part)\s*[-–—:]?\s*([A-Za-z0-9]+)?", re.I)

@dataclass
class Question:
    source_file: str
    original_number: str
    text: str
    subject: Optional[str] = None
    chapter_no: Optional[int] = None
    confidence: float = 0.0
    question_type: str = "Question"


def normalize_text(text: str) -> str:
    text = text.replace("\u00ad", "")
    text = text.replace("\u2011", "-").replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def strip_noise_lines(text: str) -> str:
    """Remove recurring exam boilerplate, page footer artifacts, and VI duplicates."""
    lines = []
    skip_visually_impaired = False
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            lines.append("")
            continue
        low = line.lower()
        if re.search(r"for visually impaired students", low):
            skip_visually_impaired = True
            continue
        # A long dashed divider ends the VI block in most CBSE PDFs.
        if skip_visually_impaired:
            if re.fullmatch(r"[-_ .]{5,}", line):
                skip_visually_impaired = False
            # Skip duplicate VI material until the next normal numbered Q appears.
            else:
                continue
        if "there is no change in the question paper design" in low:
            continue
        if low in {"science - code no. 086", "science – code no. 086"}:
            continue
        if re.fullmatch(r"\*?\s*\d+\s*\*?", line):
            # Lone page numbers / stray mark values are not useful source text.
            continue
        lines.append(line)
    return "\n".join(lines)


def infer_section_subjects(text: str) -> dict[str, str]:
    """Infer section->subject mapping from the paper's own instructions/headings."""
    low = re.sub(r"\s+", " ", text.lower())
    mapping: dict[str, str] = {}
    patterns = [
        ("a", "biology"), ("b", "chemistry"), ("c", "physics"),
    ]
    # Prefer explicit statements in the paper.
    for sec, subj in patterns:
        if re.search(rf"section\s*[-–—]?\s*{sec}\b[^.]{0,80}\b{subj}\b", low):
            mapping[sec.upper()] = subj
    # Also handle headings like "Section A - Biology".
    for m in re.finditer(r"section\s*[-–—]?\s*([abc])\b.{0,40}\b(biology|chemistry|physics)\b", low):
        mapping[m.group(1).upper()] = m.group(2)
    return mapping


def section_at_or_before(line: str, current: Optional[str]) -> Optional[str]:
    m = re.match(r"^\s*(?:section|part)\s*[-–—:]?\s*([ABC])\b", line, flags=re.I)
    return m.group(1).upper() if m else current


def subject_chapters(subject: Optional[str]) -> Optional[set[int]]:
    if subject == "biology":
        return {5, 6, 7, 8, 13}
    if subject == "chemistry":
        return {1, 2, 3, 4}
    if subject == "physics":
        return {9, 10, 11, 12}
    return None


def is_likely_question_start(line: str, following_text: str = "") -> bool:
    m = QUESTION_START_RE.match(line)
    if not m:
        return False
    num = int(m.group(1))
    if not (1 <= num <= 100):
        return False
    tail = (m.group(2) or "").strip()
    # Reject obvious numbered list entries inside tables / instructions when
    # there is no question-like sentence and a previous line is likely an item.
    if not tail and len(following_text.strip()) < 4:
        return False
    return True


def extract_questions(text: str, source_file: str) -> list[Question]:
    text = strip_noise_lines(normalize_text(text))
    lines = text.splitlines()
    section_subjects = infer_section_subjects(text)
    starts: list[tuple[int, str, Optional[str]]] = []
    current_section: Optional[str] = None
    for i, line in enumerate(lines):
        current_section = section_at_or_before(line, current_section)
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        if is_likely_question_start(line, nxt):
            m = QUESTION_START_RE.match(line)
            subj = section_subjects.get(current_section or "")
            starts.append((i, m.group(1), subj))
    questions: list[Question] = []
    for idx, (start_i, qno, subj) in enumerate(starts):
        end_i = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block_lines = lines[start_i:end_i]
        block = "\n".join(block_lines).strip()
        block = re.sub(r"^\s*(?:Q(?:uestion)?\s*)?\d{1,3}\s*[\.)]?\s*", "", block, count=1, flags=re.I)
        if len(re.sub(r"[^A-Za-z0-9\u0900-\u097F]+", "", block)) < 8:
            continue
        questions.append(Question(source_file=source_file, original_number=qno, text=block, subject=subj))
    return questions


def latin_words(text: str) -> list[str]:
    return re.findall(r"[A-Za-z]{2,}", text.lower())


def script_counts(text: str) -> tuple[int, int, int]:
    dev = len(re.findall(r"[\u0900-\u097F]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    other = len(re.findall(r"\d|[^\w\s]", text, flags=re.UNICODE))
    return dev, latin, other


def english_confidence(text: str) -> float:
    """Return 0..1 confidence that text is English *prose*, not just Latin symbols."""
    clean = re.sub(r"\s+", " ", text.lower()).strip()
    dev, latin, _ = script_counts(clean)
    words = latin_words(clean)
    if not words:
        return 0.0

    substantive = [w for w in words if len(w) >= 3 and w not in STRUCTURAL_ONLY]
    function_hits = sum(w in ENGLISH_FUNCTION_WORDS for w in words)
    content_hits = sum(w in COMMON_ENGLISH_CONTENT for w in substantive)

    # Strong Hindi/Devanagari content wins unless there is substantial English prose.
    if dev >= 6:
        # Pure variable/symbol English noise should not rescue Hindi.
        if function_hits >= 4 and len(substantive) >= 8:
            return 0.80
        if function_hits >= 2 and len(substantive) >= 14 and latin > dev * 1.4:
            return 0.65
        return 0.05

    # No Devanagari: require meaningful English words, not just x, A, B, Na, Cl,
    # units or formula fragments.
    if len(substantive) >= 7 and function_hits >= 2:
        return 0.98
    if len(substantive) >= 4 and function_hits >= 2:
        return 0.90
    if len(substantive) >= 3 and (function_hits >= 2 or content_hits >= 2):
        return 0.82
    if len(substantive) >= 2 and function_hits >= 2:
        return 0.72
    return 0.15


def is_english_question(text: str, threshold: float = 0.72) -> bool:
    return english_confidence(text) >= threshold


def clean_question(text: str) -> str:
    # Collapse spaces while keeping meaningful line breaks.
    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    out: list[str] = []
    prev_blank = False
    for line in lines:
        if not line:
            if not prev_blank:
                out.append("")
            prev_blank = True
        else:
            out.append(line)
            prev_blank = False
    return "\n".join(out).strip()


def classify_question(text: str, allowed_chapters: Optional[set[int]] = None) -> tuple[Optional[int], float, dict[int, float]]:
    t = re.sub(r"\s+", " ", text.lower())
    scores: dict[int, float] = {}
    for ch in CHAPTERS:
        if allowed_chapters is not None and ch["no"] not in allowed_chapters:
            continue
        score = 0.0
        for key, weight in ch["keywords"].items():
            # Word-ish matching for short tokens; phrase matching for multiword terms.
            if " " in key or "-" in key or "'" in key:
                hits = t.count(key.lower())
            else:
                hits = len(re.findall(r"\b" + re.escape(key.lower()) + r"\b", t))
            if hits:
                score += weight * min(hits, 3)
        scores[ch["no"]] = score

    if not scores:
        return None, 0.0, scores
    ordered = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best_no, best = ordered[0]
    second = ordered[1][1] if len(ordered) > 1 else 0.0

    if best <= 0:
        return None, 0.0, scores
    # Confidence is conservative when the top chapter barely beats the runner-up.
    margin = (best - second) / max(best, 1.0)
    coverage = min(best / 16.0, 1.0)
    confidence = max(0.0, min(1.0, 0.55 * coverage + 0.45 * margin))

    # Tiny single-keyword matches are usually unsafe for chapter assignment.
    if best < 6:
        return None, confidence, scores
    return best_no, confidence, scores


def detect_question_type(text: str) -> str:
    t = re.sub(r"\s+", " ", text.lower())
    if re.search(r"assertion\s*\(?a\)?\s*[:\-]", t) and re.search(r"reason\s*\(?r\)?\s*[:\-]", t):
        return "Assertion–Reason"
    if re.search(r"case study|read the following|study the passage|passage given", t):
        return "Case/Source-Based"

    # Descriptive questions often have A/B/C/D sub-parts. Do not call those MCQs.
    descriptive_cues = (
        "attempt either", "sub-part", "give a reason", "justify", "describe",
        "how and why", "with reason", "give reasons", "write the reactions",
        "draw the", "explain the", "what percentage",
    )
    has_descriptive_cue = any(cue in t for cue in descriptive_cues)

    labels = re.findall(r"(?:^|\s)([a-d])[\.)]\s+", t)
    distinct_labels = set(labels)
    if not has_descriptive_cue and {"a", "b", "c", "d"}.issubset(distinct_labels):
        return "MCQ"

    if re.search(r"attempt either|internal choice|\bor\b", t) and re.search(r"(?:^|\s)(?:a|b|c|d)[\.)]\s+", t):
        return "Question with Sub-parts"
    if re.search(r"\([ivx]+\)|\b[ivx]+[\.)]", t):
        return "Question with Sub-parts"
    # Marks patterns are useful but not always preserved by extraction.
    if re.search(r"\b[1-5]\s*marks?\b", t):
        return "Descriptive"
    return "Question"


def extract_pdf(path: Path, use_ocr: bool = True) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed. Install the requirements first.")
    doc = fitz.open(path)
    pages: list[str] = []
    try:
        for page in doc:
            txt = page.get_text("text") or ""
            # If the page has almost no selectable text, optionally OCR the page.
            if use_ocr and len(re.sub(r"\s+", "", txt)) < 40 and pytesseract and Image:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                txt = pytesseract.image_to_string(img, lang="eng+hin", config="--psm 6") or txt
            pages.append(txt)
    finally:
        doc.close()
    return "\n".join(pages)


def extract_docx(path: Path) -> str:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Install the requirements first.")
    doc = Document(path)
    chunks = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            chunks.append(" ".join(cell.text for cell in row.cells))
    return "\n".join(chunks)


def extract_image(path: Path, use_ocr: bool = True) -> str:
    if Image is None:
        raise RuntimeError("Pillow is not installed. Install the requirements first.")
    if not use_ocr or pytesseract is None:
        raise RuntimeError("OCR is required for image files. Install pytesseract and Tesseract OCR.")
    img = Image.open(path)
    return pytesseract.image_to_string(img, lang="eng+hin", config="--psm 6") or ""


def extract_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def process_file(path: Path, use_ocr: bool) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path, use_ocr=use_ocr)
    if suffix == ".docx":
        return extract_docx(path)
    if suffix in {".txt", ".md"}:
        return extract_txt(path)
    if suffix in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
        return extract_image(path, use_ocr=use_ocr)
    raise ValueError(f"Unsupported file type: {path.name}")


def iter_input_files_from_zip(zip_path: Path, temp_dir: Path) -> Iterable[Path]:
    allowed = {".pdf", ".docx", ".txt", ".md", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = Path(info.filename).name
            if Path(name).suffix.lower() not in allowed:
                continue
            # Avoid zip-slip: write every member to a controlled temp folder.
            safe_name = re.sub(r"[^A-Za-z0-9._ -]+", "_", name)
            dest = temp_dir / safe_name
            with zf.open(info, "r") as src, dest.open("wb") as dst:
                shutil.copyfileobj(src, dst)
            yield dest


def build_questions(zip_path: Path, use_ocr: bool, english_threshold: float, min_chapter_conf: float, max_files: int = 30) -> tuple[list[Question], Counter, list[str]]:
    all_q: list[Question] = []
    stats = Counter()
    errors: list[str] = []
    with tempfile.TemporaryDirectory(prefix="cbse_science_sorter_") as td:
        temp_dir = Path(td)
        for path in iter_input_files_from_zip(zip_path, temp_dir):
            if stats["files"] >= max_files:
                raise ValueError(f"Maximum {max_files} question-paper files are supported per run. The ZIP contains more supported files.")
            stats["files"] += 1
            try:
                raw = process_file(path, use_ocr=use_ocr)
                qs = extract_questions(raw, path.name)
                stats["candidate_questions"] += len(qs)
                for q in qs:
                    q.text = clean_question(q.text)
                    conf = english_confidence(q.text)
                    if conf < english_threshold:
                        stats["ignored_non_english"] += 1
                        continue
                    ch_no, ch_conf, _ = classify_question(q.text, subject_chapters(q.subject))
                    if ch_no is None or ch_conf < min_chapter_conf:
                        q.chapter_no = None
                        q.confidence = ch_conf
                        q.question_type = detect_question_type(q.text)
                        all_q.append(q)
                        stats["review"] += 1
                    else:
                        q.chapter_no = ch_no
                        q.confidence = ch_conf
                        q.question_type = detect_question_type(q.text)
                        all_q.append(q)
                        stats["sorted"] += 1
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
                stats["errors"] += 1
    return all_q, stats, errors


def chapter_name(no: int) -> str:
    for c in CHAPTERS:
        if c["no"] == no:
            return c["name"]
    return ""


def add_question(doc, q: Question, local_no: int, include_source: bool):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(f"{local_no}. ")
    r.bold = True
    r2 = p.add_run(f"[{q.question_type}] ")
    r2.italic = True
    r2.font.color.rgb = RGBColor(90, 90, 90)

    # Write question as a single paragraph with manual line breaks to preserve sub-parts.
    parts = q.text.split("\n")
    for i, part in enumerate(parts):
        if i:
            r = p.add_run("\n")
        r = p.add_run(part)
        r.font.size = Pt(10.5)

    if include_source:
        sp = doc.add_paragraph()
        sp.paragraph_format.left_indent = Inches(0.28)
        rr = sp.add_run(f"Source: {q.source_file} — Original Q{q.original_number}")
        rr.font.size = Pt(8.5)
        rr.italic = True
        rr.font.color.rgb = RGBColor(120, 120, 120)


def create_docx(questions: list[Question], output_path: Path, title: str, include_source: bool, include_review: bool = True):
    if Document is None:
        raise RuntimeError("python-docx is not installed. Install the requirements first.")
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.65)
    sec.right_margin = Inches(0.65)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("English-language questions only • Chapter-wise • All question types")
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(90, 90, 90)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Language rule: Hindi/Devanagari questions containing only variables, symbols, units or short English fragments are excluded.")
    r.italic = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(110, 110, 110)

    by_chapter: dict[int, list[Question]] = defaultdict(list)
    review: list[Question] = []
    for q in questions:
        if q.chapter_no is None:
            review.append(q)
        else:
            by_chapter[q.chapter_no].append(q)

    # Contents / overview
    doc.add_heading("Overview", level=1)
    for c in CHAPTERS:
        n = len(by_chapter.get(c["no"], []))
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"Chapter {c['no']}: {c['name']} — {n} question(s)")
    if review:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"Review / Unclassified — {len(review)} question(s)")

    for c in CHAPTERS:
        qs = by_chapter.get(c["no"], [])
        if not qs:
            continue
        doc.add_page_break()
        h = doc.add_heading(f"Chapter {c['no']}: {c['name']}", level=1)
        h.paragraph_format.keep_with_next = True
        counts = Counter(q.question_type for q in qs)
        summary = ", ".join(f"{k}: {v}" for k, v in counts.most_common())
        p = doc.add_paragraph()
        r = p.add_run(f"{len(qs)} question(s) • {summary}")
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(100, 100, 100)
        for local_no, q in enumerate(qs, 1):
            add_question(doc, q, local_no, include_source)

    if review and include_review:
        doc.add_page_break()
        doc.add_heading("Review / Unclassified", level=1)
        p = doc.add_paragraph("These questions passed the English-language filter but did not achieve a strong enough chapter match. They are kept here so nothing is silently lost.")
        p.runs[0].font.size = Pt(9.5)
        for local_no, q in enumerate(review, 1):
            add_question(doc, q, local_no, include_source)

    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run(f"CBSE Class 10 Science Question Sorter • {datetime.now():%d %b %Y}")
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor(125, 125, 125)

    doc.save(output_path)


def run_sort(zip_path: Path, output_path: Path, use_ocr: bool = True, english_threshold: float = 0.72,
             min_chapter_conf: float = 0.42, include_source: bool = False) -> tuple[Counter, list[str]]:
    questions, stats, errors = build_questions(zip_path, use_ocr, english_threshold, min_chapter_conf)
    # Exact duplicate removal is intentionally not automatic: question banks often
    # need duplicates across different papers. Use the UI option when desired.
    create_docx(
        questions,
        output_path,
        title="CBSE Class 10 Science — English Questions, Chapter-wise",
        include_source=include_source,
        include_review=True,
    )
    return stats, errors


def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("CBSE Class 10 Science – English Question Sorter")
    root.geometry("700x520")
    root.minsize(620, 460)

    zip_var = tk.StringVar()
    out_var = tk.StringVar(value="CBSE_Class10_Science_English_Chapterwise.docx")
    ocr_var = tk.BooleanVar(value=True)
    source_var = tk.BooleanVar(value=False)
    threshold_var = tk.DoubleVar(value=0.72)
    conf_var = tk.DoubleVar(value=0.42)
    status_var = tk.StringVar(value="Choose a ZIP of question papers to begin.")

    frm = ttk.Frame(root, padding=18)
    frm.pack(fill="both", expand=True)

    title = ttk.Label(frm, text="CBSE Class 10 Science – English Question Sorter", font=("TkDefaultFont", 16, "bold"))
    title.pack(anchor="w", pady=(0, 4))
    ttk.Label(frm, text="Sorts ALL English-language questions chapter-wise, including MCQs, Assertion–Reason, descriptive and case-based questions.", wraplength=640).pack(anchor="w", pady=(0, 14))

    row = ttk.Frame(frm); row.pack(fill="x", pady=5)
    ttk.Label(row, text="Question-paper ZIP:", width=18).pack(side="left")
    ttk.Entry(row, textvariable=zip_var).pack(side="left", fill="x", expand=True, padx=5)
    def choose_zip():
        f = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")])
        if f: zip_var.set(f)
    ttk.Button(row, text="Browse…", command=choose_zip).pack(side="left")

    row = ttk.Frame(frm); row.pack(fill="x", pady=5)
    ttk.Label(row, text="Output DOCX:", width=18).pack(side="left")
    ttk.Entry(row, textvariable=out_var).pack(side="left", fill="x", expand=True, padx=5)
    def choose_out():
        f = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word document", "*.docx")], initialfile=Path(out_var.get()).name)
        if f: out_var.set(f)
    ttk.Button(row, text="Save as…", command=choose_out).pack(side="left")

    opts = ttk.LabelFrame(frm, text="Sorting rules", padding=10)
    opts.pack(fill="x", pady=14)
    ttk.Checkbutton(opts, text="Use OCR for image-only PDF pages", variable=ocr_var).grid(row=0, column=0, sticky="w", pady=3)
    ttk.Checkbutton(opts, text="Include source filename + original question number", variable=source_var).grid(row=1, column=0, sticky="w", pady=3)
    ttk.Label(opts, text="English threshold (higher = stricter):").grid(row=2, column=0, sticky="w", pady=(8, 2))
    ttk.Scale(opts, from_=0.55, to=0.95, variable=threshold_var, orient="horizontal", length=340).grid(row=2, column=1, sticky="w", padx=10)
    ttk.Label(opts, textvariable=tk.StringVar(value="Default 0.72")).grid(row=2, column=2, sticky="w")
    ttk.Label(opts, text="Chapter-confidence threshold:").grid(row=3, column=0, sticky="w", pady=3)
    ttk.Scale(opts, from_=0.25, to=0.75, variable=conf_var, orient="horizontal", length=340).grid(row=3, column=1, sticky="w", padx=10)
    ttk.Label(opts, text="Default 0.42").grid(row=3, column=2, sticky="w")

    ttk.Label(frm, textvariable=status_var, wraplength=640).pack(anchor="w", pady=10)

    progress = ttk.Progressbar(frm, mode="indeterminate")
    progress.pack(fill="x", pady=(0, 12))

    def run():
        z = zip_var.get().strip()
        if not z or not Path(z).is_file():
            messagebox.showerror("Missing ZIP", "Please choose a valid ZIP file containing the question papers.")
            return
        out = Path(out_var.get().strip() or "CBSE_Class10_Science_English_Chapterwise.docx")
        try:
            status_var.set("Sorting… English/Hindi filtering and chapter classification are running.")
            progress.start(10)
            root.update_idletasks()
            stats, errors = run_sort(Path(z), out, use_ocr=ocr_var.get(), english_threshold=threshold_var.get(), min_chapter_conf=conf_var.get(), include_source=source_var.get())
            progress.stop()
            status_var.set(f"Done — {stats.get('sorted', 0)} sorted, {stats.get('review', 0)} sent to Review/Unclassified, {stats.get('ignored_non_english', 0)} non-English candidates ignored.")
            detail = "Output:\n" + str(out.resolve())
            if errors:
                detail += "\n\nFiles with errors:\n" + "\n".join(errors[:12])
            messagebox.showinfo("Sorter finished", detail)
        except Exception as exc:
            progress.stop()
            status_var.set("Error while processing.")
            messagebox.showerror("Sorter error", str(exc))

    ttk.Button(frm, text="SORT QUESTIONS", command=run).pack(anchor="w", pady=6, ipadx=10, ipady=4)
    ttk.Label(frm, text="Important: the English filter is based on real English prose, not the presence of A/B/C/D, variables, formulas, symbols or isolated English letters in Hindi questions.", wraplength=640, foreground="#555555").pack(anchor="w", pady=(14, 0))

    root.mainloop()


def main_cli():
    import argparse
    parser = argparse.ArgumentParser(description="Sort English CBSE Class 10 Science questions chapter-wise from a ZIP.")
    parser.add_argument("zip", type=Path, help="ZIP containing question papers")
    parser.add_argument("-o", "--output", type=Path, default=Path("CBSE_Class10_Science_English_Chapterwise.docx"))
    parser.add_argument("--no-ocr", action="store_true", help="Do not OCR image-only PDF pages")
    parser.add_argument("--include-source", action="store_true", help="Add source filename/original question number under each question")
    parser.add_argument("--english-threshold", type=float, default=0.72)
    parser.add_argument("--chapter-confidence", type=float, default=0.42)
    args = parser.parse_args()
    stats, errors = run_sort(
        args.zip, args.output,
        use_ocr=not args.no_ocr,
        english_threshold=args.english_threshold,
        min_chapter_conf=args.chapter_confidence,
        include_source=args.include_source,
    )
    print("Done.")
    print(dict(stats))
    if errors:
        print("Errors:")
        for e in errors:
            print(" -", e)
    print("Output:", args.output.resolve())


if __name__ == "__main__":
    import sys
    # No command-line arguments => open the GUI. Use CLI arguments for automation.
    if len(sys.argv) == 1:
        launch_gui()
    else:
        main_cli()
