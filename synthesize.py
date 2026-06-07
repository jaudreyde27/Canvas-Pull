"""
Synthesize course files into a study guide using the Claude API.

Reads from:
  - downloads/  (output from pull_canvas.py / the pull workflow)
  - source-files/ (manually uploaded files)

Writes:
  - study_guide.md

Required env: ANTHROPIC_API_KEY
Optional env:  COURSE_NAME
"""

import os
import sys
import base64
from pathlib import Path

import anthropic

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from pptx import Presentation as PptxPresentation
except ImportError:
    PptxPresentation = None

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
COURSE_NAME = os.environ.get("COURSE_NAME", "Course")
MAP_MODEL = "claude-haiku-4-5-20251001"      # fast + cheap for per-file extraction
REDUCE_MODEL = "claude-sonnet-4-6"           # high quality for final synthesis

if not API_KEY:
    print("ERROR: ANTHROPIC_API_KEY is not set.")
    sys.exit(1)

client = anthropic.Anthropic(api_key=API_KEY)

# ── Format description (condensed from user's example study guide) ──────────

FORMAT_DESCRIPTION = """
Structure the study guide EXACTLY like this:

# [COURSE NAME] — Study Guide

## 1. [Topic / Module Name]

### Core Definitions
**Key Term**: "Optional quoted framing." Full definition — what it is, how it's calculated, what it means strategically.

**Another Term**: Definition with any critical distinctions inline.

Key distinction: [clarify how two related concepts differ — use "NOT" for emphasis]

### [Sub-concept or Framework Name]
- Bullet for a key point or driver
- Bullet using arrow notation for causal relationships: Low X → high Y
- Bullet with an inline example. Example: [real-world case in brackets]

**Exam note**: [exam-specific guidance, traps, or required approach — include whenever material suggests a testable nuance]

### [Another Sub-concept]
(repeat pattern)

## 2. [Next Topic]
(repeat)

───
STYLE RULES:
- Bold every key term on first use
- Arrow notation (→) for all causal / directional relationships
- "Key distinction:" prefix for clarifications between commonly confused concepts
- "Exam note:" / "Exam trap:" / "Exam approach:" callouts whenever the material suggests a testable nuance
- Inline examples in brackets: Example: [...]
- Keep every bullet tight — one idea per line, no filler
- Number all top-level sections sequentially (1, 2, 3 …)
"""

FORMAT_EXAMPLE = """
EXAMPLE (from a different course — match this density and style exactly):

## 1. Value Creation & Value Capture

### Core Definitions
**Value Creation**: "Making the pie bigger." The total surplus generated in a transaction, calculated as the gap between the Buyer's Willingness to Pay (WTP) and the Supplier's Opportunity Cost (OC).

**Value Capture**: "Claiming your slice." Retaining a share of the surplus as profit by overcoming rivalry and vertical bargaining power.

Key distinction: Value creation and value capture are NOT the same thing. Price changes and competitive dynamics affect value capture — not value creation.

### Cost-Quality Frontier
- The frontier is defined by firms with the best cost/quality tradeoffs — being inside it is a real competitive disadvantage
- Two approaches: raise perceived quality (↑ WTP) or reduce cost (↓ OC)
- More expensive does NOT mean more value captured; lowest cost does NOT either — both are creation, not capture

**Exam note**: When a new entrant matches an incumbent's position exactly, price competition begins immediately — the overlap is total.
"""


# ── File extraction helpers ──────────────────────────────────────────────────

def extract_pdf_text(path: Path) -> str:
    if fitz is None:
        return f"[PDF could not be read — PyMuPDF unavailable]"
    doc = fitz.open(str(path))
    pages = [page.get_text() for page in doc]
    text = "\n".join(pages)
    return text[:40000]  # cap per file


def extract_pptx_text(path: Path) -> str:
    if PptxPresentation is None:
        return "[PPTX could not be read — python-pptx unavailable]"
    prs = PptxPresentation(str(path))
    lines = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                lines.append(shape.text.strip())
    return "\n".join(lines)[:40000]


def encode_image(path: Path) -> tuple[str, str]:
    media = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
             ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp"}
    mt = media.get(path.suffix.lower(), "image/jpeg")
    data = base64.standard_b64encode(path.read_bytes()).decode()
    return data, mt


# ── Map phase: extract key points from one file ──────────────────────────────

def extract_key_points(path: Path) -> str:
    name = path.name
    suffix = path.suffix.lower()
    image_exts = {".png", ".jpg", ".jpeg", ".gif", ".webp"}

    print(f"  → {name}")

    if suffix == ".pdf":
        text = extract_pdf_text(path)
        content = [{"type": "text", "text": (
            f"Course file: '{name}'\n\n"
            "Extract every key concept, definition, framework, formula, and important point "
            "from the text below as comprehensive notes.\n"
            "- Preserve all terminology exactly\n"
            "- Include any instructor emphasis or exam tips\n"
            "- Group related ideas together\n"
            "- Be thorough — nothing testable should be omitted\n\n"
            f"{text}"
        )}]

    elif suffix in {".pptx", ".ppt"}:
        text = extract_pptx_text(path)
        content = [{"type": "text", "text": (
            f"Course slide deck: '{name}'\n\n"
            "Extract every key concept, definition, framework, and important point "
            "from the slide text below.\n"
            "- Preserve all terminology exactly\n"
            "- Group by topic/slide section\n\n"
            f"{text}"
        )}]

    elif suffix in image_exts:
        img_data, media_type = encode_image(path)
        content = [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_data}},
            {"type": "text", "text": (
                f"This is a course slide/image: '{name}'. "
                "Extract all key concepts, definitions, frameworks, and important points visible. "
                "Preserve all terminology exactly. Format as grouped bullet points."
            )},
        ]
    else:
        return ""

    resp = client.messages.create(
        model=MAP_MODEL,
        max_tokens=2048,
        messages=[{"role": "user", "content": content}],
    )
    return resp.content[0].text


# ── Reduce phase: synthesize all notes into final study guide ────────────────

def synthesize(all_notes: list[tuple[str, str]]) -> str:
    print("\nSynthesizing into study guide …")

    combined = "\n\n".join(
        f"=== SOURCE: {fname} ===\n{notes}"
        for fname, notes in all_notes
    )

    # Trim if very large (> 150k chars ≈ 37k tokens)
    if len(combined) > 150000:
        combined = combined[:150000] + "\n\n[... additional material truncated for context limits ...]"

    prompt = (
        f'You are creating a comprehensive study guide for "{COURSE_NAME}".\n\n'
        f"{FORMAT_DESCRIPTION}\n\n"
        f"{FORMAT_EXAMPLE}\n\n"
        "Now synthesize the course notes below into a complete study guide in that exact style.\n"
        "Requirements:\n"
        "- Organize content into logical numbered sections by topic\n"
        "- Bold every key term on first use\n"
        "- Include Exam notes wherever the material suggests testable nuances\n"
        "- Be comprehensive but tight — every sentence earns its place\n"
        "- Do not repeat content; consolidate when the same idea appears in multiple files\n\n"
        "COURSE NOTES:\n\n"
        f"{combined}"
    )

    resp = client.messages.create(
        model=REDUCE_MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text


# ── Main ─────────────────────────────────────────────────────────────────────

SUPPORTED = {".pdf", ".pptx", ".ppt", ".png", ".jpg", ".jpeg", ".gif", ".webp"}


def collect_files(*dirs: Path) -> list[Path]:
    files = []
    for d in dirs:
        if d.exists():
            for f in sorted(d.rglob("*")):
                if f.is_file() and f.suffix.lower() in SUPPORTED:
                    files.append(f)
    return files


def main() -> None:
    files = collect_files(Path("downloads"), Path("source-files"))

    if not files:
        print("No supported files found in downloads/ or source-files/")
        print("Run the 'Pull Canvas Files' workflow first, or add files to source-files/")
        sys.exit(1)

    print(f"Found {len(files)} file(s)\n")

    all_notes: list[tuple[str, str]] = []
    for f in files:
        try:
            notes = extract_key_points(f)
            if notes.strip():
                all_notes.append((f.name, notes))
        except Exception as exc:
            print(f"  [error] {f.name}: {exc}")

    if not all_notes:
        print("No content could be extracted from any file.")
        sys.exit(1)

    guide = synthesize(all_notes)

    out = Path("study_guide.md")
    out.write_text(guide, encoding="utf-8")
    size = len(guide)
    print(f"\nDone — study_guide.md ({size:,} chars)")


if __name__ == "__main__":
    main()
