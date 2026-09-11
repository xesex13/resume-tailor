#!/usr/bin/env python3
"""
Terry Thomas Tyler Tristan Tucker Tobias Thaddeus Maximus Tommy The Resume Tailor
(TTTTMTTRT)

Local CLI that tailors master_resume.docx and base_cover_letter.docx to a pasted
job description using Gemini Flash, editing the existing .docx files in place
so fonts/margins/styling survive.
"""

import argparse
import json
import logging
import os
import random
import re
import subprocess
import sys
import time

logging.getLogger("google_genai.models").setLevel(logging.ERROR)

import colorama
from colorama import Fore, Style
from dotenv import load_dotenv
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.text.hyperlink import Hyperlink

colorama.init(autoreset=True)

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


class TailorError(Exception):
    """Clean, user-facing error raised for any Gemini API / parsing failure."""


def log(message, ok=True):
    icon = f"{Fore.GREEN}[✔]{Style.RESET_ALL}" if ok else f"{Fore.RED}[✘]{Style.RESET_ALL}"
    try:
        print(f"{icon} {message}")
    except UnicodeEncodeError:
        fallback = "[OK]" if ok else "[FAIL]"
        print(f"{fallback} {message}")


def warn(message):
    try:
        print(f"{Fore.YELLOW}[!]{Style.RESET_ALL} {message}")
    except UnicodeEncodeError:
        print(f"[!] {message}")

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
MASTER_RESUME = "master_resume.docx"
BASE_COVER_LETTER = "base_cover_letter.docx"
ANTHROPIC_MODEL = "claude-sonnet-5"
ANTHROPIC_API_VERSION = "2023-06-01"
GEMINI_MODEL = "gemini-3.6-flash"
GROQ_MODEL = "openai/gpt-oss-120b"
REMOVE_TOKEN = "__REMOVE__"

MARGIN_INCHES = 1.0
MARGIN_FALLBACK_INCHES = 0.75

RESUME_MIN_CHARS = 2200
RESUME_MAX_CHARS = 2400
BULLET_MIN_CHARS = 130
BULLET_MAX_CHARS = 165

COVER_MIN_WORDS = 270
COVER_MAX_WORDS = 310

HEADER_STYLE_NAMES = {"Title", "Heading 1", "Heading 2", "Heading 3"}
HEADER_FONT_SIZE = Pt(10.5)
BODY_FONT_SIZE = Pt(10.5)
HEADER_FONT_COLOR = RGBColor(0, 0, 0)

MAX_OUTPUT_TOKENS = 8192
ANTHROPIC_MAX_OUTPUT_TOKENS = 32768
GEMINI_MAX_OUTPUT_TOKENS = 16384

PROVIDER_ANTHROPIC = "anthropic"
PROVIDER_GEMINI = "gemini"
PROVIDER_GROQ = "groq"
PROVIDER_FREE_FALLBACK = "free_fallback"

PROVIDER_LABELS = {
    PROVIDER_ANTHROPIC: "Anthropic (Claude Sonnet 5)",
    PROVIDER_GEMINI: "Gemini (Gemini 3.6 Flash)",
    PROVIDER_GROQ: "Groq (GPT-OSS 120B)",
    PROVIDER_FREE_FALLBACK: "Completely Free Tier with Automatic Fallback (Gemini 3.6 Flash -> Groq)",
}

REQUIRED_KEYS_BY_PROVIDER = {
    PROVIDER_ANTHROPIC: ["ANTHROPIC_API_KEY"],
    PROVIDER_GEMINI: ["GEMINI_API_KEY"],
    PROVIDER_GROQ: ["GROQ_API_KEY"],
    PROVIDER_FREE_FALLBACK: ["GEMINI_API_KEY", "GROQ_API_KEY"],
}

GEMINI_FREE_TIER_CALL_CAP = 20
_gemini_call_count = 0
_provider_override = None

RAINBOW_COLORS = [
    Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.BLUE, Fore.MAGENTA, Fore.CYAN,
]

TOMMY_ART = [
    "        .-\"\"\"\"-.        ",
    "       /  o   o \\       ",
    "      |    ^     |   __ ",
    "      |  \\___/   |  |  |",
    "       \\        /   |__|",
    "     .--'------'--.  o  ",
    "    /  T O M M Y   \\    ",
    "   |   :::::::::::   |   ",
    "   |   :::::::::::   |   ",
    "    \\_______________/    ",
    "     |    |    |    |    ",
    "     |____|    |____|    ",
]


def rainbow_print(text):
    words = text.split(" ")
    out = []
    for w in words:
        color = random.choice(RAINBOW_COLORS)
        out.append(f"{color}{w}{Style.RESET_ALL}")
    print(" ".join(out))


def tommy_intro():
    rainbow_print("INITIALIZING TOMMY-CORE ENGINE...")
    for line in TOMMY_ART:
        rainbow_print(line)
        time.sleep(0.08)


def print_provider_menu():
    print()
    print(f"{Fore.CYAN}Select your LLM engine provider:{Style.RESET_ALL}")
    print("  1) Anthropic (Claude Sonnet 5)")
    print("     Paid API, highest intelligence, zero truncation, strict formatting.")
    print("  2) Gemini (Gemini 3.6 Flash)")
    print("     Free tier -- capped at 20 API calls, smart and fast.")
    print("  3) Groq (GPT-OSS 120B)")
    print("     Free tier -- up to 30 req/min & 14,400 req/day, ultra-fast, lower formatting retention.")
    print("  4) Completely Free Tier with Automatic Fallback (Gemini 3.6 Flash -> Groq)")
    print("     Runs Gemini first, seamlessly reroutes to Groq on rate limit or error.")
    print()


def select_provider_menu():
    choice_map = {
        "1": PROVIDER_ANTHROPIC,
        "2": PROVIDER_GEMINI,
        "3": PROVIDER_GROQ,
        "4": PROVIDER_FREE_FALLBACK,
    }
    print_provider_menu()
    while True:
        choice = input("Enter choice [1-4]: ").strip()
        if choice in choice_map:
            return choice_map[choice]
        warn("Invalid choice -- enter a number from 1 to 4.")


def collect_required_keys(provider):
    """Load whichever API keys the chosen provider needs from .env, prompting
    for any that are missing. Returns None (signalling "re-select provider")
    if the user declines to supply a missing key."""
    load_dotenv(ENV_PATH, override=True)
    keys = {}
    new_lines = []
    for key_name in REQUIRED_KEYS_BY_PROVIDER[provider]:
        value = os.environ.get(key_name) or None
        if not value:
            warn(f"{key_name} not found in .env.")
            entered = input(f"Enter {key_name} now (or press Enter to re-select a provider): ").strip()
            if not entered:
                return None
            value = entered
            os.environ[key_name] = value
            new_lines.append(f"{key_name}={value}")
        keys[key_name] = value

    if new_lines:
        with open(ENV_PATH, "a", encoding="utf-8") as f:
            f.write("\n".join(new_lines) + "\n")

    return keys


def build_clients_for_provider(provider, keys):
    """Construct whichever provider clients the run needs. Construction is
    local/offline for all three SDKs -- no network call happens here."""
    clients = {}
    if "ANTHROPIC_API_KEY" in keys:
        import anthropic
        clients["anthropic"] = anthropic.Anthropic(
            api_key=keys["ANTHROPIC_API_KEY"],
            default_headers={"anthropic-version": ANTHROPIC_API_VERSION},
        )
    if "GEMINI_API_KEY" in keys:
        from google import genai
        clients["gemini"] = genai.Client(api_key=keys["GEMINI_API_KEY"])
    if "GROQ_API_KEY" in keys:
        from groq import Groq
        clients["groq"] = Groq(api_key=keys["GROQ_API_KEY"])
    return clients


def bootup():
    """Interactive bootup menu: pick a provider, collect its API key(s),
    re-prompting until both are satisfied."""
    while True:
        provider = select_provider_menu()
        keys = collect_required_keys(provider)
        if keys is None:
            continue
        clients = build_clients_for_provider(provider, keys)
        log(f"Provider selected: {PROVIDER_LABELS[provider]}")
        return provider, clients


ANTHROPIC_NO_SAMPLING_MODELS = {"claude-sonnet-5"}


def call_anthropic(client, prompt, system_instruction=""):
    kwargs = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": ANTHROPIC_MAX_OUTPUT_TOKENS,
        "temperature": 0.2,
        "messages": [{"role": "user", "content": prompt}],
    }
    if system_instruction:
        kwargs["system"] = system_instruction

    if ANTHROPIC_MODEL in ANTHROPIC_NO_SAMPLING_MODELS:
        # Sonnet 5's adaptive thinking mode rejects explicit sampling params --
        # only the model's own default temperature/top_p/top_k are allowed.
        kwargs.pop("temperature", None)
        kwargs.pop("top_p", None)
        kwargs.pop("top_k", None)
        kwargs["thinking"] = {"type": "adaptive"}

    # Long adaptive-thinking generations can exceed 10 minutes, which the
    # non-streaming endpoint rejects outright -- stream instead and collect
    # the final message.
    with client.messages.stream(**kwargs) as stream:
        response = stream.get_final_message()
    # Extract text safely across all content block types (adaptive thinking
    # mode interleaves ThinkingBlock entries ahead of the TextBlock(s)).
    text_blocks = [
        block.text for block in response.content
        if getattr(block, "type", None) == "text" or hasattr(block, "text")
    ]
    return "".join(text_blocks).strip()


def call_gemini(client, prompt, system_instruction="", edits_paragraph_ids=None):
    from google.genai import types

    config_kwargs = {
        "temperature": 0.2,
        "max_output_tokens": GEMINI_MAX_OUTPUT_TOKENS,
        "response_mime_type": "application/json",
    }

    if edits_paragraph_ids:
        # Force the exact {company_name, pillars, edits} shape via structured
        # output instead of trusting free-form JSON -- without this, the model
        # intermittently returns a bare paragraph-id map or a JSON array
        # instead of the wrapper object, or echoes its own internal JSON-mode
        # instructions as literal text. The Gemini Developer API (unlike
        # Vertex/Enterprise mode) doesn't support "additionalProperties" for
        # an open-ended object, so enumerate the paragraph IDs as fixed
        # properties instead -- we already know them, they're our own
        # document's paragraph IDs, not something the model invents.
        edits_properties = {pid: types.Schema(type="STRING") for pid in edits_paragraph_ids}
        edits_schema = types.Schema(
            type="OBJECT",
            properties={
                "company_name": types.Schema(type="STRING"),
                "pillars": types.Schema(type="ARRAY", items=types.Schema(type="STRING")),
                "edits": types.Schema(
                    type="OBJECT",
                    properties=edits_properties,
                    required=list(edits_paragraph_ids),
                ),
            },
            required=["company_name", "pillars", "edits"],
        )
        config_kwargs["response_schema"] = edits_schema

    if system_instruction:
        config_kwargs["system_instruction"] = system_instruction
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(**config_kwargs),
    )
    return (getattr(response, "text", None) or "").strip()


def call_groq(client, prompt, system_instruction=""):
    messages = []
    if system_instruction:
        messages.append({"role": "system", "content": system_instruction})
    messages.append({"role": "user", "content": prompt})
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=messages,
        temperature=0.2,
        max_completion_tokens=MAX_OUTPUT_TOKENS,
    )
    return (response.choices[0].message.content or "").strip()


def is_anthropic_auth_error(exc):
    status = getattr(exc, "status_code", None)
    return status == 401 or exc.__class__.__name__ == "AuthenticationError" or "401" in str(exc)


def resolve_anthropic_auth_failure(clients):
    """401 from Anthropic: warn cleanly and let the user pick a fallback
    provider interactively instead of crashing the run."""
    log("Anthropic API authentication failed. Check ANTHROPIC_API_KEY in .env", ok=False)
    print("Fall back to another provider:")
    print("  1) Gemini (Gemini 3.6 Flash)")
    print("  2) Groq (GPT-OSS 120B)")
    choice_map = {"1": PROVIDER_GEMINI, "2": PROVIDER_GROQ}
    while True:
        choice = input("Enter choice [1-2]: ").strip()
        if choice in choice_map:
            fallback_provider = choice_map[choice]
            break
        warn("Invalid choice -- enter 1 or 2.")

    keys = collect_required_keys(fallback_provider)
    if keys is None:
        raise TailorError("No fallback provider key supplied -- cannot proceed.")
    clients.update(build_clients_for_provider(fallback_provider, keys))
    return fallback_provider


def generate_llm_response(clients, provider, prompt, system_instruction="", edits_paragraph_ids=None):
    """Centralized generation router. Dispatches to whichever provider was
    selected at bootup, enforcing the Gemini free-tier call cap and running
    the Gemini -> Groq automatic fallback when PROVIDER_FREE_FALLBACK is
    active."""
    global _gemini_call_count, _provider_override

    if provider == PROVIDER_ANTHROPIC and _provider_override:
        provider = _provider_override

    if provider == PROVIDER_ANTHROPIC:
        try:
            text = call_anthropic(clients["anthropic"], prompt, system_instruction)
        except Exception as exc:
            if is_anthropic_auth_error(exc):
                _provider_override = resolve_anthropic_auth_failure(clients)
                return generate_llm_response(
                    clients, _provider_override, prompt, system_instruction, edits_paragraph_ids
                )
            raise TailorError(f"Anthropic API call failed ({ANTHROPIC_MODEL}): {exc}") from exc

        if not text:
            warn("Anthropic returned empty text content. Triggering fallback to Gemini 3.6 Flash...")
            if "gemini" not in clients:
                gemini_keys = collect_required_keys(PROVIDER_GEMINI)
                if gemini_keys is None:
                    raise TailorError(
                        "Anthropic returned empty content and no GEMINI_API_KEY available for fallback."
                    )
                clients.update(build_clients_for_provider(PROVIDER_GEMINI, gemini_keys))
            return generate_llm_response(
                clients, PROVIDER_GEMINI, prompt, system_instruction, edits_paragraph_ids
            )

        return text

    if provider == PROVIDER_GEMINI:
        if _gemini_call_count >= GEMINI_FREE_TIER_CALL_CAP:
            raise TailorError(
                f"Gemini free-tier limit reached ({GEMINI_FREE_TIER_CALL_CAP} calls max). "
                "Re-run and pick Groq or the automatic fallback option."
            )
        try:
            _gemini_call_count += 1
            return call_gemini(clients["gemini"], prompt, system_instruction, edits_paragraph_ids)
        except Exception as exc:
            raise TailorError(f"Gemini API call failed ({GEMINI_MODEL}): {exc}") from exc

    if provider == PROVIDER_GROQ:
        try:
            return call_groq(clients["groq"], prompt, system_instruction)
        except Exception as exc:
            raise TailorError(f"Groq API call failed ({GROQ_MODEL}): {exc}") from exc

    if provider == PROVIDER_FREE_FALLBACK:
        raw = None
        if _gemini_call_count < GEMINI_FREE_TIER_CALL_CAP:
            try:
                _gemini_call_count += 1
                raw = call_gemini(clients["gemini"], prompt, system_instruction, edits_paragraph_ids)
            except Exception:
                warn(f"Gemini rate limit reached ({GEMINI_FREE_TIER_CALL_CAP} req max). Falling back to Groq...")
        else:
            warn(f"Gemini rate limit reached ({GEMINI_FREE_TIER_CALL_CAP} req max). Falling back to Groq...")

        if raw is not None:
            return raw

        try:
            return call_groq(clients["groq"], prompt, system_instruction)
        except Exception as exc:
            raise TailorError(f"Groq fallback API call failed ({GROQ_MODEL}): {exc}") from exc

    raise TailorError(f"Unknown provider selection: {provider}")


def clean_and_parse_json(raw_text):
    """Strip markdown code fences and surrounding chatter from an LLM
    response, then parse the remaining JSON payload."""
    if not raw_text or not raw_text.strip():
        raise ValueError("LLM returned an empty response string.")

    cleaned = re.sub(r"^```(?:json)?\s*", "", raw_text.strip(), flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```$", "", cleaned, flags=re.MULTILINE).strip()

    # Models sometimes leak preamble/thinking prose containing a stray "{" or
    # "[" before the real payload (e.g. "considering {A, B} pillars..."), and
    # Gemini can emit multiple concatenated top-level JSON values in one
    # response (e.g. a stray "{}" stub before the real payload). Try every
    # "{"/"[" position in the string rather than assuming the first one
    # starts valid JSON, and keep the largest value that successfully parses.
    decoder = json.JSONDecoder()
    length = len(cleaned)
    values = []
    idx = 0
    while idx < length:
        ch = cleaned[idx]
        if ch not in "{[":
            idx += 1
            continue
        try:
            obj, end_idx = decoder.raw_decode(cleaned, idx)
            values.append(obj)
            idx = end_idx
        except json.JSONDecodeError:
            idx += 1

    if not values:
        preview = cleaned[:200].replace("\n", "\\n")
        raise ValueError(f"LLM response contained no valid JSON. Got: {preview!r}")

    data = max(values, key=lambda v: len(json.dumps(v))) if len(values) > 1 else values[0]

    if isinstance(data, list):
        return {"items": data}
    return data


JSON_RETRY_ATTEMPTS = 3


def call_llm_json(clients, provider, prompt, system_instruction="", validate=None, edits_paragraph_ids=None):
    """Route through generate_llm_response and always surface a clean
    TailorError instead of a raw traceback. Both Anthropic (adaptive
    thinking) and Gemini occasionally return empty or degenerate output
    (e.g. Gemini echoing its own internal JSON-mode instructions instead of
    following them, or the wrong top-level shape) -- these are transient
    decoding glitches, so retry the whole generate+parse+validate round a
    few times before giving up. `validate`, if given, receives the parsed
    dict and must return the (possibly normalized) dict or raise ValueError."""
    last_exc = None
    for attempt in range(1, JSON_RETRY_ATTEMPTS + 1):
        raw = generate_llm_response(clients, provider, prompt, system_instruction, edits_paragraph_ids)
        try:
            data = clean_and_parse_json(raw)
            if validate is not None:
                data = validate(data)
            return data
        except ValueError as exc:
            last_exc = TailorError(str(exc))
        except json.JSONDecodeError as exc:
            last_exc = TailorError(f"LLM returned malformed JSON: {exc}")
        if attempt < JSON_RETRY_ATTEMPTS:
            warn(f"LLM returned unusable output (attempt {attempt}/{JSON_RETRY_ATTEMPTS}). Retrying...")
    raise last_exc


NBSP_CHARS = (" ", " ")


def sanitize_text(text):
    """Clean raw LLM text before it reaches python-docx: normalize non-
    breaking spaces and tighten spacing around percent signs."""
    for ch in NBSP_CHARS:
        text = text.replace(ch, " ")
    text = re.sub(r"(\d)\s+%", r"\1%", text)
    return text


def sanitize_edits(edits):
    return {pid: (sanitize_text(v) if v != REMOVE_TOKEN else v) for pid, v in edits.items()}


def sanitize_company_name(name):
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return name or "Company"


def iter_editable_paragraphs(doc):
    """Yield (id, paragraph) for every paragraph in the body and in tables."""
    for i, p in enumerate(doc.paragraphs):
        yield f"p{i}", p
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for pi, p in enumerate(cell.paragraphs):
                    yield f"t{ti}_r{ri}_c{ci}_p{pi}", p


def build_paragraph_map(doc):
    return {pid: p.text for pid, p in iter_editable_paragraphs(doc) if p.text.strip()}


def get_flat_runs(paragraph):
    """All runs in document order, including ones nested inside <w:hyperlink>.

    paragraph.runs only sees top-level <w:r> children -- it silently skips text
    inside hyperlinks (e.g. mailto:/tel: contact links), which left hyperlink
    text untouched by edits while a new run added the same content again,
    producing duplicated visible text.
    """
    flat = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            flat.extend(item.runs)
        else:
            flat.append(item)
    return flat


def apply_edits(doc, edits):
    to_remove = []
    for pid, p in list(iter_editable_paragraphs(doc)):
        if pid not in edits:
            continue
        new_text = edits[pid]
        if new_text == REMOVE_TOKEN:
            to_remove.append(p)
            continue
        runs = get_flat_runs(p)
        if not runs:
            continue
        runs[0].text = new_text
        for extra_run in runs[1:]:
            extra_run.text = ""
    for p in to_remove:
        p._element.getparent().remove(p._element)


def dedupe_edits(paragraph_map, edits):
    """Guard against the model echoing a paragraph's original text twice back-to-back."""
    cleaned = {}
    for pid, new_text in edits.items():
        if new_text == REMOVE_TOKEN:
            cleaned[pid] = new_text
            continue
        original = paragraph_map.get(pid, "").strip()
        if original and new_text.count(original) >= 2:
            idx = new_text.find(original)
            new_text = new_text[: idx + len(original)]
        cleaned[pid] = new_text
    return cleaned


def set_margins(doc, inches):
    for section in doc.sections:
        section.left_margin = Inches(inches)
        section.right_margin = Inches(inches)
        section.top_margin = Inches(inches)
        section.bottom_margin = Inches(inches)


def _is_header_style_text(style_name, text):
    if style_name not in HEADER_STYLE_NAMES:
        return False
    if style_name == "Heading 2":
        # master_resume.docx sometimes carries "Heading 2" forward onto bullet
        # lines typed directly after a role-title line. Real role/section
        # title rows use a trailing tab (to right-align the date range) and
        # pipe separators ("Title | Company | Location"); bullets don't.
        return "\t" in text or "|" in text
    return True


def is_header_paragraph(p, text):
    style_name = p.style.name if p.style else ""
    return _is_header_style_text(style_name, text)


def enforce_resume_typography(doc, body_font_pt=None, header_font_pt=None, line_spacing=1.15):
    """Ted Rogers layout lock: uniform Calibri, bold headers, RGB black for
    every run (avoiding Word's default blue hyperlink styling), 0pt before /
    3pt after -- applied to existing runs so borders, tables, and other
    document-level styling are untouched. Font size/line spacing are
    overridable so the one-page squeeze can tighten layout without ever
    touching the text itself."""
    body_size = body_font_pt if body_font_pt is not None else BODY_FONT_SIZE
    header_size = header_font_pt if header_font_pt is not None else HEADER_FONT_SIZE
    for _, p in iter_editable_paragraphs(doc):
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        pf.line_spacing = line_spacing
        is_header = is_header_paragraph(p, p.text)
        for run in get_flat_runs(p):
            run.font.name = "Calibri"
            run.font.size = header_size if is_header else body_size
            run.font.bold = is_header
            run.font.color.rgb = HEADER_FONT_COLOR


ONE_PAGE_MARGIN_STEPS_INCHES = [1.0, 0.85, 0.75, 0.6, 0.5]
ONE_PAGE_LINE_SPACING_STEPS = [1.15, 1.1, 1.0]
ONE_PAGE_FONT_STEPS_PT = [10.5, 10.0, 9.5, 9.0]
WD_PROPERTY_PAGES = 14


def count_word_pages(doc_path):
    """Ask an actual Word instance how many pages the saved .docx renders to
    -- python-docx has no layout engine, so this is the only way to know the
    real page count rather than guessing from character totals.
    ComputedStatistics isn't reachable via win32com's dynamic dispatch on this
    Word build, so use the repaginated BuiltInDocumentProperties page count
    instead -- same underlying layout engine, just a different accessor."""
    import win32com.client as win32

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        wd_doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
        try:
            wd_doc.Repaginate()
            return int(wd_doc.BuiltInDocumentProperties(WD_PROPERTY_PAGES).Value)
        finally:
            wd_doc.Close(False)
    finally:
        word.Quit()


def enforce_one_page(doc_path):
    """Strictly guarantee the saved resume renders to a single page in Word.
    Never touches the text -- only tightens margins, line spacing, and font
    size (in that priority order, mildest first) and re-measures the real
    page count until it fits or the floor settings are exhausted."""
    try:
        pages = count_word_pages(doc_path)
    except Exception as exc:
        warn(f"Could not verify page count via Word ({exc}) -- skipping one-page enforcement.")
        return

    if pages <= 1:
        log("Resume confirmed as one page.")
        return

    warn(f"Resume rendered as {pages} pages -- tightening layout (never trimming text)...")
    for margin_in in ONE_PAGE_MARGIN_STEPS_INCHES:
        for line_spacing in ONE_PAGE_LINE_SPACING_STEPS:
            for font_pt in ONE_PAGE_FONT_STEPS_PT:
                doc = Document(doc_path)
                set_margins(doc, margin_in)
                enforce_resume_typography(
                    doc,
                    body_font_pt=Pt(font_pt),
                    header_font_pt=Pt(max(font_pt, 10.5)),
                    line_spacing=line_spacing,
                )
                doc.save(doc_path)
                pages = count_word_pages(doc_path)
                if pages <= 1:
                    log(
                        f"Resume fit to one page (margins {margin_in}in, "
                        f"font {font_pt}pt, line spacing {line_spacing})."
                    )
                    return

    warn(
        f"Resume still renders as {pages} pages even at the tightest layout "
        "floor -- prune content in job.txt-driven edits or trim the master "
        "resume manually."
    )


def enforce_cover_typography(doc):
    """Mirror the resume's typography lock exactly: uniform 10.5pt Calibri,
    bold name header, RGB black for every run, 1.15 line spacing, 0pt before /
    3pt after, no block indents. The first paragraph is the applicant's name
    row (matching the resume's Title-style header); everything else is
    body/contact text."""
    for idx, (_, p) in enumerate(iter_editable_paragraphs(doc)):
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        pf.line_spacing = 1.15
        pf.left_indent = Pt(0)
        is_name_header = idx == 0
        for run in get_flat_runs(p):
            run.font.name = "Calibri"
            run.font.size = HEADER_FONT_SIZE if is_name_header else BODY_FONT_SIZE
            run.font.bold = is_name_header
            run.font.color.rgb = HEADER_FONT_COLOR


def resume_body_chars(edits):
    return sum(len(v) for v in edits.values() if v != REMOVE_TOKEN)


TERMINAL_PUNCTUATION = (".", "!", "?", ":")


def enforce_terminal_punctuation(doc, edits):
    """Guard against LLM truncation leaving a bullet mid-sentence: any
    non-header line without terminal punctuation gets a period appended.
    Header/role-title rows (dates, company names) are left untouched."""
    fixed = dict(edits)
    for pid, p in iter_editable_paragraphs(doc):
        if pid not in fixed:
            continue
        text = fixed[pid]
        if text == REMOVE_TOKEN:
            continue
        if is_header_paragraph(p, text):
            continue
        stripped = text.rstrip()
        if stripped and not stripped.endswith(TERMINAL_PUNCTUATION):
            fixed[pid] = stripped + "."
    return fixed


COOP_ELIGIBILITY_NOTE = "Co-op Availability: Jan 2027 - Dec 2027"


def inject_coop_eligibility(doc, edits):
    """Deterministically appends the candidate's co-op eligibility window to a
    non-header line inside the Education section, regardless of whether the
    model remembered to include it."""
    result = dict(edits)
    meta = [(pid, p.style.name if p.style else "", p.text) for pid, p in iter_editable_paragraphs(doc)]

    in_education = False
    target_pid = None
    for pid, style, original_text in meta:
        if style == "Heading 1":
            in_education = original_text.strip().lower() == "education"
            continue
        if not in_education:
            continue
        new_text = result.get(pid)
        if new_text is None or new_text == REMOVE_TOKEN:
            continue
        if not _is_header_style_text(style, new_text):
            target_pid = pid
            break

    if target_pid is None:
        return result

    current = result[target_pid].rstrip()
    if COOP_ELIGIBILITY_NOTE in current:
        return result
    if current.endswith(TERMINAL_PUNCTUATION):
        current = current[:-1]
    result[target_pid] = f"{current} | {COOP_ELIGIBILITY_NOTE}."
    return result


def cover_letter_words(edits):
    return sum(len(v.split()) for v in edits.values() if v != REMOVE_TOKEN)


RESUME_RULES = f"""
You are running a strict Job-Matching Framing Engine to tailor a resume for
Toronto Metropolitan University (TMU) Co-op standards. Work through two
explicit stages before you write anything.

STAGE 1 -- JOB POSTING ANALYSIS
Extract the 4-5 core operational pillars from the job posting (for example, a
pharma QA co-op posting's pillars might be: GMP compliance/procedures, project
lifecycle tracking, stakeholder coordination, process improvement/workshops,
MS Office/Power BI analytics). Report these verbatim in the "pillars" field of
your JSON response -- they drive every framing decision below.

STAGE 2 -- CONTEXT-ACTION-RESULT (CAR) BULLET FRAMING
For every bullet you write:
1. Open with a high-impact, domain-tailored power verb -- Coordinated,
   Executed, Oversaw, Standardized, Optimized, Audited, Streamlined, etc.
   Never open with a weak verb like "Helped" or "Worked on".
2. Synthesize the candidate's actual past duties to directly mirror the
   Stage 1 pillars (e.g. reframe general admin/process work as "project
   documentation, schedule maintenance, and cross-functional reporting" when
   the posting's pillar is project lifecycle tracking).
3. Quantify the result wherever the source material supports it: metrics,
   efficiency gains, accuracy percentages, team size, project scope, timeline.

HARD STRUCTURAL RULES (Ted Rogers layout compliance -- non-negotiable):
- Select EXACTLY 3 core experience roles: the 3 most relevant to the Stage 1
  pillars, kept in reverse-chronological order relative to each other.
- Write EXACTLY 3 bullets for each of those 3 roles. Never 2, never 4.
- Every single bullet must be {BULLET_MIN_CHARS}-{BULLET_MAX_CHARS} characters
  long (roughly 1.5-2 lines at 10.5pt) -- long enough that no line ends with a
  single orphaned word, short enough that it never wraps to a 3rd line. Count
  characters yourself and rework the wording until each bullet lands inside
  this window.
- Never cut a bullet off mid-sentence. Every bullet is a complete sentence
  ending in a period -- if a rewritten bullet would run past the character
  window, shorten the wording, don't truncate it.
- If you select the ECA Technologies or Combat Engineer roles, keep the real
  substance of their bullets intact: cross-functional stakeholder alignment
  (e.g. coordinating input across technical and business teams),
  documentation work (e.g. drafting compliant proposal packages, technical
  narratives), and compliance/safety details (e.g. safety compliance during
  field operations, regulatory/procurement compliance). Reframe the wording
  for the target job, but don't drop these specifics to make room for
  something else.
- No Summary of Qualifications, Objective, Photo, Personal Details, High
  School, or References sections.
- Present tense for the current role, past tense for completed roles.
- Never use "TRSM" or "Ryerson" -- always spell out "Toronto Metropolitan
  University".
- If the role is legal-focused: prioritize legal research, policy, and
  academic writing; drop tech/software engineering details.
- If the role is administrative/operations: emphasize workflow coordination,
  client service, and process automation.
- If the role is technical: surface codebases, frameworks, and developer
  tools.

STAGE 3 -- WHOLE-DOCUMENT PAGE BUDGET (applies to every section, not just
Work Experience -- this is the most common way this resume overflows to a
2nd page, so treat it as seriously as the bullet rules above):
- The master resume's Projects, Technical Skills & Certifications, Academic
  Legal/Regulatory Research, and Extracurricular sections all carry far more
  entries than fit on one printed page. You must actively prune them, not
  just pass them through:
  - Projects: select only the 1-2 entries most relevant to the Stage 1
    pillars. Set every paragraph belonging to a dropped project (its title
    AND every Purpose/What/How/Why/Outcome line under it) to the exact
    string "{REMOVE_TOKEN}". For the project(s) you keep, you may still
    condense their Purpose/How/Outcome lines to the most job-relevant
    sentence.
  - Technical Skills & Certifications: rewrite each category line to lead
    with the skills that match the pillars; REMOVE_TOKEN entire category
    lines that have no relevance to this job.
  - Academic Legal/Regulatory Research: keep only the 0-1 entries relevant
    to this job's pillars (0 if the job isn't legal/compliance-flavored).
    REMOVE_TOKEN the rest, title line included.
  - Extracurricular: keep only 1-2 lines that support the pillars (leadership,
    stakeholder management, etc.), REMOVE_TOKEN the rest.
  - If a whole section becomes empty after pruning, REMOVE_TOKEN its section
    heading paragraph too.
- Never achieve the page budget by cutting a bullet or sentence off
  mid-thought -- prune whole entries with "{REMOVE_TOKEN}" instead of
  shortening a kept line below its character floor.
- After pruning, the ENTIRE resume (every section combined) must read as one
  printed page at 10.5pt, 1-inch margins, 1.15 line spacing. If you are still
  unsure it fits, prune further rather than leaving it long.
"""

COVER_LETTER_RULES = f"""
You are running the Job-Matching Framing Engine on a cover letter for a
Toronto Metropolitan University (TMU) Co-op student.

RECIPIENT BLOCK & SALUTATION (fill in the existing header paragraphs, in this
exact order -- do not add or remove paragraphs):
1. Date: the current submission date.
2. Hiring manager info: their name if the job posting names one, otherwise
   the literal string "Hiring Manager"; then the employer name; then the
   employer address (city, province/state, and postal code if known).
3. Subject line: exactly "RE: Application for [Job Title] - [Job ID if
   available]" -- omit the "- [Job ID]" part entirely if the posting gives no
   job ID, don't invent one.
4. Salutation: "Dear Hiring Manager," (or "Dear [Name]," if a specific name
   was found). Never use "To Whom It May Concern" or "Sir/Madam".

BODY -- exactly 3-4 paragraphs:
1. Opening (Hook & Context): state the exact job title and company name, and
   that the applicant is an undergraduate student in the Law and Business
   Co-op program at Toronto Metropolitan University. Include a specific,
   researched hook -- something concrete about why this candidate wants THIS
   role at THIS company, not a generic "I am excited to apply" line.
2. Body (1-2 STAR paragraphs): draw on 2-3 relevant experiences (work,
   academic, or leadership) and connect the actions directly to what this
   employer needs, using Situation/Task -> Action -> Result. Do not quote or
   closely paraphrase resume bullet text verbatim -- tell it as a narrative,
   not a bullet list.
3. Closing (Call to Action): tailored to this employer and role -- restate
   the value proposition briefly, thank the reader, and express a clear wish
   to discuss the opportunity in an interview.

STRUCTURAL RULES (Ted Rogers layout compliance -- non-negotiable):
- Keep the header block (date, hiring manager, company, address, job ID,
  salutation) and the signature block (name, title, contact info) intact --
  only rewrite the field values inside them. Never delete these paragraphs.
- The applicant's own name and personal contact-info line (email, phone,
  city) at the very top of the letter is the candidate's identity block --
  it must exactly mirror the resume header (name, email, phone, location,
  LinkedIn). Copy it through byte-for-byte, unchanged. Never add a second
  email/phone, never prepend or append anything to it, and never invent a
  different contact method for the applicant.
- End with the formal sign-off already in the template: "Sincerely," followed
  by the applicant's full name on its own line. Leave this exactly as-is.
- Rewrite paragraph text in place; do not add extra line breaks or new
  paragraphs beyond what already exists in the template.
- The ENTIRE letter (header + body paragraphs + closing) must total no more
  than {COVER_MIN_WORDS}-{COVER_MAX_WORDS} words. This is a hard ceiling, not
  a target to approach from above -- be concise.
- Never use "TRSM" or "Ryerson" -- always spell out "Toronto Metropolitan
  University".
"""


def request_edits(clients, provider, rules, job_description, paragraph_map, extra_context=""):
    prompt = f"""{rules}

{extra_context}

JOB DESCRIPTION:
{job_description}

Below is a JSON object mapping paragraph IDs to their CURRENT text, in document
order. Your "edits" object MUST contain an entry for EVERY single ID listed
below -- no exceptions, including Technical Skills, Projects, Legal Research,
and Extracurricular paragraphs. Leaving an ID out is not a valid way to "keep
it as-is"; if you want to keep a paragraph unchanged, copy its exact current
text as the value. For each ID, decide the NEW text for that paragraph:
- Actively re-evaluate every paragraph against the rules above (including the
  Stage 3 page-budget pruning) -- don't default to echoing the input.
- Keep text that is relevant and rewrite it to satisfy the rules above.
- If a paragraph is irrelevant, redundant, or must be cut for the Stage 3 page
  budget, set its value to the exact string "{REMOVE_TOKEN}".
- Do not invent new paragraph IDs. Only use IDs that appear in the input.
- Do not leak any placeholder tokens, brackets, or raw XML into the new text.

CURRENT PARAGRAPHS:
{json.dumps(paragraph_map, indent=2)}

Respond with ONLY a JSON object of this exact shape, no markdown fences:
{{
  "company_name": "<company name extracted from the job description>",
  "pillars": ["<core operational pillar 1>", "<pillar 2>", "..."],
  "edits": {{ "<paragraph id>": "<new text or {REMOVE_TOKEN}>", ... }}
}}
"""
    def normalize_edits(data):
        """The Gemini structured-output schema pins this shape, but Anthropic
        and Groq get no such constraint, so tolerate the shapes models
        actually drift into and fold them into {"edits": {...}}."""
        if "edits" in data:
            return data

        if data and set(data.keys()) <= set(paragraph_map.keys()):
            # Bare paragraph-id -> text map, no wrapper object at all.
            return {"edits": data}

        items = data.get("items") if isinstance(data, dict) else None
        if isinstance(items, list) and items:
            edits = {}
            for item in items:
                if isinstance(item, dict):
                    if len(item) == 1:
                        edits.update(item)
                        continue
                    pid = item.get("id") or item.get("paragraph_id") or item.get("pid")
                    text = item.get("text") or item.get("new_text") or item.get("value")
                    if pid is not None and text is not None:
                        edits[pid] = text
                elif isinstance(item, (list, tuple)) and len(item) == 2:
                    edits[item[0]] = item[1]
            if edits:
                return {"edits": edits}

        raise ValueError(f"LLM response missing required 'edits' key. Got keys: {list(data.keys())}")

    return call_llm_json(
        clients, provider, prompt, validate=normalize_edits, edits_paragraph_ids=list(paragraph_map.keys())
    )


JOB_FILE = "job.txt"


def read_job_description():
    if not os.path.exists(JOB_FILE):
        with open(JOB_FILE, "w", encoding="utf-8") as f:
            f.write("")

    input(f"Paste/update the job description in '{JOB_FILE}', then press Enter to continue...")

    with open(JOB_FILE, "r", encoding="utf-8") as f:
        return f.read().strip()


def tailor(job_description=None):
    tommy_intro()

    if not os.path.exists(MASTER_RESUME):
        log(f"Missing {MASTER_RESUME} in current folder.", ok=False)
        sys.exit(1)
    if not os.path.exists(BASE_COVER_LETTER):
        log(f"Missing {BASE_COVER_LETTER} in current folder.", ok=False)
        sys.exit(1)

    if not job_description:
        job_description = read_job_description()
    if not job_description:
        log("No job description provided. Aborting.", ok=False)
        sys.exit(1)

    try:
        provider, clients = bootup()

        log("Processing master resume...")
        resume_doc = Document(MASTER_RESUME)
        resume_map = build_paragraph_map(resume_doc)
        resume_result = request_edits(clients, provider, RESUME_RULES, job_description, resume_map)
        resume_result["edits"] = sanitize_edits(resume_result["edits"])
        resume_result["edits"] = dedupe_edits(resume_map, resume_result["edits"])
        resume_result["edits"] = inject_coop_eligibility(resume_doc, resume_result["edits"])

        pillars = resume_result.get("pillars") or []
        if pillars:
            log("Job pillars: " + ", ".join(pillars))

        resume_result["edits"] = enforce_terminal_punctuation(resume_doc, resume_result["edits"])
        resume_chars = resume_body_chars(resume_result["edits"])
        log(f"Resume body: {resume_chars} chars (cap {RESUME_MIN_CHARS}-{RESUME_MAX_CHARS})")

        margin_inches = MARGIN_INCHES
        if resume_chars > RESUME_MAX_CHARS:
            margin_inches = MARGIN_FALLBACK_INCHES
            log(f"Content ran long ({resume_chars} chars) -- scaling margins to {margin_inches}in instead of truncating")

        apply_edits(resume_doc, resume_result["edits"])
        set_margins(resume_doc, margin_inches)
        enforce_resume_typography(resume_doc)
        company_name = sanitize_company_name(resume_result.get("company_name", "Company"))
        resume_out = f"resume_{company_name}.docx"
        resume_doc.save(resume_out)
        log(f"Complete -- saved {resume_out}")
        enforce_one_page(resume_out)

        log("Processing cover letter...")
        cover_doc = Document(BASE_COVER_LETTER)
        cover_map = build_paragraph_map(cover_doc)
        tailored_resume_text = "\n".join(
            v for v in resume_result["edits"].values() if v != REMOVE_TOKEN
        )
        cover_result = request_edits(
            clients,
            provider,
            COVER_LETTER_RULES,
            job_description,
            cover_map,
            extra_context=f"TAILORED RESUME CONTENT FOR CONTEXT:\n{tailored_resume_text}",
        )
        cover_result["edits"] = sanitize_edits(cover_result["edits"])
        cover_result["edits"] = dedupe_edits(cover_map, cover_result["edits"])
        cover_words = cover_letter_words(cover_result["edits"])
        log(f"Cover letter body: {cover_words} words (cap {COVER_MIN_WORDS}-{COVER_MAX_WORDS})")

        cover_margin_inches = MARGIN_INCHES
        if cover_words > COVER_MAX_WORDS:
            cover_margin_inches = MARGIN_FALLBACK_INCHES
            log(f"Cover letter ran long ({cover_words} words) -- scaling margins to {cover_margin_inches}in instead of truncating")

        apply_edits(cover_doc, cover_result["edits"])
        set_margins(cover_doc, cover_margin_inches)
        enforce_cover_typography(cover_doc)
        cover_out = f"cover_letter_{company_name}.docx"
        cover_doc.save(cover_out)
        log(f"Complete -- saved {cover_out}")
    except TailorError as exc:
        log(str(exc), ok=False)
        sys.exit(1)
    except PermissionError as exc:
        log(f"Could not write output file -- close it if it's open in Word: {exc}", ok=False)
        sys.exit(1)
    except Exception as exc:
        log(f"Unexpected error: {exc}", ok=False)
        sys.exit(1)

    rainbow_print("SUCCESS! TOMMY HAS APPROVED YOUR SUBMISSION FOR THE WAGE MEATGRINDER.")


def run(cmd):
    print(f"{Fore.CYAN}$ {' '.join(cmd)}{Style.RESET_ALL}")
    subprocess.run(cmd, check=True)


def git_init_and_push():
    rainbow_print("INITIALIZING GIT REPOSITORY & PUSHING TO xesex13/resume-tailor...")
    remote_url = "https://github.com/xesex13/resume-tailor.git"

    if not os.path.exists(".git"):
        run(["git", "init"])

    run(["git", "add", "main.py", "stress_test.py", ".gitignore", "README.md", "requirements.txt"])

    result = subprocess.run(["git", "diff", "--cached", "--quiet"])
    if result.returncode != 0:
        commit_message = (
            "Initial commit: TTTTMTTRT resume tailor\n\n"
            "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
        )
        run(["git", "commit", "-m", commit_message])
    else:
        rainbow_print("NOTHING NEW TO COMMIT.")

    remotes = subprocess.run(["git", "remote"], capture_output=True, text=True).stdout.split()
    if "origin" in remotes:
        run(["git", "remote", "set-url", "origin", remote_url])
    else:
        run(["git", "remote", "add", "origin", remote_url])

    run(["git", "branch", "-M", "main"])
    run(["git", "push", "-u", "origin", "main"])
    rainbow_print("PUSH COMPLETE.")


def main():
    parser = argparse.ArgumentParser(prog="TTTTMTTRT")
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("tailor", help="Tailor resume + cover letter to a pasted job description")
    sub.add_parser("git-init", help="Init git repo and push to xesex13/resume-tailor")
    sub.add_parser("stress-test", help="Run the 70-job automated stress test")

    args = parser.parse_args()

    if args.command == "git-init":
        git_init_and_push()
    elif args.command == "stress-test":
        subprocess.run([sys.executable, "stress_test.py"], check=True)
    else:
        tailor()


if __name__ == "__main__":
    main()
