#!/usr/bin/env python3
"""
Terry Thomas Tyler Tristan Tucker Tobias Thaddeus Maximus Tommy The Resume Tailor
(TTTTMTTRT)

Local CLI that tailors master_resume.docx and base_cover_letter.docx to a pasted
job description using Gemini 2.5 Flash, editing the existing .docx files in place
so fonts/margins/styling survive.
"""

import argparse
import json
import os
import random
import re
import subprocess
import sys
import time

import colorama
from colorama import Fore, Style
from dotenv import load_dotenv
from docx import Document

colorama.init(autoreset=True)

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
MASTER_RESUME = "master_resume.docx"
BASE_COVER_LETTER = "base_cover_letter.docx"
GEMINI_MODEL = "gemini-2.5-flash"
REMOVE_TOKEN = "__REMOVE__"

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


def ensure_api_key():
    load_dotenv(ENV_PATH)
    key = os.environ.get("GEMINI_API_KEY")
    if key:
        return key
    key = input("FEED ME A GEMINI API KEY (Get a free one at aistudio.google.com): ").strip()
    with open(ENV_PATH, "a", encoding="utf-8") as f:
        f.write(f"GEMINI_API_KEY={key}\n")
    os.environ["GEMINI_API_KEY"] = key
    return key


def get_client():
    from google import genai
    return genai.Client(api_key=ensure_api_key())


def call_gemini_json(client, prompt):
    response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
    text = response.text.strip()
    text = re.sub(r"^```(json)?", "", text.strip())
    text = re.sub(r"```$", "", text.strip())
    return json.loads(text.strip())


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


def apply_edits(doc, edits):
    to_remove = []
    for pid, p in list(iter_editable_paragraphs(doc)):
        if pid not in edits:
            continue
        new_text = edits[pid]
        if new_text == REMOVE_TOKEN:
            to_remove.append(p)
            continue
        if not p.runs:
            continue
        p.runs[0].text = new_text
        for extra_run in p.runs[1:]:
            extra_run.text = ""
    for p in to_remove:
        p._element.getparent().remove(p._element)


RESUME_RULES = """
You are tailoring a resume for Toronto Metropolitan University (TMU) Co-op
standards. Rules you MUST follow:
- No Summary of Qualifications, Objective, Photo, Personal Details, High School,
  or References sections.
- Reverse-chronological order. Strong action verbs. Quantified impact
  (Action Verb + Context + Quantified Result).
- Present tense for current roles, past tense for completed roles.
- Never use "TRSM" or "Ryerson" -- always spell out "Toronto Metropolitan
  University".
- Select only the 3-4 most relevant roles/projects for the target job so the
  resume fits exactly on one page with no dead space and no page-two spillover.
- If the role is legal-focused: prioritize legal research, policy, and academic
  writing; drop tech/software engineering details.
- If the role is administrative/operations: emphasize workflow coordination,
  client service, and process automation.
- If the role is technical: surface codebases, frameworks, and developer tools.
"""

COVER_LETTER_RULES = """
You are tailoring a cover letter for a Toronto Metropolitan University (TMU)
Co-op student. Rules you MUST follow:
- Header block: date, hiring manager, company, address, job ID, salutation.
- Opening paragraph: the hook, TMU Co-op major/year, genuine interest in the role.
- 1-3 body paragraphs using the STAR method (Situation, Task, Action, Result),
  each linked to a specific job requirement.
- Closing paragraph with a clear call to action.
- Target length is about 3/4 of a page. Never use "TRSM" or "Ryerson" -- always
  spell out "Toronto Metropolitan University".
"""


def request_edits(client, rules, job_description, paragraph_map, extra_context=""):
    prompt = f"""{rules}

{extra_context}

JOB DESCRIPTION:
{job_description}

Below is a JSON object mapping paragraph IDs to their CURRENT text, in document
order. For each ID, decide the NEW text for that paragraph:
- Keep text that is relevant and rewrite it to satisfy the rules above.
- If a paragraph is irrelevant and should be deleted entirely, set its value to
  the exact string "{REMOVE_TOKEN}".
- Do not invent new paragraph IDs. Only use IDs that appear in the input.
- Do not leak any placeholder tokens, brackets, or raw XML into the new text.

CURRENT PARAGRAPHS:
{json.dumps(paragraph_map, indent=2)}

Respond with ONLY a JSON object of this exact shape, no markdown fences:
{{
  "company_name": "<company name extracted from the job description>",
  "edits": {{ "<paragraph id>": "<new text or {REMOVE_TOKEN}>", ... }}
}}
"""
    return call_gemini_json(client, prompt)


def read_job_description():
    rainbow_print("PASTE THE JOB DESCRIPTION BELOW. Finish with a blank line then Ctrl+D (or an empty line + Enter twice on Windows):")
    lines = []
    blank_streak = 0
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "":
            blank_streak += 1
            if blank_streak >= 2 and lines:
                break
        else:
            blank_streak = 0
        lines.append(line)
    return "\n".join(lines).strip()


def tailor(job_description=None):
    tommy_intro()

    if not os.path.exists(MASTER_RESUME):
        print(f"{Fore.RED}Missing {MASTER_RESUME} in current folder.{Style.RESET_ALL}")
        sys.exit(1)
    if not os.path.exists(BASE_COVER_LETTER):
        print(f"{Fore.RED}Missing {BASE_COVER_LETTER} in current folder.{Style.RESET_ALL}")
        sys.exit(1)

    client = get_client()

    if not job_description:
        job_description = read_job_description()
    if not job_description:
        print(f"{Fore.RED}No job description provided. Aborting.{Style.RESET_ALL}")
        sys.exit(1)

    rainbow_print("PARSING MASTER RESUME... PRUNING IRRELEVANT EXPERIENCES...")
    resume_doc = Document(MASTER_RESUME)
    resume_map = build_paragraph_map(resume_doc)
    resume_result = request_edits(client, RESUME_RULES, job_description, resume_map)
    apply_edits(resume_doc, resume_result["edits"])
    company_name = sanitize_company_name(resume_result.get("company_name", "Company"))
    resume_out = f"resume_{company_name}.docx"
    resume_doc.save(resume_out)
    rainbow_print(f"SAVED {resume_out}")

    rainbow_print("DRAFTING COVER LETTER WITH STAR METHOD...")
    cover_doc = Document(BASE_COVER_LETTER)
    cover_map = build_paragraph_map(cover_doc)
    tailored_resume_text = "\n".join(
        v for v in resume_result["edits"].values() if v != REMOVE_TOKEN
    )
    cover_result = request_edits(
        client,
        COVER_LETTER_RULES,
        job_description,
        cover_map,
        extra_context=f"TAILORED RESUME CONTENT FOR CONTEXT:\n{tailored_resume_text}",
    )
    apply_edits(cover_doc, cover_result["edits"])
    cover_out = f"cover_letter_{company_name}.docx"
    cover_doc.save(cover_out)
    rainbow_print(f"SAVED {cover_out}")

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
