# TTTTMTTRT

**T**erry **T**homas **T**yler **T**ristan **T**ucker **T**obias **T**haddeus **M**aximus **T**ommy **T**he **R**esume **T**ailor

A local CLI that tailors your master resume and cover letter to any job posting
using Gemini 2.5 Flash, editing your existing `.docx` files in place so your
fonts, margins, and styling never break.

## Install

```bash
pip install google-genai python-docx colorama python-dotenv
```

## Setup

Place these two files in the project root:

- `master_resume.docx` -- your complete work/academic/project history
- `base_cover_letter.docx` -- your cover letter template

On first run, if no `GEMINI_API_KEY` is found, you'll be prompted for one
(free at [aistudio.google.com](https://aistudio.google.com)) and it's saved
to a local `.env` file automatically.

## Usage

Tailor a resume + cover letter to a job posting (paste the job description
when prompted):

```bash
python main.py tailor
```

Outputs are written to the project root as:

- `resume_[Company_Name].docx`
- `cover_letter_[Company_Name].docx`

Run the 70-job automated stress test (14 simulated days x 5 postings/day)
across legal, administrative, software engineering, customer service, and
corporate finance roles:

```bash
python main.py stress-test
```

Initialize git and push this project to `xesex13/resume-tailor`:

```bash
python main.py git-init
```

## Resume & Cover Letter Rules

- No Summary of Qualifications, Objective, Photo, Personal Details, High
  School, or References sections.
- Reverse-chronological order, strong action verbs, quantified impact.
- Selects the 3-4 most relevant roles/projects per posting to fit exactly one
  page.
- Never uses "TRSM" or "Ryerson" -- always "Toronto Metropolitan University".
- Cover letters follow the STAR method (Situation, Task, Action, Result) and
  target ~3/4 page.
