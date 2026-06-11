# ResumeIQ — AI Resume Analyzer

A full-stack Flask application that uses Google Gemini 2.5 Flash to analyze resumes, score ATS compatibility, and generate optimized PDF resumes from LaTeX.

## Setup

### 1. Prerequisites
- Python 3.9+
- pdflatex (install via `sudo apt install texlive-latex-base texlive-latex-extra` on Ubuntu)
- A Google Gemini API key (get one free at https://aistudio.google.com)

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the application
```bash
python app.py
```

Then open your browser at: **http://localhost:5000**

## Usage

1. **Enter your Gemini API key** at the top of the page
2. **Upload your resume** (PDF or DOCX) — you'll get a "Uploaded Successfully" alert
3. **Fill in your details** — name, job role, job description, preferred location
4. **Click "Analyze My Resume"** — Gemini AI will analyze everything and show:
   - ATS Compatibility Score (visual ring chart)
   - Job Alignment Score (visual ring chart)
   - Missing & matched keywords
   - Strengths and improvements
   - Overall summary
5. **Click "Generate Optimized Resume"** — AI generates LaTeX code, compiles it to PDF, and lets you download it

## Features
- PDF and DOCX resume parsing
- Gemini 2.5 Flash AI analysis
- Visual ATS and Job Alignment scores
- Keyword gap analysis
- AI-generated, ATS-optimized LaTeX resume
- PDF compilation via pdflatex
- Downloadable final resume named after the candidate

## Notes
- API keys are never stored server-side beyond the session
- Session data is in-memory; restart clears all sessions
- For production, use a database and Redis for sessions
