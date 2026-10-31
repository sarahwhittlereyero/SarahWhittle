#!/usr/bin/env python3
"""Build Sarah Whittle's CV from cv.json.

    .venv/bin/python build.py   # writes index.html, Sarah-Whittle-CV.pdf and Sarah-Whittle-CV.docx
"""
import json
import subprocess
from html import escape
from pathlib import Path

ROOT = Path(__file__).parent
PDF_NAME = "Sarah-Whittle-CV.pdf"
DOCX_NAME = "Sarah-Whittle-CV.docx"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

FONTS = "https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;800&family=Lato:ital,wght@0,400;0,700;1,400&display=swap"

CSS = """
:root{
  --navy:#183a59; --navy-line:#2f5375; --side-text:#eef3f8; --side-muted:#b9c8d7;
  --ink:#2b2d30; --muted:#5a6068; --page:#e7ebef; --paper:#ffffff;
  --head:'Poppins',system-ui,-apple-system,sans-serif; --body:'Lato',system-ui,-apple-system,sans-serif;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--page);color:var(--ink);font-family:var(--body);font-size:14.5px;line-height:1.5}
a{color:inherit;text-decoration:none}
a:hover{text-decoration:underline}
ul{list-style:none}

.sheet{max-width:900px;margin:32px auto;background:var(--paper);box-shadow:0 4px 20px rgba(24,58,89,.12);
  display:grid;grid-template-columns:300px 1fr;grid-template-areas:"side head" "side main"}

/* Sidebar */
.side{grid-area:side;background:var(--navy);color:var(--side-text);padding:56px 30px 40px}
.side .block{margin-bottom:30px}
.side .muted{color:var(--side-muted)}

/* Section headings: round icon + spaced caps */
.sec-head{display:flex;align-items:center;gap:12px;margin-bottom:14px}
.sec-head h2{font-family:var(--head);font-weight:600;font-size:1.05rem;letter-spacing:.2em;text-transform:uppercase}
.icon{width:34px;height:34px;border-radius:50%;display:grid;place-items:center;flex-shrink:0}
.icon svg{width:18px;height:18px}
.side .icon{background:var(--side-text);color:var(--navy)}
.main .icon{background:var(--navy);color:#fff}

.contact li{display:flex;align-items:center;gap:12px;margin-bottom:8px;word-break:break-word}
.contact svg{width:16px;height:16px;flex-shrink:0;color:var(--side-text)}
.downloads{display:flex;gap:8px;margin-top:14px}
.downloads a{border:1px solid var(--side-muted);border-radius:999px;padding:3px 12px;font-family:var(--head);font-size:.8rem}
.downloads a:hover{background:var(--side-text);color:var(--navy);text-decoration:none}
.edu{margin-bottom:14px}
.edu-school{font-weight:700}
.edu-degree{font-size:.93em}
.edu-note{font-size:.85em;margin-top:2px}
.bullets li{position:relative;padding-left:16px;margin-bottom:6px;font-family:var(--head);font-size:.88rem;font-weight:400}
.bullets li::before{content:"";position:absolute;left:2px;top:.62em;width:5px;height:5px;border-radius:50%;background:currentColor}
.lang-level{display:block;font-family:var(--body);font-size:.85rem;color:var(--side-muted)}

/* Header */
.head{grid-area:head;padding:56px 44px 8px}
h1{font-family:var(--head);font-weight:800;font-size:2.7rem;white-space:nowrap;line-height:1.02;text-transform:uppercase;letter-spacing:.01em;color:var(--ink)}
.subtitle{font-family:var(--head);font-size:1.15rem;letter-spacing:.08em;margin-top:14px}
.school{color:var(--muted);font-size:.95rem;margin:1px 0 6px}
.subtitle.role{margin-top:0}

/* Main column */
.main{grid-area:main;padding:24px 44px 40px}
.main section{margin-bottom:28px}
.about{text-align:justify;hyphens:auto}

.timeline{position:relative;padding-left:22px}
.timeline::before{content:"";position:absolute;left:4px;top:8px;bottom:6px;width:1.5px;background:var(--navy)}
.job{position:relative;margin-bottom:16px}
.job::before{content:"";position:absolute;left:-22px;top:6px;width:9px;height:9px;border-radius:50%;background:var(--navy)}
.job-title{font-family:var(--head);font-weight:500;font-size:1rem;line-height:1.35}
.job-meta{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;font-family:var(--head);font-size:.86rem}
.job-org{font-style:italic}
.job-dates{color:var(--muted);white-space:nowrap}
.job-body{color:var(--muted);font-size:.92rem;margin-top:4px}
.job-body li{position:relative;padding-left:14px;margin-bottom:3px}
.job-body li::before{content:"";position:absolute;left:2px;top:.6em;width:4px;height:4px;border-radius:50%;background:var(--navy)}

.lead{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}
.lead-org{font-family:var(--head);font-weight:600;font-size:.95rem;line-height:1.3;margin-bottom:4px}
.lead p{color:var(--muted);font-size:.88rem}

@media screen and (max-width:760px){
  .sheet{margin:0;grid-template-columns:1fr;grid-template-areas:"head" "side" "main";box-shadow:none}
  .head{padding:32px 16px 20px}
  .side{padding:28px 16px 8px}
  .main{padding:28px 16px}
  h1{font-size:2.1rem;white-space:normal}
  .lead{grid-template-columns:1fr}
}

@page{size:A4;margin:0}
@media print{
  html{font-size:12.4px}
  body{background:#fff;font-size:9.3pt;line-height:1.42}
  .bullets li{margin-bottom:3px}
  .sheet{margin:0;max-width:none;width:210mm;min-height:297mm;box-shadow:none;grid-template-columns:72mm 1fr;
    -webkit-print-color-adjust:exact;print-color-adjust:exact}
  .side{padding:11mm 8mm 8mm}
  .head{padding:11mm 11mm 1mm}
  .main{padding:4mm 11mm 6mm}
  .main section{margin-bottom:4.5mm}
  .sec-head{margin-bottom:2.5mm}
  .icon{width:28px;height:28px}
  .icon svg{width:15px;height:15px}
  .job{margin-bottom:2.6mm}
  .subtitle{margin-top:2.5mm}
  .side .block{margin-bottom:7mm}
  h1{font-size:28pt}
  .job,.edu,.lead>div{break-inside:avoid}
  .downloads{display:none}
}
"""

ICONS = {
    "user": '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-7 8-7s8 2.6 8 7z"/>',
    "briefcase": '<path d="M9 4h6a2 2 0 0 1 2 2v1h3a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2h3V6a2 2 0 0 1 2-2zm0 3h6V6H9z"/>',
    "cap": '<path d="M12 3 1 8.5l11 5.5 9-4.5V16h2V8.5z"/><path d="M5.5 12.6V16c0 1.9 2.9 3.5 6.5 3.5s6.5-1.6 6.5-3.5v-3.4L12 15.9z"/>',
    "star": '<path d="M12 2.5l2.8 6.3 6.9.6-5.2 4.6 1.6 6.8L12 17.3l-6.1 3.5 1.6-6.8L2.3 9.4l6.9-.6z"/>',
    "chat": '<path d="M4 3h16a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H9l-5 4v-4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/>',
    "heart": '<path d="M12 21s-8.5-5.4-8.5-11.2A4.8 4.8 0 0 1 12 6.6a4.8 4.8 0 0 1 8.5 3.2C20.5 15.6 12 21 12 21z"/>',
    "phone": '<path d="M6.6 10.8a15.1 15.1 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.25c1.1.37 2.3.57 3.6.57a1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1c0 1.25.2 2.45.57 3.57a1 1 0 0 1-.25 1z"/>',
    "mail": '<path d="M2 6a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2zm2.4 0L12 11.2 19.6 6zM20 8.2l-8 5.4-8-5.4V18h16z"/>',
    "linkedin": '<path d="M4 3a2 2 0 1 1 0 4 2 2 0 0 1 0-4zM2.2 8.5h3.6V21H2.2zM8.6 8.5h3.4v1.7c.5-.9 1.7-2 3.6-2 3.8 0 4.5 2.5 4.5 5.7V21h-3.6v-6.3c0-1.5 0-3.4-2.1-3.4s-2.4 1.6-2.4 3.3V21H8.6z"/>',
}


def icon(name):
    return f'<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">{ICONS[name]}</svg>'


def e(s):
    return escape(s or "")


def head(title, icon_name):
    return f'<div class="sec-head"><span class="icon">{icon(icon_name)}</span><h2>{title}</h2></div>'


def render(cv):
    c = cv["contact"]
    linkedin = (f'<a href="{e(c["linkedin_url"])}">{e(c["linkedin_label"])}</a>'
                if c.get("linkedin_url") else e(c["linkedin_label"]))
    contact = "".join(f"<li>{icon(i)}<span>{t}</span></li>" for i, t in [
        ("phone", f'<a href="tel:{e(c["phone"].replace(" ", ""))}">{e(c["phone"])}</a>'),
        ("mail", f'<a href="mailto:{e(c["email"])}">{e(c["email"])}</a>'),
        ("linkedin", linkedin),
    ])

    edu = "".join(
        f'<div class="edu"><div class="edu-school">{e(x["institution"])}</div>'
        f'<div class="edu-degree">{e(x["degree"])}</div><div class="muted edu-degree">{e(x["dates"])}</div>'
        + (f'<div class="muted edu-note">{e(x["note"])}</div>' if x.get("note") else "") + "</div>"
        for x in cv["education"])
    skills = "".join(f"<li>{e(s)}</li>" for s in cv["skills"])
    langs = "".join(f'<li>{e(l["language"])}<span class="lang-level">{e(l["level"])}</span></li>' for l in cv["languages"])

    def job(x):
        dates = " · ".join(filter(None, [x.get("type"), x.get("dates")]))
        body = ("<ul>" + "".join(f"<li>{e(b)}</li>" for b in x["bullets"]) + "</ul>"
                if x.get("bullets") else f'<p>{e(x.get("summary"))}</p>')
        return (f'<div class="job"><div class="job-title">{e(x["role"])}</div>'
                f'<div class="job-meta"><span class="job-org">{e(x["organisation"])}</span><span class="job-dates">{e(dates)}</span></div>'
                f'<div class="job-body">{body}</div></div>')

    jobs = "".join(job(x) for x in cv["experience"])
    lead = "".join(f'<div><div class="lead-org">{e(x["organisation"])}</div><p>{e(x["summary"])}</p></div>' for x in cv["leadership"])
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(cv["name"])} — CV</title>
<meta name="description" content="{e(cv["name"])} — {e(cv["headline"])}, {e(cv["school"])}. {e(cv["current_role"])}.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="{FONTS}" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="sheet">
<header class="head">
<h1>{e(cv['name'])}</h1>
<div class="subtitle">{e(cv["headline"])}</div>
<div class="school">{e(cv["school"])}</div>
<div class="subtitle role">{e(cv["current_role"])}</div>
</header>
<aside class="side">
<div class="block">{head("Contact", "phone")}<ul class="contact">{contact}</ul>
<div class="downloads"><a href="{PDF_NAME}">Download PDF</a><a href="{DOCX_NAME}">Word</a></div></div>
<div class="block">{head("Education", "cap")}{edu}</div>
<div class="block">{head("Skills", "star")}<ul class="bullets">{skills}</ul></div>
<div class="block">{head("Languages", "chat")}<ul class="bullets">{langs}</ul></div>
</aside>
<main class="main">
<section>{head("About Me", "user")}<p class="about">{e(cv["profile"])}</p></section>
<section>{head("Experience", "briefcase")}<div class="timeline">{jobs}</div></section>
<section>{head("Leadership & Volunteering", "heart")}<div class="lead">{lead}</div></section>
</main>
</div>
</body>
</html>
"""


def build_pdf():
    """Print index.html to PDF with headless Chrome (uses the page's print styles)."""
    if not Path(CHROME).exists():
        print(f"skipped {PDF_NAME}: Google Chrome not found")
        return
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                    "--virtual-time-budget=8000", f"--print-to-pdf={ROOT / PDF_NAME}",
                    (ROOT / "index.html").as_uri()],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print(PDF_NAME)


def build_docx(cv):
    try:
        from cv_docx import write_docx
    except ImportError:
        print(f"skipped {DOCX_NAME}: python-docx not installed (see README)")
        return
    write_docx(cv, ROOT / DOCX_NAME)
    print(DOCX_NAME)


def main():
    cv = json.loads((ROOT / "cv.json").read_text(encoding="utf-8"))
    (ROOT / "index.html").write_text(render(cv), encoding="utf-8")
    print("index.html")
    build_pdf()
    build_docx(cv)


if __name__ == "__main__":
    main()
