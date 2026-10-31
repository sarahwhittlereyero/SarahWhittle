#!/usr/bin/env python3
"""Build Sarah Whittle's CV from cv.json.

    .venv/bin/python build.py   # writes index.html, Sarah-Whittle-CV.pdf and Sarah-Whittle-CV.docx
"""
import json
import re
import subprocess
from html import escape
from pathlib import Path

ROOT = Path(__file__).parent
PDF_NAME = "Sarah-Whittle-CV.pdf"
DOCX_NAME = "Sarah-Whittle-CV.docx"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

FONTS = "https://fonts.googleapis.com/css2?family=Nunito:wght@600;700;800&family=Nunito+Sans:ital,opsz,wght@0,6..12,400;0,6..12,600;0,6..12,700;1,6..12,400&display=swap"

CSS = """
/* Layout: banner with name and downloads, quick-facts strip, then a tabbed timeline beside a side column */
:root{
  --bg:#f7f9fc; --surface:#ffffff; --ink:#26303d; --muted:#64707f; --line:#e3e8ef;
  --hero:#2d5d8a; --hero-2:#3a6c9c; --on-hero:#ffffff; --accent:#2f74b5; --accent-soft:#e8f1fa;
  --display:'Nunito',system-ui,-apple-system,sans-serif; --body:'Nunito Sans',system-ui,-apple-system,sans-serif;
}
@media screen and (prefers-color-scheme:dark){:root{
  --bg:#111821; --surface:#19232f; --ink:#e7edf4; --muted:#a0abb9; --line:#2a3644;
  --hero:#1f4366; --hero-2:#2a5279; --accent:#86b6ea; --accent-soft:#1d2f45; color-scheme:dark}}
*{margin:0;padding:0;box-sizing:border-box}
a{color:inherit}
ul,ol{list-style:none}
:focus-visible{outline:2px solid var(--accent);outline-offset:3px;border-radius:4px}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
body{background:var(--bg);color:var(--ink);font-family:var(--body);font-size:16px;line-height:1.6}

/* Banner */
.hero{background:var(--hero);color:var(--on-hero);position:relative;overflow:hidden}
.hero::after{content:"";position:absolute;right:-120px;top:-120px;width:420px;height:420px;border-radius:50%;border:60px solid var(--hero-2);opacity:.7}
.hero-in{position:relative;z-index:1;max-width:1040px;margin:0 auto;padding:72px 16px 96px}
.eyebrow{font-family:var(--display);font-weight:700;font-size:.8rem;letter-spacing:.22em;text-transform:uppercase;opacity:.75}
h1{font-family:var(--display);font-weight:800;font-size:clamp(2.6rem,8vw,4.8rem);line-height:1.05;letter-spacing:-.01em;margin:8px 0 18px}
.sub{font-family:var(--display);font-weight:700;font-size:clamp(1.1rem,2.4vw,1.35rem)}
.sub2{opacity:.8;margin-top:2px}
.btns{display:flex;flex-wrap:wrap;gap:12px;margin-top:30px}
.btn{font-family:var(--display);font-weight:700;font-size:.95rem;padding:10px 20px;border-radius:999px;text-decoration:none;border:1.5px solid rgba(255,255,255,.5);transition:transform .15s,background .2s}
.btn:hover{transform:translateY(-2px);background:rgba(255,255,255,.12)}
.btn.solid{background:#fff;color:var(--hero);border-color:#fff}
.btn.solid:hover{background:#eef4fb}

/* Quick facts */
.stats{max-width:1040px;margin:-48px auto 0;padding:0 16px;position:relative;z-index:2}
.stats-in{background:var(--surface);border:1px solid var(--line);border-radius:18px;display:grid;grid-template-columns:repeat(4,1fr);box-shadow:0 10px 30px rgba(45,93,138,.10)}
.stat{padding:20px 22px;border-right:1px solid var(--line)}
.stat:last-child{border:0}
.stat b{display:block;font-family:var(--display);font-weight:800;font-size:1.7rem;color:var(--accent);line-height:1.1}
.stat span{font-size:.88rem;color:var(--muted)}

/* Body */
.wrap{max-width:1040px;margin:0 auto;padding:56px 16px 72px;display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:40px}
h2{font-family:var(--display);font-weight:800;font-size:1.05rem;letter-spacing:.12em;text-transform:uppercase;margin-bottom:14px}
.about{font-size:1.08rem;margin-bottom:40px;max-width:65ch}
.tabs{display:flex;gap:6px;border-bottom:1px solid var(--line);margin-bottom:24px;overflow-x:auto}
.tab{font-family:var(--display);font-weight:700;font-size:1rem;padding:10px 16px;border:0;background:none;color:var(--muted);cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-1px;white-space:nowrap}
.tab:hover{color:var(--ink)}
.tab[aria-selected="true"]{color:var(--ink);border-color:var(--accent)}
.panel-title,.print-only{display:none}
.mini{display:flex;flex-direction:column;padding:1.2mm 0;border-bottom:1px solid var(--line)}
.mini:last-child{border:0}
.mini b{font-family:var(--display);font-weight:800;line-height:1.25}
.mini span{color:var(--muted)}
.mini .mini-date{color:var(--accent);font-weight:700}
.mini .mini-note{font-size:.92em}

.tl{position:relative;padding-left:28px}
.tl::before{content:"";position:absolute;left:6px;top:8px;bottom:8px;width:2px;background:var(--line)}
.tl li{position:relative;margin-bottom:18px}
.dot{position:absolute;left:-28px;top:20px;width:14px;height:14px;border-radius:50%;background:var(--surface);border:3px solid var(--accent)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px 20px;transition:border-color .2s,box-shadow .2s}
.card:hover{border-color:var(--accent);box-shadow:0 6px 18px rgba(45,93,138,.10)}
.card-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}
h3{font-family:var(--display);font-weight:800;font-size:1.06rem;line-height:1.35}
.org{color:var(--muted);font-style:italic}
.date{font-size:.88rem;color:var(--accent);font-weight:700;margin:4px 0 6px}
.pill{font-family:var(--display);font-size:.72rem;font-weight:800;letter-spacing:.06em;text-transform:uppercase;padding:4px 10px;border-radius:999px;background:var(--accent-soft);color:var(--accent);white-space:nowrap}
.pill.full{background:var(--accent);color:var(--surface)}
.card li{position:relative;padding-left:16px;margin-bottom:5px;color:var(--muted)}
.card li::before{content:"";position:absolute;left:2px;top:.68em;width:6px;height:6px;border-radius:50%;background:var(--accent)}
.card p{color:var(--muted)}

/* Side column */
.box{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:20px;margin-bottom:20px}
.box h2{font-size:.88rem;margin-bottom:10px}
.contact a{display:block;padding:5px 0;text-decoration:none;color:var(--accent);overflow-wrap:anywhere}
.contact a:hover{text-decoration:underline}
.skills li{padding:8px 0;border-bottom:1px solid var(--line);font-size:.94rem;transition:padding .15s,color .15s}
.skills li:last-child{border:0}
.skills li:hover{padding-left:6px;color:var(--accent)}
.lang{display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid var(--line);flex-wrap:wrap}
.lang:last-child{border:0}
.lang-name{font-weight:700}
.lang-level{color:var(--muted);font-size:.9rem}
footer{text-align:center;color:var(--muted);font-size:.85rem;padding:0 16px 40px}

@media screen and (max-width:860px){
  .wrap{grid-template-columns:minmax(0,1fr)}
  .stats-in{grid-template-columns:repeat(2,1fr)}
  .stat:nth-child(2){border-right:0}
  .stat:nth-child(-n+2){border-bottom:1px solid var(--line)}
  .hero-in{padding:56px 16px 88px}
}

/* One A4 page for the PDF: every tab's content shown, interactive parts hidden */
@page{size:A4;margin:0}
@media print{
  html{font-size:11px}
  body{background:#fff;font-size:9pt;line-height:1.4;-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .hero::after,.btns,.stats,.tabs,footer{display:none}
  .hero-in{max-width:none;padding:7mm 13mm 5mm}
  h1{font-size:28pt;margin:1mm 0 2mm}
  .eyebrow{font-size:7.5pt}
  .sub{font-size:12pt}
  .wrap{max-width:none;padding:6mm 13mm 0;grid-template-columns:minmax(0,1fr) 60mm;gap:8mm}
  .about{font-size:9.3pt;margin-bottom:3mm;max-width:none}
  h2{font-size:9.5pt;margin-bottom:2mm}
  #p-edu,#p-vol{display:none!important}
  .print-only,.panel-title{display:block}
  .box{font-size:8.3pt;line-height:1.35}
  .tl li{margin-bottom:1.6mm}
  .dot{top:3.2mm;width:10px;height:10px;border-width:2.5px;left:-26px}
  .card{padding:1.6mm 3mm;border-radius:8px;box-shadow:none!important}
  .tl::before{top:3mm;bottom:3mm}
  h3{font-size:10pt}
  .date{margin:0.5mm 0 1mm;font-size:8.5pt}
  .card li{margin-bottom:.6mm}
  .box{padding:3.4mm 4mm;margin-bottom:3.5mm;border-radius:8px}
  .box h2{font-size:8.5pt;margin-bottom:1.5mm}
  .contact a{padding:.5mm 0}
  .lang{padding:.9mm 0}
  .skills li{padding:.7mm 0}
  .vol-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:4mm}
  .vol-grid .mini{border:0;padding:0;font-size:8.3pt;line-height:1.35}
  .vol-print{margin-top:3mm}
  .card,.box{break-inside:avoid}
}
"""


def e(s):
    return escape(s or "")


def timeline(items):
    return '<ol class="tl">' + "".join(f'<li><div class="dot"></div><div class="card">{x}</div></li>' for x in items) + "</ol>"


def job(x):
    kind = x.get("type") or "Full-time"
    body = ("<ul>" + "".join(f"<li>{e(b)}</li>" for b in x["bullets"]) + "</ul>"
            if x.get("bullets") else f'<p>{e(x.get("summary"))}</p>')
    return (f'<div class="card-head"><div><h3>{e(x["role"])}</h3><div class="org">{e(x["organisation"])}</div></div>'
            f'<span class="pill{"" if x.get("type") else " full"}">{e(kind)}</span></div>'
            f'<div class="date">{e(x["dates"])}</div>{body}')


def render(cv):
    c = cv["contact"]
    years = [int(y) for x in cv["experience"] for y in re.findall(r"\b(?:19|20)\d{2}\b", x.get("dates", ""))]
    since = f"roles since {min(years)}" if years else "roles"
    work = timeline(job(x) for x in cv["experience"])
    vol = timeline(f'<h3>{e(x["organisation"])}</h3><p>{e(x["summary"])}</p>' for x in cv["leadership"])
    edu = timeline(
        f'<h3>{e(x["degree"])}</h3><div class="org">{e(x["institution"])}</div><div class="date">{e(x["dates"])}</div>'
        + (f'<p>{e(x["note"])}</p>' if x.get("note") else "")
        for x in cv["education"])
    edu_print = "".join(
        f'<div class="mini"><b>{e(x["degree"])}</b><span>{e(x["institution"])}</span><span class="mini-date">{e(x["dates"])}</span>'
        + (f'<span class="mini-note">{e(x["note"])}</span>' if x.get("note") else "") + "</div>"
        for x in cv["education"])
    vol_print = "".join(f'<div class="mini"><b>{e(x["organisation"])}</b><span>{e(x["summary"])}</span></div>' for x in cv["leadership"])
    skills = "".join(f"<li>{e(s)}</li>" for s in cv["skills"])
    langs = "".join(f'<div class="lang"><div class="lang-name">{e(l["language"])}</div><div class="lang-level">{e(l["level"])}</div></div>'
                    for l in cv["languages"])
    linkedin = (f'<a href="{e(c["linkedin_url"])}">LinkedIn: {e(c["linkedin_label"])}</a>' if c.get("linkedin_url") else "")
    linkedin_btn = (f'<a class="btn" href="{e(c["linkedin_url"])}">LinkedIn</a>' if c.get("linkedin_url") else "")
    degree_year = max(re.findall(r"\d{4}", cv["education"][0]["dates"]), default="")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(cv["name"])} — CV</title>
<meta name="description" content="{e(cv["name"])} — {e(cv["current_role"])}. {e(cv["headline"])}, {e(cv["school"])}.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="{FONTS}" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="hero"><div class="hero-in">
  <div class="eyebrow">Curriculum Vitae</div>
  <h1>{e(cv["name"])}</h1>
  <p class="sub">{e(cv["current_role"])}</p>
  <p class="sub2">{e(cv["headline"])} · {e(cv["school"])}</p>
  <div class="btns"><a class="btn solid" href="{PDF_NAME}">Download PDF</a><a class="btn" href="{DOCX_NAME}">Word version</a>{linkedin_btn}</div>
</div></header>
<div class="stats"><div class="stats-in">
  <div class="stat"><b>{len(cv["experience"])}</b><span>{since}</span></div>
  <div class="stat"><b>{len(cv["languages"])}</b><span>languages</span></div>
  <div class="stat"><b>MSc</b><span>RSM, expected {degree_year}</span></div>
  <div class="stat"><b>{len(cv["leadership"])}</b><span>volunteer roles</span></div>
</div></div>
<div class="wrap">
<main>
  <h2>About me</h2><p class="about">{e(cv["profile"])}</p>
  <div class="tabs" role="tablist" aria-label="Background">
    <button class="tab" role="tab" id="tab-work" aria-selected="true" aria-controls="p-work">Experience</button>
    <button class="tab" role="tab" id="tab-edu" aria-selected="false" aria-controls="p-edu">Education</button>
    <button class="tab" role="tab" id="tab-vol" aria-selected="false" aria-controls="p-vol">Volunteering</button>
  </div>
  <div role="tabpanel" id="p-work" aria-labelledby="tab-work"><h2 class="panel-title">Experience</h2>{work}</div>
  <section class="print-only vol-print"><h2>Leadership &amp; Volunteering</h2><div class="vol-grid">{vol_print}</div></section>
  <div role="tabpanel" id="p-edu" aria-labelledby="tab-edu" hidden><h2 class="panel-title">Education</h2>{edu}</div>
  <div role="tabpanel" id="p-vol" aria-labelledby="tab-vol" hidden><h2 class="panel-title">Leadership &amp; Volunteering</h2>{vol}</div>
</main>
<aside>
  <div class="box contact"><h2>Contact</h2><a href="mailto:{e(c["email"])}">{e(c["email"])}</a><a href="tel:{e(c["phone"].replace(" ", ""))}">{e(c["phone"])}</a>{linkedin}</div>
  <div class="box"><h2>Skills</h2><ul class="skills">{skills}</ul></div>
  <div class="box"><h2>Languages</h2>{langs}</div>
  <div class="box print-only"><h2>Education</h2>{edu_print}</div>
</aside>
</div>
<footer>© 2026 {e(cv["name"])}</footer>
<script>
const tabs=[...document.querySelectorAll('[role=tab]')];
tabs.forEach(t=>t.addEventListener('click',()=>tabs.forEach(x=>{{const on=x===t;x.setAttribute('aria-selected',on);document.getElementById(x.getAttribute('aria-controls')).hidden=!on;}})));
</script>
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
