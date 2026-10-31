# Sarah Whittle

Personal CV. View it live at: https://sarahwhittlereyero.github.io/SarahWhittle/

The CV comes in three versions, all built from the same content:

- **Website:** `index.html`, hosted with GitHub Pages
- **PDF:** `Sarah-Whittle-CV.pdf` (one A4 page)
- **Word:** `Sarah-Whittle-CV.docx` (one A4 page)

## Updating the CV

All CV content lives in `cv.json`. Edit it, then rebuild all three versions:

```bash
.venv/bin/python build.py
```

Then commit and push. The live site updates a minute or two later.

The PDF is printed with Google Chrome, so Chrome must be installed. The Word file needs `python-docx`. To set that up once on a new computer:

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx
```

## Files

| File | What it is |
|---|---|
| `cv.json` | All CV content: profile, experience, education, skills, languages |
| `build.py` | Builds the website and PDF from `cv.json` |
| `cv_docx.py` | Builds the Word version |
| `index.html` | Generated website. Don't edit by hand, as `build.py` overwrites it |

## GitHub Pages

The site is served from the `main` branch, root folder (**Settings** > **Pages** > **Deploy from a branch**).
