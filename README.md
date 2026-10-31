# Sarah Whittle

Personal CV website. View it live at: https://sarahwhittle.github.io/SarahWhittle/

## Setup

This site is a single `index.html` file designed to be hosted with [GitHub Pages](https://pages.github.com/).

To enable it:
1. Go to your repository **Settings** > **Pages**
2. Under **Source**, select **Deploy from a branch**
3. Choose **main** branch and **/ (root)** folder
4. Click **Save**

Your site will be live within a minute or two.

## Updating the CV

All CV content lives in `cv.json`. After editing it, regenerate the page:

```bash
python3 build.py
```
