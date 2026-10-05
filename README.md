# SCB Video Library - fan website

A free, static website: search/filter every video, calendar, history heatmaps, stats, and a download
of the Excel tracker. No server, no accounts, nothing stored about visitors. Updates itself weekly.

## Publish it (free, ~10 minutes, GitHub Pages)
1. Create a free GitHub account, then a **public** repository (e.g. `scb-library`).
2. Upload everything in this folder (keep the `.github/workflows/pages.yml` path). Do **not** upload your personal
   `SCB_Library.xlsx` or `api_key.txt` (the included `.gitignore` blocks them).
3. Repo **Settings > Pages > Source: GitHub Actions**.
4. Repo **Settings > Secrets and variables > Actions > New repository secret**: name `YT_API_KEY`, value = your API key.
   (Use a NEW key restricted to "YouTube Data API v3"; your current one was pasted in a chat.)
5. **Actions tab > "Refresh data and deploy site" > Run workflow.** After ~2 minutes your site is live at
   `https://<your-username>.github.io/<repo-name>/`.

It then rebuilds every Monday from fresh YouTube data. Cost: $0 (a run uses ~150 of your 10,000 daily API units).
GitHub pauses scheduled runs after 60 days with no repo activity - if that happens, click "Run workflow" once.

## How it fits with your personal workbook
Your private `SCB_Library.xlsx` (with your watch marks) stays on your computer and is untouched. The website build starts
from scratch on GitHub's servers each time, so none of your personal data can reach the site.

## What the site includes
Overview, Library (search titles **and** descriptions as two separate lists; click any date to open it in the calendar),
Calendar (Sunday-first, prev/next), History, Stats, and an Excel-tracker page with download + update instructions.
Zero-length (0:00) placeholder videos are left out. Descriptions are shortened and de-duplicated (URLs and sponsor
boilerplate removed) so they can be searched without bloating the page.

## Files
- `site_template.html` - the page (edit this to change the look); `build_site.py` - injects data, builds the Excel download
- `scb_library.py` - the same engine as your desktop updater
- `preview/` - a ready-made copy built from your current data (open `index.html`). Its links fall back to YouTube searches
  because your current workbook has no video IDs yet; the published site has real links.

## Before you go public - please check
- Ask the Carlins (or at least keep the "unofficial fan project" footer and don't use their logos).
- Read YouTube's API Services Terms and Developer Policies. As I understand them: keep data fresh (the weekly job does), show links to
  YouTube's Terms and Google's Privacy Policy (the footer does), and don't imply endorsement. Rules change - verify current wording.
- Don't add ads, accounts or tracking without re-checking those terms.
