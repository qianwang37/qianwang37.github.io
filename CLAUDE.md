# Qian Wang's Personal Website

## Project Overview
Academic personal website hosted on GitHub Pages at https://qianwang37.github.io
GitHub repo: https://github.com/qianwang37/qianwang37.github.io

## Site Structure
- `index.html` — About page (photo, bio, contact links)
- `research.html` — Publications and working papers
- `cv.html` — CV download link
- `resources.html` — Career and research resources
- `miscellaneous.html` — Personal fun facts
- `style.css` — All styling
- `sitemap.xml` / `robots.txt` — SEO files

## Content Workflow
Content is managed via Word documents in this folder:
- `About.docx` → updates `index.html`
- `Research.docx` → updates `research.html`
- `Resources.docx` → updates `resources.html`
- `Miscellaneous.docx` → updates `miscellaneous.html`
- `Qian Wang CV latest.pdf` → linked from `cv.html`

When the user says "update the docs", read all .docx files using Word COM (`New-Object -ComObject Word.Application`), compare against the current HTML, and update any changes.

## Deployment
After editing files, commit and push to deploy:
```
git add <files>
git commit -m "description"
git push
```
Changes go live at https://qianwang37.github.io within 1-2 minutes.

## Git Setup Note
On a fresh terminal, Git may not be on PATH. Refresh it with:
```powershell
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
```

## SEO
- Google Search Console is verified (meta tag in index.html)
- Sitemap submitted to Google
- Structured data (JSON-LD Person schema) in index.html
- Open Graph tags in index.html
