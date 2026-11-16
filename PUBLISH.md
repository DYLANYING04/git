# Publish To GitHub

Create a new empty public repository, then run from this directory:

```powershell
git init
git add .
git commit -m "feat: add self-contained Nanyan proposal skill"
git branch -M main
git remote add origin https://github.com/<your-account>/<your-repo>.git
git push -u origin main
```

Before pushing, verify:

```powershell
git status --short
git ls-files | Select-String 'reports|\.xlsx$|\.docx$|\.jpg$|\.png$'
```

The second command should show only the bundled template path if you intentionally keep it tracked. Never commit questionnaires, contacts, generated reports, attachments, cached search results, or monthly configs containing private data.
