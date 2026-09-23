# Aptronix Service Dashboard — repo setup

## What's in this folder

```
.
├── index.html                          <- the dashboard itself
├── data.json                           <- pre-built from your current source/master.xlsx
├── scripts/
│   └── build_data.py                   <- rebuilds data.json from master.xlsx
├── source/
│   └── master.xlsx                     <- your data
└── .github/
    └── workflows/
        └── update-dashboard.yml        <- auto-rebuild on every push
```

This is everything — unzip straight into a new, empty repo folder and every path already lines up.

## One-time setup

### 1. Create the repo and push everything
```
git init
git add .
git commit -m "Initial dashboard setup"
git branch -M main
git remote add origin https://github.com/<you>/<repo-name>.git
git push -u origin main
```
`master.xlsx` is ~25MB, right at the browser upload limit — use `git push` (CLI or GitHub Desktop) for this initial push, not the drag-and-drop web uploader.

### 2. Let the GitHub Action write back to the repo
The workflow rebuilds `data.json` and commits it back automatically — that commit-back needs write permission, which is off by default:

**Settings → Actions → General → Workflow permissions → "Read and write permissions" → Save.**

Skip this and the Action will run but fail silently at the commit step.

### 3. Connect Cloudflare Pages to the repo
In Cloudflare: **Workers & Pages → Create → Pages → Connect to Git** → pick this repo. When it asks for build settings:
- **Framework preset:** None
- **Build command:** *(leave blank)*
- **Build output directory:** `/`

This is a static site with no build step — Cloudflare just needs to serve the files as they are.

That's it. Cloudflare Pages deploys automatically on every push to `main` — no extra config on the Cloudflare side beyond this one-time connection.

## Why your "update and it just appears on Cloudflare" requirement is already satisfied

The chain is two pushes, both automatic once you've done the one push yourself:

1. **You** update `source/master.xlsx` locally and `git push`.
2. **GitHub Actions** notices the change, runs `build_data.py`, and commits the new `data.json` back to `main` — that's a second, automatic push.
3. **Cloudflare Pages** redeploys on *every* push to `main`, so it fires again for that second commit too.

So from your side, the whole workflow is: update the file, push it, wait about a minute, refresh the Cloudflare link. Nothing else to trigger by hand.

## If something doesn't update

- **Pushed but the site didn't change:** check the repo's **Actions** tab — the "Update Dashboard Data" run should be green. If it's red, open it; the most common cause is step 2 above not being done.
- **Action succeeded but Cloudflare still looks stale:** check **Cloudflare → your project → Deployments** — it should show a new deployment matching the Action's commit. If Cloudflare didn't pick it up, the Git connection may need reconnecting.
- **Numbers look wrong, not just stale:** check the Action's log output directly — `build_data.py` prints warnings for things like unmatched Ship-To IDs or missing sheets, which show up there even when the run itself succeeds (green).
