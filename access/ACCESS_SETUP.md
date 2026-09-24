# Setting up email-login access tiers (Cloudflare Zero Trust)

## What you're building

One Cloudflare Access **Application** per role, each pointing at its own
folder on your site, each gated by its own email allow-list:

```
yourdomain.pages.dev/admin/...              -> full company data
yourdomain.pages.dev/arm/naveen-sukka/...    -> Naveen's 7 centres only
yourdomain.pages.dev/arm/hithayathullah-pm/... -> Hitha's 7 centres only
yourdomain.pages.dev/arm/radhamani-munda/... -> Radhamani's 3 centres only
yourdomain.pages.dev/centre/<centre-name>/... -> one centre manager, one centre
```

Each folder is a complete, self-contained copy of the dashboard with its
own pre-filtered `data.json` sitting next to it — nothing a restricted
viewer's browser ever loads contains another region's data, so there's
nothing to leak even if someone inspects network traffic.

## One-time setup

### 1. Fill in real emails
Open `access/roles.csv`. Replace every `YOUR-EMAIL-HERE` with a real email
address. For centre managers, uncomment (remove the leading `# `) whichever
centre rows you actually want to grant access to — leave the rest commented
out, since a centre with no active row here simply gets no bundle built.

### 2. Push it
Commit and push `access/roles.csv` as normal. The GitHub Action now runs
`build_access_bundles.py` automatically after every `data.json` rebuild,
producing all the scoped folders under `access/site/` and committing them
back — same auto-deploy chain as everything else.

### 3. Check the generated policy list
After the Action runs, open `access/cloudflare-policies.md` in your repo —
it's auto-generated and lists exactly which emails need to go on which
Cloudflare Access policy. Keep this open while you do step 4.

### 4. Create the Cloudflare Access Applications
In Cloudflare: **Zero Trust → Access → Applications → Add an application →
Self-hosted.** For each entry in `cloudflare-policies.md`:

- **Application name:** whatever's readable, e.g. "Aptronix — Naveen (ARM)"
- **Session duration:** your call (e.g. 24 hours is reasonable for daily use)
- **Application domain:** your Pages domain, with the path from the doc —
  e.g. `yourdomain.pages.dev/arm/naveen-sukka`
- **Policy:** Create a policy, action **Allow**, include rule **Emails** →
  paste the exact email(s) listed for that entry.

Repeat once per bundle (5 applications for the test roster we validated;
however many you actually configure in `roles.csv`).

### 5. Share the right link with each person
Each person gets *their own URL* — the path from their Application, not the
root domain. Naveen's link is `yourdomain.pages.dev/arm/naveen-sukka/`, not
your main link. When they open it, Cloudflare shows a "verify your email"
screen, sends a one-time code, and once entered, they land straight on
their own scoped dashboard.

## What you don't need to do

- **No passwords to manage** — Cloudflare Access handles the email
  verification itself (one-time code to their inbox).
- **No per-person data files to hand-build** — the script derives every
  bundle from `roles.csv` + your existing `data.json`. Add a row, push,
  it appears.
- **The root domain and `/admin/` behave independently** — you can leave
  your current main link (whichever path you've been sharing) exactly as
  it is; adding tiers here doesn't change or break it.

## Changing someone's access later

- **Add a person:** add a row to `roles.csv`, push. Then add their email to
  the matching existing Cloudflare policy (or create a new Application if
  it's a brand-new scope) — the bundle folder either already exists or gets
  created on the next push.
- **Remove a person:** delete their row from `roles.csv`, push, and remove
  their email from the matching Cloudflare policy. Removing the last email
  from a policy doesn't delete the bundle folder or the Application — do
  that manually in Cloudflare if a role is being retired entirely.
- **A centre changes ARM, or a centre closes:** just fix it in Location
  Master in `master.xlsx` as normal — the next rebuild picks it up
  everywhere, bundles included.

## Note on the 50-user free tier

Cloudflare Zero Trust's free tier counts *unique people (by email)*, not
policies or applications — so one person with access to two bundles (e.g.
an ARM who's also a backup centre manager somewhere) still only counts
once. With 1 Admin + 3 ARMs + up to 17 centre managers, you're at 21 people
maximum right now, comfortably under the 50-user limit even if you added a
backup/secondary viewer for every single role.
