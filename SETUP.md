# Running the Market Desk from GitHub

Four files live in the repo root, plus one workflow:

    index.html                      the dashboard
    mintec_refresh.py               the refresh script
    materials.csv                   the material list and Mintec codes
    .github/workflows/refresh.yml   the hourly job

## 1. Add your credentials

Repo → Settings → Secrets and variables → Actions → New repository secret.

    MINTEC_CLIENT_ID        from Mintec Analytics → User profile → APIs Access
    MINTEC_CLIENT_SECRET    same screen

Secrets are write-only. Nobody can read them back, including you, and they
never appear in the repo or the logs.

## 2. Turn on Pages

Repo → Settings → Pages → Source: **Deploy from a branch**, branch `main`,
folder `/ (root)`. Your URL appears within a minute:

    https://<your-account>.github.io/<repo-name>/?cast=1&reload=60

## 3. Check the job

Actions tab → **Refresh market desk** → **Run workflow**. Watch the log. You
should see each material come back with settled and forward point counts, then
either a commit or "No price changes this hour."

After that it runs itself at five past every hour.

## Point the screen at it

Any device with a browser. Nothing to install.

    Chrome kiosk:
    chrome.exe --kiosk --app="https://<account>.github.io/<repo>/?cast=1&reload=60"

    iPad: open the URL in Safari, Share → Add to Home Screen, launch from there.

Over HTTPS the fullscreen button works properly, which it does not from a
local file. The page reloads hourly and picks up whatever the job last pushed.

## Worth knowing

- **Licensing.** GitHub Pages is public on free plans. That publishes Mintec
  price data to anyone with the URL. Confirm with Mintec that this is allowed,
  or use a private repo on a Team plan where Pages is access-controlled.
- **Timing.** Scheduled Actions are best-effort and can run several minutes
  late when GitHub is busy. Fine for an hourly commodity feed.
- **Dormancy.** GitHub disables scheduled workflows in repos with no activity
  for 60 days. The hourly commits count as activity, so this only bites if
  prices stop changing entirely.
- **Cost.** Public repos get unlimited Actions minutes. Private repos get 2,000
  free minutes a month; this job uses roughly one minute per run, so about 730
  a month — inside the free allowance, but keep an eye on it if you add more.
- **A bad hour is harmless.** If Mintec is unreachable the job fails, commits
  nothing, and the dashboard keeps serving the last good data.
