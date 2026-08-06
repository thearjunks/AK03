# AK stc Dashboards

Local dashboards for STC labels, devices, and labs, including Strapi refresh,
comparison, language checks, and Excel exports.

## Requirements

- Node.js 18 or newer
- npm

## Install

```powershell
npm.cmd install
```

## Start the labels dashboard

```powershell
npm.cmd start
```

Open <http://localhost:4555/dashboard>.

Dashboard pages, exports, and API routes require login. Create a local `.env`
file from `.env.example` and set `DASHBOARD_USERNAME` and
`DASHBOARD_PASSWORD`. The session expires after 12 hours, and failed login
attempts are rate limited.

The production `start` command also launches this dashboard and uses the
`PORT` environment variable supplied by the hosting platform.

## Start the devices and labs dashboard

```powershell
npm.cmd run start:devices
```

Open <http://localhost:4321/>.

Strapi credentials are entered at runtime and are not stored in this repository.

## Scheduled labels deployment

GitHub Actions runs the labels fetch and Hostinger deployment once each Sunday through Thursday at **10:00 AM Kuwait time**. Friday, Saturday, and Kuwait public holidays are skipped before any data is fetched.

The workflow is `.github/workflows/scheduled-labels-deployment.yml`. It fetches the latest Strapi labels, commits the snapshot to `main`, waits for Hostinger to serve the same snapshot, and verifies the production `/labels` route.

Add this encrypted repository secret under **GitHub > Settings > Secrets and variables > Actions** before the first scheduled run:

- `DASHBOARD_PASSWORD`: the password used by the protected dashboard
- `SMTP_PASSWORD`: a Gmail app password for `thearjunks@gmail.com`

Set matching `DASHBOARD_USERNAME` and `DASHBOARD_PASSWORD` environment
variables in Hostinger so the production dashboard and scheduled verification
use the same credentials.

The workflow checks an automatically updated Kuwait holiday calendar. For last-minute official government announcements, set the optional repository variable `KUWAIT_HOLIDAY_DATES` to comma-separated `YYYY-MM-DD` dates. If the holiday calendar cannot be checked, deployment fails closed and no data is fetched or published.

Notifications are sent to `thearjunks@gmail.com` for data-fetch start/completion, dashboard update, GitHub commit start/completion, Hostinger deployment start/completion, and the final deployment summary. Each email includes Kuwait time, status, details, and relevant logs. The workflow also retains a downloadable audit artifact for 90 days.
