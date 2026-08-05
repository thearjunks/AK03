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

The production `start` command also launches this dashboard and uses the
`PORT` environment variable supplied by the hosting platform.

## Start the devices and labs dashboard

```powershell
npm.cmd run start:devices
```

Open <http://localhost:4321/>.

Strapi credentials are entered at runtime and are not stored in this repository.

## Scheduled labels deployment

GitHub Actions runs the labels fetch and Hostinger deployment every Sunday through Thursday at:

- 10:00 AM Kuwait time
- 3:00 PM Kuwait time

The workflow is `.github/workflows/scheduled-labels-deployment.yml`. It fetches the latest Strapi labels, commits the snapshot to `main`, waits for Hostinger to serve the same snapshot, and verifies the production `/labels` route.

Add this encrypted repository secret under **GitHub > Settings > Secrets and variables > Actions** before the first scheduled run:

- `SMTP_PASSWORD`: a Gmail app password for `thearjunks@gmail.com`

Notifications are sent to `thearjunks@gmail.com` for fetch success, fetch failure, deployment success, and deployment failure. Failure emails include the available response and runtime logs.
