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
