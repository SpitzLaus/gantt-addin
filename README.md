# Gantt Chart Builder – PowerPoint Add-in

A lightweight PowerPoint task-pane add-in that lets you build **Gantt charts** the
ThinkCell way: enter tasks with start/end dates and progress, then insert them as
**native, fully editable PowerPoint shapes** (rectangles, text boxes and grid lines).

> **Willst du das Add-in mit Kollegen teilen – ohne Serverstart und ohne
> Installation?** Siehe **[DISTRIBUTION.md](DISTRIBUTION.md)** (einmal hosten,
> dann nur `install-addin.cmd` ausführen).

## Features

- Task list editor (name, start, end, % progress, color)
- Configurable time axis: weeks / months / quarters
- Progress overlay on each bar
- Automatic month/quarter/week grid lines and labels
- Output is plain PowerPoint shapes → move, restyle or animate them freely
- One-click "Load sample" to see an example

## Project structure

```
ThinkCell/
├── manifest.xml      Office Add-in manifest (Presentation host)
├── taskpane.html     Task pane UI
├── taskpane.css      Styling
├── taskpane.js       Gantt logic + Office.js rendering
├── assets/           Icons
└── package.json      Dev scripts
```

## Requirements

- PowerPoint on Windows/Mac (Microsoft 365) **or** PowerPoint on the web
- Requires the **PowerPointApi 1.4+** requirement set (`addGeometricShape`, `addLine`)
- Node.js (for the local dev server and sideloading tools)

## Run it locally

### Node.js on a machine without admin rights (portable)

If `npm` is not recognized and you cannot install with admin rights, use the
portable Node.js build (already set up for this project under
`%LOCALAPPDATA%\nodejs`). Behind a corporate proxy, download with your Windows
credentials:

```powershell
$ver='v24.18.0'; $zip="$env:TEMP\node.zip"
$wc = New-Object System.Net.WebClient
$wc.Proxy = [System.Net.WebRequest]::DefaultWebProxy
$wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$wc.DownloadFile("https://nodejs.org/dist/$ver/node-$ver-win-x64.zip", $zip)
Expand-Archive $zip -DestinationPath $env:LOCALAPPDATA -Force
Rename-Item "$env:LOCALAPPDATA\node-$ver-win-x64" "$env:LOCALAPPDATA\nodejs"
[Environment]::SetEnvironmentVariable('Path', "$env:LOCALAPPDATA\nodejs;" + [Environment]::GetEnvironmentVariable('Path','User'), 'User')
```

Open a **new** terminal afterwards so the PATH change takes effect.

### Steps

1. Install dependencies:

   ```powershell
   npm install
   ```

2. Install the local dev HTTPS certificate (once):

   ```powershell
   npm run certs
   ```

3. Start the HTTPS server. Easiest: **double-click `start-gantt.bat`**
   (keeps its own window open), or run:

   ```powershell
   npm run dev
   ```

   The files are served at `https://localhost:3000`. Leave the window open.

4. Sideload the add-in manually in PowerPoint:
   - Open PowerPoint (desktop) → **Insert** tab → **Add-ins / Get Add-ins**
     → **My Add-ins** → **Upload My Add-in**.
   - Choose `manifest.xml` from this folder.
   - (Web PowerPoint: **Insert → Office Add-ins → Upload My Add-in**.)

5. In PowerPoint open the **Home** tab → **Gantt** group → **Gantt Chart** to open
   the panel.

> `npm start` (automatic sideloading via `office-addin-debugging`) also works on
> unrestricted machines, but manual upload is the reliable path on locked-down
> corporate devices.

## Use it

1. Set the chart **Title** and the overall **Start / End** of the time axis.
2. Choose the axis granularity (Weeks / Months / Quarters).
3. Add tasks — each needs a name, a start date and an end date. Optionally set a
   progress % and a bar color.
4. Click **Insert Gantt chart**. The chart is drawn on the currently selected slide.

Because everything is a normal shape, you can select any bar and change its color,
size or position, or apply PowerPoint animations.

## Notes / troubleshooting

- "addGeometricShape is not a function" → your PowerPoint build is older than the
  1.4 requirement set. Update Office or use PowerPoint on the web.
- Icons in the manifest point to PNGs; the included `assets/icon.svg` is a source
  you can export to `icon-16.png`, `icon-32.png`, `icon-80.png` if you want custom
  ribbon icons.
