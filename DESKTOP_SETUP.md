# Excel AI Assistant — Desktop Setup Guide

> Works in **Excel Desktop** (Windows) via the same Office Add-in manifest.  
> The add-in runs in Excel's embedded WebView2 browser, which connects to the
> local Python HTTP bridge server (plain HTTP on localhost — no certificate needed).

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows 10/11 with Excel 2016+ | Microsoft 365 works best |
| Python 3.8+ | Added to `PATH` |
| OpenCode CLI | Installed via npm (`opencode --version` works); `serve` auto-starts |

---

## One-Time Setup (run once per machine)

### Step 1 — Run the setup script as Administrator (optional / legacy cert step)

Right-click `setup_desktop.ps1` → **Run with PowerShell as Administrator**

> The bridge server runs plain HTTP on localhost, so Excel Desktop needs no
> certificate. This step is optional/legacy — it only generates and trusts a
> self-signed localhost certificate for environments that enable HTTPS.

This will:
1. Generate a self-signed SSL certificate for `localhost` (if not already present)
2. **Trust it** in the Windows Certificate Store (only needed if you switch the bridge to HTTPS)
3. Print sideloading instructions

> **Why admin?** The Windows certificate store (Local Machine → Trusted Root) requires elevated permissions.

### Step 2 — Sideload the manifest into Excel

Your manifest is already registered via `register_catalog.py`
(registry key `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\<add-in ID>`).
If you ever need to re-register, run:

```
python register_catalog.py
```

then **fully quit Excel** (Task Manager → no `EXCEL.EXE` left) and reopen it.

**Where to find the add-in (Desktop Excel):**

1. Open Excel with your workbook
2. Go to **Insert → My Add-ins → MY ADD-INS** tab
   (newer UI: **Home → Add-ins → My Add-ins**)
3. You should see **AI Assistant** listed there → click **Add**
4. If it's not visible, click the **Refresh** button in that dialog
5. The **AI Assistant** button will then appear in the Home ribbon → click to open the taskpane

> ⚠️ Do NOT use Developer tab → Excel Add-ins → Browse — that only loads
> legacy COM/VBA add-ins (`.xlam`), not Office Web Add-ins like this one.

**Alternative — Shared Folder Catalog (only if the above fails):**

1. Share a local folder (folder **Properties → Sharing → Share**) and note its
   network UNC path, e.g. `\\YOUR-PC\Addins`
   (a plain local path like `C:\...` will NOT work here — Excel requires a UNC path)
2. In Excel: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
3. Paste the UNC path as **Catalog Url**, tick **Show in Menu**, restart Excel
4. Copy `manifest.xml` into the shared folder
5. **Home → Add-ins → Advanced → SHARED FOLDER → AI Assistant**

---

## Every Session

1. **Double-click `start_server.bat`** — starts the bridge server on `http://localhost:3000`
   (OpenCode `serve` on `127.0.0.1:4096` auto-starts on the first chat request)
2. Open Excel with your workbook
3. Open the taskpane via **Insert → My Add-ins → AI Assistant** (Home ribbon button after first add)

> Keep the `start_server.bat` window open while using the add-in.

---

## Troubleshooting

### Taskpane blank or won't load

Make sure `start_server.bat` is running and `http://localhost:3000/taskpane.html`
opens in a normal browser. Desktop Excel uses plain HTTP on localhost — no certificate needed.

### Taskpane shows "Could not connect to local bridge server"

The Python server isn't running. Start `start_server.bat` and try again.

### Add-in not showing after sideloading

- Fully quit Excel first (check Task Manager for lingering `EXCEL.EXE`)
- Re-run `python register_catalog.py`, reopen Excel
- Look under **Insert → My Add-ins → MY ADD-INS** (not Developer → Browse,
  not Shared Folder) and press **Refresh** in that dialog

### "Failed to call OpenCode server"

The bridge auto-starts `opencode serve --port 4096` on first chat. If it keeps failing:
```
opencode --version
```
then manually run `opencode serve --port 4096` in a terminal and retry chat.

---

## Architecture (Desktop vs Web)

```
Excel Desktop
  └── WebView2 (embedded browser)
        └── taskpane.html @ http://localhost:3000
              ├── Office.js → reads/writes active workbook directly
              └── fetch → http://localhost:3000/api/chat
                    └── bridge_server.py (Python HTTP server)
                          └── opencode serve @ 127.0.0.1:4096 (auto-started)
                                └── opencode/* models
```

The same add-in code works in both **Excel Web** (browser) and **Excel Desktop** (WebView2).  
There is no certificate requirement — the bridge runs plain HTTP on localhost
(`setup_desktop.ps1` cert step is optional/legacy).

---

## To Uninstall

Remove the trusted certificate from the Windows store:

```powershell
.\setup_desktop.ps1 -Uninstall
```

Then remove the add-in from Excel via **Insert → My Add-ins → manage**.
