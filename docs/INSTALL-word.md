# Installing Recorde in Word

Recorde is a web-hosted Word add-in: the only thing installed on your computer
is a small `manifest.xml` that tells Word where the add-in lives
(`https://hyperion-git.github.io/recorde/`). No admin rights, no installer
binary, nothing written outside Word's add-in registry and cache. Documents
keep working without the add-in — equations are ordinary pictures with the
LaTeX stored inside them.

This file ships in `recorde-word-addin-<version>.zip` together with
`manifest.xml`, the installer scripts and `Recorde-Showcase.docx`.

## What you need

- **Word for Microsoft 365** on Windows or Mac (current channel), or **Word on
  the web**. Vector (SVG) equations need a Microsoft 365 desktop build since
  2019; baseline-aligned inline math needs Word 2507 or later (Windows) /
  16.99 (Mac). Word on the web inserts high-resolution raster pictures and
  cannot align the baseline. Perpetual/LTSC versions (2016/2019/2021/2024) are
  untested.
- Network access to `hyperion-git.github.io` on first use; Word caches the pane
  afterwards. MathJax and all fonts are served from that site — nothing is
  fetched from a CDN, and nothing leaves the document (see `PRIVACY.md`).

## Option A — installer script (Windows or Mac desktop)

Extract the whole zip (the scripts look for `manifest.xml` next to them).

**Windows:** double-click `install-windows.cmd`. It copies the manifest to
`%LOCALAPPDATA%\Recorde\manifest.xml` and registers it for your user account
under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` — the mechanism
Microsoft's own developer tooling uses for sideloading. Windows may show a
"publisher could not be verified" prompt for a script downloaded from the
internet: choose *Run*. Close Word completely (check for `WINWORD.EXE` in Task
Manager) and start it again.

**Mac:** in Terminal, `sh install-mac.command` (double-clicking works too unless
Gatekeeper blocks a downloaded script). It copies the manifest to
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`, Word's sideload
folder. Quit Word completely and start it again; Recorde is listed under
**Home → Add-ins** (*Developer Add-ins*).

**Uninstall:** `uninstall-windows.cmd` / `sh uninstall-mac.command`, then
restart Word.

## Option B — Upload My Add-in (any platform, including Word on the web)

1. In Word: **Home → Add-ins → More Add-ins** (older builds: *Get Add-ins*) →
   **My Add-ins** → **Upload My Add-in** → choose `manifest.xml`.
   Word on the web: **Home → Add-ins → More Settings → Upload My Add-in**.
2. The **Recorde** button appears on the Home tab and opens the task pane.

On the web the upload is stored in the browser: clearing site data or switching
browsers means uploading again.

## Option C — organisation-wide

An administrator deploys the same `manifest.xml` through the Microsoft 365 admin
centre (*Settings → Integrated apps → Upload custom apps*). Users then find
Recorde under **Add-ins → Admin Managed**; nothing to install locally.

## Check it works

Open `Recorde-Showcase.docx` from the zip: a three-page paper whose 76 equations
were produced without Word. Click any equation — the pane loads its LaTeX; edit
and **Update**. Then, in a new document, type some text, press the Recorde
button, enter `E = mc^2`, choose *Inline*, **Insert**.

The last line of the pane's **Settings** panel reads
`Host: Word <version> (<platform>) · insert: svg|png · baseline shift: yes|no · fields: yes|no`
and tells you which features this Word build supports.

| Capability | Word Windows/Mac (Microsoft 365) | Word on the web |
|---|---|---|
| Render, preview, click-to-edit, numbering (table/inline), recovery | ✅ | ✅ |
| Vector SVG equations | ✅ | ❌ raster PNG (300 dpi) instead |
| Inline math on the text baseline | ✅ (2507+ / 16.99+) | ❌ sits slightly high |
| Numbering with Word SEQ fields | ✅ | ❌ option disabled |

## Updating

Each release ships a new `manifest.xml`. Run the installer again (Option A) or
upload the new file (Option B); an existing registration is replaced in place.
Word and its embedded browser cache the pane: every release uses versioned URLs
so the new manifest loads the new pane, but if it still looks stale:

- Windows: **File → Options → Trust Center → Trust Center Settings → Trusted
  Add-in Catalogs → "Next time Office starts, clear all previously-started web
  add-ins cache"**, restart Word. Last resort: close Word and delete the
  contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (this removes every
  sideloaded add-in's cache; do not delete single files inside it).
- Mac: in the pane, use the personality menu (top-right of the task pane) →
  **Clear Web Cache**; or quit Word and clear the folders listed in Microsoft's
  [Clear the Office cache](https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache).
- Any platform with Node.js: `npx office-addin-cache clear`.

## Troubleshooting

- **No Recorde button after restart.** Word was still running (tray /
  Task Manager); or the manifest was uploaded and registered twice with
  different files — remove one (Option B: My Add-ins → *…* → Remove; Option A:
  uninstall script).
- **"This add-in could not be started" / blank pane.** Check the network: the
  pane loads from `https://hyperion-git.github.io/recorde/`. Right-click in the
  pane → *Inspect* opens DevTools; the Console shows the error.
- **Inline math sits high.** The host line says `baseline shift: no`: this Word
  build lacks the API (needs Microsoft 365 2507+/16.99+, not LTSC, not web).
- Anything else: <https://github.com/hyperion-git/recorde/issues>.
