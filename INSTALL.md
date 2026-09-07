# INSTALL.md — Setting up `landsdetail2img.py`

This document walks you through installing everything `landsdetail2img.py` needs to run, on **Windows**, **macOS**, and **Linux**.

There are two separate things to install:

1. **Python packages** (via `pip`) — required on every platform.
2. **LibreOffice** — required only if you use the default `--engine libreoffice` (recommended, highest-fidelity rendering). If you pass `--engine pillow` instead, you can skip the LibreOffice install entirely.

---

## 1. Prerequisites

- **Python 3.9+** installed and available on your `PATH` (check with `python --version` or `python3 --version`).
- Internet access to download packages/installers (or an offline copy of the installers if your machine is air-gapped).

---

## 2. Create a virtual environment and install the Python dependencies

From the folder containing `requirements.txt`, first create a virtual environment (this keeps these packages isolated from the rest of your system's Python setup):

**Windows (PowerShell / Command Prompt):**
```powershell
python3 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> 💡 You'll need to re-run the `activate` command (the second line above) every time you open a new terminal window to work on this project — you'll know it worked because your terminal prompt will be prefixed with `(.venv)`.

This installs `python-docx`, `Pillow`, `lxml`, and `PyMuPDF` into the virtual environment only.

---

## 3. Install LibreOffice (required for `--engine libreoffice`, the default)

The script calls LibreOffice's command-line tool (`soffice`) behind the scenes to render your tables with full Word-like fidelity. You only need the free, official LibreOffice suite — no special edition or paid version.

### 🪟 Windows

**Option A — winget (recommended, Windows 10 1809+/11):**
```powershell
winget install --id TheDocumentFoundation.LibreOffice -e
```

**Option B — manual installer:**
1. Go to https://www.libreoffice.org/download/download-libreoffice/
2. Download the Windows installer (`.msi`) for your architecture (64-bit, unless you know you need 32-bit).
3. Double-click the downloaded file and follow the installation wizard (default options are fine).
4. LibreOffice installs by default to:
   ```
   C:\Program Files\LibreOffice\program\soffice.exe
   ```

**Verify the install** by opening a new PowerShell/Command Prompt window and running:
```powershell
soffice --version
```
If that command isn't recognized, the script will still auto-detect LibreOffice in the standard `C:\Program Files\LibreOffice\` location automatically. If you installed it somewhere non-standard, tell the script explicitly:
```powershell
python landsdetail2img.py "report.docx" -o out --soffice-path "C:\Path\To\soffice.exe"
```

### 🍎 macOS

**Option A — Homebrew (recommended):**
```bash
brew install --cask libreoffice
```

**Option B — manual installer:**
1. Go to https://www.libreoffice.org/download/download-libreoffice/
2. Download the `.dmg` for your Mac — **make sure to pick the correct one**: Apple Silicon (M1/M2/M3/M4) vs. Intel.
3. Open the `.dmg` and drag the LibreOffice icon into the **Applications** folder.
4. LibreOffice installs to:
   ```
   /Applications/LibreOffice.app/Contents/MacOS/soffice
   ```

**Verify the install:**
```bash
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```
The script auto-detects this standard path, so no extra flags are usually needed.

### 🐧 Linux

**Debian / Ubuntu (and derivatives):**
```bash
sudo apt update
sudo apt install libreoffice
```

**Fedora:**
```bash
sudo dnf install libreoffice
```

**Arch:**
```bash
sudo pacman -S libreoffice-fresh
```

**Flatpak (works on any distro with Flatpak set up):**
```bash
flatpak install flathub org.libreoffice.LibreOffice
```
> Note: if you install via Flatpak, the `soffice` binary may not be on your normal `PATH`. You'll likely need to pass `--soffice-path` pointing at the Flatpak wrapper, or install via your distro's native package manager instead for simplicity.

**Verify the install:**
```bash
soffice --version
```

---

## 4. Confirm everything works end-to-end

Make sure your virtual environment is still activated (prompt shows `(.venv)`), then run the script against any sample `.docx` file containing a table:

```bash
python landsdetail2img.py "path/to/your/report.docx" -o output_images
```

If LibreOffice was found successfully, you'll see:
```
Processing report.docx ...
  Rendering N chunk(s) via LibreOffice ...
  -> output_images/Some_Section-1.png  (...)
Done. Generated N image(s) in 'output_images/'.
```

If LibreOffice **cannot** be found, the script will print clear next steps, e.g.:
```
[!] Could not find a LibreOffice installation (the 'soffice' executable).
    Options:
      1. Install LibreOffice (free): https://www.libreoffice.org/download/
      2. If it's already installed somewhere non-standard, pass its path via:
         --soffice-path "C:\Path\To\soffice.exe"
      3. Or use the lower-fidelity, dependency-free fallback engine:
         --engine pillow
```

---

## 5. Optional: skip LibreOffice entirely

If you can't install LibreOffice (e.g. a locked-down corporate machine), you can use the pure-Python fallback engine instead — no external software required, just the pip packages from Step 2:

```bash
python landsdetail2img.py "report.docx" -o output_images --engine pillow
```

This produces slightly lower-fidelity output (font substitution instead of your exact installed fonts), but requires nothing beyond the virtual environment + `pip install -r requirements.txt` from Step 2.

---

## Troubleshooting quick reference

| Symptom | Fix |
|---|---|
| `soffice --version` not recognized after install | Restart your terminal (PATH changes need a fresh shell), or use `--soffice-path` to point directly at the executable. |
| Script says it can't find LibreOffice, but you know it's installed | Pass the exact path with `--soffice-path`, e.g. `--soffice-path "C:\Program Files\LibreOffice\program\soffice.exe"`. |
| `ModuleNotFoundError` for `docx`, `PIL`, `lxml`, or `fitz` | Make sure your virtual environment is activated (prompt shows `(.venv)`), then re-run `pip install -r requirements.txt`. |
| `'.venv' is not recognized` / activate script not found | Make sure you ran `python3 -m venv .venv` first in the same folder, and that the activation command matches your OS/shell (see Step 2 above). |
| Installed LibreOffice via Flatpak/Snap and script can't find it | Use `--soffice-path` with the wrapper path, or reinstall via your distro's native package manager (apt/dnf/pacman). |
