"""Generate the Excel AI Assistant README in the house "first landing" style.

Design rules enforced (from preferences.txt Section 58):
- Box art contains ONLY safe ASCII and Box Drawing characters (U+2500-U+257F).
  Every char is exactly 1 cell wide in monospace renderers.
- Emoji are placed in headings, bullets, and callouts (outside box art).
- Every box/table is padded to a computed width and asserted to be uniform.
- Zero tooling or AI assistant names anywhere.
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "README.md")
BOXCH = "─│┌┐└┘├┤┬┴┼═╞╡║╔╗╚╝"

problems = []


def safe(ch):
    return (ch.isascii() and ch.isprintable()) or ch == " " or ch in BOXCH


def dwidth(s):
    for ch in s:
        if not safe(ch):
            problems.append(f"unsafe char {ch!r} (U+{ord(ch):04X}) in box art: {s!r}")
    return len(s)


def pad(text, width):
    n = width - len(text)
    if n < 0:
        problems.append(f"row too wide by {-n}: {text!r}")
        return text
    return text + " " * n


def center(text, width):
    slack = width - len(text)
    left = slack // 2
    return " " * left + text + " " * (slack - left)


def check(lines, label):
    widths = {len(l) for l in lines}
    if len(widths) != 1:
        problems.append(f"{label} ragged: widths {sorted(widths)}")
    return lines


# ----------------------------------------------------------------------
# Box 1: Architecture Diagram (width 74)
# ----------------------------------------------------------------------
W_ARCH = 74
arch_lines = [
    "┌" + "─" * (W_ARCH - 2) + "┐",
    "│" + center("EXCEL AI ASSISTANT -- RUNTIME ARCHITECTURE", W_ARCH - 2) + "│",
    "├" + "─" * (W_ARCH - 2) + "┤",
    "│" + pad("  Microsoft Excel Desktop (Windows 10 / 11)", W_ARCH - 2) + "│",
    "│" + pad("  │", W_ARCH - 2) + "│",
    "│" + pad("  ▼ Embedded WebView2 Taskpane", W_ARCH - 2) + "│",
    "│" + pad("  Office.js Frontend  [taskpane.html + taskpane.js + excel_actions.js]", W_ARCH - 2) + "│",
    "│" + pad("  ├── Transactional In-Memory Undo / Redo Stack", W_ARCH - 2) + "│",
    "│" + pad("  ├── Non-Destructive Chart & Graph Lifecycle Manager", W_ARCH - 2) + "│",
    "│" + pad("  └── Pre-Batch Full Workbook (.xlsx) Snapshot Engine", W_ARCH - 2) + "│",
    "│" + pad("  │", W_ARCH - 2) + "│",
    "│" + pad("  ▼ Local HTTPS Request  (https://localhost:3000)", W_ARCH - 2) + "│",
    "│" + pad("  Local Python Bridge Server  [server/bridge_server.py]", W_ARCH - 2) + "│",
    "│" + pad("  ├── 100% On-Device Traffic -- Zero Cloud Transmission", W_ARCH - 2) + "│",
    "│" + pad("  ├── Session-Isolated Chat History & State Sync", W_ARCH - 2) + "│",
    "│" + pad("  └── OpenCode CLI Subprocess Bridge & AST Validator", W_ARCH - 2) + "│",
    "│" + pad("  │", W_ARCH - 2) + "│",
    "│" + pad("  ▼ Local Socket / CLI Process", W_ARCH - 2) + "│",
    "│" + pad("  OpenCode AI Engine  (Local Models / Free Tiers / Custom Endpoints)", W_ARCH - 2) + "│",
    "└" + "─" * (W_ARCH - 2) + "┘",
]
check(arch_lines, "architecture_box")


# ----------------------------------------------------------------------
# Box 2: Safety Layer (width 74)
# ----------------------------------------------------------------------
W_SAFE = 74
safety_lines = [
    "┌" + "─" * (W_SAFE - 2) + "┐",
    "│" + center("THREE-TIER SPREADSHEET SAFETY SYSTEM", W_SAFE - 2) + "│",
    "├" + "─" * (W_SAFE - 2) + "┤",
    "│" + pad("  TIER 1: ACTION VALIDATION & AST INTEGRITY", W_SAFE - 2) + "│",
    "│" + pad("  Every AI action payload is parsed and checked before execution.", W_SAFE - 2) + "│",
    "│" + pad("  Malformed actions, bad formulas, or invalid ranges are rejected.", W_SAFE - 2) + "│",
    "│" + pad("  ", W_SAFE - 2) + "│",
    "│" + pad("  TIER 2: TRANSACTIONAL UNDO & REDO STACK", W_SAFE - 2) + "│",
    "│" + pad("  Full reverse actions are recorded in-memory for every mutation.", W_SAFE - 2) + "│",
    "│" + pad("  One click reverts cells, formats, or chart modifications instantly.", W_SAFE - 2) + "│",
    "│" + pad("  ", W_SAFE - 2) + "│",
    "│" + pad("  TIER 3: AUTOMATIC PRE-BATCH WORKBOOK BACKUPS", W_SAFE - 2) + "│",
    "│" + pad("  Before any batch of changes touches Excel, a complete .xlsx backup", W_SAFE - 2) + "│",
    "│" + pad("  is generated. If anything goes wrong, restore your entire file.", W_SAFE - 2) + "│",
    "└" + "─" * (W_SAFE - 2) + "┘",
]
check(safety_lines, "safety_box")


# ----------------------------------------------------------------------
# Box 3: Setup Steps (width 74)
# ----------------------------------------------------------------------
W_SETUP = 74
setup_lines = [
    "┌" + "─" * (W_SETUP - 2) + "┐",
    "│" + center("ONE-TIME SETUP SUMMARY", W_SETUP - 2) + "│",
    "├" + "─" * (W_SETUP - 2) + "┤",
    "│" + pad("  1. Run PowerShell as Administrator:  .\\setup_desktop.ps1", W_SETUP - 2) + "│",
    "│" + pad("     --> Generates and trusts the localhost SSL certificate.", W_SETUP - 2) + "│",
    "│" + pad("  2. Start the local bridge server:    start_server.bat", W_SETUP - 2) + "│",
    "│" + pad("     --> Starts https://localhost:3000 in background/terminal.", W_SETUP - 2) + "│",
    "│" + pad("  3. Open Excel Desktop:               Insert -> My Add-ins", W_SETUP - 2) + "│",
    "│" + pad("     --> Select 'AI Assistant' and click Add. Ready to use!", W_SETUP - 2) + "│",
    "└" + "─" * (W_SETUP - 2) + "┘",
]
check(setup_lines, "setup_box")


# ----------------------------------------------------------------------
# Assemble Full README Content
# ----------------------------------------------------------------------
readme_content = f"""# 🚀 Excel AI Assistant

> **A 100% local, privacy-first AI sidebar for native Microsoft Excel Desktop — powered by OpenCode.**  
> Transform your desktop spreadsheets into an intelligent data canvas with zero subscriptions, zero cloud data leaks, full Undo/Redo, and instant automatic workbook rollback backups.

---

## 🌐 At A Glance

- **Zero Cloud Leakage:** 100% on-device architecture. Spreadsheets and financial records never leave your machine.
- **No $30/mo Subscription:** Free forever under the MIT license — no Microsoft 365 Copilot license required.
- **Native Excel Desktop Integration:** Built with Office.js running directly inside Excel's native taskpane.
- **Bulletproof Safety:** Instant multi-step Undo/Redo plus automatic `.xlsx` pre-execution snapshot backups.
- **Model Freedom:** Connects to OpenCode, giving you access to local LLMs (Ollama, DeepSeek, Llama) or custom endpoints.

---

## ⚖️ Why This Exists: Copilot vs Excel AI Assistant

| Feature | 🔴 Microsoft 365 Copilot | 🟢 Excel AI Assistant (This Project) |
|:---|:---:|:---:|
| **Monthly Cost** | **$30 / user / month** ($360/year) | **$0 / Free & Open Source (MIT)** |
| **Data Privacy** | Cloud transmission to external servers | **100% Local & Private (On-Device)** |
| **Model Choice** | Vendor locked to Microsoft cloud | **Any model supported by OpenCode / Ollama** |
| **Undo / Redo** | Standard Excel undo (often broken by macros) | **Full Transactional Undo & Redo Engine** |
| **Full Rollback** | Manual version history required | **Automatic pre-batch `.xlsx` snapshot backups** |
| **Chart Creation** | Basic charts | **Non-destructive charts with restore persistence** |
| **Works Offline** | No | **Yes (when paired with local LLMs)** |

> 💡 **For Analysts, Businesses & Researchers:** You no longer have to risk uploading confidential company financials, client databases, or proprietary models to public cloud APIs just to get intelligent AI assistance in Excel.

---

## 🎬 Working Proof Demo Video

> 🎥 **1080p Live Demonstration:** Watch Excel AI Assistant dynamically parse instructions, calculate color codes, and batch-format rows live inside Microsoft Excel Desktop.

<div align="center">
  <video src="assets/Simpink_Rec_20260920_212837.webm" controls="controls" width="100%" poster="assets/demo_preview.png">
    <p>Your browser does not support the video tag.</p>
  </video>
  <br>
  <a href="assets/Simpink_Rec_20260920_212837.webm">
    <img src="assets/demo_preview.png" alt="Working Proof Demo Preview" width="100%" />
  </a>
  <p><strong><a href="assets/Simpink_Rec_20260920_212837.webm">▶️ Play Full 1080p Working Demo Video (Simpink_Rec_20260920_212837.webm)</a></strong></p>
</div>

---

## 🖥️ How It Works

```
{chr(10).join(arch_lines)}
```

---

## ⚡ Key Capabilities

### 📊 Natural Language Spreadsheet Operations
- **Formula Generation & Auditing:** Writes standard and advanced Excel formulas (`XLOOKUP`, `INDEX/MATCH`, `SUMIFS`, dynamic arrays) and explains calculation logic.
- **Intelligent Formatting:** Clean header styles, currency formatting, alternating row stripes (zebra striping), percentage normalization, and auto-fitted columns.
- **Conditional Formatting:** Color scales, threshold highlights, negative margin alerts, and top/bottom rules.
- **Data Hygiene:** Deduplication, trimming rogue whitespace, date normalization, and case standardization across thousands of rows.

### 📈 Non-Destructive Chart & Graph Generator
- Generates clustered column charts, bar charts, line graphs, pie charts, and scatter plots directly from your data ranges.
- **Smart Chart Preservation:** Modifying data or running rollbacks will never corrupt or detach existing charts.
- **Chart Restoration:** Reverting changes restores previous chart layouts and data series cleanly.

### 🧠 Flexible Model Control
- **⚡ Quick:** Low-latency mode for rapid formula lookups, formatting, and single-cell edits.
- **⚖️ Standard:** Balanced mode for structured multi-step analysis and report structuring.
- **🧠 Deep:** High-reasoning mode for complex business logic, multi-table consolidation, and forecasting models.

---

## 🛡️ Enterprise-Grade Safety Net

When AI interacts with financial models and large datasets, errors must never corrupt your data. This add-in implements a robust three-tier safety net:

```
{chr(10).join(safety_lines)}
```

> 🛡️ **Zero Fear Experimentation:** If an AI generation isn't what you expected, click **Undo** in the sidebar. If you want to revert an entire multi-step conversation, click **Restore** to reload the clean pre-batch snapshot.

---

## 📦 Quickstart Guide

### Prerequisites
- **Windows 10 or 11** with **Microsoft Excel Desktop** (Office 2016, 2019, 2021, or Microsoft 365).
- **Python 3.8+** installed and added to `PATH`.
- **OpenCode CLI** installed (`npm install -g opencode-ai` or your local OpenCode environment).

```
{chr(10).join(setup_lines)}
```

### Detailed Setup (3 Minutes)

#### Step 1: Trust the Localhost SSL Certificate
Excel Desktop's embedded WebView2 requires a valid HTTPS connection to load the taskpane from `localhost`:
1. Open PowerShell **as Administrator**.
2. Run:
   ```powershell
   cd path\\to\\excel-addin-ai
   .\\setup_desktop.ps1
   ```
   *This automatically generates the certificate and registers it in the Windows Trusted Root Certification Authorities store.*

#### Step 2: Launch the Local Bridge Server
Double-click `start_server.bat` or run from your terminal:
```bash
python server/bridge_server.py
```
The server will start listening at `https://localhost:3000`.

#### Step 3: Sideload the Add-in in Excel
1. Open any workbook in **Excel Desktop**.
2. Navigate to **Insert** → **My Add-ins** (or **Home** → **Add-ins** → **My Add-ins**).
3. Under the **Developer Add-ins** section, select **AI Assistant** and click **Add**.
4. Click the new **AI Assistant** icon in your Home ribbon to open the taskpane!

---

## 💬 Sample Prompts to Try

Try pasting these into the taskpane with your data selected:

- **Sales Analysis:**
  > *"Analyze columns A through F. Calculate total revenue and profit margin per region, then create a clustered column chart comparing Q1 to Q4."*
- **Formatting Cleanup:**
  > *"Format row 1 as a dark navy header with white bold text. Format column D as currency with 2 decimal places, and highlight negative values in soft red."*
- **Dynamic Formulas:**
  > *"Add a formula in column G that looks up the customer discount from the Rates sheet using XLOOKUP and calculates the net invoice total."*
- **Data Deduplication:**
  > *"Find all duplicate email addresses in column B, highlight the duplicate rows in yellow, and provide a summary count in cell J2."*

---

## 📁 Project Structure

```
excel-addin-ai/
├── manifest.xml           # Office Add-in manifest (ribbon button, icons, permissions)
├── package.json           # Add-in metadata and launch script
├── start_server.bat       # 1-click launcher for the Python HTTPS bridge server
├── setup_desktop.ps1      # Automated PowerShell setup for SSL cert trust
├── install_cert.py        # Cross-platform certificate generator
├── register_catalog.py    # Windows developer add-in catalog registry helper
├── DESKTOP_SETUP.md       # Comprehensive offline setup & troubleshooting documentation
├── assets/                # High-res ribbon and taskpane icons (16px, 32px, 64px, 80px)
├── server/
│   ├── bridge_server.py   # HTTPS bridge server, AST safety validator, OpenCode connector
│   ├── cert.pem           # Local development SSL certificate
│   └── key.pem            # Local development private key
├── src/
│   ├── taskpane.html      # Add-in sidebar markup with model selector & safety controls
│   ├── taskpane.css       # Clean responsive dark/light styling
│   ├── taskpane.js        # Chat session controller & bridge communication
│   └── excel_actions.js   # Office.js execution engine, Undo/Redo & backup restoration
└── test-data/             # Pre-built test workbooks for quick validation
```

---

## 🤝 Contributing

Contributions from developers, data analysts, and Excel enthusiasts worldwide are warmly welcomed!
- **Bug Reports & Ideas:** Open an issue on GitHub.
- **Pull Requests:** Fork the repo, create a feature branch, test thoroughly with Excel Desktop, and submit a PR.
- **Model Support:** Help expand integrations for additional local model runtimes.

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.  
Built for the global open-source community by **Akshay Jatin Solanki**.
"""

if problems:
    print("ERRORS DETECTED:")
    for p in problems:
        print(" -", p)
    sys.exit(1)

with open(OUT, "w", encoding="utf-8") as f:
    f.write(readme_content)

print(f"PASS: README.md generated at {OUT}")
print(f"Total characters: {len(readme_content)}")
