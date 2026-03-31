# Building the .exe

This document explains how to build a standalone `qimpexp.exe` from the
source code, why Windows shows a SmartScreen warning, and what that warning
actually means.

---

## Why does Windows show a SmartScreen warning?

When you download and run `qimpexp.exe`, Windows may show this:

> **"Windows protected your PC"**
> Microsoft Defender SmartScreen prevented an unrecognized app from starting.

**This does not mean the file is malicious.** It means the executable is not
signed with an EV (Extended Validation) code signing certificate. EV
certificates cost approximately $300–500 USD per year and are issued by
certificate authorities like DigiCert or Sectigo after verifying the
publisher's identity. Without one, SmartScreen flags any new executable
until it accumulates enough download history to build a reputation.

To bypass the warning:
1. Click **"More info"**
2. Click **"Run anyway"**

If you prefer not to trust the pre-built `.exe`, you can build it yourself
from the source code in under five minutes — instructions below. You will
get the exact same result, built on your own machine, which Windows trusts
automatically.

---

## Building the .exe yourself

### Prerequisites

- Python 3.9 or later installed and on your PATH
- The QImpExp source code (clone or download the ZIP from GitHub/GitLab)

### Step 1 — Install PyInstaller and dependencies

```
pip install pyinstaller openpyxl
```

`openpyxl` is the only third-party runtime dependency. `pyinstaller` is the
build tool only — it is not bundled into the exe.

### Step 2 — Build the executable

From the `qimpexp` folder:

```
pyinstaller --onefile --console --name qimpexp --hidden-import openpyxl --version-file version.txt qimpexp.py
```

This produces `dist\qimpexp.exe`. The build takes 30–60 seconds.

### Step 3 — Run it

```
dist\qimpexp.exe
```

Or copy `qimpexp.exe` anywhere — it is fully self-contained.

---

## What PyInstaller bundles

PyInstaller packages the Python interpreter, all imported libraries
(`openpyxl` and its dependencies), and the source scripts into a single
executable. Nothing is sent anywhere — the tool reads files from your local
disk and writes results to your local disk. There is no network access,
no telemetry, and no data collection of any kind.

You can verify this by inspecting the source code on GitHub or GitLab
before building.

---

## What is bundled in the release ZIP

The release ZIP (`qimpexp_vX.Y.Z.zip`) contains:

| File | Description |
|------|-------------|
| `qimpexp.exe` | Standalone Windows executable |
| `README.FIRST.txt` | Quick-start guide |
| `sample_data/account_periods.json` | Example account period config |
| `sample_data/security_map.json` | Example security name mapping |
| `sample_data/*.qif` | Example QIF output files |

---

## Notes for the CI/CD release build (maintainer only)

### GitHub Actions

The pre-built `.exe` in GitHub Releases is produced by `.github/workflows/build_exe.yml`
on a clean `windows-latest` runner. The build command is identical to Step 2 above.
The resulting binary is attached to the release automatically.

To trigger a new GitHub release:

```
git tag -a v1.0.0 -m "v1.0.0"
git push origin v1.0.0
```

### GitLab CI

The GitLab pipeline (`.gitlab-ci.yml`) uses a `saas-windows-medium-amd64` runner
and produces the same artifact. To trigger:

```
git tag -a v1.0.0 -m "v1.0.0"
git push gitlab v1.0.0
```

---

## Verifying the pre-built .exe (optional)

The SHA-256 hash of each release binary is listed in the release notes
on GitHub and GitLab. To verify on Windows:

```
certutil -hashfile qimpexp.exe SHA256
```

Compare the output to the hash in the release notes. If they match, the
file is byte-for-byte identical to what was built by the CI workflow.
