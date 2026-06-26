# Contract Hierarchy Analyzer — GitHub + VDI Setup Guide

This guide covers three things:

- **Part A** — push this project to GitHub safely (from your current machine)
- **Part B** — download it inside the VDI
- **Part C** — make the OpenAI API work in the VDI environment

> ### ⚠️ Do this FIRST — rotate the API key
> The OpenAI key was previously hard-coded in `contract_hierarchy_analyzer.py`.
> Treat it as **compromised**:
> 1. Go to <https://platform.openai.com/api-keys>
> 2. **Revoke** the old key.
> 3. **Create a new key** and paste it into your local `.env` file (the gitignored one).
>
> The code no longer contains any key — it now reads `OPENAI_API_KEY` from a
> `.env` file or the environment, so the secret never goes to GitHub.

---

## What gets pushed vs. what stays private

| Pushed to GitHub | Stays private (gitignored) |
| --- | --- |
| `contract_hierarchy_analyzer.py` | `.env` (your real API key) |
| `requirements.txt` | `BLUE CROSS TEXAS FEDERAL CREDIT UNION/` (and any contract folder) |
| `.env.example` (template, no secret) | `output/`, `extraction_cache.json` |
| `.gitignore`, `VDI_SETUP.md` | `__pycache__/`, virtual envs |
| `PC-PH-ProductNames_Dictionary_v2.xlsx` | |

The `.gitignore` is **deny-by-default**: it ignores *everything*, then allows back
only the files in the left column. Contract PDFs and your key therefore cannot be
committed by accident. (If the product dictionary `.xlsx` is considered internal
and must not go to GitHub either, delete its `!`-line from `.gitignore`.)

---

## Part A — Push to GitHub (run on your current machine)

All commands run from the project folder in **PowerShell**:

```powershell
cd "C:\Users\Vedaansh\Desktop\hierarchy script"
```

### A1. Initialize git and stage files
```powershell
git init
git add -A
```

### A2. ✅ VERIFY before committing (most important step)
```powershell
git status
```
Confirm the list of files to be committed does **NOT** include:
- the `BLUE CROSS TEXAS FEDERAL CREDIT UNION` folder or any `.pdf`/`.PDF`
- `.env`

You should see only: `.env.example`, `.gitignore`, `VDI_SETUP.md`,
`requirements.txt`, `contract_hierarchy_analyzer.py`, and the `.xlsx`.

If a contract file or `.env` shows up, **stop** and re-check `.gitignore`.

### A3. First commit
```powershell
git config user.email "vrana@keplercannon.com"
git config user.name  "Vedaansh"
git commit -m "Initial commit: contract hierarchy analyzer (no secrets, contracts excluded)"
```

### A4. Create a PRIVATE GitHub repo and push
Create the repo at <https://github.com/new> — **set visibility to Private**
(this is legal/contract tooling; do not make it public). Name it e.g.
`contract-hierarchy-analyzer`. Then:

```powershell
git branch -M main
git remote add origin https://github.com/<your-username>/contract-hierarchy-analyzer.git
git push -u origin main
```

If prompted to authenticate, use a **Personal Access Token (PAT)** as the
password (see A5).

### A5. Create a Personal Access Token (PAT) — needed for private repos
1. <https://github.com/settings/tokens> → **Generate new token (classic)**.
2. Scope: check **`repo`**. Set an expiration.
3. Copy the token (starts with `ghp_…`). Use it as the password when git asks.
   Keep it somewhere safe — you'll reuse it in the VDI.

---

## Part B — Download in the VDI

Pick **one** of the following depending on what the VDI allows.

### Option 1 — `git clone` (preferred, if git + GitHub are reachable)
```powershell
cd <a-working-folder-in-the-vdi>
git clone https://github.com/<your-username>/contract-hierarchy-analyzer.git
cd contract-hierarchy-analyzer
```
When asked for credentials: username = your GitHub username, password = the **PAT** from A5.

To pull later updates: `git pull`.

### Option 2 — Download ZIP (if git isn't available in the VDI)
On GitHub: repo → green **Code** button → **Download ZIP** → transfer into the
VDI per your org's approved method → unzip.

> Either way, the contract folders are **not** in the repo. Copy the client
> contract folders (e.g. `BLUE CROSS TEXAS FEDERAL CREDIT UNION`) into the VDI
> project folder separately, so the layout becomes:
> ```
> contract-hierarchy-analyzer\
> ├── contract_hierarchy_analyzer.py
> ├── BLUE CROSS TEXAS FEDERAL CREDIT UNION\   <- one subfolder per client
> └── ...
> ```

---

## Part C — Make the OpenAI API work in the VDI

### C1. Python + dependencies
Confirm Python 3.9+ is available:
```powershell
python --version
```
Install dependencies (a virtual env avoids needing admin rights):
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```
If `pip` can't reach the internet directly, see **C4 (proxy)** or ask IT for the
internal package index (`pip install -r requirements.txt -i <internal-index-url>`).

### C2. Configure the backend + credentials (`.env`)
Copy the template and edit it:
```powershell
Copy-Item .env.example .env
notepad .env
```
The script auto-loads `.env` on startup — **no code edits needed**. `.env` is
gitignored, so it stays inside the VDI. Pick **one** of the two backends:

**Option A — Fiserv Foundation API (this is the VDI setup):**
```
OPENAI_BACKEND=fiserv
FOUNDATION_API_URL=https://dev-cst-cognitive-service.onefiserv.net/FoundationAPI/openai/deployments/default/chat/completions?api-version=2025-03-01-preview
FISERV_EMAIL=your.name@fiserv.com
FISERV_PURPOSE_GPT5=GPT5.1Purpose
FISERV_PURPOSE_GPT4=GPT4.1Purpose
# only if the gateway also needs a subscription key:
# OPENAI_API_KEY=...gateway-key...
```

**Option B — plain OpenAI (local dev / testing):**
```
OPENAI_BACKEND=openai
OPENAI_API_KEY=sk-...your-rotated-key...
```

### C3. How the Fiserv backend works (and what to confirm with Fiserv)
When `OPENAI_BACKEND=fiserv`, the script POSTs directly to `FOUNDATION_API_URL`
(an Azure-style **chat/completions** endpoint) instead of calling
`api.openai.com`. It sends the `X-Purpose` header chosen automatically from the
model family:
- model starts with `gpt-5` → `FISERV_PURPOSE_GPT5`
- anything else → `FISERV_PURPOSE_GPT4`

Because the gateway decides the actual model, **whatever `OPENAI_MODEL` you set
in the code is just a hint** — the real model name is read back from the API
response and shown in the run-metrics summary (see Part D).

⚠️ **Confirm these details with Fiserv's Foundation API docs** — they decide a
couple of header *names* that I had to guess in `_fiserv_headers()` (top of the
function in `contract_hierarchy_analyzer.py`):
- **Auth:** the script sends the key as the `api-key` header. If the gateway
  wants `Authorization: Bearer <token>` instead, change it there.
- **Email header:** sent as `X-Email`. Change the name if the gateway expects
  e.g. `X-User-Email`.
- **Purpose tag values:** make sure `GPT5.1Purpose` / `GPT4.1Purpose` are the
  exact strings Fiserv issued you.

If any of those differ, tell me the exact header names/values and I'll update
`_fiserv_headers()` for you.

### C4. If the VDI requires an HTTP proxy
The OpenAI SDK and pip respect standard proxy environment variables. Set them
(values from IT) before running:
```powershell
$env:HTTPS_PROXY = "http://proxy.yourcompany.com:8080"
$env:HTTP_PROXY  = "http://proxy.yourcompany.com:8080"
```
If you hit TLS/SSL certificate errors behind a corporate proxy, ask IT for the
internal root CA bundle and point requests at it:
```powershell
$env:SSL_CERT_FILE = "C:\path\to\corp-ca-bundle.pem"
$env:REQUESTS_CA_BUNDLE = "C:\path\to\corp-ca-bundle.pem"
```

### C5. Quick connectivity test
```powershell
python -c "from openai import OpenAI; import os; OpenAI(api_key=os.environ.get('OPENAI_API_KEY','')); print('client OK, key present:', bool(os.environ.get('OPENAI_API_KEY')))"
```
(With a `.env`, run it after `python contract_hierarchy_analyzer.py` does the
loading — or just run the tool; it prints a clear error if the key is missing.)

---

## Part D — Run the tool in the VDI

```powershell
# all client folders found next to the script:
python contract_hierarchy_analyzer.py

# or restrict to one client:
python contract_hierarchy_analyzer.py "BLUE CROSS TEXAS FEDERAL CREDIT UNION"
```
Open the result: `output\<ClientName>\contracts_hierarchy.html` in a browser.

### Run-metrics summary
At the end of every successful run the script prints a metrics block:
```
==============================================================
  RUN METRICS
==============================================================
  Model used (from API):  gpt-5.1-...        <- ACTUAL model the endpoint used
  API calls:              12
  Input tokens:           45,200
  Output tokens:          8,310
  Total tokens:           53,510
  Backend:                fiserv
  Run time:               00:03:42  (222.4s)
==============================================================
```
- **Model used** is read from the API response, not from `OPENAI_MODEL` — so in
  the Fiserv VDI it reflects whatever the Foundation endpoint actually served.
- **Run time** starts at 0 when the script launches and stops when it finishes
  successfully.
- Token counts come from each response's `usage` block. Contracts served from
  cache make **no** API call, so a fully-cached re-run shows `API calls: 0`.

---

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `ERROR: No OpenAI API key found.` | `OPENAI_BACKEND=openai` but key blank — see **C2**. |
| `ERROR: ...FOUNDATION_API_URL is not set.` | `OPENAI_BACKEND=fiserv` but URL missing — see **C2/C3**. |
| `401 / 403` on the Fiserv path | Wrong auth header or purpose tag — see **C3** and adjust `_fiserv_headers()`. |
| `Connection error` / timeouts | Proxy/gateway needed — see **C3 / C4**. |
| SSL `CERTIFICATE_VERIFY_FAILED` | Set `SSL_CERT_FILE` to corp CA bundle — see **C4**. |
| `ModuleNotFoundError` | Activate the venv, then `pip install -r requirements.txt`. |
| Model not found / 404 | The gateway exposes a different model — check `X-Purpose` / `OPENAI_MODEL` (C3). |
| `.docx` / `.msg` errors | Those inputs need `docx2pdf` (MS Word) / `extract-msg`; skip if unused. |
