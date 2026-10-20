# CLAUDE.md - guide for Claude and other agents in efffcorp-gst

<!-- EC-STANDARD:START (managed by EFFICIENTCORPORATES/ec-asset-index tools/gen_docs.py - edit there, not here) -->
## EC operating guide (read this before doing anything in this repo)

_Generated 2026-10-20 from `EFFICIENTCORPORATES/ec-asset-index`. If anything here is wrong, fix `products.json` / the generator there, then re-run `tools/gen_docs.py`._

### 1. What this repo is
- **Repo:** `EFFICIENTCORPORATES/efffcorp-gst` (default branch `master`), private GitHub.
- **Product:** `ecpl-products` - ECPL Products - everything ECPL builds and markets (some open source, some licensed): tools, DMS/DAM/Monitor/Security, MDUnix, CA website builder, open-source-derived products
- **Entity:** EFFICIENT CORPORATES PRIVATE LIMITED - Billing entity for ECPL apps/services, ALL courses, corporate training and Mera Kavach; first licensee of the white-label MyVibeCode LMS (decision 2026-10-20)
- **Domain list (Google Tasks):** 2 - ECPL - B2B Automation & Training | **Manager designation:** `ecpl-products-product-manager` (principal NOT created yet - ask Pranav to create it before you need credentials) | **Owner:** pranav (interim)
- _Placement confidence: **medium** - the product/entity assignment is provisional; confirm with Pranav before relying on it._
- **Top-level layout:** `LICENSE`, `MANIFEST.in`, `README.md`, `gst/`, `setup.py`

### 2. Other repos you may need (clone what the task touches)
- Registry and tools (always useful): `gh repo clone EFFICIENTCORPORATES/ec-asset-index` -> `ASSET_INDEX.json` (repo -> R2), `products.json` (all products/entities), `ENTITIES.md`, `DOMAINS.md`, `GOVERNANCE.md`, `ASSET_WORKFLOW.md`, `tools/ec_assets.py`.
- Tech shared services: `EFFICIENTCORPORATES/credential-agent` (tokens/secrets), `EFFICIENTCORPORATES/task-hub` (GitHub issues/delivery), `EFFICIENTCORPORATES/personal-assistant` (Google Tasks/Calendar, brand-entity map).
- Other repos of this product (`ecpl-products`):
  - `EFFICIENTCORPORATES/ECToolsEXE`
  - `EFFICIENTCORPORATES/EfficientToolKitV2`
  - `EFFICIENTCORPORATES/EffCorpGSTRReco`
  - `EFFICIENTCORPORATES/EfficientOCR`
  - `EFFICIENTCORPORATES/EfficientTranscript`
  - `EFFICIENTCORPORATES/EfficientDMS`
  - `EFFICIENTCORPORATES/EfficientDAM`
  - `EFFICIENTCORPORATES/EfficientMonitor`
  - `EFFICIENTCORPORATES/EfficientSecurity`
  - `EFFICIENTCORPORATES/MDUnix`
  - `EFFICIENTCORPORATES/EfficientCAWebsite`
  - `EFFICIENTCORPORATES/ectech-b2c`
  - `microsoft/markitdown`
  - `benjamin-luescher/SMSForwarderApp`
  - `PHPMailer/PHPMailer`
  - `EFFICIENTCORPORATES/efficienttoolkit-releases`
  - `EFFICIENTCORPORATES/EfficientFileLocator`
  - `EFFICIENTCORPORATES/gst_api_app`
  - `EFFICIENTCORPORATES/gst_app_automate`
  - `EFFICIENTCORPORATES/gst_scrutiny`
  - `EFFICIENTCORPORATES/bulk_mailer`
  - `EFFICIENTCORPORATES/EfficientToolkit`
  - `EFFICIENTCORPORATES/effcorp-filetools`
  - `EFFICIENTCORPORATES/effcorp_gstreco`
  - `EFFICIENTCORPORATES/effcorppdf`
  - `EFFICIENTCORPORATES/office_assistant`
  - `EFFICIENTCORPORATES/EfficientDRM`
  - `EFFICIENTCORPORATES/EfficientTracker`
  - `EFFICIENTCORPORATES/SMARTCAOffice`
  - `EFFICIENTCORPORATES/AARAM`
- Clone everything of a product in one go: see `ec-asset-index/ASSET_WORKFLOW.md` (`ec_assets.py pull <repo>` clones a repo and downloads its R2 files).

### 3. Large and binary files (Cloudflare R2) - do not say a file is "missing" before checking
- No R2 assets are recorded for this repo as of 2026-10-20. Bucket naming for any future files: `ec-efffcorp-gst-assets` (created on first push).
- Download everything for this repo, sha256-verified (from a clone of `ec-asset-index`):
  ```bash
  pip install boto3
  python ec-asset-index/tools/ec_assets.py --agent pull efffcorp-gst --dest <this-repo-folder-or-its-parent>
  ```
  (`--agent` needs a principal allowed to issue Cloudflare tokens; otherwise set `R2_ACCESS_KEY_ID`/`R2_SECRET_ACCESS_KEY` issued for you by Pranav. Never paste keys into chat or commit them.)
- New or changed file >= 5 MB: never commit it. `ec_assets.py scan <repo-path>` then `ec_assets.py --agent push <repo-path> --all`, commit the updated `R2_MANIFEST.md`. Install the guard once per clone: copy `ec-asset-index/tools/pre-push-large-files.sh` to `.git/hooks/pre-push`.
- Local-only small files (ignored/untracked): `ec_assets.py --agent backup-local <repo-path>` / `restore-local`. Passphrase is in Bitwarden (`EC_BACKUP__PROD__ARCHIVE_PASSPHRASE`).

### 4. Secrets - how you get them (never from chat, never into git)
- Secrets live in **Bitwarden Secrets Manager**, named `<REPO>__<ENV>__<VARIABLE>` (this repo's prefix: `EFFFCORP_GST__`, ENV = `PROD`/`DEV`/`LOCAL`/`SHARED`). Only secret **names** are in `ec-asset-index/secrets_map.csv`; values never are.
- You receive a value only through the **ECTPL credential agent** (`EFFICIENTCORPORATES/credential-agent`, `python ectpl_creds.py ...`), using **your own principal** (the designation of your role, e.g. the product manager above once it exists; today only `cloudflare-manager`, `github-manager` and `1lavya-product-manager` exist besides Pranav's admin), never another agent's or Pranav's identity unless Pranav explicitly authorises it for that task. Pipe the value straight into the command that needs it, e.g. `X=$(python ectpl_creds.py --as <you> get static:<NAME>) command`. Do not print, log, echo or commit it.
- If your principal cannot read a secret you need, **stop and ask Pranav** to grant it; do not look for it in `.env` files, chat history or other repos.
- Create `.env` locally from the secret names (see `.env.example` if present); `.env` is always gitignored.

### 5. How agents work here (agentic workflow)
- **Identity and scope:** act only as your designated role (see `ec-asset-index/GOVERNANCE.md`): Platform Manager (one platform), Product Manager (this product), or Process Owner (an approved n8n process). Stay inside this repo and the repos/platforms your task names. Humans and agents have separate identities.
- **Start of session:** read this file, `README.md`, the repo's own docs below, and the latest notes in `docs/sessions/` (if present). **End of session:** write `docs/sessions/YYYY-MM-DD-<topic>.md` (goal, decisions, commands run, files changed, open items, next step) and add anything lasting to `CONTEXT.md` / `docs/adr/` / `docs/runbooks/`. No secrets in notes.
- **Git:** work on a branch, small commits, open a PR; never force-push or rewrite history, never delete branches/repos, never commit files >= 5 MB or secrets. Destructive or irreversible actions (delete, force-push, drop data, change permissions) need Pranav's explicit approval in the session.
- **Ask, do not guess:** if ownership, a placement or a requirement is unclear, ask Pranav one concise question. Report outcomes faithfully (what ran, what failed).
- **Cost discipline:** use the cheapest model tier that does the job (open-source/small models for checks and routine edits; top models only for planning/hard reasoning), keep context small, cap loops/retries, and stop at your budget.
- **Commands:** run only commands needed for the task; project code, build scripts and git hooks are untrusted until reviewed. Do not install software or change system settings without approval.
<!-- EC-STANDARD:END -->
