<p align="center">
  <img src="https://sncfoundation.github.io/logos/sheethub.svg" width="96" alt="SheetHub logo">
</p>

<h1 align="center">SheetHub</h1>

<p align="center"><b>An all-in-one DevOps forge on a Google Sheet</b><br>
A <a href="https://sncfoundation.github.io">Sheet-Native Computing Foundation</a> project &#183; analog of <b>GitLab / GitHub</b></p>

---

SheetHub is an all-in-one DevOps forge whose data plane is a Google Sheet: source metadata, issues, merge requests, releases, users, and workload manifests, all in one workbook.

## The Meta-Win: Self-Hosting SNCF Stack

SheetHub hosts the workload manifests that [Sheeternetes](https://github.com/sncfoundation/sheeternetes) deploys via [Sheetlux CD](https://github.com/sncfoundation/sheetlux-cd):

```
  ┌───────────────────────┐
  │       SheetHub        │  ──►  Developer Forge (Repos, Issues, MRs, Manifests)
  │     (Google Sheet)    │
  └──────────┬────────────┘
             │  ./sheetlux-sync.sh
             ▼
  ┌───────────────────────┐
  │      Sheetlux CD      │  ──►  GitOps Continuous Deployment reconciler
  │     (./sheetlux)      │
  └──────────┬────────────┘
             │
             ▼
  ┌───────────────────────┐
  │     Sheeternetes      │  ──►  Spreadsheet Orchestration Control Plane
  │   (Apps Script + K8s) │
  └──────────┬────────────┘
             │
             ▼
  ┌───────────────────────┐
  │     Docker Pods       │  ──►  Running Services + Sheetlium DNS
  └───────────────────────┘
```

**Forge, CI, CD, and cluster are all a Sheet.**

---

## Data Plane Schema (Tabs as Tables)

| Tab | Columns / Schema | Description |
| --- | --- | --- |
| **`Repos`** | `name, description, default_branch, visibility, stars_count, created_at, updated_at` | Repository metadata and stars |
| **`Issues`** | `id, repo, number, title, body, author, state, labels, created_at, updated_at` | Sequential issue tracker per repo |
| **`MergeRequests`** | `id, repo, number, title, description, author, source_branch, target_branch, state, diff_manifest, created_at, updated_at` | Merge requests with visual spec diffs |
| **`Releases`** | `id, repo, tag_name, name, body, author, created_at, assets` | Release tags & artifact manifest registry |
| **`Users`** | `username, name, avatar_url, role, bio, created_at` | Contributor identity & profiles |
| **`Comments`** | `id, target_type, target_id, author, body, created_at` | Discussions on issues and MRs |
| **`Stars`** | `repo, username, starred_at` | Star registry |
| **`Files`** | `repo, path, branch, content, updated_at` | Manifest files and workload definitions |

---

## Quickstart

### 1. Control Plane Setup (Google Apps Script)

1. Create a blank Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Copy [`Code.gs`](Code.gs) into `Code.gs`.
4. Copy [`index.html`](index.html) into `index.html` (Files `+` -> HTML).
5. Set `TOKEN` at the top of `Code.gs`.
6. Run `setup()` once in the script editor. This auto-creates all 8 tabs, applies bold frozen headers, and seeds sample repos (`sncf/hello-web`, `sncf/sheeternetes-manifests`).
7. Click **Deploy → New deployment → Web app** (Execute as: *Me*, Who has access: *Anyone*).
8. Copy the `/exec` URL.

### 2. Local Testing (No Google Account Required)

You can run the control plane and test harness completely offline:

```bash
# Start local in-memory apiserver + Web UI
node hack/local-apiserver.js

# Run the test suite (auth enforcement, multi-repo filtering, MR lifecycle)
node hack/test.js
```
Open **`http://localhost:8788`** in your browser to view the GitLab-style forge Web UI.

---

## CLI Reference (`sheethub`)

Configure endpoint:
```bash
cat > .sheethub.env <<EOF
SHEETHUB_URL=http://localhost:8788/exec
TOKEN=CHANGE_ME_super_secret
EOF
chmod +x sheethub
```

Commands:

```bash
# List repositories
./sheethub repo list

# Query issues for a repo
./sheethub issue list sncf/hello-web

# Create an issue
./sheethub issue create sncf/hello-web "Scale to 10 replicas" "Traffic spike expected"

# Query Merge Requests
./sheethub mr list sncf/hello-web

# Inspect an MR and its manifest spec diff
./sheethub mr view 1 sncf/hello-web

# Merge an MR (updates Files tab)
./sheethub mr merge 1 sncf/hello-web

# Fetch raw workload manifest from the Files tab
./sheethub file get sncf/hello-web app.json

# Guided CLI Tour
./sheethub tour
```

---

## GitOps Integration with Sheetlux CD

Sync manifests directly from SheetHub into the Sheeternetes deployment directory:

```bash
./sheetlux-sync.sh sncf/hello-web ./gitops app.json
```

---

## License

[Apache License 2.0](LICENSE).
