<p align="center">
  <img src="https://sncfoundation.github.io/logos/sheethub.svg" width="96" alt="SheetHub logo">
</p>

<h1 align="center">SheetHub</h1>

<p align="center"><b>An all-in-one DevOps forge on a Google Sheet</b><br>
A <a href="https://sncfoundation.github.io">Sheet-Native Computing Foundation</a> project &#183; analog of <b>GitLab / GitHub</b></p>

<p align="center">
  <a href="https://github.com/sncfoundation/sheethub/actions/workflows/ci.yml"><img src="https://github.com/sncfoundation/sheethub/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
</p>

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
| **`Images`** | `name, digest, config, layers, created, size, repo, tag, author, pushed_at` | **SICF v0.1 Image Manifests**: OCI config digest + ordered layer digests + forge metadata |
| **`Layers`** | `digest, ordinal, media_type, data, size_bytes, id, created_at` | **SICF v0.1 Image Layers**: Content-addressed base64 layer chunks ($\le$ 32,767 chars/cell) |
| **`Users`** | `username, name, avatar_url, role, bio, created_at` | Contributor identity & profiles |
| **`Comments`** | `id, target_type, target_id, author, body, created_at` | Discussions on issues and MRs |
| **`Stars`** | `repo, username, starred_at` | Star registry |
| **`Files`** | `repo, path, branch, content, updated_at` | Manifest files and workload definitions |

---

## SICF v0.1 Container Registry in Cells (Air-Gap Edition)

SheetHub implements the **SICF (Sheet-Native Image Container Format v0.1)** specification:
- **OCI-Style Image Model**: Container images are modeled as an image config JSON (`digest`, `config` in base64) plus an ordered list of layer SHA-256 digests (`layers`).
- **Content-Addressed Layer Deduplication**: Layers in the `Layers` tab are content-addressed by their SHA-256 digest. Multiple images sharing common base layers (e.g. Alpine base) reuse the exact same layer rows in the spreadsheet without duplication.
- **Cell Character Limit Sharding**: Excel and Google Sheets cap a single cell at **32,767 characters**. Image layers are converted to Base64 and sharded across cell rows with `ordinal` ordering and media types (`application/vnd.oci.image.layer.v1.tar`).
- **Round-Trip with `sheetbuild`**: Images pushed to SheetHub can be directly imported and exported by `sheetbuild.py`, yielding digest-identical OCI container tarballs.
- **Sheeternetes Kubelet Resolution (`sicf:<name>`)**: Node agents in Sheeternetes clusters pull images directly from the spreadsheet control plane using `sicf:<image-name>` or `sicf:<repo>/<name>`.
- **Air-Gap Self-Hosting**: Both workload manifests *and* container images reside completely within the Google Sheet.

---

## Quickstart

### 1. Control Plane Setup (Google Apps Script)

1. Create a blank Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Copy [`Code.gs`](Code.gs) into `Code.gs`.
4. Copy [`index.html`](index.html) into `index.html` (Files `+` -> HTML).
5. Set `TOKEN` at the top of `Code.gs`.
6. Run `setup()` once in the script editor. This auto-creates all 10 tabs, applies bold frozen headers, and seeds sample repos (`sncf/hello-web`, `sncf/sheeternetes-manifests`) along with container images.
7. Click **Deploy → New deployment → Web app** (Execute as: *Me*, Who has access: *Anyone*).
8. Copy the `/exec` URL.

### 2. Local Testing (No Google Account Required)

You can run the control plane and test harness completely offline:

```bash
# Start local in-memory apiserver + Web UI
node hack/local-apiserver.js

# Run the test suite (auth enforcement, multi-repo filtering, MR lifecycle, container registry sharding)
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

# List container images stored in cells
./sheethub registry list sncf/hello-web

# Push container layer / tarball (auto-shards base64 into spreadsheet cells)
./sheethub registry push sncf/hello-web my-service:v1.0.0 ./layer.tar

# Pull and reassemble container image layer from spreadsheet cells (verifies SHA-256 digest)
./sheethub registry pull sncf/hello-web my-service:v1.0.0 ./pulled_layer.tar

# Inspect cell shard map coordinates
./sheethub registry view sncf/hello-web my-service:v1.0.0

# Delete container image and associated cell chunks
./sheethub registry delete sncf/hello-web my-service:v1.0.0

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
