/**
 * SheetHub — a GitLab-style DevOps forge on a Google Sheet.
 *
 * Part of the Sheet-Native Computing Foundation (SNCF) ecosystem.
 * Data plane: Google Sheet (tabs for Repos, Issues, MergeRequests, Releases, Users, Comments, Stars, Files, Images, Layers).
 * Control plane / apiserver: Apps Script Web App (doGet / doPost).
 * Web UI: GitLab-style forge interface served via HtmlService.
 *
 * Deploy:
 *   1) Extensions -> Apps Script, paste this file and index.html.
 *   2) Run setup() once (creates tabs with headers and seed data across 2 repos).
 *   3) Deploy -> New deployment -> Web app -> Execute as: Me, Access: Anyone.
 *   4) Copy the Web App URL into sheethub (.sheethub.env) or open directly in browser.
 */

const TOKEN = 'CHANGE_ME_super_secret';

const TABS = {
  Repos:          ['name', 'description', 'default_branch', 'visibility', 'stars_count', 'created_at', 'updated_at'],
  Issues:         ['id', 'repo', 'number', 'title', 'body', 'author', 'state', 'labels', 'created_at', 'updated_at'],
  MergeRequests:  ['id', 'repo', 'number', 'title', 'description', 'author', 'source_branch', 'target_branch', 'state', 'diff_manifest', 'created_at', 'updated_at'],
  Releases:       ['id', 'repo', 'tag_name', 'name', 'body', 'author', 'created_at', 'assets'],
  Users:          ['username', 'name', 'avatar_url', 'role', 'bio', 'created_at'],
  Comments:       ['id', 'target_type', 'target_id', 'author', 'body', 'created_at'],
  Stars:          ['repo', 'username', 'starred_at'],
  Files:          ['repo', 'path', 'branch', 'content', 'updated_at'],
  // SICF v0.1 Image Store schema
  Images:         ['name', 'digest', 'config', 'layers', 'created', 'size', 'repo', 'tag', 'author', 'pushed_at', 'id', 'updated_at'],
  Layers:         ['digest', 'ordinal', 'media_type', 'data', 'size_bytes', 'id', 'created_at'],
  // Compatibility aliases
  Registry:       ['name', 'digest', 'config', 'layers', 'created', 'size', 'repo', 'tag', 'author', 'pushed_at', 'id', 'updated_at'],
  RegistryChunks: ['digest', 'ordinal', 'media_type', 'data', 'size_bytes', 'id', 'created_at'],
};

// ---------- generic sheet <-> objects ----------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function normalizeTabName(name) {
  if (name === 'Registry' || name === 'images' || name === 'Images') return 'Images';
  if (name === 'RegistryChunks' || name === 'layers' || name === 'Layers') return 'Layers';
  return name;
}

function sanitizeCell(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'number' || typeof val === 'boolean') return val;
  const str = String(val);
  // Neutralize spreadsheet formula/CSV injection: a leading formula trigger, optionally
  // preceded by whitespace/TAB/CR (which spreadsheets and CSV importers strip before parsing).
  if (/^[\s\t\r\n]*[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

function unescapeCell(val) {
  if (typeof val === 'string' && val.length > 1 && val[0] === "'" && (val[1] === '=' || val[1] === '+' || val[1] === '-' || val[1] === '@')) {
    return val.substring(1);
  }
  return val;
}

function readTab(name) {
  const s = ss();
  const norm = normalizeTabName(name);
  let sh = s.getSheetByName(norm);
  let matchedName = norm;
  if (!sh && norm !== name) {
    sh = s.getSheetByName(name);
    matchedName = name;
  }
  // Fallback for legacy spreadsheets that still only have Registry / RegistryChunks tabs
  if (!sh) {
    if (norm === 'Images' && s.getSheetByName('Registry')) {
      sh = s.getSheetByName('Registry');
      matchedName = 'Registry';
    } else if (norm === 'Layers' && s.getSheetByName('RegistryChunks')) {
      sh = s.getSheetByName('RegistryChunks');
      matchedName = 'RegistryChunks';
    }
  }
  if (!sh) return [];
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const header = rng.shift();
  const sheetName = (sh.getName && typeof sh.getName === 'function') ? sh.getName() : matchedName;
  const isLegacyRegistry = sheetName === 'Registry';
  const isLegacyChunks = sheetName === 'RegistryChunks';

  return rng
    .filter(r => String(r[0]).trim() !== '')
    .map(row => {
      const o = {};
      header.forEach((h, i) => { o[h] = unescapeCell(row[i]); });
      // Normalize legacy Registry rows to SICF Images schema if reading from legacy tab
      if (isLegacyRegistry) {
        if (!o.layers && o.digest) o.layers = o.digest;
        if (o.size === undefined && o.size_bytes !== undefined) o.size = Number(o.size_bytes) || 0;
        if (!o.created && o.created_at) o.created = o.created_at;
        if (!o.pushed_at && (o.updated_at || o.created_at)) o.pushed_at = o.updated_at || o.created_at;
        if (o.name && o.tag && !o.name.includes(':')) o.name = o.name + ':' + o.tag;
      } else if (isLegacyChunks) {
        if (o.ordinal === undefined && o.chunk_index !== undefined) o.ordinal = Number(o.chunk_index) || 0;
        if (o.data === undefined && o.chunk_data !== undefined) o.data = o.chunk_data;
        if (!o.media_type) o.media_type = 'application/vnd.oci.image.layer.v1.tar';
      }
      return o;
    });
}

function writeTab(name, objects) {
  const s = ss();
  const actualName = normalizeTabName(name);
  let sh = s.getSheetByName(actualName);
  if (!sh && actualName !== name) {
    sh = s.getSheetByName(name);
  }
  if (!sh) {
    sh = s.insertSheet(actualName);
    const headerCols = TABS[actualName] || TABS[name];
    sh.getRange(1, 1, 1, headerCols.length).setValues([headerCols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const header = TABS[actualName] || TABS[name];
  const rows = objects.map(o => header.map(h => sanitizeCell(o[h])));
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, header.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function shortId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

function computeSha256(base64Str) {
  if (!base64Str || typeof base64Str !== 'string') {
    throw new Error('invalid or empty base64 payload');
  }
  const trimmed = base64Str.trim();
  if (trimmed.length % 4 !== 0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(trimmed)) {
    throw new Error('invalid base64 character or padding');
  }
  const bytes = Utilities.base64Decode(trimmed);
  if (!bytes || !bytes.length) {
    throw new Error('failed to decode base64 bytes');
  }
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  let hex = '';
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    const byteHex = byteVal.toString(16);
    hex += byteHex.length === 1 ? '0' + byteHex : byteHex;
  }
  return 'sha256:' + hex;
}

// ---------- one-time setup & multi-repo seed data ----------
function setupTab(name) {
  const s = ss();
  const actualName = normalizeTabName(name);
  let sh = s.getSheetByName(actualName);
  if (!sh) sh = s.insertSheet(actualName);
  const cols = TABS[actualName] || TABS[name];
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function setup() {
  const s = ss();
  // Auto-migrate legacy Registry / RegistryChunks tabs to Images / Layers if present
  const legacyRegistry = s.getSheetByName('Registry');
  const legacyChunks = s.getSheetByName('RegistryChunks');
  const imagesSheet = s.getSheetByName('Images');
  const layersSheet = s.getSheetByName('Layers');

  if (legacyRegistry && (!imagesSheet || imagesSheet.getLastRow() < 2)) {
    const oldRegistryRows = readTab('Registry');
    if (oldRegistryRows.length > 0) {
      writeTab('Images', oldRegistryRows);
    }
  }
  if (legacyChunks && (!layersSheet || layersSheet.getLastRow() < 2)) {
    const oldChunkRows = readTab('RegistryChunks');
    if (oldChunkRows.length > 0) {
      writeTab('Layers', oldChunkRows);
    }
  }

  ['Repos', 'Issues', 'MergeRequests', 'Releases', 'Users', 'Comments', 'Stars', 'Files', 'Images', 'Layers'].forEach(name => {
    setupTab(name);
  });

  const now = new Date().toISOString();

  // Seed Users
  const users = readTab('Users');
  if (!users.length) {
    writeTab('Users', [
      { username: 'prateeekbuilds', name: 'Prateek', avatar_url: 'https://github.com/prateeekbuilds.png', role: 'Maintainer', bio: 'Spreadsheet engineer & forge builder', created_at: now },
      { username: 'tym83', name: 'Timur', avatar_url: 'https://github.com/tym83.png', role: 'SNCF Chair', bio: 'Reconciling containers since cell A1', created_at: now },
      { username: 'sheetbot', name: 'SheetBot', avatar_url: 'https://sncfoundation.github.io/logos/sheeternetes.svg', role: 'CI Bot', bio: 'Apps Script automated runner', created_at: now },
    ]);
  }

  // Seed Repos (2 distinct repos)
  const repos = readTab('Repos');
  if (!repos.length) {
    writeTab('Repos', [
      {
        name: 'sncf/hello-web',
        description: 'Reference 3-tier microservice manifest for Sheeternetes',
        default_branch: 'main',
        visibility: 'public',
        stars_count: 12,
        created_at: now,
        updated_at: now
      },
      {
        name: 'sncf/sheeternetes-manifests',
        description: 'Production GitOps cluster manifests deployed via Sheetlux CD',
        default_branch: 'main',
        visibility: 'public',
        stars_count: 42,
        created_at: now,
        updated_at: now
      }
    ]);
  }

  // Seed Files / Manifests across both repos
  const files = readTab('Files');
  if (!files.length) {
    const helloSpec = JSON.stringify({
      deployments: [
        { name: 'hello-web', image: 'traefik/whoami', replicas: 3, cpu_req: 100, mem_req: 64, command: '' },
        { name: 'redis-cache', image: 'redis:alpine', replicas: 1, cpu_req: 150, mem_req: 128, command: '' }
      ]
    }, null, 2);

    const clusterSpec = JSON.stringify({
      deployments: [
        { name: 'ingress-router', image: 'traefik:v2.10', replicas: 2, cpu_req: 200, mem_req: 256, command: '' },
        { name: 'metrics-collector', image: 'prom/prometheus:latest', replicas: 1, cpu_req: 300, mem_req: 512, command: '' }
      ]
    }, null, 2);

    writeTab('Files', [
      { repo: 'sncf/hello-web', path: 'app.json', branch: 'main', content: helloSpec, updated_at: now },
      { repo: 'sncf/hello-web', path: 'README.md', branch: 'main', content: '# Hello Web\n\nA resilient web service running on Sheeternetes.', updated_at: now },
      { repo: 'sncf/sheeternetes-manifests', path: 'production.json', branch: 'main', content: clusterSpec, updated_at: now },
      { repo: 'sncf/sheeternetes-manifests', path: 'README.md', branch: 'main', content: '# Production Manifests\n\nCore infrastructure specs.', updated_at: now },
    ]);
  }

  // Seed Issues across both repos
  const issues = readTab('Issues');
  if (!issues.length) {
    writeTab('Issues', [
      {
        id: 'iss-001',
        repo: 'sncf/hello-web',
        number: 1,
        title: 'Scale hello-web to 5 replicas for traffic surge',
        body: 'Upcoming holiday campaign will increase load. We should scale hello-web from 3 to 5 replicas in app.json.',
        author: 'tym83',
        state: 'open',
        labels: 'enhancement,scaling',
        created_at: now,
        updated_at: now
      },
      {
        id: 'iss-002',
        repo: 'sncf/hello-web',
        number: 2,
        title: 'Add memory limits to redis-cache container',
        body: 'Currently redis-cache has mem_req 128MiB. Let us ensure it has proper resource reservations.',
        author: 'prateeekbuilds',
        state: 'closed',
        labels: 'performance',
        created_at: now,
        updated_at: now
      },
      {
        id: 'iss-101',
        repo: 'sncf/sheeternetes-manifests',
        number: 1,
        title: 'Provision ingress controller',
        body: 'Set up Traefik routing in production.json.',
        author: 'prateeekbuilds',
        state: 'open',
        labels: 'networking',
        created_at: now,
        updated_at: now
      }
    ]);
  }

  // Seed MergeRequests across both repos
  const mrs = readTab('MergeRequests');
  if (!mrs.length) {
    const diffHello = JSON.stringify({
      path: 'app.json',
      before: {
        deployments: [
          { name: 'hello-web', image: 'traefik/whoami', replicas: 3, cpu_req: 100, mem_req: 64, command: '' },
          { name: 'redis-cache', image: 'redis:alpine', replicas: 1, cpu_req: 150, mem_req: 128, command: '' }
        ]
      },
      after: {
        deployments: [
          { name: 'hello-web', image: 'traefik/whoami', replicas: 5, cpu_req: 100, mem_req: 64, command: '' },
          { name: 'redis-cache', image: 'redis:7-alpine', replicas: 2, cpu_req: 150, mem_req: 256, command: '' }
        ]
      }
    }, null, 2);

    const diffCluster = JSON.stringify({
      path: 'production.json',
      before: {
        deployments: [
          { name: 'ingress-router', image: 'traefik:v2.10', replicas: 2, cpu_req: 200, mem_req: 256, command: '' }
        ]
      },
      after: {
        deployments: [
          { name: 'ingress-router', image: 'traefik:v2.10', replicas: 2, cpu_req: 200, mem_req: 256, command: '' },
          { name: 'metrics-collector', image: 'prom/prometheus:latest', replicas: 1, cpu_req: 300, mem_req: 512, command: '' }
        ]
      }
    }, null, 2);

    writeTab('MergeRequests', [
      {
        id: 'mr-001',
        repo: 'sncf/hello-web',
        number: 1,
        title: 'Scale hello-web to 5 replicas & upgrade redis',
        description: 'Resolves #1. Increases replica count to handle surge traffic and bumps redis version.',
        author: 'prateeekbuilds',
        source_branch: 'scale-up-v2',
        target_branch: 'main',
        state: 'open',
        diff_manifest: diffHello,
        created_at: now,
        updated_at: now
      },
      {
        id: 'mr-101',
        repo: 'sncf/sheeternetes-manifests',
        number: 1,
        title: 'Add prometheus monitoring stack',
        description: 'Deploys Prometheus collector pod for cluster metrics.',
        author: 'tym83',
        source_branch: 'add-prom',
        target_branch: 'main',
        state: 'open',
        diff_manifest: diffCluster,
        created_at: now,
        updated_at: now
      }
    ]);
  }

  // Seed Releases across repos
  const releases = readTab('Releases');
  if (!releases.length) {
    writeTab('Releases', [
      {
        id: 'rel-001',
        repo: 'sncf/hello-web',
        tag_name: 'v1.0.0',
        name: 'v1.0.0: Initial Production Rollout',
        body: 'First stable release deployed to Sheeternetes cluster via Sheetlux CD.',
        author: 'prateeekbuilds',
        created_at: now,
        assets: 'app.json'
      },
      {
        id: 'rel-101',
        repo: 'sncf/sheeternetes-manifests',
        tag_name: 'v2.0.0',
        name: 'v2.0.0: High Availability Infra',
        body: 'Multi-node ingress and metrics monitoring configuration.',
        author: 'tym83',
        created_at: now,
        assets: 'production.json'
      }
    ]);
  }

  // Seed Comments
  const comments = readTab('Comments');
  if (!comments.length) {
    writeTab('Comments', [
      {
        id: 'comm-001',
        target_type: 'mr',
        target_id: 'mr-001',
        author: 'tym83',
        body: 'Diff looks clean! The memory bump on redis is well within node capacity. CI passed.',
        created_at: now
      },
      {
        id: 'comm-002',
        target_type: 'issue',
        target_id: 'iss-001',
        author: 'prateeekbuilds',
        body: 'MR !1 opened addressing this.',
        created_at: now
      }
    ]);
  }

  // Seed Stars
  const stars = readTab('Stars');
  if (!stars.length) {
    writeTab('Stars', [
      { repo: 'sncf/hello-web', username: 'prateeekbuilds', starred_at: now },
      { repo: 'sncf/hello-web', username: 'tym83', starred_at: now },
      { repo: 'sncf/sheeternetes-manifests', username: 'prateeekbuilds', starred_at: now },
    ]);
  }

  // Migrate legacy Registry / RegistryChunks tabs to Images / Layers if present
  const legacyRegSheet = s.getSheetByName('Registry');
  const legacyChunksSheet = s.getSheetByName('RegistryChunks');
  const hasImagesSheet = !!s.getSheetByName('Images');
  const hasLayersSheet = !!s.getSheetByName('Layers');

  if (legacyRegSheet && !hasImagesSheet) {
    const legacyRegRows = readTab('Images');
    if (legacyRegRows.length) {
      writeTab('Images', legacyRegRows);
    }
  }
  if (legacyChunksSheet && !hasLayersSheet) {
    const legacyChunkRows = readTab('Layers');
    if (legacyChunkRows.length) {
      writeTab('Layers', legacyChunkRows);
    }
  }

  // Seed Container Registry (SICF v0.1: Images & Layers tabs)
  const images = readTab('Images');
  if (!images.length) {
    const sampleChunk1 = 'H4sICN8+v2YCA2hlbGxvLXdlYi50YXIA7Z1rc9s2Ese/SpX7Yd25K3s2d2U6dZ3Ex04yN5P6ykmk5MskMRdFkdT8eAAYgCIeEtAiyfG0Y4kEQGDx211gsViA/y8vLy/n51+vrq6evX/26sWzV5eXlz+ev/rpyfPnL1++ePnq6cvnLy9eXj5//vzps6eXl6+ev3z+9OXzV49ffHn16tWzn6/+/u1/Hj368/r/vv3vf3/9n4e3/37996+/';
    const sampleChunk2 = 'mZzO5vP56eX5y/nJ9Ozn57Pz17Pz2cvz19PT6dnz8/lsNpufTk/PTmZn569Onr+cnrz86eTp2dlp+t1/AAAA///817t0wAAA';

    const layer1Digest = computeSha256(sampleChunk1 + sampleChunk2);
    const layer2Digest = computeSha256(sampleChunk1);

    const whoamiConfigObj = {
      architecture: 'amd64',
      os: 'linux',
      config: {
        Env: ['PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin'],
        Cmd: ['/whoami'],
        WorkingDir: '/',
        User: 'app'
      },
      rootfs: {
        type: 'layers',
        diff_ids: [layer1Digest]
      }
    };
    const whoamiConfigJson = JSON.stringify(whoamiConfigObj);
    const whoamiConfigB64 = Utilities.base64Encode(Utilities.newBlob(whoamiConfigJson).getBytes());
    const whoamiConfigDigest = computeSha256(whoamiConfigB64);

    const ingressConfigObj = {
      architecture: 'amd64',
      os: 'linux',
      config: {
        Env: ['PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin'],
        Cmd: ['/traefik'],
        WorkingDir: '/',
        User: 'root'
      },
      rootfs: {
        type: 'layers',
        diff_ids: [layer2Digest]
      }
    };
    const ingressConfigJson = JSON.stringify(ingressConfigObj);
    const ingressConfigB64 = Utilities.base64Encode(Utilities.newBlob(ingressConfigJson).getBytes());
    const ingressConfigDigest = computeSha256(ingressConfigB64);

    writeTab('Images', [
      {
        name: 'traefik/whoami:latest',
        digest: whoamiConfigDigest,
        config: whoamiConfigB64,
        layers: layer1Digest,
        created: now,
        size: 258 + whoamiConfigJson.length,
        repo: 'sncf/hello-web',
        tag: 'latest',
        author: 'prateeekbuilds',
        pushed_at: now,
        id: 'img-001',
        updated_at: now
      },
      {
        name: 'ingress-router:v2.10',
        digest: ingressConfigDigest,
        config: ingressConfigB64,
        layers: layer2Digest,
        created: now,
        size: 186 + ingressConfigJson.length,
        repo: 'sncf/sheeternetes-manifests',
        tag: 'v2.10',
        author: 'tym83',
        pushed_at: now,
        id: 'img-101',
        updated_at: now
      }
    ]);

    writeTab('Layers', [
      { digest: layer1Digest, ordinal: 0, media_type: 'application/vnd.oci.image.layer.v1.tar', data: sampleChunk1, size_bytes: 186, id: 'chk-001', created_at: now },
      { digest: layer1Digest, ordinal: 1, media_type: 'application/vnd.oci.image.layer.v1.tar', data: sampleChunk2, size_bytes: 72, id: 'chk-002', created_at: now },
      { digest: layer2Digest, ordinal: 0, media_type: 'application/vnd.oci.image.layer.v1.tar', data: sampleChunk1, size_bytes: 186, id: 'chk-101', created_at: now },
    ]);
  }
}

// ---------- Web App: apiserver endpoint ----------
// doGet: query forge resources or render web UI
function doGet(e) {
  const p = (e && e.parameter) || {};

  // If no parameters at all (direct browser access to web app URL), serve HTML UI
  if (!p.kind && !p.api && !p.token && Object.keys(p).length === 0) {
    try {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('SheetHub — DevOps Forge on a Spreadsheet')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      return json({ error: 'unauthorized' });
    }
  }

  // Any API / data access MUST provide valid token unconditionally (SCRI §4)
  if (p.token !== TOKEN) {
    return json({ error: 'unauthorized' });
  }

  const kind = (p.kind || 'repos').toLowerCase();
  const repoFilter = p.repo || '';

  const map = {
    repos: 'Repos',
    issues: 'Issues',
    mergerequests: 'MergeRequests',
    mrs: 'MergeRequests',
    releases: 'Releases',
    users: 'Users',
    comments: 'Comments',
    stars: 'Stars',
    files: 'Files',
    images: 'Images',
    registry: 'Images',
    containerregistry: 'Images',
    layers: 'Layers',
    registry_chunks: 'Layers',
    chunks: 'Layers',
    registry_pull: 'Images',
  };

  const tab = map[kind];
  if (!tab) return json({ error: 'unknown kind ' + kind });

  let items = readTab(tab);
  if (repoFilter && ['Issues', 'MergeRequests', 'Releases', 'Files', 'Stars', 'Images'].indexOf(tab) !== -1) {
    items = items.filter(it => String(it.repo).toLowerCase() === repoFilter.toLowerCase());
  }

  if (tab === 'Images') {
    if (p.name) items = items.filter(it => String(it.name).toLowerCase() === p.name.toLowerCase() || String(it.name).toLowerCase().startsWith(p.name.toLowerCase() + ':'));
    if (p.tag) items = items.filter(it => String(it.tag).toLowerCase() === p.tag.toLowerCase() || String(it.name).toLowerCase().endsWith(':' + p.tag.toLowerCase()));
    if (p.digest) items = items.filter(it => String(it.digest) === String(p.digest) || (it.layers && String(it.layers).split(',').includes(p.digest)));
  }

  if (p.id) {
    if (tab === 'Images') {
      items = items.filter(it => String(it.id) === String(p.id) || String(it.tag) === String(p.id) || String(it.digest) === String(p.id) || String(it.name) === String(p.id));
    } else {
      items = items.filter(it => String(it.id) === String(p.id) || String(it.number) === String(p.id));
    }
  }

  // Handle pull query (fetch image + ordered layer chunks per SICF)
  if (kind === 'registry_pull' || (tab === 'Images' && (p.pull === 'true' || p.action === 'pull'))) {
    const rawTarget = p.name || p.image || p.sicf || p.id || '';
    const cleanTarget = String(rawTarget).replace(/^sicf:/, '');
    let targetImage = items.find(i =>
      (cleanTarget && (
        i.name === cleanTarget ||
        `${i.name}:${i.tag}` === cleanTarget ||
        `${i.repo}/${i.name}:${i.tag}` === cleanTarget ||
        `${i.repo}/${i.name}` === cleanTarget ||
        i.digest === cleanTarget ||
        i.id === cleanTarget
      )) ||
      (p.digest && (i.digest === p.digest || (i.layers && String(i.layers).split(',').includes(p.digest)))) ||
      (p.name && (i.name === p.name || i.name === `${p.name}:${p.tag}`))
    );
    if (!targetImage && items.length > 0 && !cleanTarget && !p.digest) {
      targetImage = items[0];
    }
    if (!targetImage && (p.digest || cleanTarget)) {
      targetImage = readTab('Images').find(i =>
        (p.digest && (i.digest === p.digest || (i.layers && String(i.layers).split(',').includes(p.digest)))) ||
        (cleanTarget && (i.name === cleanTarget || `${i.name}:${i.tag}` === cleanTarget || i.digest === cleanTarget || `${i.repo}/${i.name}` === cleanTarget))
      );
    }
    if (!targetImage) return json({ error: 'image not found' });

    const allLayers = readTab('Layers');
    const layerDigests = String(targetImage.layers || '').split(',').map(d => d.trim()).filter(Boolean);
    const resolvedLayers = [];
    const allMatchingChunks = [];

    layerDigests.forEach(layerDigest => {
      const layerChunks = allLayers
        .filter(c => c.digest === layerDigest)
        .sort((a, b) => Number(a.ordinal !== undefined ? a.ordinal : a.chunk_index) - Number(b.ordinal !== undefined ? b.ordinal : b.chunk_index));

      resolvedLayers.push({
        digest: layerDigest,
        media_type: (layerChunks[0] && layerChunks[0].media_type) || 'application/vnd.oci.image.layer.v1.tar',
        size_bytes: layerChunks.reduce((acc, c) => acc + (Number(c.size_bytes) || 0), 0),
        chunks: layerChunks
      });
      allMatchingChunks.push(...layerChunks);
    });

    if (allMatchingChunks.length === 0) {
      const legacyChunks = allLayers
        .filter(c => c.digest === targetImage.digest)
        .sort((a, b) => Number(a.ordinal !== undefined ? a.ordinal : a.chunk_index) - Number(b.ordinal !== undefined ? b.ordinal : b.chunk_index));
      allMatchingChunks.push(...legacyChunks);
    }

    return json({
      ok: true,
      image: targetImage,
      layers: resolvedLayers,
      chunks: allMatchingChunks.map((c, i) => ({
        ...c,
        layer_digest: c.digest,
        ordinal: Number(c.ordinal !== undefined ? c.ordinal : (c.chunk_index !== undefined ? c.chunk_index : 0)),
        chunk_index: i,
        data: c.data !== undefined ? c.data : c.chunk_data,
        chunk_data: c.data !== undefined ? c.data : c.chunk_data
      }))
    });
  }

  return json({ items: items });
}

// doPost: mutation operations (unconditional auth + non-throwing tryLock)
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json({ error: 'bad json' }); }

  // Unconditional authentication enforcement (SCRI §4)
  if (body.token !== TOKEN) return json({ error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ error: 'server busy' });

  try {
    const action = body.action || 'apply';
    const now = new Date().toISOString();

    if (action === 'create_issue') {
      const repo = body.repo || 'sncf/default';
      const issues = readTab('Issues');
      const repoIssues = issues.filter(i => i.repo === repo);
      const nextNum = repoIssues.length ? Math.max(...repoIssues.map(i => Number(i.number) || 0)) + 1 : 1;
      const iss = {
        id: 'iss-' + shortId(),
        repo: repo,
        number: nextNum,
        title: body.title || 'Untitled Issue',
        body: body.body || '',
        author: body.author || 'anonymous',
        state: 'open',
        labels: body.labels || 'enhancement',
        created_at: now,
        updated_at: now
      };
      issues.push(iss);
      writeTab('Issues', issues);
      return json({ ok: true, issue: iss });
    }

    if (action === 'create_mr') {
      const repo = body.repo || 'default';
      const mrs = readTab('MergeRequests');
      const repoMrs = mrs.filter(m => m.repo === repo);
      const nextNum = repoMrs.length ? Math.max(...repoMrs.map(m => Number(m.number) || 0)) + 1 : 1;
      const mr = {
        id: 'mr-' + shortId(),
        repo: repo,
        number: nextNum,
        title: body.title || 'Untitled MR',
        description: body.description || '',
        author: body.author || 'anonymous',
        source_branch: body.source_branch || 'feature',
        target_branch: body.target_branch || 'main',
        state: 'open',
        diff_manifest: typeof body.diff_manifest === 'string' ? body.diff_manifest : JSON.stringify(body.diff_manifest || {}, null, 2),
        created_at: now,
        updated_at: now
      };
      mrs.push(mr);
      writeTab('MergeRequests', mrs);
      return json({ ok: true, mr: mr });
    }

    if (action === 'merge_mr') {
      const mrs = readTab('MergeRequests');
      // Scope lookup by repo when matching by number, or match on unique id
      const mr = mrs.find(m => m.id === body.mr_id || (body.repo && m.repo === body.repo && String(m.number) === String(body.mr_id)));
      if (!mr) return json({ error: 'MR not found' });
      mr.state = 'merged';
      mr.updated_at = now;
      writeTab('MergeRequests', mrs);

      // If diff contains updated file content, apply it to target branch in Files tab
      try {
        const diff = JSON.parse(mr.diff_manifest);
        if (diff.path && diff.after) {
          const files = readTab('Files');
          let f = files.find(x => x.repo === mr.repo && x.path === diff.path && x.branch === mr.target_branch);
          if (f) {
            f.content = typeof diff.after === 'string' ? diff.after : JSON.stringify(diff.after, null, 2);
            f.updated_at = now;
          } else {
            files.push({
              repo: mr.repo,
              path: diff.path,
              branch: mr.target_branch,
              content: typeof diff.after === 'string' ? diff.after : JSON.stringify(diff.after, null, 2),
              updated_at: now
            });
          }
          writeTab('Files', files);
        }
      } catch (err) {}

      return json({ ok: true, mr: mr });
    }

    if (action === 'put_file') {
      const files = readTab('Files');
      let f = files.find(x => x.repo === body.repo && x.path === body.path && x.branch === (body.branch || 'main'));
      if (f) {
        f.content = body.content || '';
        f.updated_at = now;
      } else {
        files.push({
          repo: body.repo,
          path: body.path,
          branch: body.branch || 'main',
          content: body.content || '',
          updated_at: now
        });
      }
      writeTab('Files', files);
      return json({ ok: true });
    }

    if (action === 'add_comment') {
      const comments = readTab('Comments');
      const comment = {
        id: 'comm-' + shortId(),
        target_type: body.target_type || 'issue',
        target_id: body.target_id,
        author: body.author || 'anonymous',
        body: body.body || '',
        created_at: now
      };
      comments.push(comment);
      writeTab('Comments', comments);
      return json({ ok: true, comment: comment });
    }

    if (action === 'push_image' || action === 'push_layer') {
      const repo = body.repo || 'sncf/default';
      let rawName = body.name || 'image';
      let tag = body.tag || 'latest';
      if (rawName.includes(':')) {
        const parts = rawName.split(':');
        rawName = parts[0];
        tag = parts[1] || tag;
      }
      const name = rawName;
      const author = body.author || 'anonymous';
      const maxChunkChars = Math.min(Number(body.max_chunk_chars) || 30000, 32767);

      // 1. Normalize and process layers
      let inputLayers = [];
      if (Array.isArray(body.layers) && body.layers.length > 0) {
        inputLayers = body.layers;
      } else {
        // Single layer input
        inputLayers = [{
          digest: body.layer_digest || body.digest,
          media_type: body.media_type || 'application/vnd.oci.image.layer.v1.tar',
          chunks: body.chunks,
          content: body.content || body.data || body.base64,
          size_bytes: body.size_bytes
        }];
      }

      const allExistingLayers = readTab('Layers');
      const processedLayers = [];
      const newLayerRows = [];
      let totalAllLayerChars = 0;
      let totalChunkCount = 0;

      for (let lIdx = 0; lIdx < inputLayers.length; lIdx++) {
        const layerInput = inputLayers[lIdx];
        let layerChunks = [];

        if (Array.isArray(layerInput.chunks) && layerInput.chunks.length > 0) {
          layerInput.chunks.forEach((c, idx) => {
            const str = typeof c === 'string' ? c : (c.data || c.chunk_data || '');
            layerChunks.push({
              ordinal: idx,
              data: str,
              size_bytes: (typeof c === 'object' && c.size_bytes) || Math.round(str.length * 0.75)
            });
          });
        } else if (layerInput.content || layerInput.data || layerInput.base64) {
          const fullB64 = layerInput.content || layerInput.data || layerInput.base64;
          let idx = 0;
          for (let i = 0; i < fullB64.length; i += maxChunkChars) {
            const slice = fullB64.substring(i, i + maxChunkChars);
            layerChunks.push({
              ordinal: idx++,
              data: slice,
              size_bytes: Math.round(slice.length * 0.75)
            });
          }
        }

        if (layerChunks.length === 0) {
          // If layer digest is already present in Layers tab, reuse it (layer deduplication)
          if (layerInput.digest && allExistingLayers.some(c => c.digest === layerInput.digest)) {
            processedLayers.push({
              digest: layerInput.digest,
              media_type: layerInput.media_type || 'application/vnd.oci.image.layer.v1.tar',
              size_bytes: Number(layerInput.size_bytes) || 0,
              reused: true
            });
            continue;
          }
          return json({ error: 'no image content or chunks provided' });
        }

        // Reject empty / zero-byte total content
        const layerChars = layerChunks.reduce((acc, c) => acc + (c.data ? c.data.length : 0), 0);
        if (layerChars === 0) {
          return json({ error: 'cannot push empty / zero-byte layer' });
        }
        totalAllLayerChars += layerChars;
        totalChunkCount += layerChunks.length;

        // Strict enforcement of cell character limit (Excel 32,767 / Google Sheets 50,000 cap)
        for (let k = 0; k < layerChunks.length; k++) {
          if (layerChunks[k].data.length > 32767) {
            return json({ error: 'chunk character limit exceeded (max 32767 characters per cell)' });
          }
        }

        // Reassemble layer payload server-side and compute authoritative content-addressed SHA-256 digest
        const reassembledPayload = layerChunks.map(c => c.data).join('');
        let computedLayerDigest;
        try {
          computedLayerDigest = computeSha256(reassembledPayload);
        } catch (err) {
          return json({ error: 'invalid base64 payload: failed to decode and compute SHA-256' });
        }

        // Verify client-provided digest if supplied for this layer, reject on tamper/mismatch
        if (layerInput.digest && String(layerInput.digest).trim().toLowerCase() !== computedLayerDigest.toLowerCase()) {
          return json({ error: `digest mismatch: computed ${computedLayerDigest} but received ${layerInput.digest}` });
        }
        if (inputLayers.length === 1 && body.digest && String(body.digest).trim().toLowerCase() !== computedLayerDigest.toLowerCase() && !body.config && !body.digest.startsWith('sha256:')) {
          return json({ error: `digest mismatch: computed ${computedLayerDigest} but received ${body.digest}` });
        }

        const layerDigest = computedLayerDigest;
        const layerMedia = layerInput.media_type || 'application/vnd.oci.image.layer.v1.tar';
        const layerSize = Number(layerInput.size_bytes) || Math.floor(layerChars * 0.75);

        processedLayers.push({
          digest: layerDigest,
          media_type: layerMedia,
          size_bytes: layerSize,
          chunks: layerChunks
        });

        // Content-addressed layer deduplication:
        // If layer already exists in Layers tab, do NOT duplicate its chunks!
        const alreadyStored = allExistingLayers.some(c => c.digest === layerDigest) || newLayerRows.some(c => c.digest === layerDigest);
        if (!alreadyStored) {
          layerChunks.forEach(chk => {
            newLayerRows.push({
              digest: layerDigest,
              ordinal: chk.ordinal,
              media_type: layerMedia,
              data: chk.data,
              size_bytes: chk.size_bytes,
              id: 'chk-' + shortId(),
              created_at: now
            });
          });
        }
      }

      // Cap total chunks and payload to prevent sheet row/cell exhaustion
      const serverBytes = Math.floor(totalAllLayerChars * 0.75);
      if (totalChunkCount > 500 || serverBytes > 50 * 1024 * 1024) {
        return json({ error: 'payload exceeds maximum allowed registry layer size (max 500 chunks / 50MB)' });
      }

      // 2. Prepare OCI Config JSON and Config Digest
      let configB64 = '';
      let configDigest = '';
      const layerDigestList = processedLayers.map(l => l.digest);

      if (body.config) {
        if (typeof body.config === 'object') {
          const cfgJson = JSON.stringify(body.config);
          configB64 = Utilities.base64Encode(Utilities.newBlob(cfgJson).getBytes());
        } else if (typeof body.config === 'string') {
          if (/^[A-Za-z0-9+/]+={0,2}$/.test(body.config.trim()) && body.config.trim().length % 4 === 0) {
            configB64 = body.config.trim();
          } else {
            configB64 = Utilities.base64Encode(Utilities.newBlob(body.config).getBytes());
          }
        }
        try {
          configDigest = computeSha256(configB64);
        } catch (e) {
          return json({ error: 'invalid config payload' });
        }
      } else {
        // Standard OCI config JSON referencing the layer diff_ids
        const defaultCfg = {
          architecture: 'amd64',
          os: 'linux',
          config: {
            Env: ['PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin'],
            Cmd: [name]
          },
          rootfs: {
            type: 'layers',
            diff_ids: layerDigestList
          }
        };
        const defaultCfgJson = JSON.stringify(defaultCfg);
        configB64 = Utilities.base64Encode(Utilities.newBlob(defaultCfgJson).getBytes());
        configDigest = computeSha256(configB64);
      }

      // If client specified expected digest for the image, verify config digest or single-layer digest
      if (body.digest && String(body.digest).trim().toLowerCase() !== configDigest.toLowerCase() && (inputLayers.length !== 1 || String(body.digest).trim().toLowerCase() !== layerDigestList[0].toLowerCase())) {
        return json({ error: `digest mismatch: computed config digest ${configDigest} but received ${body.digest}` });
      }

      // Write newly added layer rows to Layers tab
      if (newLayerRows.length > 0) {
        allExistingLayers.push(...newLayerRows);
        writeTab('Layers', allExistingLayers);
      }

      // Total size = sum of layer sizes + config size
      const totalSize = Number(body.size_bytes) || (serverBytes + Math.floor(configB64.length * 0.75));

      // 3. Upsert into Images tab
      const images = readTab('Images');
      const fullName = name.includes(':') ? name : `${name}:${tag}`;
      let imageRecord = images.find(r => r.repo === repo && (r.name === name || r.name === fullName) && (r.tag === tag || !r.tag));

      // Clean up orphaned layers if updating an existing image tag and old layers are no longer in use
      if (imageRecord && imageRecord.layers) {
        const remainingImages = images.filter(r => r !== imageRecord);
        const layersInUse = new Set();
        remainingImages.forEach(r => {
          String(r.layers || '').split(',').forEach(d => { if (d.trim()) layersInUse.add(d.trim()); });
        });
        layerDigestList.forEach(d => layersInUse.add(d));

        let currentLayers = readTab('Layers');
        const updatedLayers = currentLayers.filter(c => layersInUse.has(c.digest));
        if (updatedLayers.length !== currentLayers.length) {
          writeTab('Layers', updatedLayers);
        }
      }

      const layersStr = layerDigestList.join(',');

      // Per SICF spec: image.digest is ALWAYS the config digest (the image's content identity)
      if (imageRecord) {
        imageRecord.name = fullName;
        imageRecord.digest = configDigest;
        imageRecord.config = configB64;
        imageRecord.layers = layersStr;
        imageRecord.size = totalSize;
        imageRecord.repo = repo;
        imageRecord.tag = tag;
        imageRecord.author = author;
        imageRecord.pushed_at = now;
        imageRecord.updated_at = now;
      } else {
        imageRecord = {
          name: fullName,
          digest: configDigest,
          config: configB64,
          layers: layersStr,
          created: now,
          size: totalSize,
          repo: repo,
          tag: tag,
          author: author,
          pushed_at: now,
          id: 'img-' + shortId(),
          updated_at: now
        };
        images.push(imageRecord);
      }
      writeTab('Images', images);

      return json({
        ok: true,
        image: imageRecord,
        digest: configDigest,
        config_digest: configDigest,
        layers: layerDigestList,
        chunk_count: totalChunkCount
      });
    }

    if (action === 'delete_image') {
      const images = readTab('Images');
      const targetIdx = images.findIndex(r =>
        (body.id && r.id === body.id) ||
        (body.digest && (r.digest === body.digest || String(r.layers || '').split(',').includes(body.digest)) && (!body.repo || r.repo === body.repo)) ||
        (body.repo && r.repo === body.repo && (
          r.name === body.image ||
          `${r.name}:${r.tag}` === body.image ||
          (r.name === body.name && r.tag === body.tag) ||
          r.name === `${body.name}:${body.tag}`
        )) ||
        (!body.repo && (r.name === body.image || `${r.name}:${r.tag}` === body.image))
      );

      if (targetIdx === -1) return json({ error: 'image not found' });
      const removed = images.splice(targetIdx, 1)[0];
      writeTab('Images', images);

      // Clean up layers from Layers tab ONLY IF no other image references the layer digests
      const removedLayerDigests = String(removed.layers || '').split(',').map(d => d.trim()).filter(Boolean);
      if (removed.digest) removedLayerDigests.push(removed.digest);

      const stillUsedDigests = new Set();
      images.forEach(r => {
        String(r.layers || '').split(',').forEach(d => { if (d.trim()) stillUsedDigests.add(d.trim()); });
        if (r.digest) stillUsedDigests.add(r.digest);
      });

      let allLayers = readTab('Layers');
      const layersToKeep = allLayers.filter(c => stillUsedDigests.has(c.digest) || !removedLayerDigests.includes(c.digest));
      if (layersToKeep.length !== allLayers.length) {
        writeTab('Layers', layersToKeep);
      }

      return json({ ok: true, deleted: removed });
    }

    if (action === 'pull_image') {
      const images = readTab('Images');
      const cleanTarget = String(body.image || body.name || body.id || '').replace(/^sicf:/, '');
      const targetImage = images.find(r =>
        (body.id && r.id === body.id) ||
        (body.digest && (r.digest === body.digest || String(r.layers || '').split(',').includes(body.digest))) ||
        (body.repo && r.repo === body.repo && (
          r.name === cleanTarget ||
          `${r.name}:${r.tag}` === cleanTarget ||
          (r.name === body.name && r.tag === body.tag) ||
          r.name === `${body.name}:${body.tag}`
        )) ||
        (cleanTarget && (r.name === cleanTarget || `${r.name}:${r.tag}` === cleanTarget || r.digest === cleanTarget))
      );

      if (!targetImage) return json({ error: 'image not found' });

      const allLayers = readTab('Layers');
      const layerDigests = String(targetImage.layers || '').split(',').map(d => d.trim()).filter(Boolean);
      const resolvedLayers = [];
      const allMatchingChunks = [];

      layerDigests.forEach(layerDigest => {
        const layerChunks = allLayers
          .filter(c => c.digest === layerDigest)
          .sort((a, b) => Number(a.ordinal !== undefined ? a.ordinal : a.chunk_index) - Number(b.ordinal !== undefined ? b.ordinal : b.chunk_index));
        resolvedLayers.push({
          digest: layerDigest,
          media_type: (layerChunks[0] && layerChunks[0].media_type) || 'application/vnd.oci.image.layer.v1.tar',
          size_bytes: layerChunks.reduce((acc, c) => acc + (Number(c.size_bytes) || 0), 0),
          chunks: layerChunks
        });
        allMatchingChunks.push(...layerChunks);
      });

      if (allMatchingChunks.length === 0) {
        const legacyChunks = allLayers
          .filter(c => c.digest === targetImage.digest)
          .sort((a, b) => Number(a.ordinal !== undefined ? a.ordinal : a.chunk_index) - Number(b.ordinal !== undefined ? b.ordinal : b.chunk_index));
        allMatchingChunks.push(...legacyChunks);
      }

      return json({
        ok: true,
        image: targetImage,
        layers: resolvedLayers,
        chunks: allMatchingChunks.map((c, i) => ({
          ...c,
          layer_digest: c.digest,
          ordinal: Number(c.ordinal !== undefined ? c.ordinal : (c.chunk_index !== undefined ? c.chunk_index : 0)),
          chunk_index: i,
          data: c.data !== undefined ? c.data : c.chunk_data,
          chunk_data: c.data !== undefined ? c.data : c.chunk_data
        }))
      });
    }

    return json({ error: 'unknown action: ' + action });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
