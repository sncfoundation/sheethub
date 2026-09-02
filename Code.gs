/**
 * SheetHub — a GitLab-style DevOps forge on a Google Sheet.
 *
 * Part of the Sheet-Native Computing Foundation (SNCF) ecosystem.
 * Data plane: Google Sheet (tabs for Repos, Issues, MergeRequests, Releases, Users, Comments, Stars, Files).
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
  Registry:       ['id', 'repo', 'name', 'tag', 'digest', 'size_bytes', 'chunk_count', 'chunk_map', 'author', 'created_at', 'updated_at'],
  RegistryChunks: ['id', 'digest', 'chunk_index', 'chunk_data', 'size_bytes', 'created_at'],
};

// ---------- generic sheet <-> objects ----------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function readTab(name) {
  const s = ss();
  const sh = s.getSheetByName(name);
  if (!sh) return [];
  const rng = sh.getDataRange().getValues();
  if (!rng || rng.length < 2) return [];
  const header = rng.shift();
  return rng
    .filter(r => String(r[0]).trim() !== '')
    .map(row => {
      const o = {};
      header.forEach((h, i) => { o[h] = row[i]; });
      return o;
    });
}

function writeTab(name, objects) {
  const s = ss();
  let sh = s.getSheetByName(name);
  if (!sh) {
    sh = s.insertSheet(name);
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const header = TABS[name];
  const rows = objects.map(o => header.map(h => (o[h] === undefined || o[h] === null) ? '' : o[h]));
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, header.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function shortId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

function computeSha256(base64Str) {
  try {
    const bytes = Utilities.base64Decode(base64Str);
    const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
    let hex = '';
    for (let i = 0; i < rawHash.length; i++) {
      let byteVal = rawHash[i];
      if (byteVal < 0) byteVal += 256;
      const byteHex = byteVal.toString(16);
      hex += byteHex.length === 1 ? '0' + byteHex : byteHex;
    }
    return 'sha256:' + hex;
  } catch (err) {
    return 'sha256:unknown';
  }
}

function validateChunkMap(input, chunkCount, digest) {
  if (!input) return null;
  try {
    const parsed = typeof input === 'string' ? JSON.parse(input) : input;
    if (!Array.isArray(parsed) || parsed.length !== chunkCount) return null;
    const sanitized = parsed.map((m, idx) => ({
      index: Number(m.index) || idx,
      sheet: String(m.sheet || 'RegistryChunks').replace(/[^a-zA-Z0-9_-]/g, ''),
      cell: String(m.cell || `D${idx + 2}`).replace(/[^a-zA-Z0-9]/g, ''),
      chars: Number(m.chars) || 0,
      size_bytes: Number(m.size_bytes) || 0,
      digest: String(m.digest || digest)
    }));
    return JSON.stringify(sanitized);
  } catch (e) {
    return null;
  }
}

// ---------- one-time setup & multi-repo seed data ----------
function setupTab(name) {
  const s = ss();
  let sh = s.getSheetByName(name);
  if (!sh) sh = s.insertSheet(name);
  sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function setup() {
  Object.keys(TABS).forEach(name => {
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

  // Seed Container Registry
  const registry = readTab('Registry');
  if (!registry.length) {
    const whoamiDigest = 'sha256:7f83b1657ff1fc53b92dc18148a1d65dfc2d4b1fa3d677284addd200126d9069';
    const ingressDigest = 'sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

    const sampleChunk1 = 'H4sICN8+v2YCA2hlbGxvLXdlYi50YXIA7Z1rc9s2Ese/SpX7Yd25K3s2d2U6dZ3Ex04yN5P6ykmk5MskMRdFkdT8eAAYgCIeEtAiyfG0Y4kEQGDx211gsViA/y8vLy/n51+vrq6evX/26sWzV5eXlz+ev/rpyfPnL1++ePnq6cvnLy9eXj5//vzps6eXl6+ev3z+9OXzV49ffHn16tWzn6/+/u1/Hj368/r/vv3vf3/9n4e3/37996+/';
    const sampleChunk2 = 'mZzO5vP56eX5y/nJ9Ozn57Pz17Pz2cvz19PT6dnz8/lsNpufTk/PTmZn569Onr+cnrz86eTp2dlp+t1/AAAA///817t0wAAA';

    writeTab('Registry', [
      {
        id: 'img-001',
        repo: 'sncf/hello-web',
        name: 'traefik/whoami',
        tag: 'latest',
        digest: whoamiDigest,
        size_bytes: 49152,
        chunk_count: 2,
        chunk_map: JSON.stringify([
          { index: 0, sheet: 'RegistryChunks', cell: 'D2', chars: sampleChunk1.length, size_bytes: 24576, digest: whoamiDigest },
          { index: 1, sheet: 'RegistryChunks', cell: 'D3', chars: sampleChunk2.length, size_bytes: 24576, digest: whoamiDigest }
        ]),
        author: 'prateeekbuilds',
        created_at: now,
        updated_at: now
      },
      {
        id: 'img-101',
        repo: 'sncf/sheeternetes-manifests',
        name: 'ingress-router',
        tag: 'v2.10',
        digest: ingressDigest,
        size_bytes: 32768,
        chunk_count: 1,
        chunk_map: JSON.stringify([
          { index: 0, sheet: 'RegistryChunks', cell: 'D4', chars: sampleChunk1.length, size_bytes: 32768, digest: ingressDigest }
        ]),
        author: 'tym83',
        created_at: now,
        updated_at: now
      }
    ]);

    writeTab('RegistryChunks', [
      { id: 'chk-001', digest: whoamiDigest, chunk_index: 0, chunk_data: sampleChunk1, size_bytes: 24576, created_at: now },
      { id: 'chk-002', digest: whoamiDigest, chunk_index: 1, chunk_data: sampleChunk2, size_bytes: 24576, created_at: now },
      { id: 'chk-101', digest: ingressDigest, chunk_index: 0, chunk_data: sampleChunk1, size_bytes: 32768, created_at: now },
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
    registry: 'Registry',
    images: 'Registry',
    containerregistry: 'Registry',
    registry_chunks: 'RegistryChunks',
    chunks: 'RegistryChunks',
    registry_pull: 'Registry',
  };

  const tab = map[kind];
  if (!tab) return json({ error: 'unknown kind ' + kind });

  let items = readTab(tab);
  if (repoFilter && ['Issues', 'MergeRequests', 'Releases', 'Files', 'Stars', 'Registry'].indexOf(tab) !== -1) {
    items = items.filter(it => String(it.repo).toLowerCase() === repoFilter.toLowerCase());
  }

  if (tab === 'Registry') {
    if (p.name) items = items.filter(it => String(it.name).toLowerCase() === p.name.toLowerCase());
    if (p.tag) items = items.filter(it => String(it.tag).toLowerCase() === p.tag.toLowerCase());
    if (p.digest) items = items.filter(it => String(it.digest) === String(p.digest));
  }

  if (p.id) {
    items = items.filter(it => String(it.id) === String(p.id) || String(it.number) === String(p.id) || String(it.tag) === String(p.id) || String(it.digest) === String(p.id));
  }

  // Handle pull query (fetch image + ordered chunks)
  if (kind === 'registry_pull' || (tab === 'Registry' && p.pull === 'true')) {
    const targetImage = items[0] || (p.digest ? readTab('Registry').find(i => i.digest === p.digest) : null);
    if (!targetImage) return json({ error: 'image not found' });
    const allChunks = readTab('RegistryChunks');
    const imageChunks = allChunks
      .filter(c => c.digest === targetImage.digest)
      .sort((a, b) => Number(a.chunk_index) - Number(b.chunk_index));
    return json({ ok: true, image: targetImage, chunks: imageChunks });
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
      const repo = body.repo || 'default';
      const issues = readTab('Issues');
      const repoIssues = issues.filter(i => i.repo === repo);
      const nextNum = repoIssues.length ? Math.max(...repoIssues.map(i => Number(i.number) || 0)) + 1 : 1;
      const issue = {
        id: 'iss-' + shortId(),
        repo: repo,
        number: nextNum,
        title: body.title || 'Untitled Issue',
        body: body.body || '',
        author: body.author || 'anonymous',
        state: 'open',
        labels: body.labels || 'issue',
        created_at: now,
        updated_at: now
      };
      issues.push(issue);
      writeTab('Issues', issues);
      return json({ ok: true, issue: issue });
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
      const name = body.name || 'image';
      const tag = body.tag || 'latest';
      const author = body.author || 'anonymous';
      const rawChunks = body.chunks || [];
      const maxChunkChars = Math.min(Number(body.max_chunk_chars) || 30000, 32767);

      let chunkList = [];
      if (Array.isArray(rawChunks) && rawChunks.length > 0) {
        rawChunks.forEach((c, idx) => {
          const str = typeof c === 'string' ? c : (c.chunk_data || c.data || '');
          chunkList.push({
            index: idx,
            data: str,
            size_bytes: c.size_bytes || Math.round(str.length * 0.75)
          });
        });
      } else if (body.content || body.data || body.base64) {
        const fullBase64 = body.content || body.data || body.base64;
        let idx = 0;
        for (let i = 0; i < fullBase64.length; i += maxChunkChars) {
          const slice = fullBase64.substring(i, i + maxChunkChars);
          chunkList.push({
            index: idx++,
            data: slice,
            size_bytes: Math.round(slice.length * 0.75)
          });
        }
      }

      if (chunkList.length === 0) {
        return json({ error: 'no image content or chunks provided' });
      }

      // Strict enforcement of cell character limit (Excel 32,767 / Google Sheets 50,000 cap)
      for (let k = 0; k < chunkList.length; k++) {
        if (chunkList[k].data.length > 32767) {
          return json({ error: 'chunk character limit exceeded (max 32767 characters per cell)' });
        }
      }

      // Reassemble full payload server-side and compute authoritative content-addressed SHA-256 digest
      const reassembledPayload = chunkList.map(c => c.data).join('');
      const computedDigest = computeSha256(reassembledPayload);

      // Verify client-provided digest if supplied, reject on tamper/mismatch
      if (body.digest && String(body.digest).trim().toLowerCase() !== computedDigest.toLowerCase()) {
        return json({ error: `digest mismatch: computed ${computedDigest} but received ${body.digest}` });
      }

      const digest = computedDigest;
      const totalSize = Number(body.size_bytes) || chunkList.reduce((acc, c) => acc + (c.size_bytes || 0), 0);

      // Look up registry to handle updating existing images and orphaned chunks
      const registry = readTab('Registry');
      let imageRecord = registry.find(r => r.repo === repo && r.name === name && r.tag === tag);
      let allChunks = readTab('RegistryChunks');

      // If updating an existing image tag and its digest changed, clean up old orphaned chunks
      if (imageRecord && imageRecord.digest && imageRecord.digest !== digest) {
        const oldDigest = imageRecord.digest;
        const isOldDigestUsed = registry.some(r => r.id !== imageRecord.id && r.digest === oldDigest);
        if (!isOldDigestUsed) {
          allChunks = allChunks.filter(c => c.digest !== oldDigest);
        }
      }

      // Clean up previous chunks for this exact digest if re-pushing
      allChunks = allChunks.filter(c => c.digest !== digest);

      const chunkMap = [];
      const newChunkRows = [];

      chunkList.forEach((chk) => {
        const chunkId = 'chk-' + shortId();
        const sheetName = 'RegistryChunks';
        const cellCoord = 'D' + (allChunks.length + newChunkRows.length + 2);

        chunkMap.push({
          index: chk.index,
          sheet: sheetName,
          cell: cellCoord,
          chars: chk.data.length,
          size_bytes: chk.size_bytes,
          digest: digest
        });

        newChunkRows.push({
          id: chunkId,
          digest: digest,
          chunk_index: chk.index,
          chunk_data: chk.data,
          size_bytes: chk.size_bytes,
          created_at: now
        });
      });

      allChunks.push(...newChunkRows);
      writeTab('RegistryChunks', allChunks);

      // Validate client-provided chunk_map if any, otherwise use server-constructed chunkMap
      const validatedMap = validateChunkMap(body.chunk_map, chunkList.length, digest) || JSON.stringify(chunkMap);

      // Upsert Registry metadata
      if (imageRecord) {
        imageRecord.digest = digest;
        imageRecord.size_bytes = totalSize;
        imageRecord.chunk_count = chunkList.length;
        imageRecord.chunk_map = validatedMap;
        imageRecord.author = author;
        imageRecord.updated_at = now;
      } else {
        imageRecord = {
          id: 'img-' + shortId(),
          repo: repo,
          name: name,
          tag: tag,
          digest: digest,
          size_bytes: totalSize,
          chunk_count: chunkList.length,
          chunk_map: validatedMap,
          author: author,
          created_at: now,
          updated_at: now
        };
        registry.push(imageRecord);
      }
      writeTab('Registry', registry);

      return json({ ok: true, image: imageRecord, chunk_count: chunkList.length, digest: digest });
    }

    if (action === 'delete_image') {
      const registry = readTab('Registry');
      const targetIdx = registry.findIndex(r =>
        (body.id && r.id === body.id) ||
        (body.digest && r.digest === body.digest) ||
        (body.repo && r.repo === body.repo && (r.name + ':' + r.tag === body.image || (r.name === body.name && r.tag === body.tag)))
      );

      if (targetIdx === -1) return json({ error: 'image not found' });
      const removed = registry.splice(targetIdx, 1)[0];
      writeTab('Registry', registry);

      // Clean up chunks from chunk sheet
      let allChunks = readTab('RegistryChunks');
      allChunks = allChunks.filter(c => c.digest !== removed.digest);
      writeTab('RegistryChunks', allChunks);

      return json({ ok: true, deleted: removed });
    }

    if (action === 'pull_image') {
      const registry = readTab('Registry');
      const targetImage = registry.find(r =>
        (body.id && r.id === body.id) ||
        (body.digest && r.digest === body.digest) ||
        (body.repo && r.repo === body.repo && (r.name + ':' + r.tag === body.image || (r.name === body.name && r.tag === body.tag)))
      );

      if (!targetImage) return json({ error: 'image not found' });

      const allChunks = readTab('RegistryChunks');
      const imageChunks = allChunks
        .filter(c => c.digest === targetImage.digest)
        .sort((a, b) => Number(a.chunk_index) - Number(b.chunk_index));

      return json({ ok: true, image: targetImage, chunks: imageChunks });
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
