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
  Repos:         makeSheet(['name', 'description', 'default_branch', 'visibility', 'stars_count', 'created_at', 'updated_at']),
  Issues:        makeSheet(['id', 'repo', 'number', 'title', 'body', 'author', 'state', 'labels', 'created_at', 'updated_at']),
  MergeRequests: makeSheet(['id', 'repo', 'number', 'title', 'description', 'author', 'source_branch', 'target_branch', 'state', 'diff_manifest', 'created_at', 'updated_at']),
  Releases:      makeSheet(['id', 'repo', 'tag_name', 'name', 'body', 'author', 'created_at', 'assets']),
  Users:         makeSheet(['username', 'name', 'avatar_url', 'role', 'bio', 'created_at']),
  Comments:      makeSheet(['id', 'target_type', 'target_id', 'author', 'body', 'created_at']),
  Stars:         makeSheet(['repo', 'username', 'starred_at']),
  Files:         makeSheet(['repo', 'path', 'branch', 'content', 'updated_at']),
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
const Utilities = { getUuid: () => (++uuidN).toString(36).split('').reverse().join('').padEnd(8, '0') };
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

console.log(`\n==== SheetHub Test Results: ${pass} passed, ${fail} failed ====`);
process.exit(fail ? 1 : 0);
