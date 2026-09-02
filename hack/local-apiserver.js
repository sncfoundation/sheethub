#!/usr/bin/env node
// Local mock server for SheetHub — runs the real Code.gs control plane
// with an in-memory Google Sheet mock and serves both the SNCF REST API and the
// SheetHub GitLab-style Web UI locally.
//
// Usage: node hack/local-apiserver.js
// Open http://localhost:8788 in your browser.

const http = require('http');
const fs = require('fs');
const path = require('path');

const codeText = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const htmlPath = path.join(__dirname, '..', 'index.html');

// ---- In-memory Google Sheet mock ----
function makeSheet(header) { return { grid: [header.slice()] }; }
function lastNonEmpty(grid) {
  for (let r = grid.length; r >= 1; r--) {
    const row = grid[r - 1] || [];
    if (row.some(c => c !== '' && c !== undefined && c !== null)) return r;
  }
  return 1;
}

function rangeObj(sheet, row, col, numRows, numCols) {
  const o = {
    clearContent() {
      for (let i = 0; i < numRows; i++) {
        const r = row - 1 + i;
        if (!sheet.grid[r]) sheet.grid[r] = [];
        for (let j = 0; j < numCols; j++) sheet.grid[r][col - 1 + j] = '';
      }
      return o;
    },
    setValues(vals) {
      for (let i = 0; i < vals.length; i++) {
        const r = row - 1 + i;
        if (!sheet.grid[r]) sheet.grid[r] = [];
        for (let j = 0; j < vals[i].length; j++) sheet.grid[r][col - 1 + j] = vals[i][j];
      }
      return o;
    },
    setFontWeight() { return o; },
  };
  return o;
}

function makeSheetApi(sheet) {
  return {
    getDataRange: () => ({ getValues: () => sheet.grid.map(r => r.slice()) }),
    getRange: (r, c, nr, nc) => rangeObj(sheet, r, c, nr, nc),
    getLastRow: () => lastNonEmpty(sheet.grid),
    appendRow: (arr) => { sheet.grid.push(arr.slice()); },
    deleteRows: (start, num) => { sheet.grid.splice(start - 1, num); },
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
    getId: () => 'local-sheethub',
    getSheetByName: (n) => (store[n] ? makeSheetApi(store[n]) : null),
    insertSheet: (n) => {
      store[n] = store[n] || makeSheet([]);
      return makeSheetApi(store[n]);
    },
  })
};

let uuidN = 0;
const Utilities = { getUuid: () => (++uuidN).toString(36).split('').reverse().join('').padEnd(8, '0') };
const LockService = { getScriptLock: () => ({ tryLock() { return true; }, releaseLock() {} }) };
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
  codeText + '\nreturn { doGet, doPost, setup, readTab };'
)(SpreadsheetApp, Utilities, LockService, ContentService, HtmlService);

api.setup();

const PORT = process.env.PORT || 8788;

http.createServer((req, res) => {
  const u = new URL(req.url, `http://localhost:${PORT}`);

  // Serve static UI if accessing / or /index.html and no API params
  if (req.method === 'GET' && (u.pathname === '/' || u.pathname === '/index.html' || u.pathname === '/ui') && !u.searchParams.has('kind') && !u.searchParams.has('api') && !u.searchParams.has('token')) {
    if (fs.existsSync(htmlPath)) {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(fs.readFileSync(htmlPath, 'utf8'));
      return;
    }
  }

  const sendJson = (out) => {
    res.writeHead(200, {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    });
    res.end(typeof out === 'string' ? out : JSON.stringify(out));
  };

  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    });
    res.end();
    return;
  }

  if (req.method === 'GET') {
    const params = Object.fromEntries(u.searchParams);
    if (u.pathname === '/exec' && !params.kind) params.kind = 'repos';
    const e = { parameter: params };
    const out = api.doGet(e);
    sendJson(out._t || out);
  } else {
    let b = '';
    req.on('data', c => { b += c; });
    req.on('end', () => {
      try {
        const out = api.doPost({ postData: { contents: b } });
        sendJson(out._t || out);
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: String(err) }));
      }
    });
  }
}).listen(PORT, () => {
  console.log(`[local-apiserver] SheetHub Forge running on http://localhost:${PORT}`);
  console.log(`[local-apiserver] API endpoint: http://localhost:${PORT}/exec?token=CHANGE_ME_super_secret&kind=repos`);
});
