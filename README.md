<p align="center">
  <img src="https://sncfoundation.github.io/logos/sheethub.svg" width="96" alt="SheetHub logo">
</p>

<h1 align="center">SheetHub</h1>

<p align="center"><b>An all-in-one DevOps forge in one workbook</b><br>
A <a href="https://sncfoundation.github.io">Sheet-Native Computing Foundation</a> project &#183; analog of <b>GitLab</b></p>

---

**Status:** 💾 Autosaving &#183; a working prototype of the forge **metadata layer** — repos, issues and merge requests over a local CLI. Design notes and contributions are welcome.

## About

An all-in-one DevOps forge in one workbook. Part of the spreadsheet-native stack — the
Sheet stays the source of truth, and it reconciles.

## Usage

`sheethub` is a single bash script. It needs `bash` and `jq`.

```bash
# repos
sheethub repo create webapp --desc "our main app"
sheethub repo list

# issues (sequential ids per repo)
sheethub issue create webapp "Login button misaligned" --body "on mobile"
sheethub issue list webapp                 # open by default
sheethub issue list webapp --state all     # open|closed|all
sheethub issue close webapp 1

# merge requests
sheethub mr create webapp "Fix login layout" --from fix/login --to main
sheethub mr list webapp

# a GitLab-style overview
sheethub stats
```

Every entity gets a **sequential id per repo** (issue `#1`, MR `!1`, …).

## Where state lives

State is stored in a single JSON file at `${SHEETHUB_HOME:-~/.sheethub}/hub.json`.

For this MVP that **JSON store stands in for the spreadsheet tabs**: its top-level keys —
`repos`, `issues`, `mrs` — are the tabs a real SheetHub workbook would keep. The
metadata layer is real and working; the **real git-transport** (pushing and pulling actual
code through the Sheet) is a **stretch goal**.

## Get involved

- 📋 Tracking issue &amp; design: [sncfoundation/sheeternetes#35](https://github.com/sncfoundation/sheeternetes/issues/35)
- 🗺️ [SNCF Landscape](https://sncfoundation.github.io/landscape.html)
- 🧩 [All projects](https://sncfoundation.github.io/projects.html)
- ⚖️ [Governance &amp; how to contribute](https://github.com/sncfoundation/governance)
- 🎓 [Get certified (CSFE)](https://sncfoundation.github.io/certification.html)

## Status legend

Everything starts as an **Unsaved Draft**. It reconciles up the tiers from there — see the
[maturity model](https://sncfoundation.github.io/foundation.html#maturity).

---

<sub>Licensed under Apache-2.0. The SNCF does not recommend running production on a spreadsheet. If you do, please film it.</sub>
