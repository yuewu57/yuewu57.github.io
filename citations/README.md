# Citation dashboard

Public interactive dashboard for Yue Wu's OpenAlex-backed publication profile.

- `index.html` — interactive client-side dashboard.
- `data/openalex.json` — generated from the private Google Sheet; updated automatically by Apps Script.
- `data/reconciliation.json` — manually curated OpenAlex authorship-review flags, kept separate so automated refreshes do not overwrite reconciliation decisions.
- `google-apps-script-sync.gs` — copy into the bound Apps Script project that already refreshes OpenAlex.

## Automatic updates

1. Create a fine-grained GitHub personal access token scoped only to `yuewu57/yuewu57.github.io`, with repository permission **Contents: Read and write**.
2. In Apps Script → Project Settings → Script Properties, add `GITHUB_CITATION_TOKEN` with that token as its value.
3. Paste `google-apps-script-sync.gs` into the existing Apps Script project.
4. At the end of the existing successful `refreshOpenAlexData_()` function, call `syncCitationDashboardToGitHub_();`.

The sync hashes scholarly data and creates a GitHub commit only when the source data changes. No OpenAlex or GitHub token is written into the public repository.
