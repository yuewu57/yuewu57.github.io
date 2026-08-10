/**
 * Push the latest public scholarly data from the bound Google Sheet to
 * yuewu57/yuewu57.github.io/citations/data/openalex.json.
 *
 * Required Script Property:
 *   GITHUB_CITATION_TOKEN
 *
 * Create the token as a fine-grained GitHub PAT scoped only to
 * yuewu57/yuewu57.github.io with Contents: Read and write.
 *
 * Add this at the end of your existing successful refreshOpenAlexData_():
 *   syncCitationDashboardToGitHub_();
 */
const DASHBOARD_GITHUB_OWNER = 'yuewu57';
const DASHBOARD_GITHUB_REPO = 'yuewu57.github.io';
const DASHBOARD_GITHUB_BRANCH = 'master';
const DASHBOARD_DATA_PATH = 'citations/data/openalex.json';

function syncCitationDashboardToGitHub_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_CITATION_TOKEN');
  if (!token) throw new Error('Missing Script Property: GITHUB_CITATION_TOKEN');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = ss.getSheetByName('Current_Works');
  const summary = ss.getSheetByName('Summary_History');
  if (!current || !summary) {
    throw new Error('Missing Current_Works or Summary_History tab.');
  }

  const currentValues = current.getDataRange().getValues();
  const summaryValues = summary.getDataRange().getValues();
  if (currentValues.length < 2) throw new Error('Current_Works has no data rows.');

  const idx = headerIndex_(currentValues[0]);
  const required = [
    'openalex_id','doi','title','publication_date','publication_year','type',
    'cited_by_count','source','is_retracted','openalex_updated_date'
  ];
  required.forEach(k => {
    if (idx[k] === undefined) throw new Error('Current_Works missing column: ' + k);
  });

  const works = currentValues.slice(1)
    .filter(r => r[idx.openalex_id])
    .map(r => ({
      id: String(r[idx.openalex_id]),
      doi: r[idx.doi] ? String(r[idx.doi]) : null,
      title: String(r[idx.title] || ''),
      publication_date: normaliseDateForJson_(r[idx.publication_date]),
      year: Number(r[idx.publication_year] || 0) || null,
      type: String(r[idx.type] || ''),
      citations: Number(r[idx.cited_by_count] || 0),
      source: r[idx.source] ? String(r[idx.source]) : null,
      retracted: Boolean(r[idx.is_retracted]),
      openalex_updated: normaliseDateForJson_(r[idx.openalex_updated_date])
    }));

  const sidx = headerIndex_(summaryValues[0] || []);
  const summaryHistory = summaryValues.slice(1)
    .filter(r => r[sidx.snapshot_time])
    .map(r => ({
      time: normaliseDateForJson_(r[sidx.snapshot_time]),
      works: Number(r[sidx.works_count] || 0),
      citations: Number(r[sidx.cited_by_count] || 0),
      h_index: Number(r[sidx.h_index] || 0),
      i10_index: Number(r[sidx.i10_index] || 0)
    }));

  const lastRow = summaryValues.length > 1 ? summaryValues[summaryValues.length - 1] : [];
  const latest = summaryHistory.length ? summaryHistory[summaryHistory.length - 1] : null;

  const payload = {
    schema_version: 1,
    generated_at: new Date().toISOString(),
    source: {
      name: 'OpenAlex via private Google Sheet',
      openalex_author_id: sidx.openalex_author_id !== undefined ? String(lastRow[sidx.openalex_author_id] || '') : '',
      orcid: sidx.orcid !== undefined ? String(lastRow[sidx.orcid] || '') : '',
      raw_author_metrics: latest,
      note: 'OpenAlex counts can differ from Google Scholar. Reconciliation is applied separately on the public dashboard.'
    },
    works: works,
    summary_history: summaryHistory
  };

  const json = JSON.stringify(payload);
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.newBlob(json).getBytes()
  );
  const hash = Utilities.base64Encode(digest);

  if (props.getProperty('LAST_GITHUB_DASHBOARD_HASH') === hash) return;

  upsertGitHubTextFile_(
    DASHBOARD_DATA_PATH,
    json,
    'Update citation dashboard data',
    token
  );

  props.setProperty('LAST_GITHUB_DASHBOARD_HASH', hash);
  props.setProperty('LAST_GITHUB_DASHBOARD_SYNC', new Date().toISOString());
}

function headerIndex_(header) {
  const out = {};
  header.forEach((name, i) => out[String(name)] = i);
  return out;
}

function upsertGitHubTextFile_(path, content, message, token) {
  const encodedPath = path.split('/').map(encodeURIComponent).join('/');
  const api = 'https://api.github.com/repos/' +
    DASHBOARD_GITHUB_OWNER + '/' + DASHBOARD_GITHUB_REPO +
    '/contents/' + encodedPath;

  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };

  const getResponse = UrlFetchApp.fetch(
    api + '?ref=' + encodeURIComponent(DASHBOARD_GITHUB_BRANCH),
    { method: 'get', headers: headers, muteHttpExceptions: true }
  );

  const getStatus = getResponse.getResponseCode();
  let sha = null;
  if (getStatus === 200) {
    sha = JSON.parse(getResponse.getContentText()).sha;
  } else if (getStatus !== 404) {
    throw new Error(
      'GitHub read failed: HTTP ' + getStatus + ' ' + getResponse.getContentText()
    );
  }

  const body = {
    message: message,
    content: Utilities.base64Encode(Utilities.newBlob(content).getBytes()),
    branch: DASHBOARD_GITHUB_BRANCH
  };
  if (sha) body.sha = sha;

  const putResponse = UrlFetchApp.fetch(api, {
    method: 'put',
    headers: headers,
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const putStatus = putResponse.getResponseCode();
  if (putStatus !== 200 && putStatus !== 201) {
    throw new Error(
      'GitHub write failed: HTTP ' + putStatus + ' ' + putResponse.getContentText()
    );
  }
}

function normaliseDateForJson_(value) {
  if (value instanceof Date) return value.toISOString();
  return value === '' || value === null ? null : String(value);
}
