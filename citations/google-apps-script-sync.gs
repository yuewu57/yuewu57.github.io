/**
 * Yue Wu — OpenAlex citation dashboard
 * Standalone Google Apps Script.
 *
 * Script Properties required:
 *   OPENALEX_API_KEY
 *   GITHUB_CITATION_TOKEN
 *
 * Public entry points:
 *   refreshOpenAlexNow()              // manual refresh, always runs
 *   refreshOpenAlex()                 // scheduled refresh; weekdays, once/day
 *   syncCitationDashboardToGitHubNow()
 *   installCitationDashboardTrigger() // optional: installs one daily 08:00 trigger
 */

const OPENALEX_ORCID = '0000-0002-6281-2229';
const OPENALEX_AUTHOR_ID = 'A5100663751';

const DASHBOARD_GITHUB_OWNER = 'yuewu57';
const DASHBOARD_GITHUB_REPO = 'yuewu57.github.io';
const DASHBOARD_GITHUB_BRANCH = 'master';
const DASHBOARD_DATA_PATH = 'citations/data/openalex.json';

const CURRENT_WORKS_HEADERS = [
  'openalex_id','doi','title','publication_date','publication_year','type',
  'cited_by_count','source','is_retracted','openalex_updated_date'
];

const SNAPSHOT_HEADERS = [
  'snapshot_time','openalex_id','doi','title','publication_date','publication_year',
  'type','cited_by_count','source','is_retracted','openalex_updated_date'
];

const SUMMARY_HEADERS = [
  'snapshot_time','author_name','openalex_author_id','orcid','works_count',
  'cited_by_count','h_index','i10_index','openalex_author_updated_date'
];

const DATA_QUALITY_HEADERS = [
  'snapshot_time','severity','issue','openalex_id','title','details'
];

/** Manual refresh. Always runs, even if today has already been refreshed. */
function refreshOpenAlexNow() {
  return refreshOpenAlexData_({ force: true });
}

/**
 * Scheduled refresh entry point.
 * Runs only Monday-Friday and at most once per local calendar day.
 */
function refreshOpenAlex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/London';
  const now = new Date();
  const weekday = Number(Utilities.formatDate(now, tz, 'u')); // 1=Mon ... 7=Sun
  if (weekday > 5) return { skipped: true, reason: 'weekend' };

  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const last = props.getProperty('LAST_SUCCESSFUL_REFRESH');
  if (last) {
    const lastDate = new Date(last);
    if (!isNaN(lastDate.getTime())) {
      const lastDay = Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd');
      if (lastDay === today) {
        return { skipped: true, reason: 'already-refreshed-today' };
      }
    }
  }
  return refreshOpenAlexData_({ force: false });
}

/**
 * Optional one-time helper. Removes duplicate refreshOpenAlex triggers and
 * installs one daily trigger at about 08:00 in the Apps Script project timezone.
 * refreshOpenAlex() itself skips weekends.
 */
function installCitationDashboardTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshOpenAlex')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshOpenAlex')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

/** Public wrapper for a GitHub-only sync without re-fetching all OpenAlex works. */
function syncCitationDashboardToGitHubNow() {
  return syncCitationDashboardToGitHub_();
}

function refreshOpenAlexData_(options) {
  options = options || {};
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPENALEX_API_KEY');
  if (!apiKey) throw new Error('Missing Script Property: OPENALEX_API_KEY');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  const author = fetchOpenAlexAuthor_(apiKey);
  const authorId = String(author.id || '').replace('https://openalex.org/', '');
  if (authorId !== OPENALEX_AUTHOR_ID) {
    throw new Error('OpenAlex author identity check failed. Expected ' + OPENALEX_AUTHOR_ID + ' but received ' + authorId);
  }

  const authorOrcid = String(author.orcid || '').replace('https://orcid.org/', '');
  if (authorOrcid && authorOrcid !== OPENALEX_ORCID) {
    throw new Error('OpenAlex ORCID check failed. Expected ' + OPENALEX_ORCID + ' but received ' + authorOrcid);
  }

  const works = fetchAllOpenAlexWorks_(authorId, apiKey);
  const annualCitationMap = buildAnnualCitationMap_(works);

  const currentRows = works.map(workToCurrentRow_);
  const snapshotRows = works.map(w => [now].concat(workToCurrentRow_(w)));
  const qualityRows = buildDataQualityRows_(works, now);

  const summaryStats = author.summary_stats || {};
  const summaryRow = [
    now,
    String(author.display_name || 'Yue Wu'),
    authorId,
    OPENALEX_ORCID,
    Number(author.works_count || works.length || 0),
    Number(author.cited_by_count || 0),
    Number(summaryStats.h_index || 0),
    Number(summaryStats.i10_index || 0),
    toSheetDate_(author.updated_date)
  ];

  replaceSheetData_(ss, 'Current_Works', CURRENT_WORKS_HEADERS, currentRows);
  appendSheetData_(ss, 'Snapshots', SNAPSHOT_HEADERS, snapshotRows);
  appendSheetData_(ss, 'Summary_History', SUMMARY_HEADERS, [summaryRow]);
  replaceSheetData_(ss, 'Data_Quality', DATA_QUALITY_HEADERS, qualityRows);

  props.setProperty('LAST_SUCCESSFUL_REFRESH', now.toISOString());
  props.deleteProperty('LAST_OPENALEX_REFRESH_ERROR');

  try {
    syncCitationDashboardToGitHub_(annualCitationMap);
    props.deleteProperty('LAST_GITHUB_SYNC_ERROR');
  } catch (err) {
    props.setProperty('LAST_GITHUB_SYNC_ERROR', now.toISOString() + ' | ' + String(err && err.message ? err.message : err));
    throw err;
  }

  return {
    refreshed_at: now.toISOString(),
    author_id: authorId,
    works: works.length,
    citations: Number(author.cited_by_count || 0),
    h_index: Number(summaryStats.h_index || 0),
    i10_index: Number(summaryStats.i10_index || 0),
    data_quality_issues: qualityRows.length
  };
}

function fetchOpenAlexAuthor_(apiKey) {
  const url = 'https://api.openalex.org/authors/' + encodeURIComponent(OPENALEX_AUTHOR_ID) +
    '?api_key=' + encodeURIComponent(apiKey);
  return fetchJson_(url, 'OpenAlex author fetch');
}

function fetchAllOpenAlexWorks_(authorId, apiKey) {
  const out = [];
  let cursor = '*';

  do {
    const url = 'https://api.openalex.org/works' +
      '?filter=' + encodeURIComponent('author.id:' + authorId) +
      '&select=' + encodeURIComponent('id,doi,title,publication_date,publication_year,type,cited_by_count,primary_location,is_retracted,updated_date,counts_by_year') +
      '&per_page=100' +
      '&cursor=' + encodeURIComponent(cursor) +
      '&api_key=' + encodeURIComponent(apiKey);

    const body = fetchJson_(url, 'OpenAlex works fetch');
    (body.results || []).forEach(w => out.push(w));
    cursor = body.meta && body.meta.next_cursor ? body.meta.next_cursor : null;
  } while (cursor);

  out.sort((a, b) => {
    const cy = Number(b.cited_by_count || 0) - Number(a.cited_by_count || 0);
    if (cy) return cy;
    return String(b.publication_date || '').localeCompare(String(a.publication_date || ''));
  });
  return out;
}

function fetchJson_(url, label) {
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: { Accept: 'application/json' }
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(label + ' failed: HTTP ' + status + ' ' + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

function workToCurrentRow_(work) {
  const source = work && work.primary_location && work.primary_location.source
    ? work.primary_location.source.display_name
    : null;
  return [
    String(work.id || ''),
    work.doi ? String(work.doi) : '',
    String(work.title || ''),
    work.publication_date ? String(work.publication_date) : '',
    Number(work.publication_year || 0) || '',
    String(work.type || ''),
    Number(work.cited_by_count || 0),
    source ? String(source) : '',
    Boolean(work.is_retracted),
    toSheetDate_(work.updated_date)
  ];
}

function buildAnnualCitationMap_(works) {
  const out = {};
  (works || []).forEach(work => {
    out[String(work.id || '')] = (work.counts_by_year || []).map(row => ({
      year: Number(row.year),
      cited_by_count: Number(row.cited_by_count || 0)
    }));
  });
  return out;
}

function buildDataQualityRows_(works, snapshotTime) {
  const rows = [];
  const doiMap = {};
  const titleMap = {};

  (works || []).forEach(work => {
    const id = String(work.id || '');
    const title = String(work.title || '');
    const doi = work.doi ? String(work.doi) : '';
    const source = work && work.primary_location && work.primary_location.source
      ? String(work.primary_location.source.display_name || '')
      : '';

    if (!doi) rows.push([snapshotTime, 'INFO', 'Missing DOI', id, title, 'OpenAlex record has no DOI.']);
    if (!work.publication_date) rows.push([snapshotTime, 'WARN', 'Missing publication date', id, title, 'OpenAlex record has no publication_date.']);
    if (!source) rows.push([snapshotTime, 'INFO', 'Missing source', id, title, 'OpenAlex primary location has no source name.']);
    if (work.is_retracted) rows.push([snapshotTime, 'WARN', 'Retracted record', id, title, 'OpenAlex marks this work as retracted.']);

    if (doi) {
      const key = doi.toLowerCase();
      doiMap[key] = doiMap[key] || [];
      doiMap[key].push(work);
    }
    if (title) {
      const key = title.toLowerCase().replace(/\s+/g, ' ').trim();
      titleMap[key] = titleMap[key] || [];
      titleMap[key].push(work);
    }
  });

  Object.keys(doiMap).forEach(key => {
    if (doiMap[key].length > 1) {
      doiMap[key].forEach(work => rows.push([
        snapshotTime, 'WARN', 'Duplicate DOI', String(work.id || ''), String(work.title || ''),
        'The same DOI appears on ' + doiMap[key].length + ' OpenAlex records.'
      ]));
    }
  });

  Object.keys(titleMap).forEach(key => {
    if (key && titleMap[key].length > 1) {
      titleMap[key].forEach(work => rows.push([
        snapshotTime, 'INFO', 'Duplicate title', String(work.id || ''), String(work.title || ''),
        'The same normalized title appears on ' + titleMap[key].length + ' OpenAlex records.'
      ]));
    }
  });

  return rows;
}

function replaceSheetData_(ss, sheetName, headers, rows) {
  const sheet = getOrCreateSheet_(ss, sheetName);
  const neededCols = headers.length;
  if (sheet.getMaxColumns() < neededCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function appendSheetData_(ss, sheetName, headers, rows) {
  if (!rows || !rows.length) return;
  const sheet = getOrCreateSheet_(ss, sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function syncCitationDashboardToGitHub_(annualCitationMap) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_CITATION_TOKEN');
  if (!token) throw new Error('Missing Script Property: GITHUB_CITATION_TOKEN');

  const openAlexApiKey = props.getProperty('OPENALEX_API_KEY');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = ss.getSheetByName('Current_Works');
  const summary = ss.getSheetByName('Summary_History');
  if (!current || !summary) throw new Error('Missing Current_Works or Summary_History tab.');

  const currentValues = current.getDataRange().getValues();
  const summaryValues = summary.getDataRange().getValues();
  if (currentValues.length < 2) throw new Error('Current_Works has no data rows.');

  const idx = headerIndex_(currentValues[0]);
  CURRENT_WORKS_HEADERS.forEach(k => {
    if (idx[k] === undefined) throw new Error('Current_Works missing column: ' + k);
  });

  const baseWorks = currentValues.slice(1)
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
      openalex_updated: normaliseDateForJson_(r[idx.openalex_updated_date]),
      citation_counts_by_year: []
    }));

  if (!annualCitationMap) {
    annualCitationMap = openAlexApiKey
      ? fetchWorkCitationCountsByYear_(baseWorks.map(w => w.id), openAlexApiKey)
      : {};
  }

  const works = baseWorks.map(w => ({
    ...w,
    citation_counts_by_year: annualCitationMap[w.id] || []
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
    schema_version: 2,
    generated_at: new Date().toISOString(),
    source: {
      name: 'OpenAlex via private Google Sheet',
      openalex_author_id: sidx.openalex_author_id !== undefined ? String(lastRow[sidx.openalex_author_id] || '') : '',
      orcid: sidx.orcid !== undefined ? String(lastRow[sidx.orcid] || '') : '',
      raw_author_metrics: latest,
      annual_citation_window: 'OpenAlex work-level counts_by_year',
      note: 'OpenAlex counts can differ from Google Scholar. Reconciliation is applied separately on the public dashboard.'
    },
    works: works,
    summary_history: summaryHistory
  };

  const json = JSON.stringify(payload);
  upsertGitHubTextFile_(DASHBOARD_DATA_PATH, json, 'Update citation dashboard data', token);

  props.setProperty('LAST_GITHUB_DASHBOARD_SYNC', new Date().toISOString());
  props.deleteProperty('LAST_GITHUB_SYNC_ERROR');
  return { works: works.length, history_points: summaryHistory.length };
}

function fetchWorkCitationCountsByYear_(workIds, apiKey) {
  const ids = [...new Set((workIds || [])
    .map(id => String(id || '').replace('https://openalex.org/', ''))
    .filter(Boolean))];
  const out = {};
  if (!ids.length) return out;

  const chunkSize = 100;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    const url = 'https://api.openalex.org/works' +
      '?filter=' + encodeURIComponent('openalex_id:' + chunk.join('|')) +
      '&select=' + encodeURIComponent('id,counts_by_year') +
      '&per_page=' + chunk.length +
      '&api_key=' + encodeURIComponent(apiKey);

    const body = fetchJson_(url, 'OpenAlex annual citation fetch');
    (body.results || []).forEach(work => {
      out[String(work.id)] = (work.counts_by_year || []).map(row => ({
        year: Number(row.year),
        cited_by_count: Number(row.cited_by_count || 0)
      }));
    });
  }
  return out;
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
    throw new Error('GitHub read failed: HTTP ' + getStatus + ' ' + getResponse.getContentText());
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
    throw new Error('GitHub write failed: HTTP ' + putStatus + ' ' + putResponse.getContentText());
  }
}

function toSheetDate_(value) {
  if (!value) return '';
  if (value instanceof Date) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? String(value) : d;
}

function normaliseDateForJson_(value) {
  if (value instanceof Date) return value.toISOString();
  return value === '' || value === null || value === undefined ? null : String(value);
}
