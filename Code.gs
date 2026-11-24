/**
 * ILAW generator backend. Bind this script to a Google Sheet
 * (Extensions > Apps Script). Plans are logged to the "Plans" tab.
 * Script Properties required: ANTHROPIC_API_KEY, ACCESS_CODE
 */
const MODEL = 'claude-sonnet-5-5';
const SHEET_NAME = 'Plans';
const DAILY_LIMIT = 100; // max generations per day, protects your API credits

function doGet() {
  return ContentService.createTextOutput('ILAW backend is running.');
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    if (!p.code || p.code !== props.getProperty('ACCESS_CODE')) {
      return out({ error: 'Invalid access code.' });
    }
    checkLimit_(props);
    const plan = callClaude_(p, props.getProperty('ANTHROPIC_API_KEY'));
    logPlan_(p, plan);
    return out({ plan: plan });
  } catch (err) {
    return out({ error: String(err.message || err) });
  }
}

function checkLimit_(props) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const day = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
    const key = 'count_' + day;
    const used = Number(props.getProperty(key) || 0);
    if (used >= DAILY_LIMIT) throw new Error('Daily limit reached. Please try again tomorrow.');
    props.setProperty(key, String(used + 1));
  } finally {
    lock.releaseLock();
  }
}

function callClaude_(f, apiKey) {
  const prompt =
    'You are an experienced Philippine DepEd curriculum specialist. Write a lesson plan in the DepEd ILAW framework (Intentions, Learning Experience, Assessing Learning, Ways Forward) aligned to the MATATAG curriculum where applicable.\n' +
    'Subject: ' + f.subject + '\nGrade: ' + f.grade + '\nTopic or competency: ' + f.topic +
    '\nSession length: ' + f.mins + ' minutes\nWrite in: ' + f.lang +
    '\nLearner and class notes: ' + (f.ctx || 'none') + '\n\n' +
    'Rules: use Philippine context and low-cost materials where natural; make activity minutes add up to about ' + f.mins +
    '; include at least two distinct activities; objectives must be observable and measurable; do not add sections beyond the four ILAW parts; if the topic is not a formal competency, restate it as one in the "competency" field.\n' +
    'Return ONLY JSON with this shape: {"title":string,"competency":string,"intentions":{"objectives":[string],"successCriteria":[string]},"learningExperience":{"activities":[{"phase":string,"title":string,"minutes":number,"description":string,"differentiation":string}],"materials":[string]},"assessingLearning":{"quickCheck":[string],"endOfLesson":string,"rubric":[string]},"waysForward":{"reinforcement":[string],"remediation":[string],"enrichment":[string],"nextLesson":string}}';
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: MODEL, max_tokens: 3000, messages: [{ role: 'user', content: prompt }] }),
    muteHttpExceptions: true
  });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error((body.error && body.error.message) || 'AI request failed.');
  const text = body.content.map(function (c) { return c.text || ''; }).join('')
    .replace(/```json|```/g, '').trim();
  return JSON.parse(text);
}

function logPlan_(f, plan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Subject', 'Grade', 'Topic', 'Minutes', 'Language', 'Title', 'Plan JSON']);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), f.subject, f.grade, f.topic, f.mins, f.lang, plan.title || '', JSON.stringify(plan)]);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
