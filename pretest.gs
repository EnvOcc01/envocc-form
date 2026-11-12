/** Pre-test handlers. doGet/doPost are routed by Code.gs. */

const QUIZ_PRETEST_SHEET_NAME = 'pretest';

function handlePretestGet_(e) {
  const callback = e && e.parameter ? e.parameter.callback : '';
  return quizJsonResponse_({
    status: 'success',
    token: createQuizToken_(),
    questions: QUESTION_BANK.map(function (question) {
      return { id: question.id, question: question.question, choices: question.choices };
    })
  }, callback);
}

function handlePretestPost_(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (!consumeQuizToken_(params.requestToken)) return quizJsonResponse_({ status: 'error', message: 'แบบฟอร์มหมดอายุ กรุณาเปิดหน้าเว็บใหม่' });
  if (params.extra_company_url) return quizJsonResponse_({ status: 'success' });
  const result = calculateQuizResult_(params);
  const sheet = getQuizSheet_(QUIZ_PRETEST_SHEET_NAME);
  ensureQuizHeader_(sheet);
  sheet.appendRow([new Date(), params.fullName || '', params.prefix || '', params.position || '', params.agency || '', params.province || '', params.phone || '', params.email || ''].concat(result.answers, [result.score, result.total, result.percentage, '']));
  return quizJsonResponse_({ status: 'success', formType: 'pretest', ...result });
}

function calculateQuizResult_(params) {
  const results = [];
  const answers = [];
  let score = 0;
  QUESTION_BANK.forEach(function (question) {
    const raw = params['answerIndex_' + question.id];
    const selected = raw === undefined || raw === '' ? null : Number(raw);
    const validSelected = Number.isInteger(selected) && selected >= 0 && selected <= 3 ? selected : null;
    const isCorrect = validSelected !== null && validSelected === question.correctAnswer;
    if (isCorrect) score++;
    answers.push(validSelected === null ? '' : validSelected + 1);
    results.push({ id: question.id, selected: validSelected, correct: question.correctAnswer, isCorrect: isCorrect });
  });
  const total = QUESTION_BANK.length;
  return { score: score, total: total, percentage: total ? Math.round((score / total) * 100) : 0, passed: null, answers: answers, results: results };
}

function getQuizSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function ensureQuizHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(['timestamp', 'fullName', 'prefix', 'position', 'agency', 'province', 'phone', 'email'].concat(QUESTION_BANK.map(function (question) { return 'answer_' + question.id; }), ['score', 'total', 'percentage', 'passed']));
  sheet.setFrozenRows(1);
}
