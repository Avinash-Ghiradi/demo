// Code.gs
// Deploy as Web App with execute as "Anyone" and access "Anyone"

const SPREADSHEET_ID = '1thV9sTEewsNy4VTLte7sx80ZaaqzHed5IxxnzgPL5w8';

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet() {
  const initResult = initializeSheets();
  if (!initResult.success) {
    console.error('Sheet initialization failed:', initResult.error);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Padhaya Bahut, Ab Jawab Do! - Teachers Day Quiz Challenge 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// HELPERS
// ============================================================
function normalizeDifficulty(value) {
  const v = String(value || '').trim().toLowerCase();
  if (v === 'easy') return 'Easy';
  if (v === 'hard') return 'Hard';
  return 'Medium';
}

function difficultyForQuestionNumber(questionNumber) {
  const n = Number(questionNumber);
  if (n >= 1 && n <= 5) return 'Easy';
  if (n >= 6 && n <= 10) return 'Medium';
  if (n >= 11 && n <= 15) return 'Hard';
  return '';
}

function findRowByPlayer(sheet, playerName) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(playerName).trim().toLowerCase()) {
      return i + 1;
    }
  }
  return -1;
}

function ensureQuestionIds_(questionsSheet) {
  const lastRow = questionsSheet.getLastRow();
  if (lastRow < 2) return;

  if (questionsSheet.getRange('H1').getValue() !== 'Question ID') {
    questionsSheet.getRange('H1').setValue('Question ID');
    questionsSheet.getRange('H1').setFontWeight('bold');
    questionsSheet.getRange('H1').setBackground('#f0c040');
    questionsSheet.getRange('H1').setFontColor('#1a1a2e');
  }

  const ids = questionsSheet.getRange(2, 8, lastRow - 1, 1).getValues();
  let changed = false;
  for (let i = 0; i < ids.length; i++) {
    if (!ids[i][0]) {
      ids[i][0] = 'Q-' + Utilities.getUuid();
      changed = true;
    }
  }
  if (changed) {
    questionsSheet.getRange(2, 8, ids.length, 1).setValues(ids);
  }
}

function ensureGameStateColumns_(sheet) {
  const headers = [
    'Player Name', 'Current Question', 'Score', 'Lifelines Used', 'Status',
    'Last Updated', 'Timer End', 'Answered', 'Selected Option',
    'Correct Option', 'Current Question ID', 'First Padav', 'Second Padav',
    'Secured Amount', '50:50 Used', 'Audience Poll Used', 'Ask Expert Used',
    'Question Swap Used'
  ];

  if (sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f0c040')
    .setFontColor('#1a1a2e');
}

function getQuestionRecords_(questionsSheet) {
  ensureQuestionIds_(questionsSheet);
  const data = questionsSheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    records.push({
      row: i + 1,
      id: String(data[i][7] || ''),
      question: data[i][0],
      optionA: data[i][1],
      optionB: data[i][2],
      optionC: data[i][3],
      optionD: data[i][4],
      correct: String(data[i][5] || '').trim().toUpperCase(),
      difficulty: normalizeDifficulty(data[i][6])
    });
  }
  return records;
}

function getAssignmentRecords_(assignmentSheet) {
  if (!assignmentSheet || assignmentSheet.getLastRow() < 2) return [];

  const data = assignmentSheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    records.push({
      questionId: String(data[i][0]),
      playerName: String(data[i][1] || ''),
      gameQuestion: Number(data[i][2] || 0),
      difficulty: normalizeDifficulty(data[i][3]),
      assignedAt: data[i][4] || '',
      status: String(data[i][5] || 'Locked')
    });
  }
  return records;
}

// ============================================================
// SHEET INITIALIZATION
// ============================================================
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ---------------- ADMIN ----------------
    let adminSheet = ss.getSheetByName('Admin');
    if (!adminSheet) {
      adminSheet = ss.insertSheet('Admin');
      adminSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
      adminSheet.getRange('A2:B3').setValues([
        ['Admin PIN', '1234'],
        ['Created Date', new Date().toISOString()]
      ]);
      adminSheet.setColumnWidths(1, 2, 200);
    }

    const pinValue = adminSheet.getRange('B2').getValue();
    if (!pinValue || pinValue.toString().trim() === '') {
      adminSheet.getRange('B2').setValue('1234');
    }
    adminSheet.getRange('A1:B1').setFontWeight('bold')
      .setBackground('#f0c040').setFontColor('#1a1a2e');

    // ---------------- QUESTIONS ----------------
    let questionsSheet = ss.getSheetByName('Questions');
    if (!questionsSheet) {
      questionsSheet = ss.insertSheet('Questions');
      questionsSheet.getRange('A1:H1').setValues([[
        'Question', 'Option A', 'Option B', 'Option C', 'Option D',
        'Correct Answer', 'Difficulty', 'Question ID'
      ]]);

      const sampleQuestions = [
        ['What is the capital of India?', 'Mumbai', 'New Delhi', 'Kolkata', 'Chennai', 'B', 'Easy'],
        ['Which planet is known as the Red Planet?', 'Venus', 'Mars', 'Jupiter', 'Saturn', 'B', 'Easy'],
        ['What is the largest ocean on Earth?', 'Atlantic Ocean', 'Indian Ocean', 'Arctic Ocean', 'Pacific Ocean', 'D', 'Easy'],
        ['How many days are there in a leap year?', '365', '366', '364', '360', 'B', 'Easy'],
        ['Which festival is known as the festival of lights?', 'Holi', 'Eid', 'Diwali', 'Pongal', 'C', 'Easy'],

        ['Who is known as the Father of Computers?', 'Alan Turing', 'Charles Babbage', 'Bill Gates', 'Steve Jobs', 'B', 'Medium'],
        ['What is the chemical symbol for Gold?', 'Go', 'Gd', 'Au', 'Ag', 'C', 'Medium'],
        ['Who painted the Mona Lisa?', 'Michelangelo', 'Leonardo da Vinci', 'Raphael', 'Donatello', 'B', 'Medium'],
        ['Which gas is most abundant in Earth’s atmosphere?', 'Oxygen', 'Carbon Dioxide', 'Nitrogen', 'Hydrogen', 'C', 'Medium'],
        ['Which Indian state is famous for the backwaters of Alappuzha?', 'Kerala', 'Goa', 'Odisha', 'Assam', 'A', 'Medium'],

        ['Which country has the largest population in 2024?', 'India', 'China', 'USA', 'Indonesia', 'A', 'Hard'],
        ['Which is the deepest ocean trench in the world?', 'Tonga Trench', 'Mariana Trench', 'Java Trench', 'Puerto Rico Trench', 'B', 'Hard'],
        ['What is the SI unit of electric resistance?', 'Volt', 'Watt', 'Ohm', 'Ampere', 'C', 'Hard'],
        ['Which ancient civilization developed cuneiform writing?', 'Roman', 'Sumerian', 'Mayan', 'Viking', 'B', 'Hard'],
        ['Which constitutional article deals with the Right to Constitutional Remedies in India?', 'Article 14', 'Article 19', 'Article 21', 'Article 32', 'D', 'Hard']
      ];

      questionsSheet.getRange('A2:G16').setValues(sampleQuestions);
      ensureQuestionIds_(questionsSheet);
    } else {
      ensureQuestionIds_(questionsSheet);
    }

    questionsSheet.setColumnWidth(1, 400);
    questionsSheet.setColumnWidths(2, 5, 200);
    questionsSheet.setColumnWidth(8, 280);
    questionsSheet.getRange(1, 1, 1, 8).setFontWeight('bold')
      .setBackground('#f0c040').setFontColor('#1a1a2e');

    // ---------------- GAME STATE ----------------
    let gameStateSheet = ss.getSheetByName('GameState');
    if (!gameStateSheet) {
      gameStateSheet = ss.insertSheet('GameState');
    }
    ensureGameStateColumns_(gameStateSheet);

    // ---------------- PRIZE MONEY ----------------
    let prizeSheet = ss.getSheetByName('PrizeMoney');
    if (!prizeSheet) prizeSheet = ss.insertSheet('PrizeMoney');
    prizeSheet.getRange('A1:B1').setValues([['Question Level', 'Prize Amount (₹)']]);
    const prizes = [
      [1, 1000], [2, 2000], [3, 3000], [4, 5000], [5, 10000],
      [6, 20000], [7, 40000], [8, 80000], [9, 160000], [10, 320000],
      [11, 640000], [12, 1250000], [13, 2500000], [14, 5000000], [15, 10000000]
    ];
    prizeSheet.getRange('A2:B16').setValues(prizes);
    prizeSheet.getRange('A1:B1').setFontWeight('bold')
      .setBackground('#f0c040').setFontColor('#1a1a2e');

    // ---------------- QUESTION ASSIGNMENTS ----------------
    let assignmentSheet = ss.getSheetByName('QuestionAssignments');
    if (!assignmentSheet) {
      assignmentSheet = ss.insertSheet('QuestionAssignments');
      assignmentSheet.getRange('A1:F1').setValues([[
        'Question ID', 'Player Name', 'Game Question', 'Difficulty', 'Assigned At', 'Status'
      ]]);
      assignmentSheet.getRange('A1:F1').setFontWeight('bold')
        .setBackground('#f0c040').setFontColor('#1a1a2e');
      assignmentSheet.setColumnWidths(1, 1, 280);
      assignmentSheet.setColumnWidths(2, 1, 180);
      assignmentSheet.setColumnWidths(3, 1, 120);
      assignmentSheet.setColumnWidths(4, 1, 120);
      assignmentSheet.setColumnWidths(5, 1, 220);
      assignmentSheet.setColumnWidths(6, 1, 120);
    }

    // ---------------- TV CONTROL ----------------
    let tvControlSheet = ss.getSheetByName('TVControl');
    if (!tvControlSheet) {
      tvControlSheet = ss.insertSheet('TVControl');
      tvControlSheet.getRange('A1:B2').setValues([
        ['Setting', 'Value'],
        ['Locked Player', '']
      ]);
      tvControlSheet.setColumnWidths(1, 2, 200);
      tvControlSheet.getRange('A1:B1').setFontWeight('bold')
        .setBackground('#f0c040').setFontColor('#1a1a2e');
    }

    return { success: true, message: 'Sheets initialized successfully!' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ADMIN FUNCTIONS
// ============================================================
function getAdminPIN() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Admin');
    const pinValue = sheet.getRange('B2').getValue();
    return pinValue ? pinValue.toString().trim() : '1234';
  } catch (e) {
    return '1234';
  }
}

function verifyAdminPIN(pin) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Admin');
    const storedPin = sheet.getRange('B2').getValue();
    const storedPinStr = storedPin ? storedPin.toString().trim() : '1234';
    const inputPinStr = pin ? pin.toString().trim() : '';
    return storedPinStr === inputPinStr
      ? { success: true }
      : { success: false, error: 'Invalid PIN' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateAdminPIN(oldPin, newPin) {
  try {
    const verifyResult = verifyAdminPIN(oldPin);
    if (!verifyResult.success) return { success: false, error: 'Invalid current PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ss.getSheetByName('Admin').getRange('B2').setValue(String(newPin).trim());
    return { success: true, message: 'PIN updated successfully!' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// QUESTION MANAGEMENT
// ============================================================
function addQuestion(questionData, adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');

    const difficulty = normalizeDifficulty(questionData.difficulty);
    sheet.appendRow([
      questionData.question,
      questionData.optionA,
      questionData.optionB,
      questionData.optionC,
      questionData.optionD,
      String(questionData.correct || '').toUpperCase(),
      difficulty,
      'Q-' + Utilities.getUuid()
    ]);

    return { success: true, message: 'Question added successfully!' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function bulkAddQuestions(questionsData, adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };
    if (!Array.isArray(questionsData)) return { success: false, error: 'Questions must be an array' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');

    const rows = [];
    questionsData.forEach(q => {
      rows.push([
        q.question,
        q.optionA,
        q.optionB,
        q.optionC,
        q.optionD,
        String(q.correct || '').toUpperCase(),
        normalizeDifficulty(q.difficulty),
        'Q-' + Utilities.getUuid()
      ]);
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    }

    return { success: true, message: `Added ${rows.length} questions successfully!` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteQuestion(questionId, adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const questionSheet = ss.getSheetByName('Questions');
    const assignmentSheet = ss.getSheetByName('QuestionAssignments');

    const used = getAssignmentRecords_(assignmentSheet)
      .some(a => a.questionId === String(questionId));

    if (used) {
      return {
        success: false,
        error: 'This question has already been shown and is locked. It cannot be deleted.'
      };
    }

    const data = questionSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7]) === String(questionId)) {
        questionSheet.deleteRow(i + 1);
        return { success: true, message: 'Question deleted successfully!' };
      }
    }

    return { success: false, error: 'Question not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getAllQuestions() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const questionSheet = ss.getSheetByName('Questions');
    const assignmentSheet = ss.getSheetByName('QuestionAssignments');

    const usedIds = {};
    getAssignmentRecords_(assignmentSheet).forEach(a => usedIds[a.questionId] = true);

    return getQuestionRecords_(questionSheet).map(q => ({
      id: q.id,
      question: q.question,
      optionA: q.optionA,
      optionB: q.optionB,
      optionC: q.optionC,
      optionD: q.optionD,
      correct: q.correct,
      difficulty: q.difficulty,
      used: !!usedIds[q.id]
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function exportQuestions() {
  try {
    const questions = getAllQuestions();
    return { success: true, data: JSON.stringify({
      totalQuestions: questions.length,
      questions: questions
    }, null, 2) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// QUESTION LOCKING / ASSIGNMENT
// ============================================================
function claimNextQuestion(playerName, gameQuestionNumber) {
  const lock = LockService.getScriptLock();

  try {
    if (!playerName) return { success: false, error: 'Player name is required.' };

    const n = Number(gameQuestionNumber);
    if (!Number.isInteger(n) || n < 1 || n > 15) {
      return { success: false, error: 'Invalid game question number.' };
    }

    const requiredDifficulty = difficultyForQuestionNumber(n);
    if (!requiredDifficulty) {
      return { success: false, error: 'Invalid difficulty level.' };
    }

    lock.waitLock(15000);

    initializeSheets();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gameSheet = ss.getSheetByName('GameState');
    const questionSheet = ss.getSheetByName('Questions');
    const assignmentSheet = ss.getSheetByName('QuestionAssignments');

    ensureGameStateColumns_(gameSheet);

    let playerRow = findRowByPlayer(gameSheet, playerName);
    if (playerRow === -1) {
      gameSheet.appendRow([
        playerName, n, 0, 0, 'Playing', new Date().toISOString(),
        '', false, '', '', ''
      ]);
      playerRow = gameSheet.getLastRow();
    }

    const gameData = gameSheet.getRange(playerRow, 1, 1, 11).getValues()[0];
    const storedQuestionNumber = Number(gameData[1] || 0);
    const storedQuestionId = String(gameData[10] || '');

    if (storedQuestionNumber === n && storedQuestionId) {
      const records = getQuestionRecords_(questionSheet);
      const existing = records.find(q => q.id === storedQuestionId);
      if (existing) {
        return {
          success: true,
          reused: true,
          gameQuestion: n,
          difficulty: existing.difficulty,
          id: existing.id,
          question: existing.question,
          options: [existing.optionA, existing.optionB, existing.optionC, existing.optionD],
          correct: existing.correct
        };
      }
    }

    const allQuestions = getQuestionRecords_(questionSheet);
    const assignments = getAssignmentRecords_(assignmentSheet);
    const usedIds = {};
    assignments.forEach(a => usedIds[a.questionId] = true);

    const available = allQuestions.filter(q =>
      q.difficulty === requiredDifficulty && !usedIds[q.id]
    );

    if (!available.length) {
      const count = allQuestions.filter(q => q.difficulty === requiredDifficulty).length;
      const used = allQuestions.filter(q =>
        q.difficulty === requiredDifficulty && usedIds[q.id]
      ).length;

      return {
        success: false,
        error:
          `No unused ${requiredDifficulty} questions are available for Q${n}. ` +
          `Available in pool: ${count - used}. Add more ${requiredDifficulty} questions.`
      };
    }

    const selected = available[Math.floor(Math.random() * available.length)];

    assignmentSheet.appendRow([
      selected.id,
      playerName,
      n,
      requiredDifficulty,
      new Date().toISOString(),
      'Locked'
    ]);

    gameSheet.getRange(playerRow, 2).setValue(n);
    gameSheet.getRange(playerRow, 5).setValue('Playing');
    gameSheet.getRange(playerRow, 6).setValue(new Date().toISOString());
    gameSheet.getRange(playerRow, 11).setValue(selected.id);

    return {
      success: true,
      reused: false,
      gameQuestion: n,
      difficulty: selected.difficulty,
      id: selected.id,
      question: selected.question,
      options: [selected.optionA, selected.optionB, selected.optionC, selected.optionD],
      correct: selected.correct
    };

  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

function getQuestionUsage(adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const questionSheet = ss.getSheetByName('Questions');
    const assignmentSheet = ss.getSheetByName('QuestionAssignments');

    const questions = getQuestionRecords_(questionSheet);
    const questionMap = {};
    questions.forEach(q => questionMap[q.id] = q);

    const usage = getAssignmentRecords_(assignmentSheet).map(a => {
      const q = questionMap[a.questionId] || {};
      return {
        questionId: a.questionId,
        question: q.question || '(Question deleted)',
        playerName: a.playerName,
        gameQuestion: a.gameQuestion,
        difficulty: a.difficulty,
        assignedAt: a.assignedAt,
        status: a.status
      };
    });

    const counts = {
      totalQuestions: questions.length,
      usedQuestions: usage.length,
      unusedQuestions: Math.max(0, questions.length - usage.length),
      easyUsed: usage.filter(x => x.difficulty === 'Easy').length,
      mediumUsed: usage.filter(x => x.difficulty === 'Medium').length,
      hardUsed: usage.filter(x => x.difficulty === 'Hard').length
    };

    return { success: true, counts: counts, usage: usage };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function resetQuestionPool(adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('QuestionAssignments');
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    return { success: true, message: 'Question pool reset. All questions are available for a new event.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// GAME FUNCTIONS
// ============================================================
function getQuestions() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Questions');

    return getQuestionRecords_(sheet).map(q => ({
      id: q.id,
      question: q.question,
      options: [q.optionA, q.optionB, q.optionC, q.optionD],
      correct: q.correct,
      difficulty: q.difficulty
    }));
  } catch (e) {
    return { error: e.message };
  }
}

function getPrizeMoney() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PrizeMoney');
    const data = sheet.getDataRange().getValues();
    const prizes = {};
    for (let i = 1; i < data.length; i++) {
      prizes[data[i][0]] = data[i][1];
    }
    return prizes;
  } catch (e) {
    return { error: e.message };
  }
}

function saveGameState(playerName, questionIndex, score, lifelinesUsed, status, extra) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    ensureGameStateColumns_(sheet);

    let playerRow = findRowByPlayer(sheet, playerName);
    const timestamp = new Date().toISOString();
    const x = extra || {};

    if (playerRow === -1) {
      sheet.appendRow([
        playerName, questionIndex, score, lifelinesUsed, status, timestamp,
        '', false, '', '', '',
        !!x.firstPadavReached, !!x.secondPadavReached, Number(x.securedAmount || 0),
        !!x.fiftyFiftyUsed, !!x.audiencePollUsed, !!x.askExpertUsed, !!x.questionSwapUsed
      ]);
    } else {
      sheet.getRange(playerRow, 2).setValue(questionIndex);
      sheet.getRange(playerRow, 3).setValue(score);
      sheet.getRange(playerRow, 4).setValue(lifelinesUsed);
      sheet.getRange(playerRow, 5).setValue(status);
      sheet.getRange(playerRow, 6).setValue(timestamp);
      if (Object.prototype.hasOwnProperty.call(x, 'firstPadavReached')) sheet.getRange(playerRow, 12).setValue(!!x.firstPadavReached);
      if (Object.prototype.hasOwnProperty.call(x, 'secondPadavReached')) sheet.getRange(playerRow, 13).setValue(!!x.secondPadavReached);
      if (Object.prototype.hasOwnProperty.call(x, 'securedAmount')) sheet.getRange(playerRow, 14).setValue(Number(x.securedAmount || 0));
      if (Object.prototype.hasOwnProperty.call(x, 'fiftyFiftyUsed')) sheet.getRange(playerRow, 15).setValue(!!x.fiftyFiftyUsed);
      if (Object.prototype.hasOwnProperty.call(x, 'audiencePollUsed')) sheet.getRange(playerRow, 16).setValue(!!x.audiencePollUsed);
      if (Object.prototype.hasOwnProperty.call(x, 'askExpertUsed')) sheet.getRange(playerRow, 17).setValue(!!x.askExpertUsed);
      if (Object.prototype.hasOwnProperty.call(x, 'questionSwapUsed')) sheet.getRange(playerRow, 18).setValue(!!x.questionSwapUsed);
    }

    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function swapCurrentQuestion(playerName, gameQuestionNumber) {
  const lock = LockService.getScriptLock();
  try {
    if (!playerName) return { success: false, error: 'Player name is required.' };
    const n = Number(gameQuestionNumber);
    if (!Number.isInteger(n) || n < 1 || n > 15) return { success: false, error: 'Invalid game question number.' };

    lock.waitLock(15000);
    initializeSheets();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gameSheet = ss.getSheetByName('GameState');
    const questionSheet = ss.getSheetByName('Questions');
    const assignmentSheet = ss.getSheetByName('QuestionAssignments');
    const playerRow = findRowByPlayer(gameSheet, playerName);
    if (playerRow === -1) return { success: false, error: 'Player not found.' };

    const gameData = gameSheet.getRange(playerRow, 1, 1, 18).getValues()[0];
    const currentId = String(gameData[10] || '');
    const requiredDifficulty = difficultyForQuestionNumber(n);
    const records = getQuestionRecords_(questionSheet);
    const assignments = getAssignmentRecords_(assignmentSheet);
    const usedIds = {};
    assignments.forEach(a => usedIds[a.questionId] = true);

    const available = records.filter(q => q.difficulty === requiredDifficulty && !usedIds[q.id] && q.id !== currentId);
    if (!available.length) {
      return { success: false, error: `No unused ${requiredDifficulty} question is available for Q${n} to swap in.` };
    }

    const selected = available[Math.floor(Math.random() * available.length)];

    if (currentId) {
      const assignmentData = assignmentSheet.getDataRange().getValues();
      for (let i = 1; i < assignmentData.length; i++) {
        if (String(assignmentData[i][0]) === currentId && String(assignmentData[i][1]).toLowerCase() === String(playerName).toLowerCase() && Number(assignmentData[i][2]) === n) {
          assignmentSheet.getRange(i + 1, 6).setValue('Swapped Out');
          break;
        }
      }
    }

    assignmentSheet.appendRow([
      selected.id, playerName, n, requiredDifficulty, new Date().toISOString(), 'Locked (Swap)'
    ]);
    gameSheet.getRange(playerRow, 2).setValue(n);
    gameSheet.getRange(playerRow, 6).setValue(new Date().toISOString());
    gameSheet.getRange(playerRow, 7).setValue('');
    gameSheet.getRange(playerRow, 8).setValue(false);
    gameSheet.getRange(playerRow, 9).setValue('');
    gameSheet.getRange(playerRow, 10).setValue('');
    gameSheet.getRange(playerRow, 11).setValue(selected.id);
    gameSheet.getRange(playerRow, 18).setValue(true);

    return {
      success: true,
      gameQuestion: n,
      difficulty: selected.difficulty,
      id: selected.id,
      question: selected.question,
      options: [selected.optionA, selected.optionB, selected.optionC, selected.optionD],
      correct: selected.correct
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function startQuestionTimer(playerName, durationSeconds) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    ensureGameStateColumns_(sheet);

    const playerRow = findRowByPlayer(sheet, playerName);
    if (playerRow === -1) return { error: 'Player not found' };

    const timerEnd = Date.now() + (Number(durationSeconds) * 1000);
    sheet.getRange(playerRow, 7).setValue(timerEnd);
    sheet.getRange(playerRow, 8).setValue(false);
    sheet.getRange(playerRow, 9).setValue('');
    sheet.getRange(playerRow, 10).setValue('');

    return { success: true, timerEnd: timerEnd };
  } catch (e) {
    return { error: e.message };
  }
}

function revealAnswer(playerName, selectedOption, correctOption) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    ensureGameStateColumns_(sheet);

    const playerRow = findRowByPlayer(sheet, playerName);
    if (playerRow === -1) return { error: 'Player not found' };

    sheet.getRange(playerRow, 8).setValue(true);
    sheet.getRange(playerRow, 9).setValue(selectedOption || '');
    sheet.getRange(playerRow, 10).setValue(correctOption || '');
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function getGameState(playerName) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    ensureGameStateColumns_(sheet);

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(playerName).trim().toLowerCase()) {
        return {
          playerName: data[i][0],
          currentQuestion: Number(data[i][1] || 1),
          score: Number(data[i][2] || 0),
          lifelinesUsed: Number(data[i][3] || 0),
          status: data[i][4],
          timerEnd: data[i][6] || null,
          answered: data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE',
          selectedOption: data[i][8] || '',
          correctOption: data[i][9] || '',
          currentQuestionId: data[i][10] || '',
          firstPadavReached: data[i][11] === true || String(data[i][11]).toUpperCase() === 'TRUE',
          secondPadavReached: data[i][12] === true || String(data[i][12]).toUpperCase() === 'TRUE',
          securedAmount: Number(data[i][13] || 0),
          fiftyFiftyUsed: data[i][14] === true || String(data[i][14]).toUpperCase() === 'TRUE',
          audiencePollUsed: data[i][15] === true || String(data[i][15]).toUpperCase() === 'TRUE',
          askExpertUsed: data[i][16] === true || String(data[i][16]).toUpperCase() === 'TRUE',
          questionSwapUsed: data[i][17] === true || String(data[i][17]).toUpperCase() === 'TRUE'
        };
      }
    }
    return { error: 'Player not found' };
  } catch (e) {
    return { error: e.message };
  }
}

function resetGame(playerName) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    ensureGameStateColumns_(sheet);

    const playerRow = findRowByPlayer(sheet, playerName);
    if (playerRow === -1) return { success: true };

    sheet.getRange(playerRow, 2).setValue(1);
    sheet.getRange(playerRow, 3).setValue(0);
    sheet.getRange(playerRow, 4).setValue(0);
    sheet.getRange(playerRow, 5).setValue('Playing');
    sheet.getRange(playerRow, 6).setValue(new Date().toISOString());
    sheet.getRange(playerRow, 7, 1, 12).setValues([['', false, '', '', '', '', false, false, 0, false, false, false]]);

    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function getLeaderboard() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');
    const data = sheet.getDataRange().getValues();

    const players = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        players.push({
          name: data[i][0],
          score: Number(data[i][2] || 0),
          status: data[i][4],
          securedAmount: Number(data[i][13] || 0)
        });
      }
    }

    players.sort((a, b) => b.score - a.score);
    return players;
  } catch (e) {
    return { error: e.message };
  }
}

function clearAllGameStates(adminPin) {
  try {
    const pinCheck = verifyAdminPIN(adminPin);
    if (!pinCheck.success) return { success: false, error: 'Invalid admin PIN' };

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('GameState');

    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    unlockTV();
    return {
      success: true,
      message: 'All game states cleared. Question locks/usage history were preserved.'
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// TV LOCK CONTROL & PUBLIC TV SCREEN
// ============================================================
function lockTVToPlayer(playerName) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ss.getSheetByName('TVControl').getRange('B2').setValue(playerName || '');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function unlockTV() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ss.getSheetByName('TVControl').getRange('B2').setValue('');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getTVLockedPlayer() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const v = ss.getSheetByName('TVControl').getRange('B2').getValue();
    return v ? v.toString().trim() : '';
  } catch (e) {
    return '';
  }
}

function getPublicGameState() {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gameSheet = ss.getSheetByName('GameState');
    const questionSheet = ss.getSheetByName('Questions');

    const lockedPlayer = getTVLockedPlayer();
    if (!lockedPlayer) {
      return { active: false, message: 'Waiting for host to start...' };
    }

    const gameData = gameSheet.getDataRange().getValues();
    let row = null;
    for (let i = 1; i < gameData.length; i++) {
      if (String(gameData[i][0]).trim().toLowerCase() === lockedPlayer.toLowerCase()) {
        row = gameData[i];
        break;
      }
    }

    if (!row) {
      return { active: false, message: 'Waiting for player to start...' };
    }

    const status = row[4];
    if (status !== 'Playing') {
      if (status === 'Completed') {
        return {
          active: false,
          message: `🎉 ${row[0]} won ₹${Number(row[2] || 0).toLocaleString('en-IN')}!`
        };
      }
      if (status === 'Lost') {
        return {
          active: false,
          message: `Game Over — ${row[0]} scored ₹${Number(row[2] || 0).toLocaleString('en-IN')}`
        };
      }
      return { active: false, message: 'Waiting for next game...' };
    }

    const records = getQuestionRecords_(questionSheet);
    let question = null;

    const currentQuestionId = String(row[10] || '');
    if (currentQuestionId) {
      question = records.find(q => q.id === currentQuestionId) || null;
    }

    if (!question) {
      const qIndex = Number(row[1] || 1) - 1;
      if (qIndex >= 0 && qIndex < records.length) {
        question = records[qIndex];
      }
    }

    if (!question) {
      return { active: false, message: 'Game over or current question not found' };
    }

    return {
      active: true,
      playerName: row[0],
      currentQuestion: Number(row[1] || 1),
      score: Number(row[2] || 0),
      question: question.question,
      options: [question.optionA, question.optionB, question.optionC, question.optionD],
      difficulty: question.difficulty,
      timerEnd: row[6] || null,
      answered: row[7] === true || String(row[7]).toUpperCase() === 'TRUE',
      selectedOption: row[8] || '',
      correctOption: row[9] || '',
      firstPadavReached: row[11] === true || String(row[11]).toUpperCase() === 'TRUE',
      secondPadavReached: row[12] === true || String(row[12]).toUpperCase() === 'TRUE',
      securedAmount: Number(row[13] || 0),
      fiftyFiftyUsed: row[14] === true || String(row[14]).toUpperCase() === 'TRUE',
      audiencePollUsed: row[15] === true || String(row[15]).toUpperCase() === 'TRUE',
      askExpertUsed: row[16] === true || String(row[16]).toUpperCase() === 'TRUE',
      questionSwapUsed: row[17] === true || String(row[17]).toUpperCase() === 'TRUE',
      finalQuestion: Number(row[1] || 1) === 15,
      majorQuestion: [5, 10, 15].indexOf(Number(row[1] || 1)) !== -1
    };
  } catch (e) {
    return { active: false, message: 'Error: ' + e.message };
  }
}
