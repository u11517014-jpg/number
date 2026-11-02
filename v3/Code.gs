// GAS 後端程式碼
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('四位數幾 A 幾 B - 版本三雲端存檔')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 初始化 Setup：自動建立專用試算表
function setup() {
  const userProps = PropertiesService.getScriptProperties();
  let sheetId = userProps.getProperty('GUESS_BOOK_ID');
  
  if (!sheetId) {
    const ss = SpreadsheetApp.create('數位科技概論｜幾AB猜數字版本三存檔');
    sheetId = ss.getId();
    userProps.setProperty('GUESS_BOOK_ID', sheetId);
    
    const sheet = ss.getActiveSheet();
    sheet.setName('遊戲存檔');
    sheet.appendRow(['使用者代號', '局ID', '局數', '狀態', '開始時間', '更新時間', '猜測次數', '猜測歷程', '完整存檔JSON']);
  }
  
  const ss = SpreadsheetApp.openById(sheetId);
  Logger.log('請保留此試算表網址：' + ss.getUrl());
  return ss.getUrl();
}

function getSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('GUESS_BOOK_ID');
  if (!sheetId) throw new Error('尚未執行 setup 初始化！');
  return SpreadsheetApp.openById(sheetId).getSheetByName('遊戲存檔');
}

// 產生 0-9 四位不重複數字答案
function generateAnswer() {
  const digits = ['0','1','2','3','4','5','6','7','8','9'];
  for (let i = digits.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [digits[i], digits[j]] = [digits[j], digits[i]];
  }
  return digits.slice(0, 4).join('');
}

// 讀取指定代號的所有遊戲紀錄
function loadUserData(userId) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const normalizedId = String(userId).trim().toLowerCase();
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const userGames = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === normalizedId) {
        try {
          const gameObject = JSON.parse(data[i][8]);
          userGames.push(gameObject);
        } catch (e) {}
      }
    }
    return { success: true, games: userGames };
  } finally {
    lock.releaseLock();
  }
}

// 建立新局
function createNewGame(userId) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const normalizedId = String(userId).trim().toLowerCase();
    const sheet = getSheet();
    const userData = loadUserData(normalizedId);
    
    const gameId = 'G_' + Date.now() + '_' + Math.floor(Math.random()*1000);
    const gameNum = userData.games.length + 1;
    const answer = generateAnswer();
    const nowISO = new Date().toISOString();

    const gameObj = {
      gameId: gameId,
      gameNum: gameNum,
      answer: answer,
      startTime: nowISO,
      updateTime: nowISO,
      status: '未完成',
      guesses: []
    };

    sheet.appendRow([
      normalizedId,
      gameId,
      gameNum,
      '未完成',
      nowISO,
      nowISO,
      0,
      '',
      JSON.stringify(gameObj)
    ]);

    return { success: true, game: gameObj };
  } finally {
    lock.releaseLock();
  }
}

// 提交猜測並更新試算表
function submitGuess(userId, gameId, guessStr) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const normalizedId = String(userId).trim().toLowerCase();
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    let gameObj = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === normalizedId && data[i][1] === gameId) {
        rowIndex = i + 1;
        gameObj = JSON.parse(data[i][8]);
        break;
      }
    }

    if (!gameObj) return { success: false, message: '找不到此局紀錄！' };
    if (gameObj.status === '已猜對') return { success: false, message: '本局已完成，無法繼續作答！' };

    // 計算 幾A幾B
    let a = 0, b = 0;
    for (let i = 0; i < 4; i++) {
      if (guessStr[i] === gameObj.answer[i]) a++;
      else if (gameObj.answer.includes(guessStr[i])) b++;
    }
    const resultStr = `${a}A${b}B`;
    const nowISO = new Date().toISOString();

    gameObj.guesses.push({
      step: gameObj.guesses.length + 1,
      guessStr: guessStr,
      resultStr: resultStr,
      time: nowISO
    });

    gameObj.updateTime = nowISO;
    if (a === 4) {
      gameObj.status = '已猜對';
    }

    // 格式化文字歷程
    const historyText = gameObj.guesses.map(g => `${g.step}. ${g.guessStr}➔${g.resultStr}`).join('\n');

    // 更新列
    sheet.getRange(rowIndex, 4).setValue(gameObj.status);
    sheet.getRange(rowIndex, 6).setValue(nowISO);
    sheet.getRange(rowIndex, 7).setValue(gameObj.guesses.length);
    sheet.getRange(rowIndex, 8).setValue(historyText);
    sheet.getRange(rowIndex, 9).setValue(JSON.stringify(gameObj));

    return { success: true, game: gameObj, latestResult: { guessStr, resultStr } };
  } finally {
    lock.releaseLock();
  }
}