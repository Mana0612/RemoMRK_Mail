const SHEET_RULES = 'メール通知設定';
const SHEET_LOG = '通知ログ';
const SHEET_SYS = 'システム設定';
const SHEET_DEBUG = '抽出デバッグ';

const DEFAULT_SINCE_DATE_STR = '2025/08/01';
// 指定いただいたNotionページ（親ページ）URL/IDを初期投入
const DEFAULT_PARENT_PAGE_ID_HINT = '258bee38fcfe804cb15bf2326be5bb58';
const NOTION_API_BASE = 'https://api.notion.com/v1';
const NOTION_VERSION = '2022-06-28';

function onOpen() {
  ensureSheets();
  SpreadsheetApp.getUi()
    .createMenu('メール通知システム')
    .addItem('手動で通知実行', 'runNotificationsNow')
    .addItem('設定シートを再生成', 'regenerateSettings')
    .addItem('抽出デバッグを実行', 'runDebugSearch')
    .addItem('毎時トリガーを再作成', 'recreateHourlyTrigger')
    .addToUi();
}

function regenerateSettings() {
  const ss = SpreadsheetApp.getActive();
  [SHEET_RULES, SHEET_LOG, SHEET_SYS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  ensureSheets();
}

function recreateHourlyTrigger() {
  const handler = 'runNotificationsNow';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActive();
  // メール通知設定
  let sh = ss.getSheetByName(SHEET_RULES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RULES);
    sh.getRange(1, 1, 1, 6).setValues([[
      '件名キーワード',
      'アドレス（完全一致）',
      '処理済みGmailラベル名',
      'Notion DB ID',
      '備考（任意)',
      '除外キーワード（任意）'
    ]]);
  }
  // 通知ログ
  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.getRange(1, 1, 1, 9).setValues([[
      '実行日時', 'メール受信日時', '件名', '送信者', 'メッセージID', 'DB ID', 'ステータス', 'エラーメッセージ', '備考'
    ]]);
  } else {
    // 既存を9列ヘッダーに揃える（列名も更新）
    const cols = Math.max(9, log.getLastColumn());
    if (cols < 9) log.insertColumnsAfter(cols, 9 - cols);
    log.getRange(1, 1, 1, 9).setValues([[
      '実行日時', 'メール受信日時', '件名', '送信者', 'メッセージID', 'DB ID', 'ステータス', 'エラーメッセージ', '備考'
    ]]);
  }
  // システム設定
  let sys = ss.getSheetByName(SHEET_SYS);
  if (!sys) {
    sys = ss.insertSheet(SHEET_SYS);
    sys.getRange(1, 1, 3, 2).setValues([
      ['NOTION_TOKEN', ''],
      ['DEFAULT_PARENT_PAGE_ID', DEFAULT_PARENT_PAGE_ID_HINT],
      ['SINCE_DATE', DEFAULT_SINCE_DATE_STR]
    ]);
  } else {
    // 既存でも不足キーがあれば補完
    const map = getSystemSettings();
    if (!map['DEFAULT_PARENT_PAGE_ID']) sys.getRange('A2:B2').setValues([[
      'DEFAULT_PARENT_PAGE_ID', DEFAULT_PARENT_PAGE_ID_HINT
    ]]);
    if (!map['SINCE_DATE']) sys.getRange('A3:B3').setValues([[
      'SINCE_DATE', DEFAULT_SINCE_DATE_STR
    ]]);
    if (!map['NOTION_TOKEN']) sys.getRange('A1').setValue('NOTION_TOKEN');
  }
}

function runDebugSearch() {
  const sys = getSystemSettings();
  const sinceDate = coerceSinceDate(sys['SINCE_DATE']) || coerceSinceDate(DEFAULT_SINCE_DATE_STR);
  const ss = SpreadsheetApp.getActive();
  let dbg = ss.getSheetByName(SHEET_DEBUG);
  if (!dbg) {
    dbg = ss.insertSheet(SHEET_DEBUG);
  }
  // 毎回ヘッダーをA1:J1に再セットして列不足を自動拡張（列ズレ防止）
  dbg.getRange(1, 1, 1, 10).setValues([[
    'ルール行', 'クエリ', 'ヒット件数(上限100)', '件名キーワード', 'サンプル件名', 'サンプル送信元', 'サンプル受信日時', '備考', '抽出日', 'マッチ件数(フィルタ後)'
  ]]);

  const rules = readRules();
  let row = Math.max(2, dbg.getLastRow() + 1);
  rules.forEach(rule => {
    if (!rule.labelName) return;
    if (!rule.subjectKeyword && !rule.fromAddress) return;
    const query = buildQueryString(rule, sinceDate, rule.labelName);
    const threads = GmailApp.search(query, 0, 100);
    let sampleSub = '', sampleFrom = '', sampleDate = '';
    let matchedCount = 0;
    if (threads.length > 0) {
      // フィルタ後の件数計測＆サンプル抽出
      outer: for (let t = 0; t < threads.length; t++) {
        const messages = threads[t].getMessages();
        for (let m = 0; m < messages.length; m++) {
          const msg = messages[m];
          if (msg.getDate() < sinceDate) continue;
          if (!isRuleMatch(rule, msg)) continue;
          matchedCount++;
          if (!sampleSub) {
            const tz = Session.getScriptTimeZone();
            sampleSub = msg.getSubject() || '';
            sampleFrom = msg.getFrom() || '';
            sampleDate = Utilities.formatDate(msg.getDate(), tz, 'yyyy/MM/dd HH:mm:ss');
          }
        }
      }
    }
    const tz = Session.getScriptTimeZone();
    const sinceStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
    dbg.getRange(row, 1, 1, 10).setValues([[
      rule.rowIndex,
      query,
      threads.length,
      rule.subjectKeyword || '',
      sampleSub,
      sampleFrom,
      sampleDate,
      '',
      sinceStr,
      matchedCount
    ]]);
    row++;
  });
}

function buildQueryString(rule, sinceDate, labelName) {
  const terms = [];
  const yyyy = sinceDate.getFullYear();
  const mm = ('0' + (sinceDate.getMonth() + 1)).slice(-2);
  const dd = ('0' + sinceDate.getDate()).slice(-2);
  terms.push('after:' + yyyy + '/' + mm + '/' + dd);
  const sub = rule.subjectKeyword;
  const from = rule.fromAddress;
  // 件名はアプリ側で部分一致判定するため、検索では使わない
  if (from) terms.push('from:' + from);
  if (rule.excludeKeyword) {
    // debug側は単一文字列をそのまま使う（検索用）
    terms.push('-subject:"' + String(rule.excludeKeyword).replace(/"/g, '\\"') + '"');
  }
  return terms.join(' ');
}

function coerceSinceDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number' && !isNaN(value)) {
    const d = new Date(value);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  if (typeof value === 'string') {
    const s = normalizeDateString(value);
    let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      const da = parseInt(m[3], 10);
      const d = new Date(y, mo, da);
      if (!isNaN(d)) return d;
    }
    m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      const da = parseInt(m[3], 10);
      const d = new Date(y, mo, da);
      if (!isNaN(d)) return d;
    }
    const d2 = new Date(s);
    if (!isNaN(d2)) return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  }
  return null;
}

function normalizeDateString(str) {
  const s = String(str || '').trim();
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code >= 0xFF10 && code <= 0xFF19) {
      out += String.fromCharCode(code - 0xFF10 + 0x30);
    } else if (code === 0xFF0F) {
      out += '/';
    } else if (code === 0xFF0D) {
      out += '-';
    } else {
      out += s[i];
    }
  }
  return out.replace(/\./g, '/').replace(/\s.*/, '');
}

function runNotificationsNow() {
  const sys = getSystemSettings();
  const notionToken = sys['NOTION_TOKEN'];
  const parentPageId = getIdFromNotionUrlOrId(sys['DEFAULT_PARENT_PAGE_ID'] || DEFAULT_PARENT_PAGE_ID_HINT);
  const sinceDate = coerceSinceDate(sys['SINCE_DATE']) || coerceSinceDate(DEFAULT_SINCE_DATE_STR);

  if (!notionToken) {
    throw new Error('システム設定の NOTION_TOKEN が未設定です。');
  }
  if (!parentPageId) {
    throw new Error('システム設定の DEFAULT_PARENT_PAGE_ID が不正です。');
  }

  const rules = readRules();
  rules.forEach(rule => {
    // 入力妥当性: 件名orアドレスの少なくとも一方
    if (!rule.subjectKeyword && !rule.fromAddress) return;

    try {
      const dbId = ensureNotionDatabaseForRule(rule, parentPageId, notionToken);
      if (!dbId) throw new Error('Notion DB ID の確定に失敗');

      const threads = searchThreads(rule, sinceDate);
      threads.forEach(thread => {
        const messages = thread.getMessages();
        if (!messages || messages.length === 0) return;
        let anySuccess = false;
        messages.forEach(message => {
          const msgDate = message.getDate();
          if (msgDate < sinceDate) return;
          if (!isRuleMatch(rule, message)) return;
          if (isAlreadyLogged(message.getId())) return;

          const subject = message.getSubject() || '';
          const from = message.getFrom() || '';
          const body = (message.getPlainBody() || '').trim();

          try {
            notionCreatePage(notionToken, dbId, {
              subject: subject,
              receivedAt: message.getDate(),
              from: from,
              body: body
            });
            anySuccess = true;
            appendLog({
              when: new Date(), receivedAt: message.getDate(), subject: subject, from: from,
              messageId: message.getId(), dbId: dbId, status: 'OK', error: ''
            });
          } catch (e) {
            appendLog({
              when: new Date(), receivedAt: message.getDate(), subject: subject, from: from,
              messageId: message.getId(), dbId: dbId, status: 'ERROR', error: String(e && e.message || e)
            });
          }
        });
        // ラベル付与は廃止
      });
    } catch (e) {
      appendLog({
        when: new Date(), receivedAt: null, subject: '', from: '',
        messageId: '', dbId: rule.notionDbId || '', status: 'ERROR', error: String(e && e.message || e)
      });
    }
  });
}

function getSystemSettings() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SYS);
  if (!sh) return {};
  const last = sh.getLastRow();
  const obj = {};
  if (last >= 1) {
    const values = sh.getRange(1, 1, Math.max(3, last), 2).getValues();
    values.forEach(([k, v]) => {
      if (k) obj[String(k).trim()] = String(v || '').trim();
    });
  }
  return obj;
}

function readRules() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RULES);
  if (!sh) return [];
  // 既存プロジェクトで旧フォーマットの場合は列見出しを補完
  const headerCols = sh.getLastColumn();
  if (headerCols < 6) {
    sh.getRange(1, 6).setValue('除外キーワード（任意）');
  }
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 2 + 1, 6).getValues();
  const rules = [];
  rows.forEach((r, i) => {
    const [subjectKeyword, fromAddress, labelName, notionDbId, note, excludeKeyword] = r.map(v => String(v || '').trim());
    rules.push({
      rowIndex: i + 2,
      subjectKeyword, fromAddress, labelName, notionDbId, note, excludeKeyword
    });
  });
  return rules;
}

function searchThreads(rule, sinceDate) {
  const terms = [];
  // 受信日時: after:YYYY/MM/DD
  const yyyy = sinceDate.getFullYear();
  const mm = ('0' + (sinceDate.getMonth() + 1)).slice(-2);
  const dd = ('0' + sinceDate.getDate()).slice(-2);
  terms.push('after:' + yyyy + '/' + mm + '/' + dd);
  // 件名/アドレス
  const sub = rule.subjectKeyword;
  const from = rule.fromAddress;
  // 件名はプログラム側で部分一致判定するためGmail検索には入れない
  if (from) terms.push('from:' + from);
  if (rule.excludeKeyword) {
    terms.push('-subject:"' + String(rule.excludeKeyword).replace(/"/g, '\\"') + '"');
  }
  const query = terms.join(' ');
  // 最大件数は適度に抑制
  return GmailApp.search(query, 0, 100);
}

function getTargetMessage(thread) {
  const messages = thread.getMessages();
  if (!messages || messages.length === 0) return null;
  // スレッド単位処理: 最新メッセージを採用
  return messages[messages.length - 1];
}

function isAlreadyLogged(messageId) {
  if (!messageId) return false;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) return false;
  const last = sh.getLastRow();
  if (last < 2) return false;
  // メッセージID列は5列目
  const ids = sh.getRange(2, 5, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if ((ids[i][0] + '') === messageId) return true;
  }
  return false;
}

function isRuleMatch(rule, message) {
  const subjectKeyword = (rule.subjectKeyword || '').trim();
  const fromAddressRule = (rule.fromAddress || '').trim().toLowerCase();
  const excludeKeyword = (rule.excludeKeyword || '').trim();

  const subject = message.getSubject() || '';
  const fromRaw = message.getFrom() || '';
  const fromEmail = extractEmailAddress(fromRaw);

  let subjectOk = true;
  let fromOk = true;

  if (subjectKeyword) {
    subjectOk = normalizeForMatch(subject).indexOf(normalizeForMatch(subjectKeyword)) !== -1;
  }
  if (fromAddressRule) {
    fromOk = fromEmail === fromAddressRule;
  }

  if (excludeKeyword) {
    if (normalizeForMatch(subject).indexOf(normalizeForMatch(excludeKeyword)) !== -1) return false;
  }
  if (subjectKeyword && fromAddressRule) return subjectOk && fromOk; // AND
  if (subjectKeyword) return subjectOk; // ORの片側
  if (fromAddressRule) return fromOk; // ORの片側
  return false;
}

function extractEmailAddress(fromField) {
  const s = String(fromField || '').trim();
  const m = s.match(/<([^>]+)>/);
  const email = (m ? m[1] : s).trim().toLowerCase();
  return email;
}

function normalizeForMatch(text) {
  let s = String(text || '').toLowerCase();
  // 記号のゆれを簡易正規化
  const map = {
    '【': '[', '［': '[', '〔': '[', '『': '[', '（': '(', '〈': '(',
    '】': ']', '］': ']', '〕': ']', '』': ']', '）': ')', '〉': ')',
    '　': '', ' ': ''
  };
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    out += (map[ch] !== undefined) ? map[ch] : ch;
  }
  return out;
}

function addLabelToThread(thread, labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  thread.addLabel(label);
}

function appendLog(entry) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) return;
  const tz = Session.getScriptTimeZone();
  const whenStr = Utilities.formatDate(entry.when, tz, 'yyyy/MM/dd HH:mm:ss');
  const recvStr = entry.receivedAt ? Utilities.formatDate(entry.receivedAt, tz, 'yyyy/MM/dd HH:mm:ss') : '';
  sh.appendRow([
    whenStr,
    recvStr,
    entry.subject || '',
    entry.from || '',
    entry.messageId || '',
    entry.dbId || '',
    entry.status || '',
    entry.error || ''
  ]);
}

function ensureNotionDatabaseForRule(rule, defaultParentPageId, token) {
  if (rule.notionDbId) return rule.notionDbId;
  const dbName = rule.labelName ? ('Gmail通知: ' + rule.labelName) : 'Gmail通知DB';
  const created = notionCreateDatabase(token, defaultParentPageId, dbName);
  const dbId = created && created.id;
  if (dbId) writeDbIdBack(rule.rowIndex, dbId);
  return dbId;
}

function writeDbIdBack(rowIndex, dbId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RULES);
  if (!sh) return;
  // 4列目が Notion DB ID
  sh.getRange(rowIndex, 4).setValue(dbId);
}

function notionHeaders(token) {
  return {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json; charset=utf-8',
    'Notion-Version': NOTION_VERSION
  };
}

function notionCreateDatabase(token, parentPageId, dbName) {
  const url = NOTION_API_BASE + '/databases';
  const payload = {
    parent: { page_id: getIdFromNotionUrlOrId(parentPageId) },
    title: [ { type: 'text', text: { content: dbName } } ],
    properties: {
      '件名': { title: {} },
      '受信日時': { date: {} },
      '送信者': { rich_text: {} },
      '本文': { rich_text: {} }
    }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: notionHeaders(token),
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    return JSON.parse(res.getContentText());
  }
  throw new Error('Notion DB作成失敗: HTTP ' + code + ' ' + res.getContentText());
}

function notionCreatePage(token, databaseId, data) {
  const url = NOTION_API_BASE + '/pages';
  const properties = {
    '件名': {
      title: [{ text: { content: data.subject || '' } }]
    },
    '受信日時': {
      date: { start: (data.receivedAt || new Date()).toISOString() }
    },
    '送信者': {
      rich_text: [{ text: { content: data.from || '' } }]
    },
    '本文': {
      rich_text: [{ text: { content: (data.body || '').slice(0, 1800) } }]
    }
  };
  const payload = {
    parent: { database_id: getIdFromNotionUrlOrId(databaseId) },
    properties: properties
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: notionHeaders(token),
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) return JSON.parse(res.getContentText());
  throw new Error('Notionページ作成失敗: HTTP ' + code + ' ' + res.getContentText());
}

function getIdFromNotionUrlOrId(s) {
  if (!s) return '';
  const str = String(s).trim();
  // URL末尾のIDを抽出（ハイフン有無に対応）
  const m = str.match(/[0-9a-fA-F]{32}/);
  if (m) return m[0];
  return str;
}
