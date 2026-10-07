/**
 * AdKuru 予約の流入経路トラッカー
 *
 * - サイト（予約完了時）から doPost で「流入経路ログ」に1行追加
 * - 15分おきの syncBookings_ が予約カレンダーを読み、予定の作成時刻と
 *   ログの受信時刻を突き合わせて「予約一覧」を更新する
 *
 * 初回だけ setup() を実行 → ウェブアプリとしてデプロイ（実行ユーザー：自分／アクセス：全員）
 */

const ADK = {
  TAB_LOG: '流入経路ログ',
  TAB_BOOK: '予約一覧',
  TAB_SUM: '流入経路別',
  TAB_CONF: '設定',
  MATCH_MINUTES: 10,       // 予約完了ログと予定作成時刻の許容差
  LOOKBACK_DAYS: 60,       // 予定作成日がこの日数以内のものを対象にする
};

const ADK_LOG_HEAD = ['受信日時', '流入経路', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', '最初に来たページ', 'リファラ', '初回訪問日時', 'セッションID', '照合済み'];
const ADK_BOOK_HEAD = ['予約受付日時', '面談日時', '名前', '会社名', 'メールアドレス', '流入経路', 'utm_campaign', 'utm_content', '照合', '予定ID'];

/* ---------- 初期設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const conf = adkSheet_(ss, ADK.TAB_CONF);
  if (conf.getLastRow() === 0) {
    conf.getRange(1, 1, 4, 2).setValues([
      ['項目', '値'],
      ['予約カレンダーID', ''],
      ['予定タイトルに含まれる文字', 'SNS採用広告無料相談'],
      ['サイトURL', 'https://general0721.github.io/adkuru-service-site/'],
    ]);
    conf.getRange('A1:B1').setFontWeight('bold').setBackground('#e8ebf3');
    conf.getRange('B2').setBackground('#fff6cf').setNote('予約が入るGoogleカレンダーのID（主催者のメールアドレス）を入れてください');
    conf.setColumnWidth(1, 220).setColumnWidth(2, 420);
  }
  const log = adkSheet_(ss, ADK.TAB_LOG);
  if (log.getLastRow() === 0) adkHead_(log, ADK_LOG_HEAD);
  const book = adkSheet_(ss, ADK.TAB_BOOK);
  if (book.getLastRow() === 0) adkHead_(book, ADK_BOOK_HEAD);
  const sum = adkSheet_(ss, ADK.TAB_SUM);
  if (sum.getLastRow() === 0) {
    sum.getRange('A1').setValue('流入経路ごとの予約数（予約一覧から自動集計）').setFontWeight('bold');
    sum.getRange('A3').setFormula("=IFERROR(QUERY('予約一覧'!A2:H,\"select F, count(A) where A is not null group by F order by count(A) desc label F '流入経路', count(A) '予約数'\",0),\"まだ予約がありません\")");
    sum.setColumnWidth(1, 260);
  }
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  ss.setActiveSheet(book);

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncBookings_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncBookings_').timeBased().everyMinutes(15).create();
  return 'ok';
}

function adkSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function adkHead_(sh, head) {
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#254793').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, head.length);
}

/* ---------- サイトからの受信 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sh = adkSheet_(SpreadsheetApp.getActive(), ADK.TAB_LOG);
    const s = v => String(v == null ? '' : v).slice(0, 300);
    sh.appendRow([new Date(), s(d.src || '不明'), s(d.utm_source), s(d.utm_medium), s(d.utm_campaign), s(d.utm_content), s(d.landing), s(d.referrer), s(d.first_at), s(d.sid), '']);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function doGet() { return ContentService.createTextOutput('AdKuru booking tracker'); }

/* ---------- カレンダー照合 ---------- */
function syncBookings_() {
  const ss = SpreadsheetApp.getActive();
  const conf = ss.getSheetByName(ADK.TAB_CONF);
  const calId = String(conf.getRange('B2').getValue()).trim();
  const keyword = String(conf.getRange('B3').getValue()).trim();
  if (!calId) return '設定!B2 にカレンダーIDがありません';
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) return 'カレンダーが見つかりません：' + calId;

  const now = new Date();
  const from = new Date(now.getTime() - ADK.LOOKBACK_DAYS * 864e5);
  const to = new Date(now.getTime() + 180 * 864e5);
  const events = cal.getEvents(from, to).filter(ev => !keyword || ev.getTitle().indexOf(keyword) >= 0 || (ev.getDescription() || '').indexOf('TimeRex') >= 0)
    .filter(ev => ev.getDateCreated() >= from);

  const book = ss.getSheetByName(ADK.TAB_BOOK);
  const bookVals = book.getLastRow() > 1 ? book.getRange(2, 1, book.getLastRow() - 1, ADK_BOOK_HEAD.length).getValues() : [];
  const known = new Set(bookVals.map(r => r[9]));

  const log = ss.getSheetByName(ADK.TAB_LOG);
  const logVals = log.getLastRow() > 1 ? log.getRange(2, 1, log.getLastRow() - 1, ADK_LOG_HEAD.length).getValues() : [];

  const rows = [];
  events.sort((a, b) => a.getDateCreated() - b.getDateCreated()).forEach(ev => {
    const id = ev.getId() + '|' + ev.getStartTime().getTime();
    if (known.has(id)) return;
    const created = ev.getDateCreated();
    // 未照合のログのうち、作成時刻に一番近いもの
    let best = -1, bestDiff = ADK.MATCH_MINUTES * 60000 + 1;
    logVals.forEach((r, i) => {
      if (r[10]) return;
      const diff = Math.abs(new Date(r[0]).getTime() - created.getTime());
      if (diff < bestDiff) { best = i; bestDiff = diff; }
    });
    const g = adkGuest_(ev);
    let src = '不明（サイト外から予約）', camp = '', cont = '', state = '未照合';
    if (best >= 0) {
      const r = logVals[best];
      src = r[1]; camp = r[4]; cont = r[5]; state = '照合済み';
      r[10] = id;
      log.getRange(best + 2, 11).setValue(id);
    }
    rows.push([created, ev.getStartTime(), g.name, g.company, g.email, src, camp, cont, state, id]);
  });
  if (rows.length) {
    book.getRange(book.getLastRow() + 1, 1, rows.length, ADK_BOOK_HEAD.length).setValues(rows);
    book.getRange(2, 1, book.getLastRow() - 1, 2).setNumberFormat('yyyy/mm/dd hh:mm');
  }
  return rows.length + '件追加';
}

/** 予定の説明文・ゲストから 名前 / 会社名 / メール を拾う */
function adkGuest_(ev) {
  const desc = (ev.getDescription() || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
  const pick = labels => {
    for (const l of labels) {
      const m = desc.match(new RegExp(l + '\\s*[:：]\\s*(.+)'));
      if (m) return m[1].trim();
    }
    return '';
  };
  let email = pick(['メールアドレス', 'Email', 'E-mail']);
  if (!email) {
    const guest = ev.getGuestList().find(x => !/lic-inc\.co\.jp$/i.test(x.getEmail()));
    if (guest) email = guest.getEmail();
  }
  return { name: pick(['名前', 'お名前', 'Name']), company: pick(['会社名', 'Company']), email: email };
}

/** 手動実行用：照合をすぐ回す */
function main() { return syncBookings_(); }
