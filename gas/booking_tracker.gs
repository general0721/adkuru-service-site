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
const ADK_BOOK_HEAD = ['予約を受けた日時', '面談日時', '会社名', '名前', '経路名', '媒体', 'メールアドレス', '流入コード', 'utm_campaign', 'utm_content', '照合', '予定ID（システム用）'];
// E・F列（経路名・媒体）は数式で自動表示。スクリプトは A〜D と G〜L に書く

/* ---------- 初期設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const conf = adkSheet_(ss, ADK.TAB_CONF);
  if (conf.getLastRow() === 0) {
    conf.getRange(1, 1, 4, 2).setValues([
      ['項目', '値'],
      ['予約カレンダーID', ''],
      ['予約の目印（説明文に含まれる文字）', 'timerex.net'],
      ['サイトURL', 'https://recruit.adkuru.com/'],
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

/** A列に値がある最後の行（ARRAYFORMULA の空文字を数えない） */
function adkLastRow_(sh) {
  const v = sh.getRange('A1:A').getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1;
  return 1;
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
    if (d.t === 'pv' || d.t === 'leave' || d.t === 'cta') { adkAccess_(d); return ContentService.createTextOutput('ok'); }
    const sh = adkSheet_(SpreadsheetApp.getActive(), ADK.TAB_LOG);
    const s = v => String(v == null ? '' : v).slice(0, 300);
    sh.appendRow([new Date(), s(d.src || '不明'), s(d.utm_source), s(d.utm_medium), s(d.utm_campaign), s(d.utm_content), s(d.landing), s(d.referrer), s(d.first_at), s(d.sid), '', s(d.vid, 40)]);
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
  // 集計開始日時（これより前に作られた予約は対象外）
  if (!conf.getRange('A5').getValue()) conf.getRange('A5:B5').setValues([['集計開始日時', new Date()]]);
  const since = new Date(conf.getRange('B5').getValue());
  // 共有されたカレンダーは「参加」していないと開けないので、未参加なら参加する（表示はオフ）
  let cal = CalendarApp.getCalendarById(calId);
  if (!cal) {
    try { cal = CalendarApp.subscribeToCalendar(calId, { selected: false }); } catch (e) { return 'カレンダーに参加できません：' + calId + ' ' + e; }
  }
  if (!cal) return 'カレンダーが見つかりません：' + calId;

  const now = new Date();
  const from = new Date(now.getTime() - ADK.LOOKBACK_DAYS * 864e5);
  const to = new Date(now.getTime() + 180 * 864e5);
  const events = cal.getEvents(from, to).filter(ev => !keyword || ev.getTitle().indexOf(keyword) >= 0 || (ev.getDescription() || '').indexOf(keyword) >= 0)
    .filter(ev => ev.getDateCreated() >= from && ev.getDateCreated() >= since);

  const book = ss.getSheetByName(ADK.TAB_BOOK);
  const bookLast = adkLastRow_(book);
  const bookVals = bookLast > 1 ? book.getRange(2, 1, bookLast - 1, ADK_BOOK_HEAD.length).getValues() : [];
  const known = new Set(bookVals.map(r => r[11]));

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
    let src = '不明（サイトを通らない予約）', camp = '', cont = '', state = '未照合';
    if (best >= 0) {
      const r = logVals[best];
      src = r[1]; camp = r[4]; cont = r[5]; state = '照合済み';
      r[10] = id;
      log.getRange(best + 2, 11).setValue(id);
    }
    rows.push([created, ev.getStartTime(), g.company, g.name, null, null, g.email, src, camp, cont, state, id]);
  });
  if (rows.length) {
    book.getRange(bookLast + 1, 1, rows.length, 4).setValues(rows.map(x => x.slice(0, 4)));
    book.getRange(bookLast + 1, 7, rows.length, 6).setValues(rows.map(x => x.slice(6)));
    book.getRange(2, 1, bookLast - 1 + rows.length, 2).setNumberFormat('yyyy/mm/dd (ddd) hh:mm');
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
  let email = pick(['ゲストメールアドレス', 'メールアドレス', 'Email', 'E-mail']);
  if (!email) {
    const guest = ev.getGuestList().find(x => !/lic-inc\.co\.jp$/i.test(x.getEmail()));
    if (guest) email = guest.getEmail();
  }
  return { name: pick(['ゲスト氏名', '氏名', 'お名前', '名前', 'Name']), company: pick(['ゲスト会社名', '会社名', 'Company']), email: email };
}

/** 手動実行用：照合をすぐ回す */
function main() { return syncBookings_(); }

/* ---------- 見やすく整える（何度実行してもOK） ---------- */
const ADK_TAB_HOW = '見方';
const ADK_TAB_LINK = 'リンク作成';
const ADK_C = { navy: '#254793', ink: '#102a56', pale: '#e8eef9', band: '#f5f8ff', yellow: '#fff6cf', gray: '#9aa3b5', grayBg: '#f1f3f6', green: '#dff3ea', greenInk: '#1d7a52' };

function beautify() {
  const ss = SpreadsheetApp.getActive();
  const conf = ss.getSheetByName(ADK.TAB_CONF);
  const base = String(conf.getRange('B4').getValue() || 'https://recruit.adkuru.com/');

  /* リンク作成：広告・投稿ごとにコードを登録すると配信用URLができる */
  const link = adkSheet_(ss, ADK_TAB_LINK);
  if (link.getLastRow() === 0) {
    link.getRange('A1:E1').setValues([['経路名（シートに表示する名前）', '媒体', 'コード（英数字・記号 _ - のみ）', '配信用URL（自動）', 'メモ']]);
    link.getRange('A2:E4').setValues([
      ['Meta広告 A', 'Meta広告', 'meta_ad_A', '', '例：記入例。不要なら行ごと書き換えてください'],
      ['TikTok広告', 'TikTok広告', 'tiktok_ad', '', ''],
      ['Instagram投稿 10/7', 'Instagram投稿', 'insta_post_1007', '', ''],
    ]);
  }
  link.getRange('D2').setFormula('=ARRAYFORMULA(IF(C2:C="","","' + base + '?src="&C2:C))');
  adkStyleHead_(link, 5);
  link.getRange('A2:C').setBackground(ADK_C.yellow);
  link.getRange('D2:D').setBackground(ADK_C.grayBg).setFontColor(ADK_C.ink);
  link.setColumnWidth(1, 220).setColumnWidth(2, 140).setColumnWidth(3, 200).setColumnWidth(4, 470).setColumnWidth(5, 300);
  const media = SpreadsheetApp.newDataValidation().requireValueInList(['リスティング広告', 'Meta広告', 'TikTok広告', 'X広告', 'LINE広告', 'YouTube広告', 'Instagram投稿', 'TikTok投稿', 'X投稿', 'メール・その他'], true).setAllowInvalid(true).build();
  link.getRange('B2:B200').setDataValidation(media);
  link.setTabColor(ADK_C.navy);

  /* 予約一覧：見出し＋経路名・媒体を自動表示（E・F列） */
  const book = ss.getSheetByName(ADK.TAB_BOOK);
  book.getRange(1, 1, 1, ADK_BOOK_HEAD.length).setValues([ADK_BOOK_HEAD]);
  book.getRange('E1').setFormula('={"経路名";ARRAYFORMULA(IF(H2:H="","",IFERROR(VLOOKUP(H2:H,{\'' + ADK_TAB_LINK + '\'!C2:C,\'' + ADK_TAB_LINK + '\'!A2:A},2,FALSE),H2:H)))}');
  book.getRange('F1').setFormula('={"媒体";ARRAYFORMULA(IF(H2:H="","",IF(LEFT(H2:H,2)="不明","ー",IFERROR(VLOOKUP(H2:H,{\'' + ADK_TAB_LINK + '\'!C2:C,\'' + ADK_TAB_LINK + '\'!B2:B},2,FALSE),"未登録"))))}');
  adkStyleHead_(book, ADK_BOOK_HEAD.length);
  book.getRange('E1:F1').setBackground('#ff7b72');
  book.getRange('A2:B').setNumberFormat('yyyy/mm/dd (ddd) hh:mm');
  book.getRange('E2:F').setFontWeight('bold').setFontColor(ADK_C.ink);
  book.getRange('H2:L').setFontColor(ADK_C.gray);
  book.getRange('H1:L1').setBackground(ADK_C.gray);
  [165, 165, 200, 130, 180, 120, 230, 130, 110, 110, 90, 150].forEach((w, i) => book.setColumnWidth(i + 1, w));
  adkBand_(book, ADK_BOOK_HEAD.length);
  book.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('照合済み').setBackground(ADK_C.green).setFontColor(ADK_C.greenInk).setRanges([book.getRange('K2:K')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('不明').setFontColor(ADK_C.gray).setRanges([book.getRange('E2:F')]).build(),
  ]);
  book.setTabColor('#ff7b72');

  /* 流入経路別：経路名ごと・媒体ごと・月ごと */
  const sum = ss.getSheetByName(ADK.TAB_SUM);
  sum.clear();
  sum.getCharts().forEach(c => sum.removeChart(c));
  sum.getRange('A1').setValue('流入経路ごとの予約数').setFontSize(14).setFontWeight('bold').setFontColor(ADK_C.ink);
  sum.getRange('A2').setValue('「予約一覧」から自動で集計しています。経路名は「リンク作成」で登録した名前で表示されます。').setFontColor(ADK_C.gray);
  sum.getRange('A4:B4').setValues([['経路名', '予約数']]);
  sum.getRange('A5').setFormula("=IFERROR(QUERY('" + ADK.TAB_BOOK + "'!E2:E,\"select E, count(E) where E is not null and E <> '' group by E order by count(E) desc label count(E) ''\",0),\"まだ予約がありません\")");
  sum.getRange('D4:E4').setValues([['媒体', '予約数']]);
  sum.getRange('D5').setFormula("=IFERROR(QUERY('" + ADK.TAB_BOOK + "'!F2:F,\"select F, count(F) where F is not null and F <> '' group by F order by count(F) desc label count(F) ''\",0),\"まだ予約がありません\")");
  sum.getRange('G4:H4').setValues([['月', '予約数']]);
  sum.getRange('G5').setFormula("=IFERROR(QUERY({ARRAYFORMULA(IF('" + ADK.TAB_BOOK + "'!A2:A=\"\",\"\",TEXT('" + ADK.TAB_BOOK + "'!A2:A,\"yyyy年m月\")))},\"select Col1, count(Col1) where Col1 <> '' group by Col1 order by Col1 label count(Col1) ''\",0),\"まだ予約がありません\")");
  [sum.getRange('A4:B4'), sum.getRange('D4:E4'), sum.getRange('G4:H4')].forEach(r => r.setFontWeight('bold').setBackground(ADK_C.navy).setFontColor('#ffffff'));
  sum.setColumnWidth(1, 220).setColumnWidth(2, 80).setColumnWidth(3, 30).setColumnWidth(4, 160).setColumnWidth(5, 80).setColumnWidth(6, 30).setColumnWidth(7, 120).setColumnWidth(8, 80);
  sum.getRange('B5:B').setFontWeight('bold'); sum.getRange('E5:E').setFontWeight('bold'); sum.getRange('H5:H').setFontWeight('bold');
  const chart = sum.newChart().asBarChart().addRange(sum.getRange('A4:B30')).setNumHeaders(1)
    .setOption('title', '経路名ごとの予約数').setOption('legend', { position: 'none' }).setOption('colors', [ADK_C.navy])
    .setPosition(4, 10, 0, 0).setOption('width', 520).setOption('height', 320).build();
  sum.insertChart(chart);
  sum.setTabColor('#f7ce0f');

  /* 流入経路ログ・設定：システム用とわかるように */
  const log = ss.getSheetByName(ADK.TAB_LOG);
  log.getRange(1, 1, 1, ADK_LOG_HEAD.length).setBackground(ADK_C.gray).setFontColor('#ffffff');
  log.setTabColor(ADK_C.gray);
  conf.setTabColor(ADK_C.gray);
  conf.getRange('B2:B3').setBackground(ADK_C.yellow);

  /* 見方 */
  const how = adkSheet_(ss, ADK_TAB_HOW);
  how.clear();
  const lines = [
    ['このシートの見方', ''],
    ['', ''],
    ['見るところ', ''],
    ['予約一覧', '予約が入ると15分以内に1行追加されます。「経路名」「媒体」でどこから来た予約かがわかります。'],
    ['流入経路別', '経路名ごと・媒体ごと・月ごとの予約数をまとめています。'],
    ['', ''],
    ['広告・投稿を出すとき', ''],
    ['1. リンク作成', '黄色の欄に「経路名」「媒体」「コード」を1行ずつ入れます。コードは英数字と _ - だけ。'],
    ['2. URLをコピー', '「配信用URL（自動）」にできたURLを、その広告・投稿のリンク先に設定します。'],
    ['3. あとは自動', 'そのURLから来て予約した人は、予約一覧にその経路名で表示されます。'],
    ['', ''],
    ['補足', ''],
    ['不明（サイトを通らない予約）', 'TimeRexのURLを直接開いて予約した人です。サイトを通っていないので経路がわかりません。'],
    ['媒体が「未登録」', 'コードが「リンク作成」に登録されていません。登録すると経路名・媒体が表示されます。'],
    ['灰色のタブ', '流入経路ログ・設定はシステム用です。編集しないでください。'],
  ];
  how.getRange(1, 1, lines.length, 2).setValues(lines);
  how.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor(ADK_C.ink);
  ['A3', 'A7', 'A12'].forEach(a => how.getRange(a + ':' + a.replace('A', 'B')).setBackground(ADK_C.navy).setFontColor('#ffffff').setFontWeight('bold'));
  how.getRange('A4:A15').setFontWeight('bold').setFontColor(ADK_C.ink);
  how.setColumnWidth(1, 240).setColumnWidth(2, 640);
  how.getRange('A1:B15').setVerticalAlignment('middle').setWrap(true);
  how.setRowHeights(4, 12, 28);
  how.setTabColor(ADK_C.ink);

  /* タブの並び */
  [ADK_TAB_HOW, ADK.TAB_BOOK, ADK.TAB_SUM, ADK_TAB_LINK, ADK.TAB_LOG, ADK.TAB_CONF].forEach((n, i) => {
    ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(how);
  ss.getSheets().forEach(sh => sh.getDataRange().setFontFamily('Arial'));
  return 'ok';
}

function adkStyleHead_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground(ADK_C.navy).setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
}

function adkBand_(sh, n) {
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 200), n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(ADK_C.navy).setFirstRowColor('#ffffff').setSecondRowColor(ADK_C.band);
}


/* =====================================================================
 * アクセス計測（2026-11-05 追加）
 * - サイトの閲覧・離脱（滞在秒・どこまで読んだか）・相談ボタンを「アクセスログ」に1行ずつ追加
 * - 1時間ごとの summarizeAccess_ が「アクセス_日別」「アクセス_流入元」「アクセス_読了」「アクセス_訪問者」を作り直す
 * - 初回だけ setupAccess() を実行。doPost を変えたので「デプロイを管理」で新バージョンにする
 * ===================================================================== */
const ACC = {
  TAB_LOG: 'アクセスログ',
  TAB_DAY: 'アクセス_日別',
  TAB_SRC: 'アクセス_流入元',
  TAB_READ: 'アクセス_読了',
  TAB_VIS: 'アクセス_訪問者',
  VISITOR_ROWS: 1000,
  BOT_UA: /bot|crawl|spider|slurp|headless|lighthouse|pagespeed|preview|facebookexternalhit|embedly|monitor|uptime|python|curl|wget|phantom|puppeteer|playwright/i,
};
const ACC_SECTIONS = ['ファーストビュー', 'サービス', 'お悩み', '解決策', '3つの強み', '実績', '導入の流れ', '料金', 'よくある質問', '無料相談'];
const ACC_LOG_HEAD = ['日時', '種類', 'ページ', '流入元（今回）', '流入元（初回）', '参照元URL', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', '端末', '滞在秒', 'スクロール%', 'どこまで読んだか', 'セクション番号', 'ボタン', '訪問者ID', 'セッションID', '訪問回数', 'ブラウザ（UA）'];

function setupAccess() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  const log = adkSheet_(ss, ACC.TAB_LOG);
  if (log.getLastRow() === 0) adkHead_(log, ACC_LOG_HEAD);
  log.setTabColor(ADK_C.gray);
  // 予約ログに訪問者IDの列を足す（予約した人がどのページをどう見たかをつなぐため）
  const blog = ss.getSheetByName(ADK.TAB_LOG);
  if (blog && !blog.getRange('L1').getValue()) blog.getRange('L1').setValue('訪問者ID').setFontWeight('bold').setBackground(ADK_C.gray).setFontColor('#ffffff');
  [ACC.TAB_DAY, ACC.TAB_SRC, ACC.TAB_READ, ACC.TAB_VIS].forEach(n => adkSheet_(ss, n).setTabColor('#2f9e6e'));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'summarizeAccess_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('summarizeAccess_').timeBased().everyHours(1).create();
  // タブの並び：見方 → アクセス集計 → 予約まわり → システム用
  const order = [ADK_TAB_HOW, ACC.TAB_DAY, ACC.TAB_SRC, ACC.TAB_READ, ACC.TAB_VIS, ADK.TAB_BOOK, ADK.TAB_SUM, ADK_TAB_LINK, ACC.TAB_LOG, ADK.TAB_LOG, ADK.TAB_CONF];
  order.filter(n => ss.getSheetByName(n)).forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  adkHowAccess_(ss);
  summarizeAccess_();
  ss.setActiveSheet(ss.getSheetByName(ACC.TAB_DAY));
  return 'ok';
}

function adkAccess_(d) {
  const ua = String(d.ua || '');
  if (ACC.BOT_UA.test(ua)) return; // 念のためサーバー側でも bot を落とす
  const s = (v, n) => String(v == null ? '' : v).slice(0, n || 300);
  const kind = { pv: '閲覧', leave: '離脱', cta: '相談ボタン' }[d.t];
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(ACC.TAB_LOG) || (() => { const sh = adkSheet_(ss, ACC.TAB_LOG); adkHead_(sh, ACC_LOG_HEAD); return sh; })();
  log.appendRow([
    new Date(), kind, s(d.path, 120), s(d.src, 120), s(d.first_src, 120), s(d.ref), s(d.utm_source, 100), s(d.utm_medium, 100), s(d.utm_campaign, 100), s(d.utm_content, 100),
    s(d.device, 10), d.t === 'leave' ? Number(d.sec) || 0 : '', d.t === 'leave' ? Number(d.scroll) || 0 : '', s(d.section, 30), d.t === 'leave' ? Number(d.section_no) || '' : '',
    d.t === 'cta' ? s(d.label, 60) : '', s(d.vid, 40), s(d.sid, 40), Number(d.visits) || 1, s(ua, 250),
  ]);
}

/* ---------- 集計（1時間ごと） ---------- */
function summarizeAccess_() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(ACC.TAB_LOG);
  if (!log) return 'アクセスログがありません';
  const n = log.getLastRow() - 1;
  const v = n > 0 ? log.getRange(2, 1, n, ACC_LOG_HEAD.length).getValues().filter(r => r[0] !== '') : [];
  const tz = ss.getSpreadsheetTimeZone();
  const day = d => Utilities.formatDate(new Date(d), tz, 'yyyy/MM/dd');
  const since30 = Date.now() - 30 * 864e5;
  const avg = xs => xs.length ? Math.round(xs.reduce((p, c) => p + c, 0) / xs.length) : '';
  const pct = (a, b) => b ? Math.round(a / b * 1000) / 10 : '';

  // 予約した訪問者（流入経路ログの訪問者ID）と、日別の予約数（予約一覧）
  const blog = ss.getSheetByName(ADK.TAB_LOG);
  const bv = blog && blog.getLastRow() > 1 ? blog.getRange(2, 1, blog.getLastRow() - 1, 12).getValues() : [];
  const bookVid = new Set(bv.map(r => r[11]).filter(Boolean));
  const book = ss.getSheetByName(ADK.TAB_BOOK);
  const bookRows = book ? book.getRange(2, 1, Math.max(adkLastRow_(book) - 1, 1), 1).getValues().filter(r => r[0] !== '') : [];
  const bookDay = {};
  bookRows.forEach(r => { const k = day(r[0]); bookDay[k] = (bookDay[k] || 0) + 1; });

  // セッション（1回の訪問）単位にまとめる
  const S = {};
  v.forEach(r => {
    const sid = r[17] || ('_' + r[16] + day(r[0]));
    const x = S[sid] || (S[sid] = { at: r[0], src: '', vid: r[16], pv: 0, sec: 0, scroll: 0, sec_no: 0, cta: false, device: r[10], visits: Number(r[18]) || 1 });
    if (r[1] === '閲覧') { x.pv++; if (!x.src) x.src = r[3] || '不明'; if (!x.device) x.device = r[10]; }
    else if (r[1] === '離脱') { x.sec = Math.max(x.sec, Number(r[11]) || 0); x.scroll = Math.max(x.scroll, Number(r[12]) || 0); x.sec_no = Math.max(x.sec_no, Number(r[14]) || 0); }
    else if (r[1] === '相談ボタン') x.cta = true;
    if (!x.src && r[3]) x.src = r[3];
  });
  const sessions = Object.keys(S).map(k => S[k]).filter(x => x.pv > 0);
  const firstSeen = {};
  sessions.slice().sort((a, b) => new Date(a.at) - new Date(b.at)).forEach(x => { if (!(x.vid in firstSeen)) firstSeen[x.vid] = day(x.at); });

  // 日別
  const D = {};
  sessions.forEach(x => {
    const k = day(x.at), a = D[k] || (D[k] = { pv: 0, ses: 0, vis: new Set(), nw: new Set(), cta: 0, sec: [], price: 0, measured: 0 });
    a.pv += x.pv; a.ses++; a.vis.add(x.vid); if (firstSeen[x.vid] === k) a.nw.add(x.vid);
    if (x.cta) a.cta++;
    if (x.sec_no) { a.measured++; if (x.sec_no >= 8) a.price++; a.sec.push(x.sec); }
  });
  const drows = Object.keys(D).sort().reverse().map(k => { const a = D[k]; return [k, a.pv, a.ses, a.vis.size, a.nw.size, a.cta, pct(a.cta, a.ses), bookDay[k] || 0, avg(a.sec), pct(a.price, a.measured)]; });
  adkWriteAcc_(ss, ACC.TAB_DAY, '日別の推移（新しい順）。bot・自動ブラウザ・社内（?staff=1 を付けて開いたブラウザ）は除外済み', ['日付', '閲覧数', '訪問数', '訪問者数', 'うち初めての人', '相談ボタンを押した訪問', '相談ボタン率（%）', '予約数', '平均滞在（秒）', '料金まで読んだ割合（%）'], drows);

  // 流入元
  const R = {};
  sessions.forEach(x => {
    const a = R[x.src] || (R[x.src] = { ses: 0, ses30: 0, vis: new Set(), cta: 0, book: new Set(), sec: [], price: 0, measured: 0, sp: 0 });
    a.ses++; if (+new Date(x.at) >= since30) a.ses30++; a.vis.add(x.vid);
    if (x.cta) a.cta++;
    if (bookVid.has(x.vid)) a.book.add(x.vid);
    if (x.device === 'スマホ') a.sp++;
    if (x.sec_no) { a.measured++; if (x.sec_no >= 8) a.price++; a.sec.push(x.sec); }
  });
  const srows = Object.keys(R).map(k => { const a = R[k]; return [k, a.ses, a.ses30, a.vis.size, a.cta, pct(a.cta, a.ses), a.book.size, avg(a.sec), pct(a.price, a.measured), pct(a.sp, a.ses)]; }).sort((x, y) => y[1] - x[1]);
  adkWriteAcc_(ss, ACC.TAB_SRC, '流入元ごとの訪問（1回の訪問を1件。今回どこから来たかで数える）', ['流入元', '訪問数（全期間）', '訪問数（30日）', '訪問者数', '相談ボタンを押した訪問', '相談ボタン率（%）', '予約した人', '平均滞在（秒）', '料金まで読んだ割合（%）', 'スマホの割合（%）'], srows);

  // 読了（どのセクションまで読まれたか）
  const measured = sessions.filter(x => x.sec_no);
  const m30 = measured.filter(x => +new Date(x.at) >= since30);
  const rrows = ACC_SECTIONS.map((name, i) => {
    const all = measured.filter(x => x.sec_no >= i + 1).length, r30 = m30.filter(x => x.sec_no >= i + 1).length;
    const stop = measured.filter(x => x.sec_no === i + 1).length;
    return [i + 1, name, all, pct(all, measured.length), r30, pct(r30, m30.length), stop, pct(stop, measured.length)];
  });
  adkWriteAcc_(ss, ACC.TAB_READ, `トップページをどこまで読んだか（離脱を計測できた訪問 ${measured.length} 件のうち、そのセクションまでスクロールした割合）`, ['順番', 'セクション', '到達した訪問（全期間）', '到達率（%）', '到達した訪問（30日）', '到達率（30日・%）', 'ここで離脱した訪問', '離脱の割合（%）'], rrows);

  // 訪問者
  const P = {};
  sessions.slice().sort((a, b) => new Date(a.at) - new Date(b.at)).forEach(x => {
    const a = P[x.vid] || (P[x.vid] = { first: x.at, last: x.at, fsrc: x.src, lsrc: x.src, visits: 0, pv: 0, deep: 0, cta: 0, sec: 0, device: x.device });
    a.last = x.at; a.lsrc = x.src; a.visits++; a.pv += x.pv; a.deep = Math.max(a.deep, x.sec_no); a.sec += x.sec; if (x.cta) a.cta++;
  });
  const vrows = Object.keys(P).map(k => { const a = P[k]; return [a.last, a.first, a.fsrc, a.lsrc, a.device, a.visits, a.pv, a.sec, a.deep ? ACC_SECTIONS[a.deep - 1] : '', a.cta, bookVid.has(k) ? '予約した' : '', k]; })
    .sort((x, y) => new Date(y[0]) - new Date(x[0])).slice(0, ACC.VISITOR_ROWS);
  adkWriteAcc_(ss, ACC.TAB_VIS, `訪問者（1人＝1行、最近来た順に${ACC.VISITOR_ROWS}人まで。同じブラウザを同じ人とみなす）`, ['最後に来た日時', '初めて来た日時', '流入元（初回）', '流入元（最後）', '端末', '訪問回数', '閲覧ページ数', '合計滞在（秒）', 'いちばん深く読んだところ', '相談ボタンを押した回数', '予約', '訪問者ID'], vrows);
  if (vrows.length) ss.getSheetByName(ACC.TAB_VIS).getRange(3, 1, vrows.length, 2).setNumberFormat('yyyy/mm/dd hh:mm');
  return 'ok';
}

/** 集計タブを丸ごと書き直す（1行目=説明、2行目=見出し、3行目〜） */
function adkWriteAcc_(ss, name, note, head, rows) {
  const sh = adkSheet_(ss, name);
  sh.clear();
  sh.getRange(1, 1).setValue(note + '　最終更新：' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm')).setFontColor(ADK_C.gray);
  sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(ADK_C.navy).setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 150);
  for (let i = 2; i <= head.length; i++) sh.setColumnWidth(i, 120);
  if (rows.length) sh.getRange(3, 1, rows.length, head.length).setValues(rows).setVerticalAlignment('top');
  else sh.getRange(3, 1).setValue('まだデータがありません');
  sh.getDataRange().setFontFamily('Arial');
}

/** 「見方」タブにアクセス計測の説明を足す（何度実行してもOK） */
function adkHowAccess_(ss) {
  const how = ss.getSheetByName(ADK_TAB_HOW);
  if (!how) return;
  const start = 18;
  how.getRange(start, 1, 10, 2).clear();
  const lines = [
    ['アクセス計測（1時間ごとに更新）', ''],
    ['アクセス_日別', '日ごとの閲覧数・訪問者数・相談ボタンを押した数・予約数・平均滞在・料金まで読んだ割合。'],
    ['アクセス_流入元', 'どこから来た訪問か（Google検索・Google広告・Instagram・?src= のコードなど）ごとの数と、相談ボタン率・予約した人。'],
    ['アクセス_読了', 'トップページをどのセクションまで読んだか。離脱の多いセクションがわかります。'],
    ['アクセス_訪問者', '1人＝1行。何回来て、どこまで読んで、相談ボタンを押したか。'],
    ['除外しているもの', 'bot・自動ブラウザ、社内のブラウザ。社内の人は一度 https://recruit.adkuru.com/?staff=1 を開くと、そのブラウザは以後数えません（?staff=0 で解除）。'],
    ['アクセスログ', 'システム用（灰色）。編集しないでください。'],
  ];
  how.getRange(start, 1, lines.length, 2).setValues(lines);
  how.getRange(start, 1, 1, 2).setBackground(ADK_C.navy).setFontColor('#ffffff').setFontWeight('bold');
  how.getRange(start + 1, 1, lines.length - 1, 1).setFontWeight('bold').setFontColor(ADK_C.ink);
  how.getRange(start, 1, lines.length, 2).setVerticalAlignment('middle').setWrap(true);
}
