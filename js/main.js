/* 流入経路・アクセス計測：?src= / utm_* を覚えておき、閲覧・離脱・相談ボタン・予約完了を集計シートへ送る（見ている人には見えない）
   - bot・自動ブラウザは送らない。社内の確認は一度 ?staff=1 を付けて開くと、そのブラウザは以後計測しない（?staff=0 で解除） */
(() => {
  const ENDPOINT = 'https://script.google.com/macros/s/AKfycbzdGEwI_k-jz5YppZ_Zz5zCXF4fTM7EzhMz-xTczQOwq6KXdHp6tYYfN_C0-Llh0c9_gQ/exec'; // GAS ウェブアプリ（AdKuru 予約トラッカー）
  const KEY = 'adk_src';
  const TTL = 90 * 864e5;
  const ls = { get: (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} }, del: (k) => { try { localStorage.removeItem(k); } catch (e) {} } };
  const ss = { get: (k) => { try { return sessionStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) {} } };
  const read = () => { try { const v = JSON.parse(ls.get(KEY) || 'null'); return v && Date.now() - v.at < TTL ? v : null; } catch (e) { return null; } };
  const write = (v) => ls.set(KEY, JSON.stringify(v));
  const rid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  const q = new URLSearchParams(location.search);
  const pick = (k) => (q.get(k) || '').slice(0, 120);
  const now = new Date().toISOString();

  // 今回の流入元（URLのパラメータ → 参照元の順で判定）
  let refHost = '';
  try { refHost = document.referrer ? new URL(document.referrer).hostname : ''; } catch (e) {}
  const fromRef = (h) => {
    if (!h) return '直接';
    if (h === location.hostname) return '直接';
    if (/(^|\.)google\./.test(h)) return 'Google検索';
    if (/(^|\.)yahoo\./.test(h)) return 'Yahoo!検索';
    if (/(^|\.)bing\.com$/.test(h)) return 'Bing検索';
    if (/instagram\.com$/.test(h)) return 'Instagram';
    if (/facebook\.com$|fb\.com$/.test(h)) return 'Facebook';
    if (/tiktok\.com$/.test(h)) return 'TikTok';
    if (/^t\.co$|twitter\.com$|x\.com$/.test(h)) return 'X';
    if (/line\.me$|line-apps\.com$/.test(h)) return 'LINE';
    if (/lic-inc\.co\.jp$/.test(h)) return 'Lic.サイト・Journal';
    return 'referral:' + h;
  };
  const curSrc = pick('src') || (pick('utm_source') ? [pick('utm_source'), pick('utm_campaign')].filter(Boolean).join('/') : '') || (q.get('gclid') || q.get('gbraid') || q.get('wbraid') ? 'Google広告' : '') || (q.get('fbclid') ? fromRef(refHost) + '（広告・リンク）' : '') || fromRef(refHost);

  // 最初の流入元（90日保持）。パラメータ付きで来たら上書き
  const prev = read();
  const hit = pick('src') || pick('utm_source') || pick('utm_campaign') || q.get('gclid');
  if (hit) {
    write({ src: curSrc, utm_source: pick('utm_source'), utm_medium: pick('utm_medium'), utm_campaign: pick('utm_campaign'), utm_content: pick('utm_content'), landing: (() => { try { return decodeURIComponent(location.pathname + location.search); } catch (e) { return location.pathname + location.search; } })(), referrer: document.referrer, first_at: prev ? prev.first_at : now, at: Date.now() });
  } else if (!prev) {
    write({ src: curSrc, landing: location.pathname, referrer: document.referrer, first_at: now, at: Date.now() });
  }
  const first = read() || { src: curSrc };

  // 訪問者ID（ブラウザ単位）・セッションID（タブを閉じるまで）・訪問回数
  let vid = ls.get('adk_vid');
  if (!vid) { vid = rid(); ls.set('adk_vid', vid); }
  let sid = ss.get('adk_sid');
  if (!sid) {
    sid = rid(); ss.set('adk_sid', sid);
    ls.set('adk_visits', String((+ls.get('adk_visits') || 0) + 1));
  }
  const visits = +ls.get('adk_visits') || 1;

  // 計測しない：社内フラグ・bot・自動ブラウザ
  if (q.get('staff') === '1') ls.set('adk_staff', '1');
  if (q.get('staff') === '0') ls.del('adk_staff');
  const ua = navigator.userAgent || '';
  const isBot = navigator.webdriver || /bot|crawl|spider|slurp|headless|lighthouse|pagespeed|preview|facebookexternalhit|embedly|monitor|uptime|python|curl|wget|phantom|puppeteer|playwright/i.test(ua);
  const track = !!ENDPOINT && !isBot && ls.get('adk_staff') !== '1';
  const device = /iPad|Tablet/i.test(ua) ? 'タブレット' : (/Mobi|Android|iPhone/i.test(ua) ? 'スマホ' : 'PC');

  const send = (t, extra) => {
    if (!track) return;
    const body = JSON.stringify(Object.assign({
      t: t, path: location.pathname, src: curSrc, first_src: first.src || '', ref: document.referrer.slice(0, 300),
      utm_source: pick('utm_source'), utm_medium: pick('utm_medium'), utm_campaign: pick('utm_campaign'), utm_content: pick('utm_content'),
      device: device, vid: vid, sid: sid, visits: visits, ua: ua.slice(0, 250),
    }, extra || {}));
    try {
      if (t === 'leave' && navigator.sendBeacon) { navigator.sendBeacon(ENDPOINT, new Blob([body], { type: 'text/plain;charset=utf-8' })); return; }
      fetch(ENDPOINT, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: body, keepalive: true });
    } catch (e) {}
  };

  // どこまで読んだか（セクション単位）と最大スクロール率
  const SECTIONS = [['fv', 'ファーストビュー'], ['service', 'サービス'], ['worry', 'お悩み'], ['reason', '解決策'], ['strength', '3つの強み'], ['case', '実績'], ['flow', '導入の流れ'], ['price', '料金'], ['faq', 'よくある質問'], ['contact', '無料相談']];
  let deepest = 0, maxScroll = 0, cta = 0;
  const els = SECTIONS.map(([id]) => id === 'fv' ? document.querySelector('.fv') : document.getElementById(id));
  const measure = () => {
    const vh = window.innerHeight, doc = document.documentElement;
    const h = Math.max(doc.scrollHeight - vh, 1);
    maxScroll = Math.max(maxScroll, Math.min(100, Math.round(window.scrollY / h * 100)));
    els.forEach((el, i) => { if (el && i > deepest && el.getBoundingClientRect().top < vh * 0.6) deepest = i; });
  };
  window.addEventListener('scroll', measure, { passive: true });
  const t0 = Date.now();
  let left = false, hiddenAt = 0, hiddenMs = 0;
  const leave = () => {
    if (left) return;
    left = true;
    measure();
    const sec = Math.round((Date.now() - t0 - hiddenMs) / 1000);
    send('leave', { sec: Math.min(sec, 3600), scroll: maxScroll, section: els.some(Boolean) ? SECTIONS[deepest][1] : '', section_no: deepest + 1, cta: cta });
  };
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') { hiddenAt = Date.now(); leave(); }
    else if (hiddenAt) { hiddenMs += Date.now() - hiddenAt; hiddenAt = 0; left = false; }
  });
  window.addEventListener('pagehide', leave);

  send('pv', { title: document.title.slice(0, 120) });
  setTimeout(measure, 500);

  // 相談ボタン（カレンダーを開いた）
  window.adkTrackCta = (label) => { cta++; send('cta', { label: (label || '').slice(0, 60), section: els.some(Boolean) ? SECTIONS[deepest][1] : '' }); };

  window.adkTrackBooking = () => {
    if (!ENDPOINT) return;
    const v = read() || { src: '不明' };
    const body = JSON.stringify(Object.assign({}, v, { sid: sid, vid: vid, page: location.pathname }));
    try { fetch(ENDPOINT, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: body, keepalive: true }); } catch (e) {}
  };
})();

/* AdKuru サービスサイト */
(() => {
  const header = document.getElementById('header');
  const spbar = document.querySelector('.spbar');
  const fv = document.querySelector('.fv');
  const contact = document.getElementById('contact');

  // ヘッダーの影／スマホ固定CTA（FVを過ぎたら出し、最終フォームでは隠す）
  const onScroll = () => {
    const y = window.scrollY;
    header && header.classList.toggle('is-scrolled', y > 10);
    if (spbar && fv) {
      const pastFv = y > fv.offsetTop + fv.offsetHeight - 200;
      const atForm = contact && contact.getBoundingClientRect().top < window.innerHeight * 0.7;
      spbar.classList.toggle('is-show', pastFv && !atForm);
    }
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();

  // 今月の特別料金枠（data-taken を毎月書き換える）
  document.querySelectorAll('.slots').forEach((ul) => {
    const total = +ul.dataset.total || 3;
    const taken = Math.min(+ul.dataset.taken || 0, total);
    ul.innerHTML = Array.from({ length: total }, (_, i) =>
      `<li class="${i < taken ? 'is-taken' : ''}">${i < taken ? '済' : '空'}</li>`).join('');
    document.querySelectorAll('.js-left').forEach((el) => { el.textContent = total - taken; });
  });

  // フォーム（2ステップ）
  document.querySelectorAll('.lead-form').forEach((form) => {
    const steps = form.querySelectorAll('.lead-form__step');
    const marks = form.querySelectorAll('.steps li');
    const go = (n) => {
      steps.forEach((s) => { s.hidden = s.dataset.step !== String(n); });
      marks.forEach((m, i) => m.classList.toggle('is-on', i === n - 1));
    };

    form.querySelectorAll('.chips').forEach((group) => {
      group.addEventListener('click', (e) => {
        const btn = e.target.closest('button');
        if (!btn) return;
        group.querySelectorAll('button').forEach((b) => b.classList.toggle('is-on', b === btn));
        form.querySelector(`input[name="${group.dataset.name}"]`).value = btn.textContent;
        group.classList.remove('is-err');
      });
    });

    const check = (step) => {
      let ok = true;
      step.querySelectorAll('.chips').forEach((g) => {
        const filled = !!form.querySelector(`input[name="${g.dataset.name}"]`).value;
        g.classList.toggle('is-err', !filled);
        ok = ok && filled;
      });
      step.querySelectorAll('input[type="text"], input[type="tel"], input[type="email"]').forEach((i) => {
        let filled = i.value.trim() !== '';
        if (filled && i.type === 'email') filled = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(i.value.trim());
        i.classList.toggle('is-err', !filled);
        ok = ok && filled;
      });
      const agree = step.querySelector('input[name="agree"]');
      if (agree && !agree.checked) ok = false;
      step.querySelector('.lead-form__err').hidden = ok;
      return ok;
    };

    form.querySelector('.js-next').addEventListener('click', () => {
      if (check(steps[0])) go(2);
    });
    form.querySelector('.js-back').addEventListener('click', () => go(1));
    form.addEventListener('input', (e) => e.target.classList && e.target.classList.remove('is-err'));

    // 送信先は未設定。action が空のあいだは完了ページへ移るだけ
    form.addEventListener('submit', (e) => {
      if (!check(steps[1])) { e.preventDefault(); return; }
      if (!form.getAttribute('action')) {
        e.preventDefault();
        location.href = 'thanks/';
      }
    });
  });
})();

/* 予約ポップアップ（TimeRex のカレンダーを埋め込み） */
(() => {
  const modal = document.getElementById('bookmodal');
  if (!modal || !modal.showModal) return; // 非対応ブラウザは href どおり予約ページへ
  let inited = false, tries = 0;

  // Google 広告のコンバージョン（同じページでは各1回まで）
  const sentCv = {};
  const adsCv = (sendTo, withValue) => {
    if (!sendTo || sentCv[sendTo] || typeof window.gtag !== 'function') return;
    sentCv[sendTo] = true;
    try { window.gtag('event', 'conversion', withValue ? { send_to: sendTo, value: 1.0, currency: 'JPY' } : { send_to: sendTo }); } catch (e) {}
  };

  // embed.js は window.TimerexCalendar を定義するだけなので、開いたときに呼ぶ
  const initCalendar = () => {
    if (inited) return;
    if (typeof window.TimerexCalendar !== 'function') {
      if (tries++ < 40) setTimeout(initCalendar, 250);
      return;
    }
    try { window.TimerexCalendar({ onBookingComplete: () => { window.adkTrackBooking && window.adkTrackBooking(); adsCv(window.ADKURU_CV_DONE, true); } }); inited = true; } catch (e) { /* 予備ボタンを出したまま */ }
  };

  // カレンダーが出るまでは「予約ページを開く」を見せる
  const watchFallback = () => {
    const cal = document.getElementById('timerex_calendar');
    const fb = document.getElementById('book-fallback');
    if (!cal || !fb) return;
    const check = () => { fb.hidden = cal.children.length > 0; return fb.hidden; };
    if (check()) return;
    const t = setInterval(() => { if (check()) clearInterval(t); }, 400);
    setTimeout(() => clearInterval(t), 15000);
  };

  document.addEventListener('click', (e) => {
    const b = e.target.closest('[data-book]');
    if (!b) return;
    e.preventDefault();
    modal.showModal();
    adsCv(window.ADKURU_CV);
    window.adkTrackCta && window.adkTrackCta((b.textContent || '').trim() || b.className);
    document.body.style.overflow = 'hidden';
    initCalendar();
    watchFallback();
  });
  modal.addEventListener('close', () => { document.body.style.overflow = ''; });
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.close(); });
  document.getElementById('modalclose').addEventListener('click', () => modal.close());
})();

/* 導入の流れ（flow）：画面に入ったら左から順に表示 */
(() => {
  const list = document.querySelector('.flow');
  if (!list || !('IntersectionObserver' in window)) return;
  document.documentElement.classList.add('js-anim');
  list.querySelectorAll('li').forEach((li, i) => li.style.setProperty('--i', i));
  const io = new IntersectionObserver((entries) => {
    entries.forEach((e) => { if (e.isIntersecting) { list.classList.add('is-in'); io.disconnect(); } });
  }, { threshold: 0.4, rootMargin: '0px 0px -15% 0px' });
  io.observe(list);
})();
