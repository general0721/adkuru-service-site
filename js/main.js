/* 流入経路：?src= / utm_* を覚えておき、予約完了時に集計シートへ送る（予約者には見えない） */
(() => {
  const ENDPOINT = 'https://script.google.com/macros/s/AKfycbzdGEwI_k-jz5YppZ_Zz5zCXF4fTM7EzhMz-xTczQOwq6KXdHp6tYYfN_C0-Llh0c9_gQ/exec'; // GAS ウェブアプリ（AdKuru 予約トラッカー）
  const KEY = 'adk_src';
  const TTL = 90 * 864e5;
  const read = () => { try { const v = JSON.parse(localStorage.getItem(KEY) || 'null'); return v && Date.now() - v.at < TTL ? v : null; } catch (e) { return null; } };
  const write = (v) => { try { localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) {} };
  const q = new URLSearchParams(location.search);
  const pick = (k) => (q.get(k) || '').slice(0, 120);
  const now = new Date().toISOString();
  const prev = read();
  const hit = pick('src') || pick('utm_source') || pick('utm_campaign');
  if (hit) {
    write({ src: pick('src') || [pick('utm_source'), pick('utm_campaign')].filter(Boolean).join('/'), utm_source: pick('utm_source'), utm_medium: pick('utm_medium'), utm_campaign: pick('utm_campaign'), utm_content: pick('utm_content'), landing: (() => { try { return decodeURIComponent(location.pathname + location.search); } catch (e) { return location.pathname + location.search; } })(), referrer: document.referrer, first_at: prev ? prev.first_at : now, at: Date.now() });
  } else if (!prev) {
    let ref = '';
    try { ref = document.referrer ? new URL(document.referrer).hostname : ''; } catch (e) {}
    const self = ref && ref === location.hostname;
    write({ src: self ? '直接' : (ref ? 'referral:' + ref : '直接'), landing: location.pathname, referrer: document.referrer, first_at: now, at: Date.now() });
  }
  let sid = '';
  try { sid = sessionStorage.getItem('adk_sid') || Math.random().toString(36).slice(2, 10); sessionStorage.setItem('adk_sid', sid); } catch (e) {}
  window.adkTrackBooking = () => {
    if (!ENDPOINT) return;
    const v = read() || { src: '不明' };
    const body = JSON.stringify(Object.assign({}, v, { sid: sid, page: location.pathname }));
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

  // embed.js は window.TimerexCalendar を定義するだけなので、開いたときに呼ぶ
  const initCalendar = () => {
    if (inited) return;
    if (typeof window.TimerexCalendar !== 'function') {
      if (tries++ < 40) setTimeout(initCalendar, 250);
      return;
    }
    try { window.TimerexCalendar({ onBookingComplete: () => window.adkTrackBooking && window.adkTrackBooking() }); inited = true; } catch (e) { /* 予備ボタンを出したまま */ }
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
