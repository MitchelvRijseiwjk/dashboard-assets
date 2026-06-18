/*!
 * HugoReveal — branded iris transition (SuperOffice / Hugo)
 * Vanilla JS, no dependencies. Injects its own styles on first use.
 *
 * API
 *   HugoReveal.show(caption?)   -> opens the iris, shows Hugo (loading state)
 *   HugoReveal.reveal()         -> spins Hugo + closes the iris; returns a Promise
 *                                  that resolves once the overlay is gone
 *   HugoReveal.play(caption?)   -> show() then auto reveal() (handy for demos)
 *   HugoReveal.hide()           -> remove the overlay instantly
 *
 * Typical use around a loading screen:
 *   HugoReveal.show('Laden…');
 *   await loadYourData();
 *   await HugoReveal.reveal();   // content underneath is now visible
 *
 * Theming (set these CSS variables anywhere, e.g. on :root):
 *   --hgx-green     background / iris colour   (default #06423e)
 *   --hgx-dune      Hugo colour                (default #f2efea)
 *   --hgx-owl-size  Hugo width                 (default 120px)
 *   --hgx-z         overlay z-index            (default 99999)
 *
 * Respects prefers-reduced-motion (falls back to a simple fade).
 */
(function (global) {
  'use strict';

  var STYLE_ID = 'hgx-styles';
  var OWL =
    '<div class="hgx-spin"><svg class="hgx-owl idle" viewBox="0 0 426.6 321.2" aria-hidden="true">' +
    '<path d="M230.65,167.36s25.26-117.5-80.98-142.88C101.36,12.94,105.42,36.65,0,0c0,0,29.53,65.43,81.81,76.96,10.36,2.28,21.77,2.57,33.45,2.57h.45c1.78,0,3.56,0,5.34-.01,1.81-.01,3.61-.02,5.43-.02h.48c2.27,0,4.54,0,6.8.04.17,0,.32,0,.48,0,4.6.08,9.18.25,13.68.63.14.01.28.02.41.03,2.25.19,4.47.44,6.67.74,0,0,.03,0,.06.01v.08c.19,0,.39.01.58,0,9.82,1.44,19.12,4.19,27.38,9.37,16.95,13.04,27.82,38.23,26.58,66.67-.59,13.38-3.78,25.79-8.86,36.4.67-4.16,1.03-8.45,1.03-12.79,0-15.88-4.62-30.9-13-42.28-8.84-11.98-20.77-18.57-33.6-18.57s-24.78,6.59-33.61,18.57c-8.38,11.38-13,26.39-13,42.28,0,6.47.77,12.78,2.25,18.77-7.13-13.03-11.04-29.46-10.27-47.14,1.33-30.35,16.05-55.81,35.95-66.25-3.17-.08-6.37-.13-9.6-.13-1.91,0-3.83,0-5.74.02-1.92,0-3.83.01-5.74.01-6.51,0-13.63-.08-20.72-.68-11.91,16.53-19.58,37.95-20.66,61.61-2.49,54.45,30.88,100.07,74.52,101.89,39.86,1.66,74.5-33.79,82.08-81.44l.02-.02ZM194.54,184.78c-.46,10.57-3.16,20.32-7.42,28.45-9.86,10-22.17,15.73-35.28,15.16-10.79-.47-20.66-5.17-28.81-12.84-5.15-9.46-7.98-21.38-7.42-34.22.71-16.27,6.71-30.61,15.6-40.05,8.46,8.56,18.37,18.17,18.44,16.3.08-1.71-1.47-14.66-2.94-26.52,3.43-1.13,7.02-1.67,10.7-1.51,21.79.95,38.42,25.68,37.12,55.23"/>' +
    '<path d="M289.4,204.59c-5.22-8.92-13.57-19.14-22.43-28.25-8.73-8.99-17.95-16.9-25.08-21.4-3.46,34.3-18.31,63.87-39.23,82.33,3.58,5.85,10.98,17.83,17.34,27.24,14.82,21.9,41.04,56.7,41.04,56.7,27.76-31.48,41.93-86.94,31.44-110.67-.83-1.89-1.88-3.89-3.08-5.95"/>' +
    '<path d="M312.7,80.17c2.31-.46,4.72-.79,7.27-.95,1-.06,1.96-.11,2.89-.14,16.81-.67,24.57,2.39,36.35.07,6.34-1.24,13.85-4.04,24.57-9.83,31.75-17.13,42.82-57.77,42.82-57.77,0,0-28.88,13.93-50.6,16.83-29.06,3.9-54.84-5.17-89.93,18.75-37.12,25.3-35.16,95.98-34.77,104.91,3.18,2.36,6.57,5.13,10.03,8.21,15.26,13.55,31.94,32.87,38.24,47.15.53,1.2,1,2.47,1.43,3.79,1.34,4.19,2.16,8.93,2.47,14.08,5.39,1.65,11.01,2.52,16.79,2.45,37.3-.48,66.94-38.98,66.2-86.01-.35-22.54-7.63-42.93-19.22-58.02-1.17.38-2.28.69-3.37.98-5.24,1.37-9.69,1.83-14.01,1.83-3.08,0-6.06-.22-9.19-.45,14.47,10.59,24.22,32.28,23.86,57.25-.19,13.69-3.4,26.33-8.68,36.67.92-3.69,1.52-7.55,1.81-11.53.13-1.73.22-3.48.22-5.25,0-30.92-19.85-56.08-44.26-56.08-15.05,0-28.37,9.59-36.37,24.2,1.86-16.37,8.21-30.62,17.15-40.1,5.36-5.48,11.58-9.34,18.28-11.04M287.83,131.7c6.51,6.65,15.57,15.63,15.65,14.11.07-1.55-1.68-15.63-2.9-25.29,4.98-2.7,10.4-4.04,16.02-3.67,20.57,1.36,35.75,25.08,33.9,52.97-.42,6.28-1.68,12.23-3.59,17.68-7.64,9.1-17.45,14.57-28.1,14.48-20.38-.18-37.15-20.65-41.19-47.93,1.99-8.59,5.54-16.26,10.21-22.35"/>' +
    '</svg></div>';

  var CSS = [
    '.hgx{position:fixed;inset:0;z-index:var(--hgx-z,99999);overflow:hidden}',
    '.hgx[hidden]{display:none}',
    '.hgx-disc{position:absolute;top:50%;left:50%;width:155vmax;height:155vmax;',
      'margin:-77.5vmax 0 0 -77.5vmax;border-radius:50%;background:var(--hgx-green,#06423e);',
      'transform-origin:center;will-change:transform}',
    '.hgx-stage{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);z-index:1;',
      'display:flex;flex-direction:column;align-items:center;gap:22px}',
    '.hgx-spin{display:flex}',
    '.hgx-owl{width:var(--hgx-owl-size,120px);height:auto;display:block}',
    '.hgx-owl path{fill:var(--hgx-dune,#f2efea)}',
    '.hgx-cap{color:var(--hgx-dune,#f2efea);opacity:.72;font:560 14px/1.4 ui-sans-serif,system-ui,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;transition:opacity .35s ease}',

    '.hgx-owl.idle{animation:hgxIdle 2s ease-in-out infinite}',
    '@keyframes hgxIdle{0%,100%{transform:scale(.97)}50%{transform:scale(1)}}',

    /* entrance: iris opens from the centre, Hugo settles in */
    '.hgx.enter .hgx-disc{animation:hgxDiscIn .5s cubic-bezier(.3,0,.18,1) forwards}',
    '@keyframes hgxDiscIn{from{transform:scale(0)}to{transform:scale(1)}}',
    '.hgx.enter .hgx-owl{animation:hgxOwlIn .52s cubic-bezier(.22,.7,.3,1) .14s both}',
    '@keyframes hgxOwlIn{from{opacity:0;transform:scale(.7)}to{opacity:1;transform:scale(1)}}',

    /* exit: rotate on the wrapper, then fade + shrink on the owl, iris closes */
    '.hgx.run .hgx-disc{animation:hgxDisc .72s cubic-bezier(.42,0,.3,1) .38s forwards}',
    '@keyframes hgxDisc{from{transform:scale(1)}to{transform:scale(0)}}',
    '.hgx.run .hgx-spin{animation:hgxSpin .66s cubic-bezier(.32,0,.26,1) forwards}',
    '@keyframes hgxSpin{from{transform:rotate(0)}to{transform:rotate(360deg)}}',
    '.hgx.run .hgx-owl{animation:hgxFadeOut .5s ease .44s forwards, hgxShrink .66s cubic-bezier(.3,0,.26,1) .44s forwards}',
    '@keyframes hgxFadeOut{from{opacity:1}to{opacity:0}}',
    '@keyframes hgxShrink{from{transform:scale(1)}to{transform:scale(.04)}}',

    '@media (prefers-reduced-motion: reduce){',
      '.hgx.enter .hgx-disc,.hgx.enter .hgx-owl{animation:none}',
      '.hgx.run{animation:hgxFade .45s ease forwards}',
      '@keyframes hgxFade{to{opacity:0;visibility:hidden}}',
      '.hgx.run .hgx-disc,.hgx.run .hgx-spin,.hgx.run .hgx-owl,.hgx-owl.idle{animation:none}',
    '}'
  ].join('');

  var el = null;

  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    var s = document.createElement('style');
    s.id = STYLE_ID;
    s.textContent = CSS;
    document.head.appendChild(s);
  }

  function build(caption) {
    injectStyles();
    el = document.createElement('div');
    el.className = 'hgx';
    el.innerHTML = '<div class="hgx-disc"></div><div class="hgx-stage">' + OWL +
      '<div class="hgx-cap">' + (caption || '') + '</div></div>';
    document.body.appendChild(el);
    return el;
  }

  function show(caption) {
    if (!el) build(caption);
    el.hidden = false;
    el.classList.remove('run', 'enter');
    var owl = el.querySelector('.hgx-owl');
    owl.classList.remove('run', 'idle');
    var cap = el.querySelector('.hgx-cap');
    if (caption !== undefined) cap.textContent = caption;
    cap.style.opacity = '0';
    void el.offsetWidth;                       // restart animations cleanly
    el.classList.add('enter');                 // iris opens in
    requestAnimationFrame(function () { cap.style.opacity = '1'; });
    el._idleTimer = setTimeout(function () { owl.classList.add('idle'); }, 680);
    return el;
  }

  function reveal() {
    if (!el) show();
    return new Promise(function (resolve) {
      var owl = el.querySelector('.hgx-owl');
      var disc = el.querySelector('.hgx-disc');
      var cap = el.querySelector('.hgx-cap');
      owl.classList.remove('idle');
      el.classList.remove('enter');
      clearTimeout(el._idleTimer);
      void owl.offsetWidth;
      cap.style.opacity = '0';
      el.classList.add('run');
      var finished = false;
      var done = function () {
        if (finished) return;
        finished = true;
        el.hidden = true;
        disc.removeEventListener('animationend', onEnd);
        el.removeEventListener('animationend', onEnd);
        resolve();
      };
      var onEnd = function (e) {
        if (e.animationName === 'hgxDisc' || e.animationName === 'hgxFade') done();
      };
      disc.addEventListener('animationend', onEnd);
      el.addEventListener('animationend', onEnd);
      setTimeout(done, 2200);                  // safety fallback
    });
  }

  function hide() {
    if (el) el.hidden = true;
  }

  function play(caption) {
    show(caption);
    return new Promise(function (resolve) {
      setTimeout(function () { reveal().then(resolve); }, 820);
    });
  }

  var api = { show: show, reveal: reveal, hide: hide, play: play };

  if (typeof module === 'object' && module.exports) module.exports = api;
  else global.HugoReveal = api;

})(typeof window !== 'undefined' ? window : this);
