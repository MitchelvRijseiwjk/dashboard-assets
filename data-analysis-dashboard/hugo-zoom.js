/*!
 * HugoZoom — branded "fly into Hugo" transition (SuperOffice / Hugo)
 * Vanilla JS, no dependencies. Injects its own styles on first use.
 *
 * On entrance Hugo zooms in from large and settles; on completion Hugo spins,
 * then scales up toward the viewer and fades as the green dissolves to content
 * (as if the camera flies through Hugo into the page).
 *
 * API
 *   HugoZoom.show(caption?)   -> green cover + Hugo (loading state)
 *   HugoZoom.reveal()         -> spin + zoom-through; returns a Promise that
 *                                resolves once the overlay is gone
 *   HugoZoom.play(caption?)   -> show() then auto reveal() (handy for demos)
 *   HugoZoom.hide()           -> remove the overlay instantly
 *
 * Typical use:
 *   HugoZoom.show('Laden…');
 *   await loadYourData();
 *   await HugoZoom.reveal();
 *
 * Theming (shared with HugoReveal — set anywhere, e.g. :root):
 *   --hgx-green  cover colour   (#06423e)   --hgx-dune  Hugo colour (#f2efea)
 *   --hgx-owl-size  Hugo width  (120px)     --hgx-z     z-index     (99999)
 *
 * Respects prefers-reduced-motion (falls back to a simple fade).
 */
(function (global) {
  'use strict';

  var STYLE_ID = 'hgz-styles';
  var OWL =
    '<div class="hgz-spin"><svg class="hgz-owl idle" viewBox="0 0 426.6 321.2" aria-hidden="true">' +
    '<path d="M230.65,167.36s25.26-117.5-80.98-142.88C101.36,12.94,105.42,36.65,0,0c0,0,29.53,65.43,81.81,76.96,10.36,2.28,21.77,2.57,33.45,2.57h.45c1.78,0,3.56,0,5.34-.01,1.81-.01,3.61-.02,5.43-.02h.48c2.27,0,4.54,0,6.8.04.17,0,.32,0,.48,0,4.6.08,9.18.25,13.68.63.14.01.28.02.41.03,2.25.19,4.47.44,6.67.74,0,0,.03,0,.06.01v.08c.19,0,.39.01.58,0,9.82,1.44,19.12,4.19,27.38,9.37,16.95,13.04,27.82,38.23,26.58,66.67-.59,13.38-3.78,25.79-8.86,36.4.67-4.16,1.03-8.45,1.03-12.79,0-15.88-4.62-30.9-13-42.28-8.84-11.98-20.77-18.57-33.6-18.57s-24.78,6.59-33.61,18.57c-8.38,11.38-13,26.39-13,42.28,0,6.47.77,12.78,2.25,18.77-7.13-13.03-11.04-29.46-10.27-47.14,1.33-30.35,16.05-55.81,35.95-66.25-3.17-.08-6.37-.13-9.6-.13-1.91,0-3.83,0-5.74.02-1.92,0-3.83.01-5.74.01-6.51,0-13.63-.08-20.72-.68-11.91,16.53-19.58,37.95-20.66,61.61-2.49,54.45,30.88,100.07,74.52,101.89,39.86,1.66,74.5-33.79,82.08-81.44l.02-.02ZM194.54,184.78c-.46,10.57-3.16,20.32-7.42,28.45-9.86,10-22.17,15.73-35.28,15.16-10.79-.47-20.66-5.17-28.81-12.84-5.15-9.46-7.98-21.38-7.42-34.22.71-16.27,6.71-30.61,15.6-40.05,8.46,8.56,18.37,18.17,18.44,16.3.08-1.71-1.47-14.66-2.94-26.52,3.43-1.13,7.02-1.67,10.7-1.51,21.79.95,38.42,25.68,37.12,55.23"/>' +
    '<path d="M289.4,204.59c-5.22-8.92-13.57-19.14-22.43-28.25-8.73-8.99-17.95-16.9-25.08-21.4-3.46,34.3-18.31,63.87-39.23,82.33,3.58,5.85,10.98,17.83,17.34,27.24,14.82,21.9,41.04,56.7,41.04,56.7,27.76-31.48,41.93-86.94,31.44-110.67-.83-1.89-1.88-3.89-3.08-5.95"/>' +
    '<path d="M312.7,80.17c2.31-.46,4.72-.79,7.27-.95,1-.06,1.96-.11,2.89-.14,16.81-.67,24.57,2.39,36.35.07,6.34-1.24,13.85-4.04,24.57-9.83,31.75-17.13,42.82-57.77,42.82-57.77,0,0-28.88,13.93-50.6,16.83-29.06,3.9-54.84-5.17-89.93,18.75-37.12,25.3-35.16,95.98-34.77,104.91,3.18,2.36,6.57,5.13,10.03,8.21,15.26,13.55,31.94,32.87,38.24,47.15.53,1.2,1,2.47,1.43,3.79,1.34,4.19,2.16,8.93,2.47,14.08,5.39,1.65,11.01,2.52,16.79,2.45,37.3-.48,66.94-38.98,66.2-86.01-.35-22.54-7.63-42.93-19.22-58.02-1.17.38-2.28.69-3.37.98-5.24,1.37-9.69,1.83-14.01,1.83-3.08,0-6.06-.22-9.19-.45,14.47,10.59,24.22,32.28,23.86,57.25-.19,13.69-3.4,26.33-8.68,36.67.92-3.69,1.52-7.55,1.81-11.53.13-1.73.22-3.48.22-5.25,0-30.92-19.85-56.08-44.26-56.08-15.05,0-28.37,9.59-36.37,24.2,1.86-16.37,8.21-30.62,17.15-40.1,5.36-5.48,11.58-9.34,18.28-11.04M287.83,131.7c6.51,6.65,15.57,15.63,15.65,14.11.07-1.55-1.68-15.63-2.9-25.29,4.98-2.7,10.4-4.04,16.02-3.67,20.57,1.36,35.75,25.08,33.9,52.97-.42,6.28-1.68,12.23-3.59,17.68-7.64,9.1-17.45,14.57-28.1,14.48-20.38-.18-37.15-20.65-41.19-47.93,1.99-8.59,5.54-16.26,10.21-22.35"/>' +
    '</svg></div>';

  var CSS = [
    '.hgz{position:fixed;inset:0;z-index:var(--hgx-z,99999);overflow:hidden}',
    '.hgz[hidden]{display:none}',
    '.hgz-cover{position:absolute;inset:0;background:var(--hgx-green,#06423e);transform-origin:center;will-change:transform,opacity}',
    '.hgz-stage{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);z-index:1;',
      'display:flex;flex-direction:column;align-items:center;gap:22px}',
    '.hgz-spin{display:flex}',
    '.hgz-owl{width:var(--hgx-owl-size,120px);height:auto;display:block}',
    '.hgz-owl path{fill:var(--hgx-dune,#f2efea)}',
    '.hgz-cap{color:var(--hgx-dune,#f2efea);opacity:.72;font:560 14px/1.4 ui-sans-serif,system-ui,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;transition:opacity .35s ease}',

    '.hgz-owl.idle{animation:hgzIdle 2s ease-in-out infinite}',
    '@keyframes hgzIdle{0%,100%{transform:scale(.97)}50%{transform:scale(1)}}',

    /* entrance: green fades in, Hugo zooms in from large and settles */
    '.hgz.enter .hgz-cover{animation:hgzCoverIn .42s ease forwards}',
    '@keyframes hgzCoverIn{from{opacity:0}to{opacity:1}}',
    '.hgz.enter .hgz-owl{animation:hgzOwlIn .55s cubic-bezier(.2,.7,.3,1) .1s both}',
    '@keyframes hgzOwlIn{from{opacity:0;transform:scale(2.6)}to{opacity:1;transform:scale(1)}}',

    /* exit: spin, then fly through Hugo while the green dissolves */
    '.hgz.run .hgz-spin{animation:hgzSpin .62s cubic-bezier(.32,0,.26,1) forwards}',
    '@keyframes hgzSpin{from{transform:rotate(0)}to{transform:rotate(360deg)}}',
    '.hgz.run .hgz-owl{animation:hgzFade .5s ease .42s forwards, hgzZoom .64s cubic-bezier(.5,0,.82,.4) .42s forwards}',
    '@keyframes hgzFade{from{opacity:1}to{opacity:0}}',
    '@keyframes hgzZoom{from{transform:scale(1)}to{transform:scale(13)}}',
    '.hgz.run .hgz-cover{animation:hgzCoverOut .5s ease .56s forwards}',
    '@keyframes hgzCoverOut{from{opacity:1;transform:scale(1)}to{opacity:0;transform:scale(1.18)}}',

    '@media (prefers-reduced-motion: reduce){',
      '.hgz.enter .hgz-cover,.hgz.enter .hgz-owl{animation:none}',
      '.hgz.run .hgz-spin,.hgz.run .hgz-owl,.hgz-owl.idle{animation:none}',
      '.hgz.run .hgz-cover{animation:hgzCoverOut .45s ease forwards}',
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
    el.className = 'hgz';
    el.innerHTML = '<div class="hgz-cover"></div><div class="hgz-stage">' + OWL +
      '<div class="hgz-cap">' + (caption || '') + '</div></div>';
    document.body.appendChild(el);
    return el;
  }

  function show(caption) {
    if (!el) build(caption);
    el.hidden = false;
    el.classList.remove('run', 'enter');
    var owl = el.querySelector('.hgz-owl');
    owl.classList.remove('run', 'idle');
    var cap = el.querySelector('.hgz-cap');
    if (caption !== undefined) cap.textContent = caption;
    cap.style.opacity = '0';
    void el.offsetWidth;
    el.classList.add('enter');
    requestAnimationFrame(function () { cap.style.opacity = '1'; });
    el._idleTimer = setTimeout(function () { owl.classList.add('idle'); }, 700);
    return el;
  }

  function reveal() {
    if (!el) show();
    return new Promise(function (resolve) {
      var owl = el.querySelector('.hgz-owl');
      var cover = el.querySelector('.hgz-cover');
      var cap = el.querySelector('.hgz-cap');
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
        cover.removeEventListener('animationend', onEnd);
        resolve();
      };
      var onEnd = function (e) {
        if (e.animationName === 'hgzCoverOut') done();
      };
      cover.addEventListener('animationend', onEnd);
      setTimeout(done, 2200);
    });
  }

  function hide() { if (el) el.hidden = true; }

  function play(caption) {
    show(caption);
    return new Promise(function (resolve) {
      setTimeout(function () { reveal().then(resolve); }, 820);
    });
  }

  var api = { show: show, reveal: reveal, hide: hide, play: play };

  if (typeof module === 'object' && module.exports) module.exports = api;
  else global.HugoZoom = api;

})(typeof window !== 'undefined' ? window : this);
