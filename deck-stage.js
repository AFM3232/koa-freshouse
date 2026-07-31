// deck-stage.js
// A self-contained "PowerPoint-style" slide viewer for <x-import> decks.
// Takes the <section data-label data-screen-label data-speaker-notes> children
// of the host element and turns them into a fullscreen, scaled, navigable deck
// with keyboard/click navigation, smooth transitions, a progress bar and
// optional speaker notes.
(function () {
  'use strict';

  var STYLE_ID = 'deck-stage-styles';

  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    var style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = [
      '.ds-stage{position:fixed;inset:0;background:#05070f;display:flex;align-items:center;justify-content:center;overflow:hidden;z-index:100000;box-sizing:border-box;transition:background .4s ease;}',
      '.ds-canvas{position:relative;flex:none;box-shadow:0 30px 80px rgba(0,0,0,.55);}',
      '.ds-canvas .ds-slide{position:absolute;top:0;left:0;width:100%;height:100%;box-sizing:border-box;opacity:0;visibility:hidden;pointer-events:none;transition:opacity .5s cubic-bezier(.4,0,.2,1),transform .5s cubic-bezier(.4,0,.2,1);will-change:opacity,transform;overflow-y:auto;overflow-x:hidden;-webkit-overflow-scrolling:touch;}',
      '.ds-canvas .ds-slide::-webkit-scrollbar{width:10px;}',
      '.ds-canvas .ds-slide::-webkit-scrollbar-thumb{background:rgba(249,115,22,.55);border-radius:6px;}',
      '.ds-canvas .ds-slide::-webkit-scrollbar-track{background:rgba(0,0,0,.06);}',
      '.ds-canvas .ds-slide.ds-active{opacity:1;visibility:visible;pointer-events:auto;transform:translateX(0);z-index:2;}',
      '.ds-canvas .ds-slide.ds-from-right{transform:translateX(4%);}',
      '.ds-canvas .ds-slide.ds-from-left{transform:translateX(-4%);}',
      '.ds-canvas .ds-slide.ds-to-left{opacity:0;transform:translateX(-4%);}',
      '.ds-canvas .ds-slide.ds-to-right{opacity:0;transform:translateX(4%);}',
      '.ds-canvas .ds-slide.ds-noanim{transition:none!important;}',
      '.ds-progress{position:fixed;left:0;top:0;height:4px;background:#F97316;z-index:100002;transition:width .45s ease;box-shadow:0 0 12px rgba(249,115,22,.6);}',
      '.ds-bar{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);display:flex;align-items:center;gap:10px;background:rgba(10,14,32,.82);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);padding:8px 10px 8px 16px;border-radius:999px;z-index:100002;font-family:Arial,Helvetica,sans-serif;color:#fff;box-shadow:0 10px 30px rgba(0,0,0,.4);opacity:0;transform:translate(-50%,12px);transition:opacity .3s ease,transform .3s ease;}',
      '.ds-stage:hover .ds-bar,.ds-bar.ds-visible{opacity:1;transform:translate(-50%,0);}',
      '.ds-label{font-size:12px;letter-spacing:.03em;color:#C9D0E6;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding-right:4px;}',
      '.ds-counter{font-size:12px;font-variant-numeric:tabular-nums;color:#fff;min-width:52px;text-align:center;}',
      '.ds-btn{background:rgba(255,255,255,.1);border:none;color:#fff;width:34px;height:34px;border-radius:50%;font-size:16px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .2s,transform .15s;flex:none;}',
      '.ds-btn:hover{background:#F97316;}',
      '.ds-btn:active{transform:scale(.92);}',
      '.ds-btn[disabled]{opacity:.3;cursor:default;pointer-events:none;}',
      '.ds-sep{width:1px;height:20px;background:rgba(255,255,255,.15);flex:none;}',
      '.ds-notes{position:fixed;left:24px;bottom:90px;max-width:440px;max-height:40vh;overflow:auto;background:rgba(10,14,32,.92);color:#D9DEEF;font-size:13px;line-height:1.6;font-family:Arial,Helvetica,sans-serif;padding:16px 20px;border-radius:12px;display:none;z-index:100002;box-shadow:0 10px 30px rgba(0,0,0,.4);}',
      '.ds-notes.ds-visible{display:block;}',
      '.ds-notes strong{display:block;color:#FDBA74;font-size:11px;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;}',
      '.ds-exit{position:fixed;right:20px;top:20px;z-index:100002;}',
      '.ds-hidden-host > *:not(.ds-stage):not(.ds-progress){display:none!important;}',
      'img.lb-img{cursor:zoom-in;}',
      '.ds-lightbox{position:fixed;inset:0;background:rgba(5,7,15,.94);z-index:100010;display:none;align-items:center;justify-content:center;padding:56px;box-sizing:border-box;}',
      '.ds-lightbox.ds-visible{display:flex;}',
      '.ds-lightbox img{max-width:100%;max-height:100%;object-fit:contain;border-radius:6px;box-shadow:0 20px 70px rgba(0,0,0,.6);cursor:default;}',
      '.ds-lightbox-close{position:fixed;top:22px;right:26px;width:46px;height:46px;border-radius:50%;background:rgba(255,255,255,.12);border:none;color:#fff;font-size:22px;line-height:1;cursor:pointer;z-index:100011;display:flex;align-items:center;justify-content:center;transition:background .2s;}',
      '.ds-lightbox-close:hover{background:#F97316;}'
    ].join('\n');
    document.head.appendChild(style);
  }

  function mount(host, opts) {
    injectStyles();

    var width = opts.width || 1920;
    var height = opts.height || 1080;

    var slides = Array.prototype.slice.call(host.querySelectorAll(':scope > section'));
    if (!slides.length) return;

    // Build stage DOM
    var stage = document.createElement('div');
    stage.className = 'ds-stage';

    var canvas = document.createElement('div');
    canvas.className = 'ds-canvas';
    canvas.style.width = width + 'px';
    canvas.style.height = height + 'px';

    slides.forEach(function (sec) {
      sec.classList.add('ds-slide', 'ds-noanim');
      canvas.appendChild(sec);
    });

    stage.appendChild(canvas);

    var progress = document.createElement('div');
    progress.className = 'ds-progress';

    var bar = document.createElement('div');
    bar.className = 'ds-bar';
    bar.innerHTML =
      '<button class="ds-btn" data-act="first" title="Primera (Inicio)">«</button>' +
      '<button class="ds-btn" data-act="prev" title="Anterior (←)">‹</button>' +
      '<div class="ds-counter"><span class="ds-cur">1</span>/<span class="ds-tot">' + slides.length + '</span></div>' +
      '<button class="ds-btn" data-act="next" title="Siguiente (→)">›</button>' +
      '<button class="ds-btn" data-act="last" title="Última (Fin)">»</button>' +
      '<div class="ds-sep"></div>' +
      '<div class="ds-label"></div>' +
      '<div class="ds-sep"></div>' +
      '<button class="ds-btn" data-act="notes" title="Notas del orador (N)">📝</button>' +
      '<button class="ds-btn" data-act="fullscreen" title="Pantalla completa (F)">⛶</button>';

    var notesPanel = document.createElement('div');
    notesPanel.className = 'ds-notes';

    var lightbox = document.createElement('div');
    lightbox.className = 'ds-lightbox';
    lightbox.innerHTML = '<button class="ds-lightbox-close" aria-label="Cerrar">✕</button><img alt="">';
    var lightboxImg = lightbox.querySelector('img');

    // Hide any original page chrome/content that isn't part of the deck.
    host.classList.add('ds-hidden-host');

    document.body.appendChild(progress);
    document.body.appendChild(stage);
    document.body.appendChild(bar);
    document.body.appendChild(notesPanel);
    document.body.appendChild(lightbox);

    function openLightbox(src, alt) {
      lightboxImg.src = src;
      lightboxImg.alt = alt || '';
      lightbox.classList.add('ds-visible');
    }
    function closeLightbox() {
      lightbox.classList.remove('ds-visible');
      lightboxImg.src = '';
    }
    lightbox.addEventListener('click', function (e) {
      if (e.target === lightbox || e.target.classList.contains('ds-lightbox-close')) closeLightbox();
    });

    var idx = 0;
    var inFlight = false;
    var notesVisible = false;

    function labelFor(i) {
      var sec = slides[i];
      return sec.getAttribute('data-label') || sec.getAttribute('data-screen-label') || '';
    }
    function notesFor(i) {
      return slides[i].getAttribute('data-speaker-notes') || '';
    }

    function isFullscreen() {
      return !!(document.fullscreenElement || document.webkitFullscreenElement);
    }

    function layout() {
      stage.style.paddingBottom = '0px';
      var availW = window.innerWidth;
      var availH = window.innerHeight;
      // Never crop: always show the whole slide (letterboxed if the
      // window's aspect ratio doesn't match the deck's 16:9 canvas). The
      // letterbox color is kept in sync with each slide's own background
      // (see matchStageBackground) so the gap reads as part of the slide
      // instead of a visible black bar.
      var scale = Math.min(availW / width, availH / height);
      canvas.style.transform = 'scale(' + scale + ')';
    }

    function matchStageBackground() {
      var active = slides[idx];
      if (!active) return;
      var bg = getComputedStyle(active).backgroundColor;
      if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') {
        stage.style.background = bg;
      }
    }

    function updateChrome() {
      bar.querySelector('.ds-cur').textContent = idx + 1;
      bar.querySelector('.ds-label').textContent = labelFor(idx);
      bar.querySelector('[data-act="prev"]').disabled = idx === 0;
      bar.querySelector('[data-act="first"]').disabled = idx === 0;
      bar.querySelector('[data-act="next"]').disabled = idx === slides.length - 1;
      bar.querySelector('[data-act="last"]').disabled = idx === slides.length - 1;
      var pct = slides.length > 1 ? (idx / (slides.length - 1)) * 100 : 100;
      progress.style.width = pct + '%';
      if (notesVisible) {
        var n = notesFor(idx);
        notesPanel.innerHTML = '<strong>Notas</strong>' + (n ? n : '<em>Sin notas para esta diapositiva.</em>');
      }
      matchStageBackground();
    }

    function goTo(target, opt_instant) {
      target = Math.max(0, Math.min(slides.length - 1, target));
      if (target === idx && !opt_instant) return;
      if (inFlight) return;

      var dir = target > idx ? 1 : -1;
      var cur = slides[idx];
      var next = slides[target];

      if (opt_instant) {
        slides.forEach(function (s) { s.classList.remove('ds-active', 'ds-from-left', 'ds-from-right', 'ds-to-left', 'ds-to-right'); });
        next.classList.add('ds-active');
        idx = target;
        updateChrome();
        return;
      }

      inFlight = true;
      next.classList.remove('ds-noanim');
      cur.classList.remove('ds-noanim');
      next.classList.add(dir > 0 ? 'ds-from-right' : 'ds-from-left');
      // force reflow so the "from" position is committed before activating
      // eslint-disable-next-line no-unused-expressions
      next.getBoundingClientRect();

      requestAnimationFrame(function () {
        requestAnimationFrame(function () {
          next.classList.remove('ds-from-right', 'ds-from-left');
          next.classList.add('ds-active');
          cur.classList.remove('ds-active');
          cur.classList.add(dir > 0 ? 'ds-to-left' : 'ds-to-right');
        });
      });

      idx = target;
      updateChrome();

      setTimeout(function () {
        cur.classList.remove('ds-to-left', 'ds-to-right');
        inFlight = false;
      }, 520);
    }

    function next() { goTo(idx + 1); }
    function prev() { goTo(idx - 1); }

    function toggleNotes() {
      notesVisible = !notesVisible;
      notesPanel.classList.toggle('ds-visible', notesVisible);
      updateChrome();
    }

    function toggleFullscreen() {
      if (!document.fullscreenElement) {
        (stage.requestFullscreen || stage.webkitRequestFullscreen || function () {}).call(stage);
      } else {
        (document.exitFullscreen || document.webkitExitFullscreen || function () {}).call(document);
      }
    }

    bar.addEventListener('click', function (e) {
      var btn = e.target.closest('[data-act]');
      if (!btn) return;
      var act = btn.getAttribute('data-act');
      if (act === 'next') next();
      else if (act === 'prev') prev();
      else if (act === 'first') goTo(0);
      else if (act === 'last') goTo(slides.length - 1);
      else if (act === 'notes') toggleNotes();
      else if (act === 'fullscreen') toggleFullscreen();
    });

    document.addEventListener('keydown', function (e) {
      if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable)) return;
      if (e.key === 'Escape' && lightbox.classList.contains('ds-visible')) {
        closeLightbox();
        return;
      }
      if (lightbox.classList.contains('ds-visible')) return;
      switch (e.key) {
        case 'ArrowRight':
        case 'PageDown':
        case ' ':
          e.preventDefault();
          next();
          break;
        case 'ArrowLeft':
        case 'PageUp':
          e.preventDefault();
          prev();
          break;
        case 'Home':
          e.preventDefault();
          goTo(0);
          break;
        case 'End':
          e.preventDefault();
          goTo(slides.length - 1);
          break;
        case 'n':
        case 'N':
          toggleNotes();
          break;
        case 'f':
        case 'F':
          toggleFullscreen();
          break;
      }
    });

    // Click zones on left/right thirds of the canvas to navigate, like a
    // typical slide viewer (ignore clicks on interactive elements/links).
    canvas.addEventListener('click', function (e) {
      var zoomImg = e.target.closest('img.lb-img');
      if (zoomImg) {
        openLightbox(zoomImg.currentSrc || zoomImg.src, zoomImg.alt);
        return;
      }
      if (e.target.closest('a,button,input,textarea,select')) return;
      var rect = canvas.getBoundingClientRect();
      var relX = (e.clientX - rect.left) / rect.width;
      if (relX < 0.18) prev();
      else if (relX > 0.82) next();
    });

    var touchStartX = null;
    stage.addEventListener('touchstart', function (e) { touchStartX = e.touches[0].clientX; }, { passive: true });
    stage.addEventListener('touchend', function (e) {
      if (touchStartX === null) return;
      var dx = e.changedTouches[0].clientX - touchStartX;
      if (Math.abs(dx) > 50) { dx < 0 ? next() : prev(); }
      touchStartX = null;
    }, { passive: true });

    window.addEventListener('resize', layout);
    document.addEventListener('fullscreenchange', layout);
    document.addEventListener('webkitfullscreenchange', layout);

    layout();
    goTo(0, true);
    updateChrome();
  }

  window['deck-stage'] = { mount: mount };
})();
