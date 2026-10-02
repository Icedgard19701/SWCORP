/* Freight Bill Processor — page script.
   Moved out of templates/index.html (2026-10-01). Server values come from
   window.FB_CONFIG, set by the template just before this file loads. */
(function () {
  'use strict';

  const BASE = window.FB_CONFIG.base;
  const SID  = window.FB_CONFIG.sid;

  const s = {
    needsPriority1: false,
    carrierDone:    false,
    priority1Done:  false,
    pacejetDone:    false,
    acumaticaDone:   false,
    acumaticaFailed: false,
    pollTimer:       null,
    // Small Parcels screen: whether the master workbook was located, and
    // whether it is a hand-uploaded copy that has to be downloaded back.
    parcelMasterOk:   false,
    parcelMasterCopy: false,
  };


  const noticeToast    = document.getElementById('notice-toast');
  const noticeText     = document.getElementById('notice-text');
  const noticeCloseBtn = document.getElementById('notice-close');
  const dropzone       = document.getElementById('dropzone');
  const fileInput      = document.getElementById('file-input');
  const chooseLink     = document.getElementById('choose-link');
  const uploadHint     = document.getElementById('upload-hint');
  const zoneLoadingTxt = document.getElementById('zone-loading-text');
  const pillsWrap      = document.getElementById('pills-container');
  const pCarrier       = document.getElementById('pill-carrier');
  const pP1            = document.getElementById('pill-p1');
  const pPacejet       = document.getElementById('pill-pacejet');
  const pAcumatica     = document.getElementById('pill-acumatica');
  const scrUpload      = document.getElementById('screen-upload');
  const scrProcess     = document.getElementById('screen-processing');
  const scrComplete    = document.getElementById('screen-complete');
  const compRec        = document.getElementById('comp-records');
  const compClose      = document.getElementById('comp-close');
  const compCountdown  = document.getElementById('comp-countdown');
  const compDownloads  = document.getElementById('comp-downloads');

  // ── Event wiring ─────────────────────────────────────────────
  chooseLink.addEventListener('click', e => { e.preventDefault(); fileInput.click(); });
  dropzone.addEventListener('click', e => { if (e.target !== chooseLink) fileInput.click(); });
  // The zone is a button: Enter or Space opens the file picker.
  const pressToOpen = input => e => {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    e.preventDefault();
    // Same as the mouse: a disabled or busy zone takes nothing.
    if (e.currentTarget.matches('.disabled, .uploading')) return;
    input.click();
  };
  dropzone.addEventListener('keydown', pressToOpen(fileInput));
  fileInput.addEventListener('change', () => {
    const files = [...fileInput.files];
    fileInput.value = '';
    if (files.length) handleFiles(files);
  });
  dropzone.addEventListener('dragover',  e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', ()  => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('dragend',   ()  => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault(); dropzone.classList.remove('drag-over');
    const files = [...e.dataTransfer.files];
    if (files.length) handleFiles(files);
  });

  // Start Acumatica fetch immediately on page load
  const _initFd = new FormData(); _initFd.append('sid', SID);
  fetch(BASE + '/init', { method: 'POST', body: _initFd })
    .then(r => r.json())
    .then(data => {
      if (!data.ok) return; // server couldn't start fetch — pill stays idle
      pAcumatica.classList.add('loading');
      startPolling();
    })
    .catch(() => {}); // network failure — pill stays idle

  // Try to auto-load carrier file from shared Iris folder
  pCarrier.classList.add('loading');
  const _carrierFd = new FormData(); _carrierFd.append('sid', SID);
  fetch(BASE + '/carrier-auto', { method: 'POST', body: _carrierFd })
    .then(r => r.json())
    .then(data => {
      if (!data.found) {
        pCarrier.classList.remove('loading');
        showError('Carrier File Not Found — upload manually.');
        return;
      }
      if (data.error) { pCarrier.classList.remove('loading'); showError(data.error, data.blocking, data.blocking); return; }
      if (data.warning) showWarning(data.warning);
      markDone(pCarrier);
      s.carrierDone = true;
      s.summary = data.summary;
      if (data.needs_priority1) {
        s.needsPriority1 = true;
        showPill(pP1);
        showPill(pPacejet);
        showNotice('Two more files are needed — upload the Priority 1 '
                   + 'Pending Invoices and PaceJet BulkExport files below.');
        pP1.classList.add('active');
        pPacejet.classList.add('active');
        updateHint();
      } else {
        dropzone.classList.add('disabled');
        uploadHint.textContent = 'Waiting for Acumatica connection…';
      }
      checkAllReady();
    })
    // Anything that is not JSON lands here — a server error page, a dropped
    // connection. Without a message the pill simply stopped spinning and the
    // screen looked like it had finished, so the failure has to say something.
    .catch(() => {
      pCarrier.classList.remove('loading');
      showError('The Carrier Import File could not be loaded automatically — '
                + 'upload it manually below.');
    });

  // Ask for notification permission on page load
  if ('Notification' in window && Notification.permission === 'default') {
    Notification.requestPermission();
  }

  // ── Process multiple files sequentially ──────────────────────
  async function handleFiles(files) {
    for (const file of files) {
      await handleFile(file);
    }
  }

  // ── Auto-classify upload ──────────────────────────────────────
  function handleFile(file) {
    return new Promise(resolve => {
    hideError();
    const ext = file.name.toLowerCase();
    if (!ext.endsWith('.xlsx') && !ext.endsWith('.csv') && !ext.endsWith('.xls')) {
      showError('Invalid File Type — only .xlsx, .xls and .csv files are accepted.'); resolve(); return;
    }
    setUploading(true, 'Analyzing file…');
    const fd = new FormData();
    fd.append('sid', SID);
    fd.append('file', file);
    fd.append('scope', 'main');
    fetch(BASE + '/upload-auto', { method: 'POST', body: fd })
      .then(r => r.json())
      .then(data => {
        setUploading(false);
        if (data.error) { showError(data.error, data.blocking, data.blocking); resolve(); return; }
        if (data.warning) showWarning(data.warning);
        flashAccepted();



        if (data.type === 'carrier') {
          markDone(pCarrier);
          s.carrierDone = true;
          s.summary = data.summary;

          if (data.needs_priority1) {
            s.needsPriority1 = true;
            pillsWrap.classList.add('four-col');
            showPill(pP1);
            showPill(pPacejet);
            showNotice('Two more files are needed — upload the Priority 1 '
                       + 'Pending Invoices and PaceJet BulkExport files below.');
            if (s.priority1Done) { markDone(pP1); } else { pP1.classList.add('active'); }
            if (s.pacejetDone)   { markDone(pPacejet); } else { pPacejet.classList.add('active'); }
            updateHint();
          } else {
            dropzone.classList.add('disabled');
            uploadHint.textContent = 'Waiting for Acumatica connection…';
          }
          checkAllReady();

        } else if (data.type === 'priority1') {
          s.priority1Done = true;
          if (pP1.style.display === 'none') {
            pillsWrap.classList.add('four-col');
            showPill(pP1);
          }
          markDone(pP1);
          updateHint();
          checkAllReady();

        } else if (data.type === 'pacejet') {
          s.pacejetDone = true;
          if (pPacejet.style.display === 'none') {
            pillsWrap.classList.add('four-col');
            showPill(pPacejet);
          }
          markDone(pPacejet);
          updateHint();
          checkAllReady();

        }
        // Small-parcel files (WWEX raw export, FedEx invoice PDF, the Small
        // Parcels workbook) are rejected by /upload-auto for this screen — they
        // are imported from the Small Parcels screen only.
        resolve();
      })
      .catch(() => { setUploading(false); showError('Connection Error — could not reach the server. Please try again.'); resolve(); });
    }); // end Promise
  }

  function updateHint() {
    // The notice has served its purpose once both files are in.
    if (s.needsPriority1 && s.priority1Done && s.pacejetDone) hideNotice();
    if (!s.carrierDone) {
      uploadHint.textContent = 'Carrier Import, Priority 1, or PaceJet (.xlsx, .csv)';
      return;
    }
    if (!s.needsPriority1) return;
    if (!s.priority1Done && !s.pacejetDone) {
      uploadHint.textContent = 'Upload Priority 1 Pending Invoices and PaceJet BulkExport';
    } else if (!s.priority1Done) {
      uploadHint.textContent = 'Still needed: Priority 1 Pending Invoices (.xlsx or .csv)';
    } else if (!s.pacejetDone) {
      uploadHint.textContent = 'Still needed: PaceJet BulkExport (.csv or .xlsx)';
    } else {
      dropzone.classList.add('disabled');
      uploadHint.textContent = 'All files received. Waiting for Acumatica…';
    }
  }

  // ── Remove a loaded file (× on its pill) ─────────────────────
  // Only the Freight Bill inputs still waiting to be used: the Carrier Import
  // File, Priority 1 and PaceJet. The pill un-ticks and the drop zone takes
  // the next file. WWEX/FedEx invoices have no ×: they are appended to the
  // Small Parcels workbook the moment they land.
  function unmarkPill(pill) {
    pill.classList.remove('done', 'loading', 'popping');
  }
  function hidePill(pill) {
    unmarkPill(pill);
    pill.classList.remove('active');
    if (!_shown(pill)) return;
    const m = _morph(pill.parentElement);
    if (m.entering.delete(pill)) { _finishHide(pill); return; }      // never got on screen
    m.leaving.add(pill);
  }

  function removeFile(kind) {
    if (scrProcess.classList.contains('active')) return;   // already in use
    // All three were in and the processing screen was 1.2s away: hold it.
    if (_transitionTimer) { clearTimeout(_transitionTimer); _transitionTimer = null; }
    const fd = new FormData(); fd.append('sid', SID); fd.append('type', kind);
    fetch(BASE + '/remove-file', { method: 'POST', body: fd })
      .then(r => r.json())
      .then(d => {
        if (d.error) { showError(d.error); return; }
        if (kind === 'carrier') {
          s.carrierDone = false;
          unmarkPill(pCarrier);
          s.summary = null;
          if (s.needsPriority1) {           // what the carrier file asked for goes with it
            s.needsPriority1 = false;
            hideNotice();
            if (!s.priority1Done) hidePill(pP1);
            if (!s.pacejetDone)   hidePill(pPacejet);
          }
        } else if (kind === 'priority1') {
          s.priority1Done = false;
          if (s.needsPriority1) { unmarkPill(pP1); pP1.classList.add('active'); } else hidePill(pP1);
        } else if (kind === 'pacejet') {
          s.pacejetDone = false;            // one file serves both screens
          if (s.needsPriority1) { unmarkPill(pPacejet); pPacejet.classList.add('active'); } else hidePill(pPacejet);
          hidePill(pPacejetP);
        }
        const missing = !s.carrierDone || (s.needsPriority1 && (!s.priority1Done || !s.pacejetDone));
        if (missing) dropzone.classList.remove('disabled');
        updateHint();
      })
      .catch(() => showError('Connection Error — could not remove the file. Please try again.'));
  }

  document.querySelectorAll('[data-remove]').forEach(btn =>
    btn.addEventListener('click', e => { e.stopPropagation(); removeFile(btn.dataset.remove); }));

  // ── Acumatica polling ─────────────────────────────────────────
  function startPolling() {
    if (s.pollTimer) return;
    s.pollTimer = setInterval(() => {
      fetch(BASE + '/acumatica-status?sid=' + SID)
        .then(r => r.json())
        .then(d => {
          if (d.status === 'done') {
            clearInterval(s.pollTimer); s.pollTimer = null;
            // One fetch feeds both screens: the main flow's matching cascade and
            // the Subaccount column of the Small Parcels import file.
            markDone(pAcumatica);
            markDone(pAcumaticaP);
            s.acumaticaDone = true;
            checkAllReady();
          } else if (d.status === 'error') {
            clearInterval(s.pollTimer); s.pollTimer = null;
            pAcumatica.classList.remove('loading');  // back to empty circle
            pAcumaticaP.classList.remove('loading');
            s.acumaticaFailed = true;
            showError('Acumatica: ' + (d.error || 'Sync error.'));
          }
        }).catch(() => {});
    }, 2500);
  }

  // ── Processing screen messages ────────────────────────────────
  const PROC_MSGS = [
    'Connecting to Acumatica...',
    'Fetching open shipments...',
    'Matching sales orders...',
    'Cross-referencing PRO numbers...',
    'Syncing carrier records...',
    'Validating invoice data...',
    'Calculating freight charges...',
    'Reconciling pending invoices...',
    'Almost there...',
  ];
  const procText = document.getElementById('proc-text');
  let _msgTimer = null;
  let _msgIdx   = 0;

  // msgs defaults to the freight-bill sequence; Small Parcels passes its own.
  function startMsgCycle(msgs) {
    const list = msgs || PROC_MSGS;
    _msgIdx = 0;
    procText.classList.remove('swapping');
    procText.textContent = list[0];
    _msgTimer = setInterval(() => {
      _msgIdx = (_msgIdx + 1) % list.length;
      procText.classList.add('swapping');          // fade + 2px blur (CSS)
      setTimeout(() => {
        procText.textContent = list[_msgIdx];
        procText.classList.remove('swapping');
      }, 300);
    }, 2800);
  }

  function stopMsgCycle() {
    if (_msgTimer) { clearInterval(_msgTimer); _msgTimer = null; }
  }

  // ── Gate: all required uploads + Acumatica → process ─────────
  let _transitionTimer = null;
  let _run = null;   // the run once every file is in: {started, shown, land}

  function checkAllReady() {
    const uploadsDone = s.carrierDone &&
      (!s.needsPriority1 || (s.priority1Done && s.pacejetDone));
    if (!uploadsDone) return;

    if (_run) {
      // Already past the pills — fire if Acumatica just answered
      if (s.acumaticaDone && !_run.started) runProcess();
      return;
    }

    // Delay so user sees all pill checkmarks before the next screen takes over
    if (!_transitionTimer) {
      _transitionTimer = setTimeout(() => {
        _transitionTimer = null;
        // The pending invoices roll up while the bills are built behind them,
        // so the odometer fills the wait instead of adding to it. When it ends
        // the outcome is shown at once if it is in, else the loading screen
        // until it is.
        _run = { started: false, shown: false, land: null };
        if (s.acumaticaDone) runProcess();
        showResults(s.summary ? s.summary.total : 0, () => !!(_run && _run.land), () => {
          _run.shown = true;
          if (_run.land) { const f = _run.land; _run.land = null; f(); }
          else { showScreen('processing'); startMsgCycle(); }
        });
      }, 1200);
    }
  }

  // The outcome of /process waits for the odometer to end.
  function landRun(f) {
    if (_run && !_run.shown) { _run.land = f; return; }
    f();
  }

  function runProcess() {
    if (_run) _run.started = true;
    stopMsgCycle();
    procText.classList.remove('swapping');
    procText.textContent   = 'Building your report...';
    const _procFd = new FormData(); _procFd.append('sid', SID);
    fetch(BASE + '/process', { method: 'POST', body: _procFd })
      .then(r => r.json())
      .then(data => {
        if (data.error) { landRun(() => { _run = null; stopMsgCycle(); showScreen('upload'); showError(data.error); }); return; }
        compRec.textContent = 'Total Records Processed: ' + data.total_records;

        const compNote = document.getElementById('comp-note');
        compDownloads.innerHTML = '';  // LTL run hands nothing over by button
        if (data.copy_error) {
          compNote.textContent = 'Could not save to shared folder — downloading instead.';
          window.location.href = BASE + '/download?sid=' + (data.sid || SID) + '&file=' + encodeURIComponent(data.filename);
        } else {
          compNote.innerHTML = 'File saved to the <a href="https://swcorp.sharepoint.com/:f:/s/APTeam/IgBXo3fwVSwCRbUbYOGaV8lYARbA-L8OS3bbmYCcBhF7ZUc?e=foq1Gt" target="_blank" rel="noopener" class="comp-folder-link">Carrier Bills Import folder</a>.';
        }
        landRun(() => {
          stopMsgCycle();
          showScreen('complete');
          fireNotification(data.total_records, !data.copy_error);
          startResetCountdown();
        });
      })
      .catch(() => landRun(() => { _run = null; stopMsgCycle(); showScreen('upload'); showError('Processing Error — please try again.'); }));
  }

  // ── Toast timer ───────────────────────────────────────────────
  // Family standard: every toast lives 6s, with a thin line along its foot
  // that empties as the time runs out. The line's CSS animation IS the timer:
  // hovering or focusing the toast pauses it (CSS), and its end dismisses the
  // toast. Arming again (same toast shown again) restarts it from full.
  function armToast(el, onDone) {
    let bar = el.querySelector(':scope > .toast-timer');
    if (!bar) {
      bar = document.createElement('i');
      bar.className = 'toast-timer';
      bar.setAttribute('aria-hidden', 'true');
      el.appendChild(bar);
    }
    bar.style.animation = 'none';
    void bar.offsetWidth;                         // restart the animation
    bar.style.animation = '';
    // Stop the event here: the toast's own animationend listeners (exit
    // animation) must not see the timer's end as theirs.
    bar.onanimationend = (e) => { e.stopPropagation(); onDone(); };
  }
  function disarmToast(el) {
    const bar = el.querySelector(':scope > .toast-timer');
    if (bar) bar.onanimationend = null;
  }
  // Exit animation, then gone (notices).
  function leaveToast(el) {
    disarmToast(el);
    if (!el.classList.contains('visible')) return;
    el.classList.add('hiding');
    el.addEventListener('animationend', function done(e) {
      if (e.target !== el) return;
      el.removeEventListener('animationend', done);
      if (!el.classList.contains('hiding')) return;   // shown again meanwhile
      el.classList.remove('visible', 'hiding');
    });
  }

  // Every toast timer pauses while the tab is hidden (CSS: html.tab-hidden),
  // so a message is never missed behind another tab.
  document.addEventListener('visibilitychange', () =>
    document.documentElement.classList.toggle('tab-hidden', document.hidden));

  // When a toast joins or leaves the stack the others slide to their new
  // place instead of jumping: the stack is measured after every change and
  // each toast that moved animates from where it was (FLIP, transform only).
  (function () {
    const stack = document.querySelector('.toast-stack');
    if (!stack) return;
    const measure = () => {
      const m = new Map();
      [...stack.children].forEach(t => {
        if (getComputedStyle(t).display !== 'none') m.set(t, t.getBoundingClientRect().top);
      });
      return m;
    };
    let last = measure();
    new MutationObserver(() => {
      const now = measure();
      if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
        now.forEach((top, t) => {
          const was = last.get(t);
          if (was == null || Math.abs(was - top) < 1) return;
          t.animate([{ transform: `translateY(${was - top}px)` }, { transform: 'none' }],
                    { duration: 240, easing: 'cubic-bezier(.23, 1, .32, 1)', composite: 'add' });
        });
      }
      last = now;
    }).observe(stack, { attributes: true, attributeFilter: ['class', 'style'], subtree: true });
  })();

  // ── Results screen (odometer, then Processing Complete) ──────
  // Ported from Lowe's Invoice Reconciler's Process step. Each digit of the
  // totals rolls up to its value: a strip of 0-9 per digit slides under a
  // one-line window, digits further right spin more laps, cards and digits
  // start staggered. Then the bar under "Processing Complete" fills until the
  // rolls have ended. Then, if the run behind it is already done, the figures
  // stand READ ms to be read before its outcome; if not, the loading screen
  // follows at once (no pause in front of a wait). Any click or key stops the
  // countdown and offers Continue.
  const scrResults   = document.getElementById('screen-results');
  const resultsSums  = document.getElementById('results-sums');
  const resultsAdv   = document.getElementById('results-advance');
  const resultsGo    = document.getElementById('results-continue');
  const RESULTS_READ = 2000;
  // The screen is in view at once (it only slides a little): figures start soon.
  const RESULTS_LEAD = 160;
  const still = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function sumCard(label, value, lead) {
    const card = document.createElement('div');
    card.className = 'sum' + (lead ? ' lead' : '');
    const l = document.createElement('span'); l.className = 'sum-l'; l.textContent = label;
    const v = document.createElement('b');    v.className = 'sum-v'; v.textContent = value;
    card.append(l, v);
    return card;
  }

  // Rolls one figure; returns the ms from now until its last digit stops.
  function rollUp(out, delay) {
    const final = out.textContent;
    const chars = [...final], total = chars.filter(c => /\d/.test(c)).length;
    let pos = 0, end = 0;
    out.replaceChildren();
    const sr = document.createElement('span'); sr.className = 'vh'; sr.textContent = final;
    out.append(sr);
    const strips = [];
    chars.forEach(c => {
      if (!/\d/.test(c)) {
        const t = document.createElement('span'); t.setAttribute('aria-hidden', 'true'); t.textContent = c;
        out.append(t); return;
      }
      const laps = 1 + Math.round((pos++ / Math.max(1, total - 1)) * 2);   // 1 lap left .. 3 right
      const win = document.createElement('span'); win.className = 'roll'; win.setAttribute('aria-hidden', 'true');
      const strip = document.createElement('span'); strip.className = 'roll-strip';
      for (let k = 0; k < laps * 10 + Number(c) + 1; k++) {
        const d = document.createElement('i'); d.textContent = k % 10; strip.append(d);
      }
      win.append(strip); out.append(win); strips.push(strip);
    });
    const anims = strips.map((st, k) => st.animate(
      [{ transform: 'translateY(0)' }, { transform: `translateY(-${(st.children.length - 1) * 1.2}em)` }],
      { duration: 1500, delay: delay + k * 45, easing: 'cubic-bezier(.16, 1, .3, 1)', fill: 'both' }));
    anims.forEach(an => { const t = an.effect.getTiming(); end = Math.max(end, t.delay + t.duration); });
    Promise.all(anims.map(an => an.finished)).then(() => { out.textContent = final; }, () => {});
    return end;
  }

  function showResults(total, isDone, onDone) {
    const sm = s.summary;
    const cards = [sumCard('Pending Invoices', Number(total).toLocaleString('en-US'), true)];
    if (sm && sm.carriers) sm.carriers.forEach(c => cards.push(sumCard(c.name, Number(c.count).toLocaleString('en-US'))));
    cards.forEach((c, i) => c.style.setProperty('--i', i));   // rise in one after another
    resultsSums.replaceChildren(...cards);
    resultsGo.hidden = true;
    resultsAdv.hidden = false;
    showScreen('results');

    let rollEnd = 0;
    if (!still()) cards.forEach((card, i) => {
      rollEnd = Math.max(rollEnd, rollUp(card.querySelector('.sum-v'), RESULTS_LEAD + i * 110));
    });
    const wait = still() ? 1000 : rollEnd;
    const bar  = resultsAdv.querySelector('.advance-bar > i');
    const fill = still() ? null : bar.animate([{ transform: 'scaleX(0)' }, { transform: 'scaleX(1)' }],
                                               { duration: wait, easing: 'linear', fill: 'forwards' });
    if (!fill) bar.style.transform = 'none';

    let gone = false;
    const finish = () => {
      if (gone) return; gone = true;
      clearTimeout(timer);
      document.removeEventListener('keydown', hold, true);
      document.removeEventListener('pointerdown', hold, true);
      onDone();
    };
    let timer = setTimeout(() => {
      timer = isDone() ? setTimeout(finish, RESULTS_READ) : (finish(), null);
    }, wait);
    // A click or a key anywhere keeps the figures on screen; Continue goes on.
    const hold = e => {
      if (e.target === resultsGo) return;
      clearTimeout(timer); if (fill) fill.cancel();
      resultsAdv.hidden = true; resultsGo.hidden = false;
      document.removeEventListener('keydown', hold, true);
      document.removeEventListener('pointerdown', hold, true);
    };
    document.addEventListener('keydown', hold, true);
    document.addEventListener('pointerdown', hold, true);
    resultsGo.onclick = finish;
  }

  // ── UI helpers ────────────────────────────────────────────────
  // Pills joining or leaving a row (2 <-> 4 on Freight Bill, the carrier
  // pills on Small Parcels) slide instead of jumping. Every change made in the
  // same tick is one move: the row is measured before the first change and
  // after the last, then each pill animates its width between the two —
  // staying pills grow or shrink, a joining pill opens from nothing, a leaving
  // one closes to nothing and fades (it stays in the row until then, and its
  // gap is taken back with a negative margin, so nothing jumps at either
  // end). Instant on phones (pills wrap there) and under reduced motion.
  const MORPH_MS   = 350;
  // Ease-in-out: a layout change that starts at full speed reads as a jump.
  const MORPH_EASE = 'cubic-bezier(.45, 0, .2, 1)';
  const _morphs    = new Map();   // row -> {before, entering, leaving}
  const _shown     = p => getComputedStyle(p).display !== 'none';
  const _rowPills  = row => [...row.querySelectorAll(':scope > .pill')];
  const _noMorph   = () => window.matchMedia('(prefers-reduced-motion: reduce), (max-width: 599px)').matches;

  function _finishHide(pill) {
    pill.classList.remove('pill-show');
    pill.style.display = 'none';
  }

  function _morph(row) {
    let m = _morphs.get(row);
    if (m) return m;
    // A move still running settles where it is first, so this one starts
    // from what is on screen.
    _rowPills(row).forEach(p => p.getAnimations().forEach(a => { if (a.id === 'pill-morph') a.finish(); }));
    m = { before: new Map(), entering: new Set(), leaving: new Set() };
    _rowPills(row).forEach(p => { if (_shown(p)) m.before.set(p, p.getBoundingClientRect().width); });
    _morphs.set(row, m);
    queueMicrotask(() => { _morphs.delete(row); _playMorph(row, m); });
    return m;
  }

  function _playMorph(row, m) {
    m.leaving.forEach(p => { p.style.display = 'none'; });          // the row as it ends up
    const after = new Map(_rowPills(row).filter(_shown).map(p => [p, p.getBoundingClientRect().width]));
    if (_noMorph() || !row.offsetParent || !(m.entering.size || m.leaving.size)) {
      m.leaving.forEach(_finishHide);                                // other view, phone, reduced motion
      return;
    }
    m.leaving.forEach(p => { p.style.display = ''; });              // back in the row for its exit
    const gap = parseFloat(getComputedStyle(row).columnGap) || 0;
    _rowPills(row).filter(_shown).forEach((p, i) => {
      const cs   = getComputedStyle(p);
      const side = i === 0 ? 'marginRight' : 'marginLeft';
      const open = w => ({ width: `${w}px`, paddingLeft: cs.paddingLeft, paddingRight: cs.paddingRight, [side]: '0px', opacity: 1 });
      const shut = { width: '0px', paddingLeft: '0px', paddingRight: '0px', [side]: `-${gap}px`, opacity: 0 };
      let frames;
      if (m.entering.has(p))     frames = [shut, open(after.get(p))];
      else if (m.leaving.has(p)) frames = [open(m.before.get(p)), shut];
      else                       frames = [{ width: `${m.before.get(p) ?? after.get(p)}px` }, { width: `${after.get(p)}px` }];
      p.style.flex = '0 0 auto';
      p.style.overflow = 'hidden';
      const a = p.animate(frames, { duration: MORPH_MS, easing: MORPH_EASE, id: 'pill-morph' });
      a.finished.then(() => {
        p.style.flex = ''; p.style.overflow = '';
        if (m.leaving.has(p)) _finishHide(p);
      }, () => { p.style.flex = ''; p.style.overflow = ''; });
    });
  }

  function showPill(pill) {
    const m = _morph(pill.parentElement);
    if (m.leaving.delete(pill)) return;                              // called back mid-exit
    if (_shown(pill)) return;
    pill.style.display = '';
    pill.classList.add('pill-show');
    m.entering.add(pill);
  }

  // ── Screen change ─────────────────────────────────────────────
  // The new screen comes in from the left (always left to right) and fades
  // up: transform + opacity only, so the browser composites it without
  // repainting (the masked view-transition wipe froze ~180ms on the click and
  // ran at 15-20fps, measured). The old screen goes at once; header, menu and
  // toasts are untouched. Reduced motion: the swap just happens.
  const SCREEN_IN = [{ opacity: 0, transform: 'translateX(-28px)' }, { opacity: 1, transform: 'translateX(0)' }];
  function contentSwap(swap, el) {
    swap();
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    el.getAnimations().forEach(a => { if (a.id === 'screen-in') a.cancel(); });
    el.animate(SCREEN_IN, { duration: 280, easing: 'cubic-bezier(.23, 1, .32, 1)', id: 'screen-in' });
  }

  // Keyboard users land on the screen that opened (unless they are in the menu).
  function focusScreen(el) {
    if (!document.getElementById('side-menu').contains(document.activeElement)) el.focus({ preventScroll: true });
  }

  function showScreen(name) {
    if (parcelsOpen()) closeParcels(true);   // never leave Small Parcels stacked under a screen
    const fromUpload   = scrUpload.style.display !== 'none';

    function _activateScreen(el) {
      el.classList.add('active', 'just-entered');
      el.addEventListener('animationend', () => el.classList.remove('just-entered'), { once: true });
      focusScreen(el);
    }

    if (name === 'results') {
      // Files in -> Pending Invoices.
      contentSwap(() => {
        scrUpload.style.display = 'none';
        scrResults.classList.add('active');
        focusScreen(scrResults);
      }, scrResults);

    } else if (name === 'processing') {
      // Pending Invoices (or upload) -> processing (navy): a plain fade in.
      if (fromUpload) scrUpload.style.display = 'none';
      scrResults.classList.remove('active', 'just-entered');
      scrComplete.classList.remove('active', 'just-entered');
      scrProcess.classList.add('active', 'fading-in');
      scrProcess.addEventListener('animationend', () => scrProcess.classList.remove('fading-in'), { once: true });
      focusScreen(scrProcess);

    } else if (name === 'complete') {
      scrProcess.classList.remove('active', 'just-entered');
      scrResults.classList.remove('active', 'just-entered');
      _activateScreen(scrComplete);

    } else {
      // Back to upload (error recovery or reset)
      scrProcess.classList.remove('active', 'just-entered');
      scrResults.classList.remove('active', 'just-entered');
      scrComplete.classList.remove('active', 'just-entered');
      scrUpload.style.display = '';
      scrUpload.classList.remove('leaving');
      scrUpload.classList.add('entering');
      scrUpload.addEventListener('animationend', () => scrUpload.classList.remove('entering'), { once: true });
    }
  }

  // ── Complete screen: countdown + auto-reset ───────────────────
  let _resetTimer    = null;
  let _countdownTick = null;

  function startResetCountdown() {
    if (_countdownTick) { clearInterval(_countdownTick); _countdownTick = null; }
    let secs = 10;
    compCountdown.textContent = `Returning to home in ${secs}s…`;
    _countdownTick = setInterval(() => {
      secs--;
      if (secs > 0) {
        compCountdown.textContent = `Returning to home in ${secs}s…`;
      } else {
        clearInterval(_countdownTick);
        resetApp();
      }
    }, 1000);
  }

  function resetApp() {
    if (_resetTimer)    { clearTimeout(_resetTimer);    _resetTimer    = null; }
    if (_countdownTick) { clearInterval(_countdownTick); _countdownTick = null; }
    window.location.reload();
  }

  compClose.addEventListener('click', resetApp);

  // ── Small Parcels view (WWEX raw export, FedEx invoice PDF) ───
  const scrParcels = document.getElementById('screen-parcels');
  const pzDrop     = document.getElementById('pz-dropzone');
  const pzInput    = document.getElementById('pz-file-input');
  const pzChoose   = document.getElementById('pz-choose');
  const pzHint     = document.getElementById('pz-hint');
  const pzLoadTxt  = document.getElementById('pz-loading-text');
  const pMasterP    = document.getElementById('pill-master-p');
  const pPacejetP   = document.getElementById('pill-pacejet-p');
  const pAcumaticaP = document.getElementById('pill-acumatica-p');
  const pInvoiceP  = document.getElementById('pill-invoice-p');
  const pWWEXp     = document.getElementById('pill-wwex-p');
  const pFedexP    = document.getElementById('pill-fedex-p');
  const PZ_HINT_DEFAULT = pzHint.textContent;
  const pzNotice      = document.getElementById('pz-notice-toast');
  const pzNoticeText  = document.getElementById('pz-notice-text');
  const pzNoticeClose = document.getElementById('pz-notice-close');

  // ── "An invoice is still needed" ──────────────────────────────
  // The PaceJet export and the workbook are both helpers: neither imports a
  // single row on its own. Landing back on this screen with every pill ticked
  // and no invoice looks like a completed run, so the state is stated instead —
  // a notice that has to be closed by hand, and a pill naming the missing file.
  function showInvoiceNeeded() {
    showPill(pInvoiceP);
    pInvoiceP.classList.add('active');
    pzNoticeText.textContent = 'One more file is needed — upload the WWEX raw '
                             + 'export or a FedEx invoice to import.';
    pzNotice.classList.remove('hiding');
    pzNotice.classList.add('visible');
    armToast(pzNotice, () => leaveToast(pzNotice));
    pzHint.textContent = 'Upload the WWEX raw export (.xls, .xlsx, .csv) or a '
                       + 'FedEx invoice (.pdf) to import';
  }

  // Called the moment an invoice is accepted: the placeholder has been answered
  // by a real carrier pill, so it goes away rather than sitting alongside it.
  function clearInvoiceNeeded() {
    leaveToast(pzNotice);
    hidePill(pInvoiceP);   // the carrier pill shown next takes its place in one move
    pzHint.textContent = PZ_HINT_DEFAULT;
  }

  pzNoticeClose.addEventListener('click', () => leaveToast(pzNotice));

  // type returned by /upload-auto -> pill, /parcels-append type, target tab
  const PARCEL_KINDS = {
    wwex_raw:  { pill: pWWEXp,  type: 'wwex_raw',  tab: 'WWEX'  },
    fedex_pdf: { pill: pFedexP, type: 'fedex_pdf', tab: 'Fedex' },
  };

  // ── Additional Files Required toast ───────────────────────────
  // 6s like every toast; it also leaves at once when the files are in
  // (hideNotice). The amber pills keep showing what is still missing.
  function hideNotice() {
    leaveToast(noticeToast);
  }

  function showNotice(msg) {
    noticeText.textContent = msg;
    noticeToast.classList.remove('hiding');
    noticeToast.classList.add('visible');
    armToast(noticeToast, hideNotice);
  }

  noticeCloseBtn.addEventListener('click', hideNotice);

  // ── Generated Acumatica bills ─────────────────────────────────
  // One file per uploaded carrier file, collected across the whole batch so the
  // completion screen can hand them all over at once.
  let _importFiles = [];

  function money(n) {
    return '$' + (n || 0).toLocaleString('en-US',
      { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // Several downloads in a row: a hidden iframe per file, staggered, because a
  // single navigation can only carry one of them.
  function triggerDownload(url, delay) {
    setTimeout(() => {
      const f = document.createElement('iframe');
      f.style.display = 'none';
      f.src = url;
      document.body.appendChild(f);
      setTimeout(() => f.remove(), 60000);
    }, delay);
  }

  function collectImportFile(d) {
    if (d.status_error)  showError(d.status_error);
    if (d.import_error) { showError(d.import_error); return; }
    if (!d.import_file) return;
    _importFiles.push({
      file:        d.import_file,
      label:       (d.tab || '') + (d.invoice_number ? ' ' + d.invoice_number : ''),
      rows:        d.import_rows || 0,
      total:       d.import_total || 0,
      unmatched:   d.import_unmatched || 0,
      viaPacejet:  d.import_via_pacejet || 0,
      viaSender:   d.import_via_sender || 0,
    });
  }

  function parcelsOpen() { return scrParcels.classList.contains('active'); }

  // The tab icon is the SWCorp mark and no longer follows the view — one brand
  // icon for both screens. Kept as a no-op so the view-switch calls below stay
  // readable; delete both call sites if the per-view icon is never coming back.
  function setFavicon() {}

  // Locate the master workbook before anything can be imported into it. Runs on
  // first open, and again on any later open that found nothing — the shared
  // folder may have finished syncing in the meantime.
  function checkParcelMaster() {
    if (s.parcelMasterOk) return;
    pMasterP.classList.add('loading');
    const fd = new FormData(); fd.append('sid', SID);
    fetch(BASE + '/parcels-master', { method: 'POST', body: fd })
      // 404 = SMALL_PARCEL_ENABLED is off server-side; the body isn't JSON
      .then(r => (r.status === 404
        ? { error: 'Small Parcels is disabled — enable it on the server.' }
        : r.json()))
      .then(d => {
        pMasterP.classList.remove('loading');
        if (!d.found) {
          // Not fatal: the workbook can be uploaded by hand instead.
          showError(d.error || 'Small Parcels workbook not found — upload it manually.');
          pzHint.textContent = 'Upload the Small Parcels workbook (.xlsx) to continue';
          return;
        }
        s.parcelMasterOk  = true;
        s.parcelMasterCopy = !!d.session_copy;
        markDone(pMasterP);
        pzHint.textContent = PZ_HINT_DEFAULT;
      })
      .catch(() => {
        pMasterP.classList.remove('loading');
        showError('Connection Error — could not check the Small Parcels workbook.');
      });
  }

  // keepError = we are returning here because an import just failed, and the
  // error toast explaining why was raised a tick ago — wiping it is what left
  // the carrier pill grey with no message at all.
  function openParcels(keepError) {
    if (!keepError) hideError();
    // The fetch was already kicked off on page load; mirror wherever it got to.
    if (s.acumaticaDone)           markDone(pAcumaticaP);
    else if (!s.acumaticaFailed)   { pAcumaticaP.classList.add('loading'); startPolling(); }
    checkParcelMaster();
    document.body.classList.add('view-parcels');
    markMenuView('parcels');
    setFavicon('bill');

    // Freight Bills -> Small Parcels.
    scrUpload.classList.remove('re-entering', 'entering', 'leaving');
    contentSwap(() => {
      scrUpload.style.display = 'none';
      scrParcels.classList.add('active');
      focusScreen(scrParcels);
    }, scrParcels);
  }

  // instant = another screen (processing/complete) is taking over, so skip the
  // exit animation and just get out of the way.
  function closeParcels(instant) {
    document.body.classList.remove('view-parcels');
    markMenuView('freight');
    setFavicon('parcel');

    function finish() {
      scrParcels.classList.remove('active', 'just-entered', 'leaving');
      // Only bring the upload screen back if no other screen has taken over
      if (!scrProcess.classList.contains('active') && !scrComplete.classList.contains('active')) {
        scrUpload.style.display = '';
        scrUpload.classList.remove('re-entering');
        if (!instant) focusScreen(scrUpload);
      }
    }

    if (instant) { finish(); return; }
    contentSwap(finish, scrUpload);   // Small Parcels -> Freight Bills: same direction, left to right
  }

  // ── Side menu (header button) ─────────────────────────────────
  // Views (Freight Bills / Small Parcels) and links. Opens as a circle out of
  // the button (CSS); Esc, a click outside or picking an entry closes it, and
  // focus goes back to the button. Closed, it is inert (no Tab stops inside).
  const menuBtn   = document.getElementById('menu-toggle');
  const sideMenu  = document.getElementById('side-menu');
  const menuScrim = document.getElementById('menu-scrim');

  // The circle is centred on the button, measured (not assumed): header
  // padding and button margins change with the screen tier.
  function placeMenuCircle() {
    const r = sideMenu.getBoundingClientRect(), b = menuBtn.getBoundingClientRect();
    const cx = b.left + b.width / 2 - r.left, cy = b.top + b.height / 2 - r.top;
    sideMenu.style.setProperty('--cx', cx + 'px');
    sideMenu.style.setProperty('--cy', cy + 'px');
    // Radius that reaches the panel's far corner (bottom left).
    sideMenu.style.setProperty('--menu-r', Math.ceil(Math.hypot(cx, r.height - cy)) + 2 + 'px');
  }
  placeMenuCircle();
  window.addEventListener('resize', placeMenuCircle);

  // quick: picking an entry closes it at once and fast (the new screen is
  // already coming in; the full close kept a third of the window covered
  // for ~560ms, measured).
  function setMenu(open, quick) {
    if (open === sideMenu.classList.contains('open')) return;
    if (open) placeMenuCircle();
    if (quick) {
      document.body.classList.add('menu-quick');
      setTimeout(() => document.body.classList.remove('menu-quick'), 260);
    }
    sideMenu.classList.toggle('open', open);
    menuScrim.classList.toggle('open', open);
    document.body.classList.toggle('menu-open', open);
    sideMenu.inert = !open;
    menuBtn.setAttribute('aria-expanded', String(open));
    menuBtn.setAttribute('aria-label', open ? 'Close menu' : 'Open menu');
    if (open) (sideMenu.querySelector('[aria-current]') || sideMenu.querySelector('.menu-item')).focus({ preventScroll: true });
    else if (sideMenu.contains(document.activeElement)) menuBtn.focus();
  }

  function markMenuView(view) {
    sideMenu.querySelectorAll('[data-view]').forEach(b => {
      if (b.dataset.view === view) b.setAttribute('aria-current', 'page');
      else b.removeAttribute('aria-current');
    });
  }

  menuBtn.addEventListener('click', () => setMenu(!sideMenu.classList.contains('open')));
  menuScrim.addEventListener('click', () => setMenu(false));
  // Tabbing out of the open menu closes it (it is not a trap, but it never
  // stays open behind the focus).
  sideMenu.addEventListener('focusout', e => {
    const to = e.relatedTarget;
    if (to && !sideMenu.contains(to) && to !== menuBtn) setMenu(false);
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && sideMenu.classList.contains('open')) { e.preventDefault(); setMenu(false); }
  });
  sideMenu.querySelector('a.menu-item').addEventListener('click', () => setMenu(false));

  // Pointer glide (mouse only): the highlight slides to the entry under the
  // pointer; it appears in place on the first entry, fades when leaving.
  const menuBody  = sideMenu.querySelector('.menu-body');
  const menuGlide = sideMenu.querySelector('.menu-glide');
  if (window.matchMedia('(hover: hover) and (pointer: fine)').matches) {
    menuBody.addEventListener('pointerover', e => {
      const item = e.target.closest('.menu-item');
      if (!item) return;
      const r = item.getBoundingClientRect(), b = menuBody.getBoundingClientRect();
      const fresh = !menuGlide.classList.contains('on');
      if (fresh) menuGlide.style.transition = 'opacity 160ms ease';   // fades in place, no slide from the last spot
      menuGlide.style.height = r.height + 'px';
      menuGlide.style.transform = `translateY(${r.top - b.top}px)`;
      if (fresh) { void menuGlide.offsetWidth; menuGlide.style.transition = ''; }
      menuGlide.classList.add('on');
    });
    menuBody.addEventListener('pointerleave', () => menuGlide.classList.remove('on'));
  }

  // ── Logo: back to the main screen (Freight Bills) ────────────
  // Small Parcels -> Freight Bills; the finished run -> a fresh start (as the
  // complete screen's close). Already home, or mid-run: stay (a run is never
  // interrupted from here). Never a page reload: files already in are kept.
  document.getElementById('header-home').addEventListener('click', e => {
    e.preventDefault();
    if (sideMenu.classList.contains('open')) setMenu(false, true);
    if (scrComplete.classList.contains('active')) { resetApp(); return; }
    const midRun = scrProcess.classList.contains('active') || scrResults.classList.contains('active');
    if (midRun || _navBusy || !parcelsOpen()) return;
    _navBusy = true;
    setTimeout(() => { _navBusy = false; }, 300);
    closeParcels();
  });

  // Guard: the swap is two chained animations, so ignore picks until it settles
  let _navBusy = false;
  sideMenu.querySelectorAll('[data-view]').forEach(b => b.addEventListener('click', () => {
    // Only from the upload screens: mid-run (results, processing, complete)
    // the view stays where the run put it.
    const midRun = [scrProcess, scrResults, scrComplete].some(el => el.classList.contains('active'));
    const want = b.dataset.view === 'parcels';
    setMenu(false, true);   // closes fast while the new screen comes in underneath
    if (midRun || _navBusy || want === parcelsOpen()) return;
    _navBusy = true;
    setTimeout(() => { _navBusy = false; }, 300);   // the screen's entrance
    if (want) openParcels(); else closeParcels();
  }));

  pzChoose.addEventListener('click', e => { e.preventDefault(); pzInput.click(); });
  pzDrop.addEventListener('click', e => { if (e.target !== pzChoose) pzInput.click(); });
  pzDrop.addEventListener('keydown', pressToOpen(pzInput));
  pzInput.addEventListener('change', () => {
    const files = [...pzInput.files];
    pzInput.value = '';
    if (files.length) sendParcelFiles(files);
  });
  pzDrop.addEventListener('dragover',  e => { e.preventDefault(); pzDrop.classList.add('drag-over'); });
  pzDrop.addEventListener('dragleave', ()  => pzDrop.classList.remove('drag-over'));
  pzDrop.addEventListener('dragend',   ()  => pzDrop.classList.remove('drag-over'));
  pzDrop.addEventListener('drop', e => {
    e.preventDefault(); pzDrop.classList.remove('drag-over');
    const files = [...e.dataTransfer.files];
    if (files.length) sendParcelFiles(files);
  });

  // Several files at once — the WWEX export plus a FedEx invoice, or a batch of
  // invoices. Processed one at a time: they all write to the same workbook, and
  // each one produces its own Acumatica file.
  // The PaceJet export imports nothing by itself — it only lets later invoices
  // resolve more Subaccounts — so a batch made of nothing else never earns the
  // processing screen. Detected from the CSV header, the same column the server
  // validates on, because the WWEX raw export is also a .csv and only the
  // header tells them apart.
  function sniffPacejetCsv(file) {
    if (!file.name.toLowerCase().endsWith('.csv')) return Promise.resolve(false);
    return file.slice(0, 65536).text()
      .then(chunk => {
        const header = chunk.split(/\r?\n/, 1)[0] || '';
        return /shipmentuserfield3/i.test(header) || header.split(',').length > 60;
      })
      .catch(() => false);
  }

  async function sendParcelFiles(files) {
    // Helper files first, invoices after: the PaceJet export (.csv) and the
    // workbook itself (.xlsx) have to be in place before anything is appended,
    // since each invoice is resolved and billed as it is imported. The invoices
    // are the WWEX export (.xls) and the FedEx PDF.
    const RANK = { csv: 0, xlsx: 1, xls: 2, pdf: 3 };
    const rank = f => RANK[f.name.toLowerCase().split('.').pop()] ?? 9;
    const queue = [...files].sort((a, b) => rank(a) - rank(b));

    const isHelper   = await Promise.all(queue.map(sniffPacejetCsv));
    const helperOnly = isHelper.length > 0 && isHelper.every(Boolean);

    _importFiles = [];
    _parcelResults = [];
    if (!helperOnly) {
      showScreen('processing');
      startMsgCycle(PARCEL_PROC_MSGS);
    }

    for (let i = 0; i < queue.length; i++) {
      await sendParcelFile(queue[i], isHelper[i]);
      if (_parcelAborted) break;
    }

    // Nothing was appended and the screen never changed, so there is no
    // completion report to show and no screen to return to. Any failure has
    // already put its message on this screen.
    if (helperOnly) {
      _parcelAborted = false;
      return;
    }

    stopMsgCycle();
    // Back to Small Parcels when something failed (the error stays on screen) or
    // when the batch was only the workbook itself and there is nothing to report.
    if (_parcelAborted || !_parcelResults.length) {
      const failed = _parcelAborted;
      _parcelAborted = false;
      showScreen('upload');
      openParcels(failed);
      // Nothing was imported and nothing went wrong — the batch was helper
      // files only, so say which file would actually import something. An
      // outright failure is left to its own message.
      if (!failed) showInvoiceNeeded();
      return;
    }
    showParcelComplete();
  }

  const PARCEL_PROC_MSGS = [
    'Reading the carrier file...',
    'Matching against what is already logged...',
    'Appending to the Small Parcels workbook...',
    'Resolving customers in Acumatica...',
    'Building the Acumatica bill...',
  ];

  let _parcelResults = [];
  let _parcelAborted = false;

  function showParcelComplete() {
    const added   = _parcelResults.reduce((n, r) => n + (r.added || 0), 0);
    const skipped = _parcelResults.reduce((n, r) => n + (r.skipped_duplicates || 0), 0);
    const gaps    = _importFiles.reduce((n, f) => n + f.unmatched, 0);

    compRec.textContent = 'Total Records Processed: ' + added;

    const lines = _parcelResults.map(r => {
      const label = (r.tab || '') + (r.invoice_number ? ' ' + r.invoice_number : '');
      return r.added
        ? `${label} — ${r.added} rows logged, ${r.status || ''}`
        : `${label} — already imported, nothing new`;
    });
    if (skipped) lines.push(`${skipped} duplicate rows skipped.`);
    // Charges the invoice bills to the account rather than to a shipment (a
    // scheduled pickup, a late fee). Called out because they carry no tracking
    // number, so AP would otherwise read them as unidentified rows.
    const fees = _parcelResults.reduce((n, r) => n + (r.fee_rows_added || 0), 0);
    if (fees) {
      lines.push(`${fees === 1 ? '1 non-shipment charge' : fees + ' non-shipment charges'} `
                 + 'billed to the carrier fee subaccount.');
    }
    if (gaps) {
      lines.push(`${gaps === 1 ? '1 row needs' : gaps + ' rows need'} a Subaccount `
                 + '— highlighted in the file.');
    }
    const rescued = _importFiles.reduce((n, f) => n + f.viaPacejet, 0);
    if (rescued) {
      lines.push(`${rescued === 1 ? '1 row resolved' : rescued + ' rows resolved'} `
                 + 'through PaceJet.');
    }
    // Sender-matched rows carry the customer's Subaccount but not a shipment, so
    // they are called out separately from the ones resolved by reference.
    const bySender = _importFiles.reduce((n, f) => n + f.viaSender, 0);
    if (bySender) {
      lines.push(`${bySender === 1 ? '1 row resolved' : bySender + ' rows resolved'} `
                 + 'by sender name.');
    }
    const wantsMaster = Boolean(s.parcelMasterCopy && added);
    if (_importFiles.length || wantsMaster) {
      lines.push(_importFiles.length === 1 && !wantsMaster
        ? 'Your Acumatica bill is ready below.'
        : 'Your files are ready below.');
    }

    document.getElementById('comp-note').innerHTML = lines.join('<br>');
    buildParcelDownloads(wantsMaster);

    showScreen('complete');
    fireNotification(added, !s.parcelMasterCopy);
    // The buttons live on this screen, so the auto-reset waits until every file
    // has been taken. Nothing to take — reset as before.
    if (_pendingDownloads) {
      compCountdown.textContent = 'Download your files, then close this screen.';
    } else {
      startResetCountdown();
    }
  }

  let _pendingDownloads = 0;

  // One button per generated bill, plus a zip of all of them. Automatic
  // downloads are not used: Chrome and Edge silently block every one after the
  // first, so only the first bill of a multi-carrier run ever arrived.
  function buildParcelDownloads(wantsMaster) {
    compDownloads.innerHTML = '';
    _pendingDownloads = 0;

    const ICON = '<svg viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round">'
               + '<path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg>';

    const addBtn = (label, sub, url, opts = {}) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'comp-dl-btn' + (opts.all ? ' all' : '') + (opts.cls ? ' ' + opts.cls : '');
      b.innerHTML = ICON + '<span>' + label + '</span>'
                  + (sub ? '<span class="comp-dl-sub">' + sub + '</span>' : '');
      if (!opts.optional) _pendingDownloads++;
      b.addEventListener('click', () => {
        triggerDownload(url, 0);
        if (!b.classList.contains('done')) {
          b.classList.add('done');
          if (!opts.optional && --_pendingDownloads === 0) startResetCountdown();
        }
      });
      compDownloads.appendChild(b);
      return b;
    };

    for (const f of _importFiles) {
      addBtn(f.label || f.file, f.rows ? f.rows + ' rows' : '',
             BASE + '/parcels-import-download?sid=' + encodeURIComponent(SID)
                  + '&file=' + encodeURIComponent(f.file),
             { cls: 'bill' });
    }
    // A workbook uploaded by hand has no shared copy to write back to, so the
    // updated file has to come along too.
    if (wantsMaster) {
      addBtn('Updated Small Parcels workbook', '',
             BASE + '/parcels-download?sid=' + encodeURIComponent(SID));
    }
    // Convenience on top of the per-file buttons — one click, one file, no
    // browser blocking. It does not count towards the pending files, since the
    // individual buttons already cover them.
    if (_importFiles.length > 1) {
      addBtn('Download all bills', '.zip',
             BASE + '/parcels-import-zip?sid=' + encodeURIComponent(SID),
             { all: true, optional: true })
        .addEventListener('click', () => {
          // The zip carries every bill, so those buttons are satisfied. The
          // hand-uploaded workbook is not in it and stays pending.
          compDownloads.querySelectorAll('.comp-dl-btn.bill:not(.done)').forEach(x => {
            x.classList.add('done');
            _pendingDownloads--;
          });
          if (_pendingDownloads === 0) startResetCountdown();
        });
    }
  }

  function sendParcelFile(file, isHelper) {
    hideError();
    const name = file.name.toLowerCase();
    if (!/\.(xlsx|xls|csv|pdf)$/.test(name)) {
      showError('Invalid File Type — only .xlsx, .xls, .csv and .pdf files are accepted.');
      _parcelAborted = true;
      return;
    }
    // Until the workbook is located, the only file worth accepting is the
    // workbook itself.
    if (!s.parcelMasterOk && !name.endsWith('.xlsx')) {
      showError('Small Parcels workbook not found — upload it before importing invoices.');
      _parcelAborted = true;
      return;
    }
    // The PaceJet export takes the pill and leaves the drop zone alone: the
    // upload is the whole job, and taking over the zone reads as an import
    // being processed when none is.
    if (isHelper) {
      showPill(pPacejetP);
      pPacejetP.classList.add('loading');
    } else {
      pzLoadTxt.textContent = 'Analyzing file…';
      pzDrop.classList.add('uploading');
    }

    const fd = new FormData();
    fd.append('sid', SID);
    fd.append('file', file);
    fd.append('scope', 'parcels');
    // Returned so sendParcelFiles can await each file before starting the next.
    return fetch(BASE + '/upload-auto', { method: 'POST', body: fd })
      .then(r => r.json())
      .then(data => {
        pzDrop.classList.remove('uploading');
        pPacejetP.classList.remove('loading');
        if (data.error) {
          showError(data.error, data.blocking, data.blocking);
          _parcelAborted = true;
          return;
        }

        pzDrop.classList.add('accepted');
        pzDrop.addEventListener('animationend', () => pzDrop.classList.remove('accepted'), { once: true });

        // Optional helper file — nothing to append. Every invoice imported after
        // it can use it to resolve rows Acumatica alone cannot.
        if (data.type === 'pacejet') {
          showPill(pPacejetP);
          markDone(pPacejetP);
          // Helper file only. Whether the invoice is still to come depends on
          // the rest of the batch, so the ask is raised once the batch ends.
          return;
        }

        // The workbook itself — the manual fallback. Nothing to append yet.
        if (data.type === 'sp_master') {
          s.parcelMasterOk   = true;
          s.parcelMasterCopy = true;
          markDone(pMasterP);
          pzHint.textContent = PZ_HINT_DEFAULT;
          return;
        }

        const kind = PARCEL_KINDS[data.type];
        if (!kind) {
          showError('Wrong File — this screen only accepts the WWEX raw export, '
                    + 'a FedEx invoice PDF, or the Small Parcels workbook.');
          _parcelAborted = true;
          return;
        }

        clearInvoiceNeeded();
        showPill(kind.pill);
        kind.pill.classList.add('loading');
        pzLoadTxt.textContent = 'Appending to the ' + kind.tab + ' tab…';
        const afd = new FormData();
        afd.append('sid', SID);
        afd.append('type', kind.type);
        return fetch(BASE + '/parcels-append', { method: 'POST', body: afd })
          // 404 = SMALL_PARCEL_ENABLED is off server-side; body isn't JSON
          .then(r => (r.status === 404
            ? { error: 'Small Parcels is disabled — enable it on the server.' }
            : r.json()))
          .then(wd => {
            kind.pill.classList.remove('loading');
            // Sticky: the append is the slow step, the user may have looked
            // away, and a 6-second toast is how this failure went silent.
            if (wd.error) {
              showError(wd.error, true, wd.blocking);
              _parcelAborted = true;
              return;
            }
            markDone(kind.pill);
            _parcelResults.push(wd);
            collectImportFile(wd);
          });
      })
      .catch(() => {
        pzDrop.classList.remove('uploading');
        pPacejetP.classList.remove('loading');
        pWWEXp.classList.remove('loading');
        pFedexP.classList.remove('loading');
        _parcelAborted = true;
        showError('Connection Error — could not reach the server. Please try again.');
      });
  }

  // ── Browser notification ──────────────────────────────────────
  function fireNotification(totalRecords, savedToOneDrive) {
    if (!('Notification' in window)) return;
    const body = savedToOneDrive
      ? `${totalRecords} records · Saved to Carrier Bills Import`
      : `${totalRecords} records · Downloaded to your device`;
    const send = () => new Notification('Freight Bill — Done', { body, icon: '' });
    if (Notification.permission === 'granted') {
      send();
    } else if (Notification.permission !== 'denied') {
      Notification.requestPermission().then(p => { if (p === 'granted') send(); });
    }
  }

  function markDone(pill) {
    pill.classList.remove('active', 'loading', 'popping');
    pill.classList.add('done');
    void pill.offsetWidth;
    pill.classList.add('popping');
    pill.addEventListener('animationend', () => pill.classList.remove('popping'), { once: true });
  }

  let _pendingPills = [];

  function setUploading(on, msg) {
    dropzone.classList.toggle('uploading', on);
    if (on) {
      if (msg) zoneLoadingTxt.textContent = msg;
      // Pills that are waiting for an upload → show loading spinner
      _pendingPills = [pCarrier, pP1, pPacejet].filter(p =>
        p.style.display !== 'none' && p.classList.contains('active')
      );
      _pendingPills.forEach(p => { p.classList.remove('active'); p.classList.add('loading'); });
    } else {
      // Restore any pill that wasn't resolved (markDone already removes loading)
      _pendingPills.forEach(p => {
        if (p.classList.contains('loading')) {
          p.classList.remove('loading'); p.classList.add('active');
        }
      });
      _pendingPills = [];
    }
  }

  function flashAccepted() {
    dropzone.classList.remove('uploading');
    dropzone.classList.add('accepted');
    dropzone.addEventListener('animationend', () => dropzone.classList.remove('accepted'), { once: true });
  }

  const errFixedToast  = document.getElementById('err-fixed-toast');
  const errToastTitle  = document.getElementById('err-toast-title');
  const errToastMsg    = document.getElementById('err-toast-msg');
  const errToastClose  = document.getElementById('err-toast-close');

  function _dismissErrToast() {
    leaveToast(errFixedToast);
  }

  errToastClose.addEventListener('click', _dismissErrToast);

  // Every toast lasts 6s (family standard), so `sticky` no longer keeps it up;
  // the argument stays so callers need not change.
  // blocking = presented as an action-required notice rather than a rejected
  //            upload: amber and wider. Used when the file was fine but the
  //            import cannot go ahead until someone resolves the message.
  function showError(msg, sticky, blocking) {
    // Split into title + detail if the message has a period or dash separator
    const sepIdx = msg.search(/[—–\-]\s/);
    if (sepIdx > 0) {
      errToastTitle.textContent = msg.slice(0, sepIdx).trim();
      const detail = msg.slice(sepIdx).replace(/^[—–\-]\s*/, '').trim();
      errToastMsg.textContent   = detail.charAt(0).toUpperCase() + detail.slice(1);
    } else {
      errToastTitle.textContent = 'File Not Accepted';
      const cap = msg.charAt(0).toUpperCase() + msg.slice(1);
      errToastMsg.textContent   = cap;
    }
    errFixedToast.classList.toggle('blocking', Boolean(blocking));
    errFixedToast.classList.remove('hiding');
    errFixedToast.classList.add('visible');
    armToast(errFixedToast, _dismissErrToast);
  }

  // Soft warning (file accepted, some tabs skipped): amber toast in the stack,
  // 6s like every toast.
  const warnToast = document.getElementById('warn-toast');
  const warnText  = document.getElementById('warn-text');
  document.getElementById('warn-close').addEventListener('click', () => leaveToast(warnToast));

  function showWarning(msg) {
    warnText.textContent = msg.charAt(0).toUpperCase() + msg.slice(1);
    warnToast.classList.remove('hiding');
    warnToast.classList.add('visible');
    armToast(warnToast, () => leaveToast(warnToast));
  }
  function hideError() {
    leaveToast(warnToast);
    if (errFixedToast.classList.contains('visible')) _dismissErrToast();
  }

})();
