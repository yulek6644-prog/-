/* Скан в Таблицу — сканер штрих-кодов с надёжной отправкой в Google Таблицу.
 *
 * Главный принцип: каждый скан СНАЧАЛА сохраняется в памяти телефона (IndexedDB),
 * и только потом отправляется. Запись помечается отправленной лишь после того,
 * как скрипт таблицы подтвердил её ID. Повторная отправка не создаёт дублей —
 * скрипт отбрасывает уже записанные ID.
 */
(function () {
  'use strict';

  var APP_VERSION = '1.0.0';
  var BATCH_SIZE = 100;
  var REQUEST_TIMEOUT_MS = 45000;
  var RETRY_MIN_MS = 5000;
  var RETRY_MAX_MS = 5 * 60 * 1000;
  var HISTORY_LIMIT = 300;

  // ---------- Настройки ----------

  var DEFAULTS = {
    url: '',
    token: '',
    sheet: 'Сканы',
    device: '',
    cooldown: 2,
    beep: true,
    vibrate: true,
    confirmQty: false,
    qtyReset: true
  };

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem('settings') || '{}'); } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = k in s ? s[k] : DEFAULTS[k]; });
    return out;
  }

  function saveSettings() {
    try { localStorage.setItem('settings', JSON.stringify(settings)); } catch (e) { /* приватный режим */ }
  }

  var settings = loadSettings();

  // ---------- База на телефоне (IndexedDB) ----------

  var dbPromise = null;

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var req = indexedDB.open('scan-to-sheet', 1);
      req.onupgradeneeded = function () {
        var store = req.result.createObjectStore('scans', { keyPath: 'id' });
        store.createIndex('status', 'status');
        store.createIndex('ts', 'ts');
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
    return dbPromise;
  }

  function store(mode) {
    return openDb().then(function (db) {
      return db.transaction('scans', mode).objectStore('scans');
    });
  }

  function reqToPromise(req) {
    return new Promise(function (resolve, reject) {
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  // Запись считается сохранённой только после завершения транзакции.
  function putAll(records) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction('scans', 'readwrite');
        var st = tx.objectStore('scans');
        records.forEach(function (r) { st.put(r); });
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
        tx.onabort = function () { reject(tx.error); };
      });
    });
  }

  function getAll() {
    return store('readonly').then(function (st) { return reqToPromise(st.getAll()); });
  }

  function getByStatus(status) {
    return store('readonly').then(function (st) {
      return reqToPromise(st.index('status').getAll(status));
    });
  }

  function countByStatus(status) {
    return store('readonly').then(function (st) {
      return reqToPromise(st.index('status').count(status));
    });
  }

  function deleteIds(ids) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction('scans', 'readwrite');
        var st = tx.objectStore('scans');
        ids.forEach(function (id) { st.delete(id); });
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  // ---------- Утилиты ----------

  var $ = function (id) { return document.getElementById(id); };

  function uuid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    var b = new Uint8Array(16);
    crypto.getRandomValues(b);
    b[6] = (b[6] & 15) | 64;
    b[8] = (b[8] & 63) | 128;
    var h = Array.prototype.map.call(b, function (x) { return (x + 256).toString(16).slice(1); }).join('');
    return h.slice(0, 8) + '-' + h.slice(8, 12) + '-' + h.slice(12, 16) + '-' + h.slice(16, 20) + '-' + h.slice(20);
  }

  function fmtTime(iso) {
    var d = new Date(iso);
    return d.toLocaleDateString('ru-RU') + ' ' + d.toLocaleTimeString('ru-RU');
  }

  var toastTimer = null;
  function toast(msg, isError) {
    var t = $('toast');
    t.textContent = msg;
    t.className = 'toast show' + (isError ? ' error' : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.className = 'toast'; }, isError ? 4500 : 2200);
  }

  var audioCtx = null;
  function beep(ok) {
    if (!settings.beep) return;
    try {
      audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
      if (audioCtx.state === 'suspended') audioCtx.resume();
      var o = audioCtx.createOscillator();
      var g = audioCtx.createGain();
      o.type = 'square';
      o.frequency.value = ok ? 1800 : 300;
      g.gain.value = 0.08;
      o.connect(g);
      g.connect(audioCtx.destination);
      o.start();
      o.stop(audioCtx.currentTime + (ok ? 0.09 : 0.3));
    } catch (e) { /* звук недоступен */ }
  }

  function vibrate(pattern) {
    if (settings.vibrate && navigator.vibrate) {
      try { navigator.vibrate(pattern); } catch (e) { /* нет вибро */ }
    }
  }

  function flash() {
    var f = $('flash');
    f.classList.add('on');
    setTimeout(function () { f.classList.remove('on'); }, 60);
  }

  // ---------- Сохранение скана ----------

  var lastCode = null;
  var lastCodeAt = 0;
  var persistAsked = false;

  function addScan(code, format, source) {
    code = String(code || '').replace(/[\r\n\t]+/g, '').trim();
    if (!code) return Promise.resolve();

    var now = Date.now();
    var cooldownMs = Math.max(0, Number(settings.cooldown) || 0) * 1000;
    if (source === 'camera' && code === lastCode && now - lastCodeAt < cooldownMs) {
      return Promise.resolve();
    }
    lastCode = code;
    lastCodeAt = now;

    var qty = parseFloat(String($('qty').value).replace(',', '.'));
    if (!isFinite(qty) || qty <= 0) qty = 1;

    if (settings.confirmQty) {
      var answer = window.prompt('Количество для ' + code, String(qty));
      if (answer === null) return Promise.resolve(); // отмена — не сохраняем
      var q = parseFloat(String(answer).replace(',', '.'));
      if (isFinite(q) && q > 0) qty = q;
    }

    var rec = {
      id: uuid(),
      code: code,
      format: format || '',
      qty: qty,
      note: $('note').value.trim(),
      device: settings.device || '',
      sheet: settings.sheet || DEFAULTS.sheet,
      ts: new Date().toISOString(),
      status: 'pending',
      tries: 0,
      error: ''
    };

    if (!persistAsked && navigator.storage && navigator.storage.persist) {
      persistAsked = true;
      navigator.storage.persist().catch(function () {});
    }

    return putAll([rec]).then(function () {
      beep(true);
      vibrate(80);
      flash();
      showLast(rec);
      if (settings.qtyReset) $('qty').value = 1;
      refreshCounts();
      if (currentPage === 'history') renderHistory();
      scheduleSync(300);
    }).catch(function (err) {
      beep(false);
      vibrate([200, 100, 200]);
      toast('Не удалось сохранить скан на телефоне: ' + (err && err.message || err), true);
    });
  }

  function showLast(rec) {
    $('lastScan').hidden = false;
    $('lastCode').textContent = rec.code;
    $('lastMeta').textContent = 'Кол-во: ' + rec.qty + (rec.format ? ' · ' + rec.format : '') +
      ' · ' + new Date(rec.ts).toLocaleTimeString('ru-RU') + ' · сохранено, отправляется…';
  }

  // ---------- Отправка в таблицу ----------

  var syncing = false;
  var syncTimer = null;
  var retryDelay = RETRY_MIN_MS;
  var lastSyncError = '';
  var lastSyncOkAt = 0;

  function scheduleSync(ms) {
    clearTimeout(syncTimer);
    syncTimer = setTimeout(sync, ms);
  }

  function isConfigured() {
    return /^https:\/\/script\.google(usercontent)?\.com\//.test(settings.url || '');
  }

  function postJson(payload) {
    var ctrl = window.AbortController ? new AbortController() : null;
    var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, REQUEST_TIMEOUT_MS) : null;
    // text/plain — чтобы не было CORS preflight, который Apps Script не поддерживает.
    return fetch(settings.url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow',
      cache: 'no-store',
      signal: ctrl ? ctrl.signal : undefined
    }).then(function (res) {
      return res.text().then(function (text) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        try {
          return JSON.parse(text);
        } catch (e) {
          throw new Error('Скрипт ответил не JSON. Проверьте, что доступ веб-приложения — «Все».');
        }
      });
    }).catch(function (err) {
      if (err && err.name === 'AbortError') throw new Error('Нет ответа от Google (таймаут)');
      if (err instanceof TypeError) throw new Error('Нет связи с Google');
      throw err;
    }).finally(function () {
      if (timer) clearTimeout(timer);
    });
  }

  function sync(manual) {
    if (syncing) return Promise.resolve();
    if (!isConfigured()) {
      lastSyncError = 'Не указан адрес скрипта (Настройки)';
      updateStatus();
      if (manual) toast(lastSyncError, true);
      return Promise.resolve();
    }
    if (!navigator.onLine && !manual) {
      updateStatus();
      return Promise.resolve();
    }

    syncing = true;
    updateStatus();

    return getByStatus('pending').then(function (pending) {
      if (!pending.length) return { done: true };
      pending.sort(function (a, b) { return a.ts < b.ts ? -1 : 1; });
      var sheet = pending[0].sheet;
      var batch = pending.filter(function (r) { return r.sheet === sheet; }).slice(0, BATCH_SIZE);

      var payload = {
        token: settings.token || '',
        sheet: sheet,
        rows: batch.map(function (r) {
          return { id: r.id, code: r.code, qty: r.qty, format: r.format, device: r.device, note: r.note, ts: r.ts };
        })
      };

      return postJson(payload).then(function (resp) {
        if (!resp || resp.ok !== true) {
          throw new Error((resp && resp.error) || 'Неизвестная ошибка скрипта');
        }
        var confirmed = {};
        (resp.saved || []).concat(resp.duplicates || []).forEach(function (id) { confirmed[id] = true; });
        var sentAt = new Date().toISOString();
        var updated = [];
        batch.forEach(function (r) {
          if (confirmed[r.id]) {
            r.status = 'sent';
            r.sentAt = sentAt;
            r.error = '';
            updated.push(r);
          }
        });
        return putAll(updated).then(function () {
          return { done: pending.length <= updated.length, sent: updated.length, progressed: updated.length > 0 };
        });
      }).catch(function (err) {
        var msg = (err && err.message) || String(err);
        batch.forEach(function (r) { r.tries = (r.tries || 0) + 1; r.error = msg; });
        return putAll(batch).then(function () { throw err; });
      });
    }).then(function (result) {
      syncing = false;
      lastSyncError = '';
      lastSyncOkAt = Date.now();
      retryDelay = RETRY_MIN_MS;
      refreshCounts();
      if (currentPage === 'history') renderHistory();
      if (result && !result.done && result.progressed) {
        scheduleSync(200); // есть ещё — отправляем следующую пачку
      } else if (result && !result.done) {
        lastSyncError = 'Скрипт не подтвердил запись. Повтор…';
        scheduleSync(retryDelay);
      } else if (manual) {
        toast('Всё отправлено ✔');
      }
    }).catch(function (err) {
      syncing = false;
      lastSyncError = (err && err.message) || String(err);
      refreshCounts();
      if (currentPage === 'history') renderHistory();
      if (manual) toast('Ошибка отправки: ' + lastSyncError + '. Данные сохранены, повторим.', true);
      scheduleSync(retryDelay);
      retryDelay = Math.min(retryDelay * 2, RETRY_MAX_MS);
    });
  }

  // ---------- Индикаторы ----------

  var pendingCount = 0;

  function refreshCounts() {
    return Promise.all([countByStatus('pending'), countByStatus('sent')]).then(function (c) {
      pendingCount = c[0];
      $('pendingBadge').textContent = '⏳ ' + c[0];
      $('pendingBadge').classList.toggle('warn', c[0] > 0);
      $('sentBadge').textContent = '✔ ' + c[1];
      updateStatus();
    }).catch(function () {});
  }

  function updateStatus() {
    var dot = $('netDot');
    var txt = $('netText');
    var info;
    if (!isConfigured()) {
      dot.className = 'dot offline';
      txt.textContent = 'Не настроено';
      info = 'Укажите адрес скрипта в Настройках. Сканы сохраняются и будут отправлены позже.';
    } else if (!navigator.onLine) {
      dot.className = 'dot offline';
      txt.textContent = 'Нет сети';
      info = 'Нет интернета. Сканы сохраняются на телефоне и уйдут автоматически.';
    } else if (syncing) {
      dot.className = 'dot online';
      txt.textContent = 'Отправка…';
      info = 'Идёт отправка…';
    } else if (lastSyncError) {
      dot.className = 'dot offline';
      txt.textContent = 'Ошибка отправки';
      info = 'Последняя ошибка: ' + lastSyncError + '. Повтор автоматически.';
    } else {
      dot.className = 'dot online';
      txt.textContent = pendingCount ? 'Онлайн' : 'Всё отправлено';
      info = lastSyncOkAt ? 'Последняя успешная отправка: ' + new Date(lastSyncOkAt).toLocaleTimeString('ru-RU') : '';
    }
    $('syncInfo').textContent = info;
  }

  // ---------- История ----------

  function renderHistory() {
    return getAll().then(function (all) {
      all.sort(function (a, b) { return a.ts < b.ts ? 1 : -1; });
      var list = $('historyList');
      list.textContent = '';
      $('historyEmpty').hidden = all.length > 0;
      all.slice(0, HISTORY_LIMIT).forEach(function (r) {
        var li = document.createElement('li');
        var ic = document.createElement('span');
        ic.className = 'ic';
        ic.textContent = r.status === 'sent' ? '✅' : '⏳';
        ic.title = r.status === 'sent' ? 'В таблице' : 'Ждёт отправки';
        var body = document.createElement('div');
        body.className = 'body';
        var code = document.createElement('div');
        code.className = 'code';
        code.textContent = r.code + (r.qty !== 1 ? '  × ' + r.qty : '');
        var meta = document.createElement('div');
        meta.className = 'meta';
        meta.textContent = fmtTime(r.ts) + (r.format ? ' · ' + r.format : '') + ' · лист «' + r.sheet + '»' + (r.note ? ' · ' + r.note : '');
        body.appendChild(code);
        body.appendChild(meta);
        if (r.status !== 'sent' && r.error) {
          var e = document.createElement('div');
          e.className = 'err';
          e.textContent = 'Попыток: ' + r.tries + ' · ' + r.error;
          body.appendChild(e);
        }
        var del = document.createElement('button');
        del.className = 'del';
        del.textContent = '✕';
        del.title = 'Удалить с телефона';
        del.onclick = function () {
          var warn = r.status === 'sent'
            ? 'Удалить запись только с телефона? В таблице она останется.'
            : 'Этот скан ЕЩЁ НЕ ОТПРАВЛЕН. Удалить его насовсем?';
          if (confirm(warn)) deleteIds([r.id]).then(function () { refreshCounts(); renderHistory(); });
        };
        li.appendChild(ic);
        li.appendChild(body);
        li.appendChild(del);
        list.appendChild(li);
      });
      if (all.length > HISTORY_LIMIT) {
        var more = document.createElement('li');
        more.className = 'hint';
        more.textContent = 'Показаны последние ' + HISTORY_LIMIT + ' из ' + all.length + '. Полный список — в CSV.';
        list.appendChild(more);
      }
    });
  }

  function exportCsv() {
    getAll().then(function (all) {
      all.sort(function (a, b) { return a.ts < b.ts ? -1 : 1; });
      var esc = function (v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; };
      var lines = [['Дата и время', 'Штрих-код', 'Количество', 'Тип', 'Устройство', 'Комментарий', 'Лист', 'Статус', 'ID'].map(esc).join(';')];
      all.forEach(function (r) {
        // ="…" — чтобы Excel не превращал длинные коды в 4,6E+12
        lines.push([fmtTime(r.ts), '=' + esc(r.code), r.qty, r.format, r.device, r.note, r.sheet,
          r.status === 'sent' ? 'отправлен' : 'не отправлен', r.id].map(function (v, i) {
          return i === 1 ? v : esc(v);
        }).join(';'));
      });
      var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'scans-' + new Date().toISOString().slice(0, 10) + '.csv';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    });
  }

  function clearSent() {
    getByStatus('sent').then(function (sent) {
      if (!sent.length) { toast('Нет отправленных записей'); return; }
      if (!confirm('Убрать с телефона ' + sent.length + ' отправленных записей? В таблице они останутся.')) return;
      deleteIds(sent.map(function (r) { return r.id; })).then(function () {
        refreshCounts();
        renderHistory();
      });
    });
  }

  // ---------- Камера и распознавание ----------

  var detector = null;
  var stream = null;
  var scanning = false;
  var detectBusy = false;
  var torchOn = false;
  var wakeLock = null;
  var resumeOnShow = false;

  var FORMATS = ['ean_13', 'ean_8', 'upc_a', 'upc_e', 'code_128', 'code_39', 'code_93', 'codabar', 'itf',
    'qr_code', 'data_matrix', 'pdf417', 'aztec'];

  function createDetector() {
    if (detector) return Promise.resolve(detector);

    function useZxing() {
      var api = window.BarcodeDetectionAPI;
      if (!api) throw new Error('Модуль распознавания не загрузился');
      api.prepareZXingModule({
        overrides: {
          locateFile: function (path, prefix) {
            return path.slice(-5) === '.wasm' ? new URL('vendor/' + path, location.href).href : prefix + path;
          }
        },
        fireImmediately: true
      });
      return api.BarcodeDetector.getSupportedFormats().then(function (sup) {
        detector = new api.BarcodeDetector({ formats: FORMATS.filter(function (f) { return sup.indexOf(f) >= 0; }) });
        return detector;
      });
    }

    // Встроенный распознаватель Android/Chrome — быстрее. Иначе — ZXing (iPhone и др.).
    if ('BarcodeDetector' in window && window.BarcodeDetector.getSupportedFormats) {
      return window.BarcodeDetector.getSupportedFormats().then(function (sup) {
        var fmts = FORMATS.filter(function (f) { return sup.indexOf(f) >= 0; });
        if (fmts.indexOf('ean_13') < 0 || fmts.indexOf('code_128') < 0) return useZxing();
        detector = new window.BarcodeDetector({ formats: fmts });
        return detector;
      }).catch(useZxing);
    }
    return Promise.resolve().then(useZxing);
  }

  function startCamera() {
    if (scanning) return Promise.resolve();
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      $('cameraMsg').textContent = 'Камера недоступна. Откройте приложение по https и разрешите доступ к камере.';
      $('cameraMsg').hidden = false;
      return Promise.resolve();
    }
    $('cameraMsg').textContent = 'Включаю камеру…';
    $('cameraMsg').hidden = false;
    $('btnStart').disabled = true;

    return createDetector().then(function () {
      return navigator.mediaDevices.getUserMedia({
        audio: false,
        video: {
          facingMode: { ideal: 'environment' },
          width: { ideal: 1280 },
          height: { ideal: 720 }
        }
      });
    }).then(function (s) {
      stream = s;
      var video = $('video');
      video.srcObject = s;
      return video.play();
    }).then(function () {
      scanning = true;
      $('cameraMsg').hidden = true;
      $('btnStart').textContent = '■ Стоп';
      $('btnStart').classList.remove('primary');
      setupTrack();
      requestWakeLock();
      loop();
    }).catch(function (err) {
      stopCamera();
      var name = err && err.name;
      $('cameraMsg').textContent = name === 'NotAllowedError'
        ? 'Нет доступа к камере. Разрешите камеру в настройках браузера для этого сайта.'
        : 'Не удалось включить камеру: ' + ((err && err.message) || err);
      $('cameraMsg').hidden = false;
    }).finally(function () {
      $('btnStart').disabled = false;
    });
  }

  function setupTrack() {
    var track = stream && stream.getVideoTracks()[0];
    if (!track || !track.getCapabilities) return;
    var caps = {};
    try { caps = track.getCapabilities(); } catch (e) { caps = {}; }
    if (caps.focusMode && caps.focusMode.indexOf('continuous') >= 0) {
      track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(function () {});
    }
    $('btnTorch').hidden = !caps.torch;
  }

  function toggleTorch() {
    var track = stream && stream.getVideoTracks()[0];
    if (!track) return;
    torchOn = !torchOn;
    track.applyConstraints({ advanced: [{ torch: torchOn }] }).catch(function () { torchOn = false; });
  }

  function stopCamera() {
    scanning = false;
    if (stream) {
      stream.getTracks().forEach(function (t) { t.stop(); });
      stream = null;
    }
    $('video').srcObject = null;
    $('btnStart').textContent = '▶ Старт';
    $('btnStart').classList.add('primary');
    $('btnTorch').hidden = true;
    torchOn = false;
    $('cameraMsg').textContent = 'Камера выключена';
    $('cameraMsg').hidden = false;
    if (wakeLock) { wakeLock.release().catch(function () {}); wakeLock = null; }
  }

  function requestWakeLock() {
    if (!('wakeLock' in navigator)) return;
    navigator.wakeLock.request('screen').then(function (l) { wakeLock = l; }).catch(function () {});
  }

  function loop() {
    if (!scanning) return;
    var video = $('video');
    if (!detectBusy && video.readyState >= 2 && video.videoWidth) {
      detectBusy = true;
      detector.detect(video).then(function (codes) {
        if (codes && codes.length && scanning) {
          var c = codes[0];
          return addScan(c.rawValue, c.format, 'camera');
        }
      }).catch(function () { /* пропускаем неудачный кадр */ }).finally(function () {
        detectBusy = false;
      });
    }
    setTimeout(loop, 120);
  }

  // ---------- Настройки: экран и ссылка для других ----------

  function fillSettingsForm() {
    $('setUrl').value = settings.url;
    $('setToken').value = settings.token;
    $('setSheet').value = settings.sheet;
    $('setDevice').value = settings.device;
    $('setCooldown').value = settings.cooldown;
    $('setBeep').checked = !!settings.beep;
    $('setVibrate').checked = !!settings.vibrate;
    $('setConfirm').checked = !!settings.confirmQty;
    $('qtyReset').checked = !!settings.qtyReset;
  }

  function readSettingsForm() {
    settings.url = $('setUrl').value.trim();
    settings.token = $('setToken').value.trim();
    settings.sheet = $('setSheet').value.trim() || DEFAULTS.sheet;
    settings.device = $('setDevice').value.trim();
    var cd = parseFloat(String($('setCooldown').value).replace(',', '.'));
    settings.cooldown = isFinite(cd) && cd >= 0 ? cd : DEFAULTS.cooldown;
    settings.beep = $('setBeep').checked;
    settings.vibrate = $('setVibrate').checked;
    settings.confirmQty = $('setConfirm').checked;
  }

  function testConnection() {
    readSettingsForm();
    if (!isConfigured()) {
      toast('Адрес должен начинаться с https://script.google.com/', true);
      return;
    }
    toast('Проверяю…');
    var url = settings.url + (settings.url.indexOf('?') >= 0 ? '&' : '?') + 'token=' + encodeURIComponent(settings.token);
    fetch(url, { redirect: 'follow', cache: 'no-store' }).then(function (r) { return r.text(); }).then(function (t) {
      var j;
      try { j = JSON.parse(t); } catch (e) { throw new Error('Ответ не JSON — проверьте доступ «Все» и что URL заканчивается на /exec'); }
      if (!j.ok) throw new Error(j.error || 'ошибка');
      toast('Связь есть ✔ Таблица: «' + j.spreadsheet + '»');
    }).catch(function (err) {
      toast('Нет связи: ' + ((err && err.message) || err), true);
    });
  }

  function b64urlEncode(str) {
    return btoa(unescape(encodeURIComponent(str))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function b64urlDecode(str) {
    str = str.replace(/-/g, '+').replace(/_/g, '/');
    while (str.length % 4) str += '=';
    return decodeURIComponent(escape(atob(str)));
  }

  function shareLink() {
    readSettingsForm();
    var cfg = { url: settings.url, token: settings.token, sheet: settings.sheet };
    var link = location.origin + location.pathname + '#cfg=' + b64urlEncode(JSON.stringify(cfg));
    var ta = $('shareLink');
    ta.value = link;
    ta.hidden = false;
    if (navigator.share) {
      navigator.share({ title: 'Скан в Таблицу', text: 'Сканер штрих-кодов в нашу таблицу', url: link }).catch(function () {});
    } else if (navigator.clipboard) {
      navigator.clipboard.writeText(link).then(function () { toast('Ссылка скопирована'); }, function () { ta.select(); });
    } else {
      ta.select();
    }
  }

  function applyConfigFromLink() {
    var m = location.hash.match(/cfg=([A-Za-z0-9_-]+)/);
    if (!m) return;
    try {
      var cfg = JSON.parse(b64urlDecode(m[1]));
      if (cfg && cfg.url && confirm('Применить настройки таблицы из ссылки?')) {
        settings.url = String(cfg.url);
        settings.token = String(cfg.token || '');
        settings.sheet = String(cfg.sheet || DEFAULTS.sheet);
        saveSettings();
        if (!settings.device) {
          var name = prompt('Как подписывать сканы с этого телефона? (имя или номер устройства)', '');
          if (name) { settings.device = name.trim(); saveSettings(); }
        }
        toast('Настройки применены');
      }
    } catch (e) {
      toast('Ссылка с настройками повреждена', true);
    }
    history.replaceState(null, '', location.pathname + location.search);
  }

  // ---------- Навигация ----------

  var currentPage = 'scan';

  function showPage(name) {
    currentPage = name;
    document.querySelectorAll('.page').forEach(function (p) { p.classList.toggle('active', p.id === 'page-' + name); });
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.getAttribute('data-page') === name); });
    if (name === 'history') renderHistory();
    if (name === 'settings') fillSettingsForm();
  }

  // ---------- Запуск ----------

  function init() {
    applyConfigFromLink();
    fillSettingsForm();
    $('versionInfo').textContent = 'Версия ' + APP_VERSION;

    document.querySelectorAll('.tab').forEach(function (t) {
      t.onclick = function () { showPage(t.getAttribute('data-page')); };
    });

    $('btnStart').onclick = function () {
      // Звук на iOS можно включить только по нажатию — инициализируем здесь.
      try {
        audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
        audioCtx.resume();
      } catch (e) { /* нет звука */ }
      if (scanning) { resumeOnShow = false; stopCamera(); } else { startCamera(); }
    };
    $('btnTorch').onclick = toggleTorch;

    $('manualForm').onsubmit = function (e) {
      e.preventDefault();
      var v = $('manualCode').value;
      addScan(v, 'вручную', 'manual').then(function () { $('manualCode').value = ''; });
    };

    $('qtyReset').onchange = function () { settings.qtyReset = this.checked; saveSettings(); };

    $('btnSync').onclick = function () { clearTimeout(syncTimer); retryDelay = RETRY_MIN_MS; sync(true); };
    $('btnExport').onclick = exportCsv;
    $('btnClearSent').onclick = clearSent;

    $('btnSave').onclick = function () {
      readSettingsForm();
      saveSettings();
      toast('Сохранено');
      retryDelay = RETRY_MIN_MS;
      updateStatus();
      scheduleSync(100);
    };
    $('btnTest').onclick = testConnection;
    $('btnShare').onclick = shareLink;

    window.addEventListener('online', function () { retryDelay = RETRY_MIN_MS; updateStatus(); scheduleSync(500); });
    window.addEventListener('offline', updateStatus);

    // Камеру освобождаем при сворачивании, при возвращении — включаем снова.
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) {
        if (scanning) { resumeOnShow = true; stopCamera(); }
      } else {
        if (resumeOnShow) { resumeOnShow = false; startCamera(); }
        scheduleSync(500);
      }
    });

    // Страховочная периодическая попытка — вдруг таймер повтора был потерян.
    setInterval(function () { if (!syncing && pendingCount > 0) sync(); }, 60000);

    refreshCounts().then(function () { scheduleSync(1000); });

    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    }
  }

  if (!window.indexedDB) {
    document.body.innerHTML = '<p style="padding:16px">Этот браузер не поддерживает локальное хранилище. Откройте в Chrome или Safari.</p>';
    return;
  }
  init();
})();
