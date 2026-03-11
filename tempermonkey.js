// ==UserScript==
// @name         NovelAI Vibe Batch Commit-Strict
// @namespace    local.nai.vibe.batch.commitstrict
// @version      1.0.0
// @description  Strict per-card vibe extraction/downloading with commit verification for long virtualized lists
// @match        https://novelai.net/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(() => {
  'use strict';

  const CONFIG = {
    defaultValues: '0.01',
    pollMs: 250,
    shortMs: 100,
    afterScrollMs: 180,
    afterFocusMs: 120,
    afterSetMs: 700,
    afterActionMs: 1200,
    extractTimeoutMs: 180000,
    scanStepRatio: 0.72,
    maxScanLoops: 400,
    setCommitRetries: 4,
    extractRetries: 2,
    continueOnError: true,
  };

  let running = false;
  let stopRequested = false;
  let panel = null;
  let pendingDownloadPrefix = '';

  // -----------------------------
  // download filename prefix
  // -----------------------------
  (function patchAnchorDownloadName() {
    try {
      const originalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function (...args) {
        try {
          if (pendingDownloadPrefix && this.hasAttribute('download')) {
            const oldName = this.getAttribute('download') || '';
            const prefix = String(pendingDownloadPrefix).replace(/[^\d.]/g, '');
            if (oldName && prefix && !oldName.startsWith(prefix + '_')) {
              this.setAttribute('download', `${prefix}_${oldName}`);
            }
          }
        } catch (err) {
          console.warn('[NAI CommitStrict] rename hook failed', err);
        }
        return originalClick.apply(this, args);
      };
    } catch (err) {
      console.warn('[NAI CommitStrict] anchor patch failed', err);
    }
  })();

  // -----------------------------
  // generic helpers
  // -----------------------------
  function log(...args) {
    console.log('[NAI CommitStrict]', ...args);
    setStatus(args.join(' '));
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function textOf(el) {
    return (el?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function isVisible(el) {
    if (!el || !(el instanceof Element)) return false;
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return (
      rect.width > 0 &&
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden'
    );
  }

  function makeEl(tag, props = {}, style = {}) {
    const el = document.createElement(tag);
    Object.assign(el, props);
    Object.assign(el.style, style);
    return el;
  }

  function formatValue(n) {
    return Number(n).toFixed(2);
  }

  function clamp01(n) {
    return Math.max(0.01, Math.min(1, Number(n)));
  }

  function approxEqual(a, b, eps = 0.011) {
    return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < eps;
  }

  function parseValues(raw) {
    return String(raw)
      .split(/[,\s]+/)
      .map(s => s.trim())
      .filter(Boolean)
      .map(Number)
      .filter(n => Number.isFinite(n) && n >= 0.01 && n <= 1)
      .map(n => Number(n.toFixed(2)));
  }

  function setStatus(msg) {
    if (panel?.status) panel.status.textContent = msg;
  }

  function syncPanelState() {
    if (!panel) return;
    panel.start.disabled = running;
    panel.stop.disabled = !running;
  }

  function setNativeValue(input, value) {
    const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (desc?.set) {
      desc.set.call(input, String(value));
    } else {
      input.value = String(value);
    }

    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new InputEvent('input', { bubbles: true, data: String(value) }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function fireRealClick(el) {
    if (!el) return;
    el.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  }

  async function waitUntil(fn, timeoutMs, intervalMs = CONFIG.pollMs) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (stopRequested) throw new Error('已手动停止');
      const result = await fn();
      if (result) return result;
      await sleep(intervalMs);
    }
    throw new Error('等待超时');
  }

  // -----------------------------
  // card detection / virtual list
  // -----------------------------
  function looksLikeCard(el) {
    if (!el || !(el instanceof Element)) return false;
    if (!isVisible(el)) return false;

    const txt = textOf(el);
    if (!txt.includes('Reference Strength')) return false;
    if (!txt.includes('Information Extracted')) return false;

    const idInputs = Array.from(el.querySelectorAll('input[type="text"]')).filter(isVisible);
    const numberInputs = Array.from(el.querySelectorAll('input[type="number"]')).filter(isVisible);
    const rangeInputs = Array.from(el.querySelectorAll('input[type="range"]')).filter(isVisible);

    if (idInputs.length !== 1) return false;
    if (numberInputs.length < 2) return false;
    if (rangeInputs.length < 2) return false;

    return true;
  }

  function findNearestCard(startEl) {
    let el = startEl;
    while (el && el !== document.body) {
      if (looksLikeCard(el)) return el;
      el = el.parentElement;
    }
    return null;
  }

  function getVisibleCardMap() {
    const map = new Map();
    const idInputs = Array.from(document.querySelectorAll('input[type="text"]')).filter(isVisible);

    for (const input of idInputs) {
      const id = (input.value || '').trim();
      if (!id) continue;
      if (map.has(id)) continue;

      const card = findNearestCard(input);
      if (!card) continue;
      if (!isVisible(card)) continue;

      map.set(id, card);
    }

    return map;
  }

  function collectVisibleCards() {
    return Array.from(getVisibleCardMap(), ([id, card]) => ({ id, card }));
  }

  function findScrollParent(startEl) {
    const candidates = [];

    let el = startEl;
    while (el && el !== document.body) {
      const style = getComputedStyle(el);
      const overflowY = style.overflowY;
      const canScroll = (overflowY === 'auto' || overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 40;
      if (canScroll) {
        candidates.push(el);
      }
      el = el.parentElement;
    }

    // 优先在卡片祖先链里选滚动容器，避免被页面整体滚动条“抢走”。
    if (candidates.length) {
      // 嵌套 overflow 时，选择可滚动距离最大的祖先容器。
      candidates.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));
      return candidates[0];
    }

    return document.scrollingElement || document.documentElement;
  }

  async function findVibeListContainer() {
    const visible = collectVisibleCards();
    if (!visible.length) {
      throw new Error('当前页面没有找到任何可见的 Vibe 卡片');
    }
    return findScrollParent(visible[0].card);
  }

  async function harvestAllCardIds() {
    const list = await findVibeListContainer();
    const seen = new Set();

    const originalTop = list.scrollTop;
    list.scrollTop = 0;
    await sleep(CONFIG.afterScrollMs);

    for (let i = 0; i < CONFIG.maxScanLoops; i++) {
      const visibleMap = getVisibleCardMap();
      for (const id of visibleMap.keys()) {
        seen.add(id);
      }

      const atBottom = list.scrollTop + list.clientHeight >= list.scrollHeight - 4;
      if (atBottom) break;

      const step = Math.max(120, Math.floor(list.clientHeight * CONFIG.scanStepRatio));
      const nextTop = Math.min(list.scrollTop + step, list.scrollHeight);

      if (nextTop === list.scrollTop) break;
      list.scrollTop = nextTop;
      await sleep(CONFIG.afterScrollMs);
    }

    list.scrollTop = originalTop;
    await sleep(CONFIG.afterScrollMs);

    return { ids: Array.from(seen), list };
  }

  async function ensureCardVisible(id, list) {
    const tryVisible = () => getVisibleCardMap().get(id) || null;

    let card = tryVisible();
    if (card) return card;

    const originalTop = list.scrollTop;
    list.scrollTop = 0;
    await sleep(CONFIG.afterScrollMs);

    for (let i = 0; i < CONFIG.maxScanLoops; i++) {
      card = tryVisible();
      if (card) return card;

      const atBottom = list.scrollTop + list.clientHeight >= list.scrollHeight - 4;
      if (atBottom) break;

      const step = Math.max(120, Math.floor(list.clientHeight * CONFIG.scanStepRatio));
      const nextTop = Math.min(list.scrollTop + step, list.scrollHeight);

      if (nextTop === list.scrollTop) break;
      list.scrollTop = nextTop;
      await sleep(CONFIG.afterScrollMs);
    }

    list.scrollTop = originalTop;
    await sleep(CONFIG.afterScrollMs);

    throw new Error(`无法在列表中定位卡片 ${id}`);
  }

  async function waitForCardState(id, list, predicate, timeoutMs, message) {
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      if (stopRequested) throw new Error('已手动停止');
      const card = await ensureCardVisible(id, list);
      if (predicate(card)) return card;
      await sleep(CONFIG.pollMs);
    }

    throw new Error(message || `等待卡片 ${id} 状态超时`);
  }

  // -----------------------------
  // card internals
  // -----------------------------
  function getInfoControls(card) {
    const numbers = Array.from(card.querySelectorAll('input[type="number"]')).filter(isVisible);
    const ranges = Array.from(card.querySelectorAll('input[type="range"]')).filter(isVisible);

    return {
      number: numbers[1] || null,
      range: ranges[1] || null,
    };
  }

  function getCurrentInfo(card) {
    const { number, range } = getInfoControls(card);
    const raw = number?.value ?? range?.value;
    const n = Number(raw);
    return Number.isFinite(n) ? n : NaN;
  }

  function isPending(card) {
    return textOf(card).includes('Encoding required');
  }

  function getHeaderButtons(card) {
    const header = card?.firstElementChild;
    if (!header) return [];
    return Array.from(header.querySelectorAll('button')).filter(isVisible);
  }

  // header buttons usually [trash, action, check]
  function getActionButton(card) {
    const buttons = getHeaderButtons(card);
    return buttons[1] || null;
  }

  function getActionMode(card) {
    const btn = getActionButton(card);
    if (!btn) return 'unknown';

    const txt = textOf(btn);

    if (/anlas/i.test(txt)) return 'extract';
    if (/\d/.test(txt) && txt.length <= 12) return 'extract';

    return 'download';
  }

  function isDownloadReady(card) {
    return getActionMode(card) === 'download';
  }

  function getCardFocusTarget(card) {
    return (
      card.querySelector('input[type="text"]') ||
      card.querySelector('input[type="number"]') ||
      card.querySelector('input[type="range"]') ||
      card
    );
  }

  async function focusCard(card, list) {
    if (list && card) {
      card.scrollIntoView({ block: 'center', inline: 'nearest' });
      await sleep(CONFIG.afterScrollMs);
    }

    const target = getCardFocusTarget(card);
    if (target) {
      fireRealClick(target);
      await sleep(CONFIG.afterFocusMs);
      target.focus?.();
      await sleep(CONFIG.afterFocusMs);
    }
  }

  async function jiggleCard(card, list) {
    if (!list || !card) return;
    const before = list.scrollTop;
    list.scrollTop = Math.max(0, before - 80);
    await sleep(CONFIG.afterScrollMs);
    card.scrollIntoView({ block: 'center', inline: 'nearest' });
    await sleep(CONFIG.afterScrollMs);
  }

  // -----------------------------
  // setting value with strict commit
  // -----------------------------
  async function applyInfoValue(card, targetText) {
    const { number, range } = getInfoControls(card);
    if (!number || !range) {
      throw new Error('找不到 Information Extracted 控件');
    }

    number.focus();
    number.select?.();
    setNativeValue(number, targetText);
    number.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    number.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true }));
    await sleep(80);
    number.blur();

    range.focus?.();
    setNativeValue(range, targetText);
    await sleep(80);
    range.blur?.();

    number.focus();
    setNativeValue(number, targetText);
    await sleep(80);
    number.blur();
  }

  async function waitValueStable(id, target, list) {
    await waitForCardState(
      id,
      list,
      (card) => {
        const { number, range } = getInfoControls(card);
        if (!number || !range) return false;

        const n1 = Number(number.value);
        const n2 = Number(range.value);

        return approxEqual(n1, target) && approxEqual(n2, target);
      },
      7000,
      `卡片 ${id} 的 Information Extracted 没有稳定更新`
    );

    await sleep(CONFIG.afterSetMs);
  }

  async function forceCommitTarget(id, target, oldValue, list) {
    const targetText = formatValue(target);

    for (let attempt = 1; attempt <= CONFIG.setCommitRetries; attempt++) {
      let card = await ensureCardVisible(id, list);
      await focusCard(card, list);

      // 抖一下再回目标，逼 React 内部状态刷新
      const nudge = clamp01(target <= 0.98 ? target + 0.01 : target - 0.01);
      const nudgeText = formatValue(nudge);

      await applyInfoValue(card, nudgeText);
      await waitValueStable(id, nudge, list);

      card = await ensureCardVisible(id, list);
      await applyInfoValue(card, targetText);
      await waitValueStable(id, target, list);

      card = await ensureCardVisible(id, list);
      await focusCard(card, list);
      await sleep(300);

      const currentValue = getCurrentInfo(card);
      const changed = !approxEqual(oldValue, target);

      const committed =
        approxEqual(currentValue, target) &&
        (
          !changed ||      // 原来就等于目标值
          isPending(card) // 改值后必须进入待提取，避免误用旧缓存结果
        );

      if (committed) {
        return card;
      }

      await jiggleCard(card, list);
    }

    throw new Error(`卡片 ${id} 的 ${targetText} 未能被网页真正接受`);
  }

  // -----------------------------
  // extract / download
  // -----------------------------
  async function clickCardAction(id, list) {
    const card = await ensureCardVisible(id, list);
    const btn = getActionButton(card);
    if (!btn) throw new Error(`卡片 ${id} 找不到主动作按钮`);

    await focusCard(card, list);
    fireRealClick(btn);
  }

  async function extractIfNeeded(id, target, oldValue, list) {
    const targetText = formatValue(target);
    let card = await ensureCardVisible(id, list);
    const changed = !approxEqual(oldValue, target);

    if (!changed && isDownloadReady(card)) {
      log(`卡片 ${id} 的 ${targetText} 已可下载`);
      return;
    }

    // 改值后必须是 pending；如果直接是 download，通常仍是旧缓存结果
    if (changed && !isPending(card)) {
      throw new Error(`卡片 ${id} 改成 ${targetText} 后未进入待提取状态`);
    }

    for (let attempt = 1; attempt <= CONFIG.extractRetries; attempt++) {
      log(`卡片 ${id} 开始提取 ${targetText}（第 ${attempt} 次）`);
      await clickCardAction(id, list);

      try {
        await waitForCardState(
          id,
          list,
          (c) => isDownloadReady(c),
          CONFIG.extractTimeoutMs,
          `卡片 ${id} 等待提取完成超时`
        );

        await sleep(CONFIG.afterActionMs);
        card = await ensureCardVisible(id, list);

        if (isDownloadReady(card)) {
          return;
        }
      } catch (err) {
        if (attempt >= CONFIG.extractRetries) throw err;
      }
    }

    throw new Error(`卡片 ${id} 提取后仍未进入可下载状态`);
  }

  async function downloadCard(id, value, list) {
    const prefix = formatValue(value);
    const card = await ensureCardVisible(id, list);

    if (!isDownloadReady(card)) {
      throw new Error(`卡片 ${id} 当前不是可下载状态`);
    }

    pendingDownloadPrefix = prefix;
    try {
      await clickCardAction(id, list);
      await sleep(CONFIG.afterActionMs);
    } finally {
      pendingDownloadPrefix = '';
    }
  }

  async function processOne(id, value, list) {
    const target = Number(value);
    const targetText = formatValue(target);

    log(`开始处理 ${id} -> ${targetText}`);

    let card = await ensureCardVisible(id, list);
    const oldValue = getCurrentInfo(card);

    card = await forceCommitTarget(id, target, oldValue, list);

    const committedValue = getCurrentInfo(card);
    if (!approxEqual(committedValue, target)) {
      throw new Error(`卡片 ${id} 提交后显示值不是 ${targetText}`);
    }

    await extractIfNeeded(id, target, oldValue, list);

    card = await ensureCardVisible(id, list);
    const beforeDownloadValue = getCurrentInfo(card);

    if (!approxEqual(beforeDownloadValue, target)) {
      throw new Error(`卡片 ${id} 下载前值异常：当前 ${beforeDownloadValue}，目标 ${targetText}`);
    }

    if (!isDownloadReady(card)) {
      throw new Error(`卡片 ${id} 下载前仍不是可下载状态`);
    }

    await downloadCard(id, target, list);
    log(`已下载 ${id} -> ${targetText}`);
  }

  // -----------------------------
  // main
  // -----------------------------
  async function runBatch() {
    if (running) return;

    const values = parseValues(panel.values.value);
    if (!values.length) {
      alert('请输入至少一个 0.01 到 1.00 之间的值，例如：0.01');
      return;
    }

    running = true;
    stopRequested = false;
    syncPanelState();

    const failures = [];
    const successes = [];

    try {
      const { ids, list } = await harvestAllCardIds();

      if (!ids.length) {
        throw new Error('没有扫描到任何 Vibe 卡片');
      }

      log(`扫描到 ${ids.length} 张卡片：${ids.join(', ')}`);

      for (const value of values) {
        const targetText = formatValue(value);
        log(`开始处理这一轮值：${targetText}`);

        for (const id of ids) {
          if (stopRequested) throw new Error('已手动停止');

          try {
            await processOne(id, value, list);
            successes.push(`${id}@${targetText}`);
          } catch (err) {
            console.error(err);
            failures.push(`${id}@${targetText}: ${err.message}`);
            log(`失败但继续：${id} -> ${targetText}`);
            if (!CONFIG.continueOnError) throw err;
          }

          await sleep(500);
        }
      }

      if (failures.length) {
        log(`完成：成功 ${successes.length} 项，失败 ${failures.length} 项。失败明细请看控制台。`);
        console.warn('[NAI CommitStrict] failures:', failures);
      } else {
        log(`全部完成：共成功 ${successes.length} 项`);
      }
    } catch (err) {
      console.error(err);
      log(`任务结束：${err.message}`);
    } finally {
      running = false;
      syncPanelState();
    }
  }

  function stopBatch() {
    stopRequested = true;
    log('已请求停止');
  }

  async function scanCards() {
    try {
      const { ids } = await harvestAllCardIds();
      log(`扫描到 ${ids.length} 张卡片：${ids.join(', ')}`);
    } catch (err) {
      console.error(err);
      log(`扫描失败：${err.message}`);
    }
  }

  // -----------------------------
  // UI
  // -----------------------------
  function mountPanel() {
    if (document.getElementById('nai-vibe-batch-commitstrict-panel')) return;

    const root = makeEl('div', { id: 'nai-vibe-batch-commitstrict-panel' }, {
      position: 'fixed',
      right: '16px',
      bottom: '16px',
      zIndex: '2147483647',
      width: '440px',
      background: 'rgba(18,18,22,0.96)',
      color: '#fff',
      border: '1px solid rgba(255,255,255,0.14)',
      borderRadius: '12px',
      padding: '14px',
      boxShadow: '0 10px 30px rgba(0,0,0,0.35)',
      fontSize: '13px',
      fontFamily: 'ui-sans-serif, system-ui, sans-serif',
      lineHeight: '1.45'
    });

    const title = makeEl('div', { textContent: 'NAI Vibe Batch Commit-Strict' }, {
      fontSize: '15px',
      fontWeight: '700',
      marginBottom: '8px'
    });

    const desc = makeEl('div', {
      textContent: '逐卡片滚动、逐卡片强制提交新值。只有确认网页真的接受了新值，才会继续提取和下载。'
    }, {
      marginBottom: '10px',
      opacity: '0.9'
    });

    const label = makeEl('div', {
      textContent: 'Values (0.01 - 1.00, comma separated)'
    }, {
      marginBottom: '6px'
    });

    const values = makeEl('input', {
      type: 'text',
      value: CONFIG.defaultValues
    }, {
      width: '100%',
      boxSizing: 'border-box',
      padding: '8px 10px',
      borderRadius: '8px',
      border: '1px solid rgba(255,255,255,0.16)',
      background: 'rgba(255,255,255,0.08)',
      color: '#fff',
      marginBottom: '10px'
    });

    const row = makeEl('div', {}, {
      display: 'flex',
      gap: '8px',
      marginBottom: '10px'
    });

    const start = makeEl('button', { textContent: 'Start' }, {
      flex: '1',
      padding: '8px 10px',
      borderRadius: '8px',
      border: 'none',
      cursor: 'pointer'
    });

    const stop = makeEl('button', { textContent: 'Stop', disabled: true }, {
      flex: '1',
      padding: '8px 10px',
      borderRadius: '8px',
      border: 'none',
      cursor: 'pointer'
    });

    const scan = makeEl('button', { textContent: 'Scan Cards' }, {
      flex: '1',
      padding: '8px 10px',
      borderRadius: '8px',
      border: 'none',
      cursor: 'pointer'
    });

    const status = makeEl('div', { textContent: '脚本已加载' }, {
      whiteSpace: 'pre-wrap',
      minHeight: '72px',
      background: 'rgba(255,255,255,0.06)',
      borderRadius: '8px',
      padding: '8px'
    });

    start.addEventListener('click', runBatch);
    stop.addEventListener('click', stopBatch);
    scan.addEventListener('click', scanCards);

    row.append(start, stop, scan);
    root.append(title, desc, label, values, row, status);
    document.body.appendChild(root);

    panel = { root, values, start, stop, scan, status };
    syncPanelState();
  }

  function init() {
    if (!document.body) return;
    mountPanel();
  }

  new MutationObserver(() => {
    if (!document.getElementById('nai-vibe-batch-commitstrict-panel')) {
      init();
    }
  }).observe(document.documentElement, {
    childList: true,
    subtree: true
  });

  init();

  window.__naiVibeBatchCommitStrict = {
    scanCards,
    runBatch,
    stopBatch,
    collectVisibleCards,
    harvestAllCardIds,
  };

  log('脚本已加载');
})();
