// ==UserScript==
// @name         NovelAI Vibe Batch Commit-Strict
// @namespace    local.nai.vibe.batch.commitstrict
// @version      1.0.10
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
    modeSettleTimeoutMs: 1800,
    // 若改值后始终直接可下载，说明前端可能仍复用旧缓存；开启后会阻止该次下载，避免“文件名是目标值、内容却是旧值”。
    strictCacheGuard: false,
    continueOnError: true,
    // 直接设值模式：不做探针抖值/绕过值，只提交目标值。
    directSetOnly: true,
    // 提取按钮默认单击，避免双触发导致重复提取。
    aggressiveExtractClick: false,
    // 用户要求：不执行滚动扫描，只处理当前已挂载(可见)卡片。
    noScrollTraversal: true,
  };

  let running = false;
  let stopRequested = false;
  let panel = null;
  let pendingDownloadPrefix = '';

  // -----------------------------
  // download filename prefix
  // -----------------------------
  function renameDownloadAnchor(anchor) {
    if (!pendingDownloadPrefix || !anchor || !anchor.hasAttribute('download')) return;
    try {
      const oldName = anchor.getAttribute('download') || '';
      const prefix = String(pendingDownloadPrefix).replace(/[^a-zA-Z0-9._-]/g, '');
      if (oldName && prefix && !oldName.startsWith(prefix + '_')) {
        anchor.setAttribute('download', `${prefix}_${oldName}`);
      }
    } catch (err) {
      console.warn('[NAI CommitStrict] rename hook failed', err);
    }
  }

  (function patchAnchorDownloadName() {
    try {
      // 拦截 anchor.click() — 最常见的编程式下载触发方式。
      const originalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function (...args) {
        renameDownloadAnchor(this);
        return originalClick.apply(this, args);
      };

      // 兜底：捕获阶段监听 click 事件，拦截通过 dispatchEvent 触发的下载。
      document.addEventListener('click', (e) => {
        const target = e.target?.closest?.('a[download]');
        if (target) renameDownloadAnchor(target);
      }, true);
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

  function pickProbeValue(target, delta = 0.02) {
    const up = clamp01(target + delta);
    if (!isSameCommittedValue(up, target)) return up;
    return clamp01(target - delta);
  }

  function pickBypassValue(target) {
    // 对 0.01 这类极小值，直接切到中间值更容易打破“直接可下载旧缓存”状态。
    if (target <= 0.05) return 0.35;
    if (target >= 0.95) return 0.65;
    return target < 0.5 ? 0.75 : 0.25;
  }

  function approxEqual(a, b, eps = 0.011) {
    return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < eps;
  }

  function isSameCommittedValue(a, b, eps = 0.0005) {
    return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) <= eps;
  }

  function hasMeaningfulChange(oldValue, target) {
    // 旧值不可读时按“已变化”处理，避免在瞬时读值失败时错误复用旧下载结果。
    if (!Number.isFinite(oldValue)) return true;
    return !isSameCommittedValue(oldValue, target);
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
    el.focus?.();
    const Ptr = typeof PointerEvent !== 'undefined' ? PointerEvent : MouseEvent;
    el.dispatchEvent(new Ptr('pointerdown', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    // 单次 click 即可触发 React 事件委托（React 在 root 监听原生 click）。
    // 注意：不再追加 el.click()，因为会导致 React handler 双重触发，
    // 从而引起提取按钮连按两次、下载双重触发等严重问题。
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  }

  function clickElementAtCenter(el) {
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const topEl = document.elementFromPoint(x, y);
    if (!topEl) return;
    fireRealClick(topEl);
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

  function getCardId(card) {
    const idInput = card?.querySelector('input[type="text"]');
    return (idInput?.value || '').trim();
  }

  async function waitCardBindingStable(id, list, initialCard, timeoutMs = 1600) {
    const start = Date.now();
    let card = initialCard;
    let stableHits = 0;

    while (Date.now() - start < timeoutMs) {
      const mapped = getVisibleCardMap().get(id) || null;
      if (mapped && mapped !== card) {
        card = mapped;
        stableHits = 0;
      }

      const { number, range } = getInfoControls(card || null);
      const idOk = card && isVisible(card) && getCardId(card) === id;
      const controlsOk = Boolean(number && range);

      if (idOk && controlsOk) {
        stableHits += 1;
        if (stableHits >= 2) return card;
      } else {
        stableHits = 0;
      }

      await sleep(CONFIG.pollMs);
    }

    throw new Error(`卡片 ${id} 在滚动后未稳定挂载`);
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
    if (CONFIG.noScrollTraversal) {
      const ids = Array.from(getVisibleCardMap().keys());
      return { ids, list: null };
    }

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
    if (card) {
      if (CONFIG.noScrollTraversal || !list) return card;
      return waitCardBindingStable(id, list, card);
    }

    if (CONFIG.noScrollTraversal || !list) {
      throw new Error(`当前视图找不到卡片 ${id}（noScrollTraversal=true）`);
    }

    const originalTop = list.scrollTop;
    list.scrollTop = 0;
    await sleep(CONFIG.afterScrollMs);

    for (let i = 0; i < CONFIG.maxScanLoops; i++) {
      card = tryVisible();
      if (card) return waitCardBindingStable(id, list, card);

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
      let card;
      try {
        card = await ensureCardVisible(id, list);
      } catch (_err) {
        // Card may briefly disappear during React remount; keep polling
        // instead of propagating the error immediately.
        await sleep(CONFIG.pollMs);
        continue;
      }
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
    const raw = number?.value ?? range?.value ?? number?.getAttribute('value') ?? range?.getAttribute('value');
    const n = Number(raw);
    return Number.isFinite(n) ? n : NaN;
  }

  function hasVisibleText(root, needle) {
    if (!root || !needle) return false;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode;
      const txt = (node.nodeValue || '').trim();
      if (!txt || !txt.includes(needle)) continue;
      const parent = node.parentElement;
      if (parent && isVisible(parent)) return true;
    }
    return false;
  }

  function isPending(card) {
    return hasVisibleText(card, 'Encoding required');
  }

  function requiresExtraction(card) {
    // 有些界面文案或语言环境下不会出现 "Encoding required"，
    // 但主按钮会切回提取模式（显示 Anlas 价格）。
    return isPending(card) || getActionMode(card) === 'extract';
  }

  function getHeaderButtons(card) {
    if (!card) return [];

    // 策略 1：传统位置——firstElementChild 通常是 header/toolbar。
    const first = card.firstElementChild;
    if (first) {
      const btns = Array.from(first.querySelectorAll('button')).filter(isVisible);
      if (btns.length >= 2) return btns;
    }

    // 策略 2：遍历 card 所有直接子元素，找含有最多可见 button 的那个区域。
    // 典型 header/toolbar 比其他区域（滑条、缩略图）拥有更多按钮。
    let best = [];
    for (const child of card.children) {
      if (child === first) continue; // 已检查
      const btns = Array.from(child.querySelectorAll('button')).filter(isVisible);
      if (btns.length > best.length) best = btns;
    }
    if (best.length >= 2) return best;

    // 策略 3：整张卡片范围搜索（最后手段）。
    return Array.from(card.querySelectorAll('button')).filter(isVisible);
  }

  // header buttons usually [trash, action, check].
  // 但按钮数量/顺序在提取进行中可能变化，因此加内容兜底。
  function getActionButton(card) {
    const buttons = getHeaderButtons(card);
    if (buttons.length === 0) return null;

    // --- 内容优先匹配（最可靠）---

    // 1. 含 "anlas" 文字 → 一定是提取/动作按钮。
    for (const btn of buttons) {
      if (/anlas/i.test(textOf(btn))) return btn;
    }

    // 2. 纯数字文字（Anlas 费用）。
    for (const btn of buttons) {
      if (/^\d+$/.test(textOf(btn).trim())) return btn;
    }

    // 3. 含 SVG 的按钮中，跳过第一个（通常是 trash/删除），取第二个（动作按钮）。
    const svgButtons = buttons.filter((b) => b.querySelector('svg'));
    if (svgButtons.length >= 2) return svgButtons[1];
    // 如果只有一个含 SVG 的按钮，它可能就是动作按钮（如删除按钮不含 SVG）。
    if (svgButtons.length === 1) return svgButtons[0];

    // --- 位置兜底 ---
    if (buttons[1]) return buttons[1];

    return buttons[0] || null;
  }

  function getActionMode(card) {
    const btn = getActionButton(card);
    if (!btn) return 'unknown';

    const txt = textOf(btn);
    const hasSvg = btn.querySelector('svg') !== null;

    // React remount 时文本可能为空。
    if (!txt) {
      return hasSvg ? 'download' : 'unknown';
    }

    // 明确含 "anlas" → 提取按钮。
    if (/anlas/i.test(txt)) return 'extract';

    // 按钮含 SVG 图标 → 优先判定为下载（提取按钮通常只有文字/数字，无图标）。
    if (hasSvg) return 'download';

    // 纯数字文本（如 "5"、"10"、"0"）→ Anlas 费用，视为提取。
    // 注意：只匹配纯数字，不匹配含字母/混合文本（如 "Download"、"2.5MB"）。
    if (/^\d+$/.test(txt.trim())) return 'extract';

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

    if (CONFIG.directSetOnly) {
      const card = await ensureCardVisible(id, list);
      await focusCard(card, list);
      await applyInfoValue(card, targetText);
      await waitValueStable(id, target, list);
      const committedCard = await ensureCardVisible(id, list);
      const currentValue = getCurrentInfo(committedCard);
      if (!approxEqual(currentValue, target)) {
        throw new Error(`卡片 ${id} 直接设值后未稳定到 ${targetText}`);
      }
      return committedCard;
    }

    for (let attempt = 1; attempt <= CONFIG.setCommitRetries; attempt++) {
      let card = await ensureCardVisible(id, list);
      await focusCard(card, list);

      // 抖一下再回目标，逼 React 内部状态刷新
      const nudge = pickProbeValue(target, 0.05);
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
      const changed = hasMeaningfulChange(oldValue, target);

      const committed =
        approxEqual(currentValue, target) &&
        (
          !changed ||
          requiresExtraction(card) ||
          isDownloadReady(card)
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

  async function waitForModeSettleAfterCommit(id, target, list, timeoutMs = CONFIG.modeSettleTimeoutMs) {
    const start = Date.now();
    let lastMode = null;
    let stableCount = 0;

    while (Date.now() - start < timeoutMs) {
      let card;
      try {
        card = await ensureCardVisible(id, list);
      } catch (_err) {
        // Card may briefly disappear during React remount; keep polling.
        await sleep(CONFIG.pollMs);
        continue;
      }
      const valueNow = getCurrentInfo(card);
      const mode = getActionMode(card);

      // 只在目标值显示稳定时判断模式，避免读到虚拟列表切换的瞬时旧节点。
      if (approxEqual(valueNow, target)) {
        if (mode === 'extract') return card;
        if (mode === lastMode) {
          stableCount += 1;
        } else {
          lastMode = mode;
          stableCount = 1;
        }

        // 连续多次稳定为 download，视为前端已结算到该状态。
        if (mode === 'download' && stableCount >= 3) {
          return card;
        }
      }

      await sleep(CONFIG.pollMs);
    }

    return ensureCardVisible(id, list);
  }

  async function clickCardAction(id, list, options = {}) {
    const { aggressive = false, forDownload = false } = options;
    const card = await ensureCardVisible(id, list);
    const btn = getActionButton(card);
    if (!btn) throw new Error(`卡片 ${id} 找不到主动作按钮`);

    await focusCard(card, list);

    if (aggressive) {
      // 激进模式：先尝试中心点击（绕过层叠遮挡），若 elementFromPoint
      // 未命中按钮自身或其子元素，则回退到直接 fireRealClick。
      const rect = btn.getBoundingClientRect();
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const topEl = document.elementFromPoint(cx, cy);
      if (topEl && btn.contains(topEl)) {
        fireRealClick(topEl);
      } else {
        fireRealClick(btn);
      }
      return;
    }

    if (forDownload) {
      // 下载动作只触发一次 click 事件，且使用原生 .click()，
      // 确保内部创建的 <a download> 走 HTMLAnchorElement.prototype.click 路径，
      // 从而被文件名重命名补丁拦截。
      btn.click();
      return;
    }

    // 提取动作使用完整事件分发，保证 React 合成事件能响应。
    fireRealClick(btn);
  }

  async function extractIfNeeded(id, target, oldValue, list) {
    const targetText = formatValue(target);
    let card = await ensureCardVisible(id, list);
    const changed = hasMeaningfulChange(oldValue, target);

    if (!changed && isDownloadReady(card) && !CONFIG.strictCacheGuard) {
      log(`卡片 ${id} 的 ${targetText} 已可下载`);
      return;
    }

    if (changed) {
      card = await waitForModeSettleAfterCommit(id, target, list);
    }

    // When strictCacheGuard is on, a !changed card that is already
    // download-ready must also go through the cache-bypass verification,
    // because the existing download may reflect stale content.
    if ((changed || CONFIG.strictCacheGuard) && isDownloadReady(card)) {
      log(`卡片 ${id} 的 ${targetText} 变更后直接可下载，执行二次改值校验以避免下载旧缓存`);

      const probe = pickProbeValue(target, target <= 0.05 ? 0.20 : 0.08);
      await forceCommitTarget(id, probe, target, list);
      card = await forceCommitTarget(id, target, probe, list);
      card = await waitForModeSettleAfterCommit(id, target, list);

      // 若仍直接可下载，尝试多组绕过值，直到至少出现一次“待提取”并完成提取。
      if (isDownloadReady(card)) {
        const bypassCandidates = [
          pickBypassValue(target),
          clamp01(target <= 0.50 ? 0.88 : 0.12),
          clamp01(target <= 0.50 ? 0.66 : 0.34),
        ];

        let forcedFresh = false;
        for (const bypass of bypassCandidates) {
          if (isSameCommittedValue(bypass, target)) continue;
          log(`卡片 ${id} 的 ${targetText} 仍直接可下载，执行缓存绕过：${formatValue(bypass)} -> ${targetText}`);

          card = await forceCommitTarget(id, bypass, target, list);
          card = await waitForModeSettleAfterCommit(id, bypass, list);

          if (requiresExtraction(card)) {
            log(`卡片 ${id} 在绕过值 ${formatValue(bypass)} 进入待提取，执行强制提取`);
            await clickCardAction(id, list, { aggressive: CONFIG.aggressiveExtractClick });
            await waitForCardState(
              id,
              list,
              (c) => isDownloadReady(c),
              CONFIG.extractTimeoutMs,
              `卡片 ${id} 绕过值 ${formatValue(bypass)} 提取超时`
            );
            await sleep(CONFIG.afterActionMs);
            forcedFresh = true;
          } else {
            log(`卡片 ${id} 在绕过值 ${formatValue(bypass)} 仍直接可下载，继续尝试下一组绕过值`);
          }

          card = await forceCommitTarget(id, target, bypass, list);
          card = await waitForModeSettleAfterCommit(id, target, list);

          if (requiresExtraction(card)) {
            break;
          }
        }

        if (isDownloadReady(card)) {
          if (CONFIG.strictCacheGuard) {
            const hint = forcedFresh
              ? '已在绕过值完成提取，但回到目标值后仍直接可下载'
              : '所有绕过值均未进入待提取状态';
            throw new Error(`卡片 ${id} 改成 ${targetText} 后仍直接可下载（${hint}），已阻止下载`);
          }

          log(`卡片 ${id} 的 ${targetText} 缓存绕过后仍可直接下载，按已提交值继续（strictCacheGuard=false）`);
          return;
        }
      }
    }

    // 改值后必须是 pending（!changed 只因 strictCacheGuard 到达此处时，
    // 卡片可能仍是 download-ready，此时无需提取——缓存校验已在上方完成）。
    if (!requiresExtraction(card)) {
      if (changed) {
        throw new Error(`卡片 ${id} 改成 ${targetText} 后未进入待提取状态`);
      }
      // !changed + strictCacheGuard path: card passed the cache-bypass block
      // above and is still download-ready — nothing more to extract.
      return;
    }

    // 点击提取的总次数上限（包括 stale 重试）。
    const maxClicks = CONFIG.extractRetries * 3;
    let clicks = 0;

    for (let attempt = 1; attempt <= CONFIG.extractRetries; attempt++) {
      clicks++;
      log(`卡片 ${id} 开始提取 ${targetText}（第 ${attempt} 次, click#${clicks}）`);
      await clickCardAction(id, list, { aggressive: CONFIG.aggressiveExtractClick });

      try {
        let diagTimer = 0;
        const clickTime = Date.now();
        const STALE_EXTRACT_MS = 15000; // 15 秒内仍是 extract → 点击未生效
        let retriedStale = false;

        await waitForCardState(
          id,
          list,
          (c) => {
            const mode = getActionMode(c);
            if (mode === 'download') return true;

            diagTimer += 1;

            // Stale guard：点击后 15 秒仍为 extract 模式 → 点击未生效，立即重新点击。
            if (
              !retriedStale &&
              mode === 'extract' &&
              Date.now() - clickTime > STALE_EXTRACT_MS &&
              clicks < maxClicks
            ) {
              retriedStale = true;
              clicks++;
              log(`[诊断] ${id} 点击后 15 秒仍为 extract 模式，重新点击（click#${clicks}）`);
              // 异步重点击——不阻塞 predicate，下一轮 poll 会看到新状态。
              clickCardAction(id, list, { aggressive: CONFIG.aggressiveExtractClick });
            }

            // 每 ~5 秒输出一次诊断日志。
            if (diagTimer % 20 === 0) {
              const allBtns = getHeaderButtons(c);
              const btn = getActionButton(c);
              const btnTxt = textOf(btn);
              const hasSvg = btn?.querySelector('svg') !== null;
              const pending = isPending(c);
              const btnSummary = allBtns.map((b, i) => `[${i}]"${textOf(b)}"`).join(', ');
              log(`[诊断] ${id} 等待中：mode=${mode}, btnText="${btnTxt}", hasSvg=${hasSvg}, pending=${pending}, buttons(${allBtns.length}): ${btnSummary || '(无)'}`);
            }
            return false;
          },
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

    pendingDownloadPrefix = `${prefix}_${id}`;
    try {
      await clickCardAction(id, list, { forDownload: true });
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

      const scope = CONFIG.noScrollTraversal ? '（仅当前可见）' : '';
      log(`扫描到 ${ids.length} 张卡片${scope}：${ids.join(', ')}`);

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
      const scope = CONFIG.noScrollTraversal ? '（仅当前可见）' : '';
      log(`扫描到 ${ids.length} 张卡片${scope}：${ids.join(', ')}`);
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
      textContent: CONFIG.noScrollTraversal
        ? '仅处理当前可见卡片；直接提交新值并校验，再提取与下载。'
        : '逐卡片滚动、逐卡片强制提交新值。只有确认网页真的接受了新值，才会继续提取和下载。'
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
