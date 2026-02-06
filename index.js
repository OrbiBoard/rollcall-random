const path = require('path');
const url = require('url');
const { shell } = require('electron');

let pluginApi = null;
const state = {
  eventChannel: 'rollcall-random',
  students: [],
  picked: new Set(),
  currentName: '',
  pickCount: 1,
  noRepeat: true,
  recent: [],
  recentLimit: 20,
  backgroundBase: '',
  floatSettingsBase: '',
  history: [] // { name: string, timestamp: number }
};

function loadHistory() {
  if (!pluginApi) return;
  try {
    const data = pluginApi.store.get('history');
    if (Array.isArray(data)) {
      state.history = data;
    }
  } catch (e) {
    state.history = [];
  }
}

function saveHistory(name) {
  if (!pluginApi) return;
  try {
    const now = Date.now();
    const fiveDaysAgo = now - 5 * 24 * 60 * 60 * 1000;

    // Add new record
    state.history.push({ name, timestamp: now });

    // Prune records older than 5 days
    state.history = state.history.filter(h => h.timestamp >= fiveDaysAgo);

    pluginApi.store.set('history', state.history);
  } catch (e) {
    console.error('Failed to save history', e);
  }
}

function getStats(name) {
  const now = Date.now();
  const fiveDaysAgo = now - 5 * 24 * 60 * 60 * 1000;

  // 1. Last 3 picks
  const myPicks = state.history
    .filter(h => h.name === name)
    .sort((a, b) => b.timestamp - a.timestamp);
  const last3 = myPicks.slice(0, 3).map(h => h.timestamp);

  // 2. Last 5 days stats
  const recentPicks = state.history.filter(h => h.timestamp >= fiveDaysAgo);
  const myRecentCount = recentPicks.filter(h => h.name === name).length;
  const totalRecentCount = recentPicks.length;
  const probability = totalRecentCount > 0 ? (myRecentCount / totalRecentCount) : 0;

  return {
    last3,
    recentCount: myRecentCount,
    recentTotal: totalRecentCount,
    probability
  };
}

function emitUpdate(target, value) {
  try { pluginApi.emit(state.eventChannel, { type: 'update', target, value }); } catch (e) { }
}

async function ensureStudents() {
  try {
    let res = await pluginApi.call('profiles-students', 'getStudents');
    res = res?.result || res;
    const list = Array.isArray(res?.students) ? res.students : [];
    state.students = list.filter((s) => String((s && s.name) || '').trim() !== '');
  } catch (e) { state.students = []; }
}

function pickOne(batchExcluded) {
  const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
  const unique = Array.from(new Set(names));

  // 基础池：先排除本轮已选（确保单次批量抽选中不重复）
  let basePool = unique.filter((n) => !batchExcluded || !batchExcluded.has(n));
  if (basePool.length === 0) return ''; // 没得选了

  let pool = [];
  if (state.noRepeat) {
    // 进一步排除全局已选
    pool = basePool.filter((n) => !state.picked.has(n));
    if (pool.length === 0) {
      // 全部抽完，重置全局记录
      state.picked.clear();
      // 重置后，依然基于基础池（排除本轮已选）
      pool = basePool;
    }
  } else {
    pool = basePool;
  }

  const idx = Math.floor(Math.random() * pool.length);
  const name = pool[idx] || '';

  if (name) {
    state.currentName = name;
    if (state.noRepeat) {
      state.picked.add(name);
      // 保持 recent 兼容性（虽然逻辑主要依赖 picked）
      state.recent.push(name);
      if (state.recent.length > state.recentLimit) state.recent.shift();
    }
  }
  return name;
}


const functions = {
  openRollcallTemplate: async () => {
    try {
      await ensureStudents();
      const bgFile = path.join(__dirname, 'background', 'rollcall.html');
      const floatFile = path.join(__dirname, 'float', 'settings.html');
      state.backgroundBase = url.pathToFileURL(bgFile).href;
      state.floatSettingsBase = url.pathToFileURL(floatFile).href;
      const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
      const uniqueCount = new Set(names).size;
      const initBg = state.backgroundBase + '?channel=' + encodeURIComponent(state.eventChannel) + '&caller=rollcall-random&max=' + uniqueCount + '&name=';
      const params = {
        title: '随机点名',
        icon: 'ri-shuffle-line',
        eventChannel: state.eventChannel,
        subscribeTopics: [state.eventChannel],
        callerPluginId: 'rollcall-random',
        floatingSizePercent: 48,
        floatingWidth: 520,
        floatingHeight: 360,
        centerItems: [{ id: 'start-roll', text: '开始抽选', icon: 'ri-shuffle-line' }],
        leftItems: [
          { id: 'openSettings', text: '抽选设置', icon: 'ri-settings-3-line' },
          { id: 'openExternal', text: '外部抽选', icon: 'ri-external-link-line' }
        ],
        backgroundUrl: initBg,
        floatingUrl: null,
        backgroundTargets: { rollcall: state.backgroundBase },
        floatingBounds: 'left'
      };
      await pluginApi.call('ui-lowbar', 'openTemplate', [params]);
      return true;
    } catch (e) { return { ok: false, error: e?.message || String(e) }; }
  },
  onLowbarEvent: async (payload = {}) => {
    try {
      if (!payload || typeof payload !== 'object') return true;
      if (payload.type === 'click') {
        if (payload.id === 'start-roll') {
          await ensureStudents();
          const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
          const unique = Array.from(new Set(names));

          const count = Math.max(1, state.pickCount || 1);
          const actualCount = Math.min(count, unique.length); // 限制不超过总人数
          const finalNames = [];
          const batchSet = new Set();
          
          for (let k = 0; k < actualCount; k++) {
            const name = pickOne(batchSet);
            if (name) {
              finalNames.push(name);
              batchSet.add(name);
            }
          }

          const seq = [];
          const exclude = state.noRepeat ? new Set(state.recent) : new Set();
          const pool = state.noRepeat ? unique.filter((n) => !exclude.has(n)) : unique;
          const basePool = pool.length ? pool : unique;
          const steps = basePool.length ? Math.min(5, basePool.length) : 0;
          for (let i = 0; i < steps; i++) { const j = Math.floor(Math.random() * basePool.length); seq.push(basePool[j] || ''); }

          let seat = null;
          if (finalNames.length === 1) {
            try { seat = await functions._getSeatingContext(finalNames[0]); } catch (e) { seat = null; }
          }

          try { pluginApi.emit(state.eventChannel, { type: 'animate.pick', names: seq, final: finalNames, stepMs: 40, seat }); } catch (e) { }
        }
      } else if (payload.type === 'config.count') {
        const c = parseInt(payload.count, 10);
        if (!isNaN(c) && c >= 1) state.pickCount = c;
      } else if (payload.type === 'left.click') {
        if (payload.id === 'openSettings') {
          emitUpdate('floatingBounds', 'left');
          emitUpdate('floatingBounds', { width: 520, height: 360 });
          const u = new URL(state.floatSettingsBase);
          u.searchParams.set('channel', state.eventChannel);
          u.searchParams.set('caller', 'rollcall-random');
          u.searchParams.set('noRepeat', state.noRepeat ? '1' : '0');
          u.searchParams.set('recentLimit', String(state.recentLimit || 20));
          emitUpdate('floatingUrl', u.href);
        } else if (payload.id === 'openExternal') {
          emitUpdate('floatingBounds', 'left');
          emitUpdate('floatingBounds', { width: 360, height: 480 });
          const extFile = path.join(__dirname, 'float', 'external.html');
          const u = new URL(url.pathToFileURL(extFile).href);
          u.searchParams.set('channel', state.eventChannel);
          u.searchParams.set('caller', 'rollcall-random');
          emitUpdate('floatingUrl', u.href);
        }
      } else if (payload.type === 'float.settings') {
        const v = String(payload.noRepeat || '').trim();
        if (v === '1' || v === '0') state.noRepeat = (v === '1');
        const rl = Number(payload.recentLimit);
        if (Number.isFinite(rl)) {
          const k = Math.max(1, Math.min(100, Math.floor(rl)));
          state.recentLimit = k;
          if (state.recent.length > state.recentLimit) state.recent = state.recent.slice(-state.recentLimit);
        }
        if (payload.resetPicked) { state.recent = []; state.picked.clear(); state.currentName = ''; }
      }
      return true;
    } catch (e) { return { ok: false, error: e?.message || String(e) }; }
  },
  _getSeatingContext: async (finalName) => {
    try {
      let res = await pluginApi.call('profiles-seating', 'getConfig');
      res = res?.result || res;
      const cfg = res?.config || {};
      const rows = Array.isArray(cfg.rows) ? cfg.rows : [];
      const cols = Array.isArray(cfg.cols) ? cfg.cols : [];
      const seats = (cfg && typeof cfg.seats === 'object') ? cfg.seats : {};
      const name = String(finalName || '').trim();
      if (!name) return { found: false };
      let foundKey = '';
      for (const k of Object.keys(seats || {})) { const v = seats[k]; if (v && String(v.name || '').trim() === name) { foundKey = k; break; } }
      if (!foundKey) return { found: false };
      const parts = foundKey.split('-');
      if (parts.length !== 2) return { found: false };
      const rowId = parts[0]; const colId = parts[1];
      const ri = rows.findIndex(r => String(r?.id || '') === rowId);
      const ci = cols.findIndex(c => String(c?.id || '') === colId);
      if (ri < 0 || ci < 0) return { found: false };
      const isAisleRow = (i) => (rows[i]?.type || 'row') === 'aisle';
      const isAisleCol = (i) => (cols[i]?.type || 'col') === 'aisle';
      const seatKey = (rIdx, cIdx) => { const r = rows[rIdx]?.id; const c = cols[cIdx]?.id; return (r && c) ? `${r}-${c}` : ''; };
      const occupantAt = (rIdx, cIdx) => { const key = seatKey(rIdx, cIdx); const o = key ? seats[key] : null; return o && typeof o === 'object' ? String(o.name || '') : ''; };
      let rowNumber = 0; for (let i = 0; i <= ri; i++) { if (!isAisleRow(i)) rowNumber++; }
      let colNumber = 0; for (let j = 0; j <= ci; j++) { if (!isAisleCol(j)) colNumber++; }
      const collectLeft = () => { const arr = []; let c = ci - 1; while (c >= 0 && arr.length < 2) { if (!isAisleCol(c)) { const nm = occupantAt(ri, c); if (nm) arr.push(nm); } c--; } return arr; };
      const collectRight = () => { const arr = []; let c = ci + 1; while (c < cols.length && arr.length < 2) { if (!isAisleCol(c)) { const nm = occupantAt(ri, c); if (nm) arr.push(nm); } c++; } return arr; };
      const collectFront = () => { const arr = []; let r = ri - 1; while (r >= 0 && arr.length < 2) { if (!isAisleRow(r)) { const nm = occupantAt(r, ci); if (nm) arr.push(nm); } r--; } return arr; };
      const collectBack = () => { const arr = []; let r = ri + 1; while (r < rows.length && arr.length < 2) { if (!isAisleRow(r)) { const nm = occupantAt(r, ci); if (nm) arr.push(nm); } r++; } return arr; };
      return { found: true, pos: { row: rowNumber, col: colNumber }, neighbors: { left: collectLeft(), right: collectRight(), front: collectFront(), back: collectBack() } };
    } catch (e) { return { found: false }; }
  },
  openUrl: async (targetUrl) => {
    try {
      if (!targetUrl || typeof targetUrl !== 'string') return { ok: false };
      await shell.openExternal(targetUrl);
      return { ok: true };
    } catch (e) { return { ok: false, error: e?.message }; }
  }

};

// 供预加载 quickAPI 调用的窗口控制函数

const init = async (api) => {
  pluginApi = api;
  loadHistory();
};

module.exports = { name: '随机点名', version: '0.1.0', init, functions: { ...functions, getVariable: async (name) => { const k = String(name || ''); if (k === 'currentName') return String(state.currentName || ''); return ''; }, listVariables: () => ['currentName'] } };
