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
  history: [],
  fairMode: true,
  genderBalance: false,
  coldStartWeight: 1.5,
  maxGapThreshold: 3,
  minCandidatePool: 3
};

function loadHistory() {
  if (!pluginApi || !pluginApi.store) return;
  try {
    const data = pluginApi.store.get('history');
    if (Array.isArray(data)) {
      state.history = data;
    }
  } catch (e) {
    console.error('loadHistory error:', e);
    state.history = [];
  }
}

function loadFairSettings() {
  if (!pluginApi || !pluginApi.store) return;
  try {
    const settings = pluginApi.store.get('fairSettings');
    if (settings !== null && settings !== undefined && typeof settings === 'object') {
      if (typeof settings.fairMode === 'boolean') state.fairMode = settings.fairMode;
      if (typeof settings.genderBalance === 'boolean') state.genderBalance = settings.genderBalance;
      if (typeof settings.coldStartWeight === 'number' && Number.isFinite(settings.coldStartWeight)) {
        state.coldStartWeight = Math.max(1, Math.min(3, settings.coldStartWeight));
      }
      if (typeof settings.maxGapThreshold === 'number' && Number.isFinite(settings.maxGapThreshold)) {
        state.maxGapThreshold = Math.max(1, Math.min(10, settings.maxGapThreshold));
      }
      if (typeof settings.minCandidatePool === 'number' && Number.isFinite(settings.minCandidatePool)) {
        state.minCandidatePool = Math.max(1, Math.min(20, settings.minCandidatePool));
      }
    }
  } catch (e) {
    console.error('loadFairSettings error:', e);
  }
}

function saveHistory(name) {
  if (!pluginApi) return;
  try {
    const now = Date.now();
    const fiveDaysAgo = now - 5 * 24 * 60 * 60 * 1000;
    state.history.push({ name, timestamp: now });
    state.history = state.history.filter(h => h.timestamp >= fiveDaysAgo);
    pluginApi.store.set('history', state.history);
  } catch (e) {
    console.error('Failed to save history', e);
  }
}

function getPickCountByName(name) {
  return state.history.filter(h => h.name === name).length;
}

function getAllPickStats() {
  const names = state.students.map(s => String(s.name || '').trim()).filter(n => !!n);
  const unique = Array.from(new Set(names));
  const stats = {};
  unique.forEach(name => {
    stats[name] = getPickCountByName(name);
  });
  return stats;
}

function calculateAveragePickCount(stats) {
  const values = Object.values(stats);
  if (values.length === 0) return 0;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function calculateWeights(stats, options = {}) {
  const { coldStartWeight = 1.5, avgCount = 0 } = options;
  const weights = {};
  const names = Object.keys(stats);

  names.forEach(name => {
    const count = stats[name];
    let weight;

    if (count === 0) {
      weight = coldStartWeight * 10;
    } else if (count <= avgCount) {
      weight = (avgCount - count + 1) * coldStartWeight;
    } else {
      weight = 1 / (count - avgCount + 1);
    }

    weights[name] = Math.max(0.1, weight);
  });

  return weights;
}

function buildCandidatePool(stats, options = {}) {
  const {
    avgCount = 0,
    maxGapThreshold = 3,
    minCandidatePool = 3,
    batchExcluded = null
  } = options;

  let candidates = Object.keys(stats);

  if (batchExcluded && batchExcluded.size > 0) {
    candidates = candidates.filter(n => !batchExcluded.has(n));
  }

  if (candidates.length === 0) return [];

  const counts = candidates.map(n => stats[n]);
  const minCount = Math.min(...counts);
  const maxCount = Math.max(...counts);
  const gap = maxCount - minCount;

  if (gap > maxGapThreshold) {
    const threshold = maxCount - Math.floor(gap / 2);
    let filtered = candidates.filter(n => stats[n] <= threshold);

    if (filtered.length >= minCandidatePool) {
      candidates = filtered;
    }
  }

  let avgFiltered = candidates.filter(n => stats[n] <= avgCount);
  if (avgFiltered.length >= minCandidatePool) {
    candidates = avgFiltered;
  }

  if (candidates.length < minCandidatePool) {
    const sortedByCount = Object.keys(stats)
      .filter(n => !batchExcluded || !batchExcluded.has(n))
      .sort((a, b) => stats[a] - stats[b]);
    candidates = sortedByCount.slice(0, Math.max(minCandidatePool, sortedByCount.length));
  }

  return candidates;
}

function selectByWeight(candidates, weights) {
  if (candidates.length === 0) return '';
  if (candidates.length === 1) return candidates[0];

  const totalWeight = candidates.reduce((sum, name) => sum + (weights[name] || 1), 0);
  let random = Math.random() * totalWeight;

  for (const name of candidates) {
    random -= (weights[name] || 1);
    if (random <= 0) {
      return name;
    }
  }

  return candidates[candidates.length - 1];
}

function calculateProbabilities(candidates, weights) {
  const probs = {};
  if (candidates.length === 0) return probs;

  const totalWeight = candidates.reduce((sum, name) => sum + (weights[name] || 1), 0);

  candidates.forEach(name => {
    probs[name] = (weights[name] || 1) / totalWeight;
  });

  return probs;
}

function getGenderStats(name) {
  const student = state.students.find(s => String(s.name || '').trim() === name);
  return student?.gender || '未选择';
}

function applyGenderBalance(candidates, weights, stats) {
  if (!state.genderBalance || candidates.length < 2) return weights;

  const adjustedWeights = { ...weights };
  const genderCounts = { '男': 0, '女': 0, '未选择': 0 };
  const genderTotalWeight = { '男': 0, '女': 0, '未选择': 0 };

  candidates.forEach(name => {
    const gender = getGenderStats(name);
    const w = weights[name] || 1;
    if (genderCounts[gender] !== undefined) {
      genderCounts[gender]++;
      genderTotalWeight[gender] += w;
    }
  });

  if (genderCounts['男'] === 0 || genderCounts['女'] === 0) return weights;

  const maleAvgWeight = genderTotalWeight['男'] / genderCounts['男'];
  const femaleAvgWeight = genderTotalWeight['女'] / genderCounts['女'];

  if (maleAvgWeight > 0 && femaleAvgWeight > 0) {
    const ratio = maleAvgWeight / femaleAvgWeight;
    const balanceFactor = 1.2;

    if (ratio > balanceFactor) {
      candidates.forEach(name => {
        if (getGenderStats(name) === '女') {
          adjustedWeights[name] = (weights[name] || 1) * ratio;
        }
      });
    } else if (ratio < 1 / balanceFactor) {
      candidates.forEach(name => {
        if (getGenderStats(name) === '男') {
          adjustedWeights[name] = (weights[name] || 1) / ratio;
        }
      });
    }
  }

  return adjustedWeights;
}

function getStats(name) {
  const now = Date.now();
  const fiveDaysAgo = now - 5 * 24 * 60 * 60 * 1000;

  const myPicks = state.history
    .filter(h => h.name === name)
    .sort((a, b) => b.timestamp - a.timestamp);
  const last3 = myPicks.slice(0, 3).map(h => h.timestamp);

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

function pickOneFair(batchExcluded) {
  const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
  const unique = Array.from(new Set(names));

  let basePool = unique.filter((n) => !batchExcluded || !batchExcluded.has(n));
  if (basePool.length === 0) return '';

  if (!state.fairMode) {
    return pickOneSimple(batchExcluded);
  }

  const stats = {};
  basePool.forEach(name => {
    stats[name] = getPickCountByName(name);
  });

  const avgCount = calculateAveragePickCount(stats);

  let candidates = buildCandidatePool(stats, {
    avgCount,
    maxGapThreshold: state.maxGapThreshold,
    minCandidatePool: state.minCandidatePool,
    batchExcluded
  });

  if (candidates.length === 0) {
    candidates = basePool;
  }

  let weights = calculateWeights(stats, {
    coldStartWeight: state.coldStartWeight,
    avgCount
  });

  if (state.genderBalance) {
    weights = applyGenderBalance(candidates, weights, stats);
  }

  const name = selectByWeight(candidates, weights);

  if (name) {
    state.currentName = name;
    if (state.noRepeat) {
      state.picked.add(name);
      state.recent.push(name);
      if (state.recent.length > state.recentLimit) state.recent.shift();
    }
  }
  return name;
}

function pickOneSimple(batchExcluded) {
  const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
  const unique = Array.from(new Set(names));

  let basePool = unique.filter((n) => !batchExcluded || !batchExcluded.has(n));
  if (basePool.length === 0) return '';

  let pool = [];
  if (state.noRepeat) {
    pool = basePool.filter((n) => !state.picked.has(n));
    if (pool.length === 0) {
      state.picked.clear();
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
      state.recent.push(name);
      if (state.recent.length > state.recentLimit) state.recent.shift();
    }
  }
  return name;
}

function pickOne(batchExcluded) {
  if (state.fairMode) {
    return pickOneFair(batchExcluded);
  }
  return pickOneSimple(batchExcluded);
}

function getFairPickInfo(batchExcluded = null) {
  const names = state.students.map((s) => String(s.name || '').trim()).filter((n) => !!n);
  const unique = Array.from(new Set(names));

  let basePool = unique.filter((n) => !batchExcluded || !batchExcluded.has(n));
  if (basePool.length === 0) {
    return { candidates: [], weights: {}, probabilities: {}, stats: {} };
  }

  const stats = {};
  basePool.forEach(name => {
    stats[name] = getPickCountByName(name);
  });

  const avgCount = calculateAveragePickCount(stats);

  let candidates = buildCandidatePool(stats, {
    avgCount,
    maxGapThreshold: state.maxGapThreshold,
    minCandidatePool: state.minCandidatePool,
    batchExcluded
  });

  if (candidates.length === 0) {
    candidates = basePool;
  }

  let weights = calculateWeights(stats, {
    coldStartWeight: state.coldStartWeight,
    avgCount
  });

  if (state.genderBalance) {
    weights = applyGenderBalance(candidates, weights, stats);
  }

  const probabilities = calculateProbabilities(candidates, weights);

  return {
    candidates,
    weights,
    probabilities,
    stats,
    avgCount
  };
}

function getPickedPersonStats(name, probabilities, stats, candidates) {
  if (!name) return null;
  
  const pickCount = stats[name] || 0;
  
  const sortedByProb = [...candidates].sort((a, b) => (probabilities[b] || 0) - (probabilities[a] || 0));
  const rank = sortedByProb.indexOf(name) + 1;
  const totalCandidates = candidates.length;
  
  const recentLimit = 100;
  const recentPicks = state.history.slice(-recentLimit);
  const recentCount = recentPicks.filter(h => h.name === name).length;
  const totalCount = recentPicks.length + 1;
  const actualRecentCount = recentCount + 1;
  const recentProbability = actualRecentCount / totalCount;
  
  let consecutiveMisses = 0;
  if (pickCount === 0) {
    consecutiveMisses = state.history.length;
  } else {
    for (let i = state.history.length - 1; i >= 0; i--) {
      if (state.history[i].name === name) {
        break;
      }
      consecutiveMisses++;
    }
  }
  
  let luckLevel = 'normal';
  let luckText = '正常概率';
  if (recentProbability >= 0.3) {
    luckLevel = 'high';
    luckText = '高概率命中';
  } else if (recentProbability >= 0.15) {
    luckLevel = 'above-avg';
    luckText = '概率较高';
  } else if (recentProbability <= 0.03) {
    luckLevel = 'miracle';
    luckText = '奇迹发生';
  } else if (recentProbability <= 0.08) {
    luckLevel = 'lucky';
    luckText = '锦鲤附体';
  }
  
  if (consecutiveMisses >= 30) {
    luckLevel = 'long-wait';
    luckText = '守得云开';
  } else if (consecutiveMisses >= 15) {
    luckLevel = 'patience';
    luckText = '终于轮到';
  }
  
  return {
    name,
    probability: recentProbability,
    probabilityPercent: (recentProbability * 100).toFixed(1),
    pickCount,
    recentCount: actualRecentCount,
    recentLimit: totalCount,
    consecutiveMisses,
    rank,
    totalCandidates,
    luckLevel,
    luckText
  };
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
          { id: 'showProbabilities', text: '查看概率', icon: 'ri-pie-chart-line' },
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
          const actualCount = Math.min(count, unique.length);
          const finalNames = [];
          const batchSet = new Set();

          let pickedPersonStats = null;
          
          for (let k = 0; k < actualCount; k++) {
            const fairInfoBeforePick = getFairPickInfo(batchSet);
            
            const name = pickOne(batchSet);
            if (name) {
              finalNames.push(name);
              batchSet.add(name);
              
              if (k === 0 && fairInfoBeforePick && fairInfoBeforePick.candidates.length > 0) {
                pickedPersonStats = getPickedPersonStats(
                  name,
                  fairInfoBeforePick.probabilities,
                  fairInfoBeforePick.stats,
                  fairInfoBeforePick.candidates
                );
              }
              
              saveHistory(name);
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

          try { pluginApi.emit(state.eventChannel, { type: 'animate.pick', names: seq, final: finalNames, stepMs: 40, seat, pickedPersonStats }); } catch (e) { }
        }
      } else if (payload.type === 'config.count') {
        const c = parseInt(payload.count, 10);
        if (!isNaN(c) && c >= 1) state.pickCount = c;
      } else if (payload.type === 'left.click') {
        if (payload.id === 'openSettings') {
          emitUpdate('floatingBounds', 'left');
          emitUpdate('floatingBounds', { width: 520, height: 480 });
          const u = new URL(state.floatSettingsBase);
          u.searchParams.set('channel', state.eventChannel);
          u.searchParams.set('caller', 'rollcall-random');
          u.searchParams.set('noRepeat', state.noRepeat ? '1' : '0');
          u.searchParams.set('recentLimit', String(state.recentLimit || 20));
          u.searchParams.set('fairMode', state.fairMode ? '1' : '0');
          u.searchParams.set('genderBalance', state.genderBalance ? '1' : '0');
          u.searchParams.set('coldStartWeight', String(state.coldStartWeight || 1.5));
          u.searchParams.set('maxGapThreshold', String(state.maxGapThreshold || 3));
          u.searchParams.set('minCandidatePool', String(state.minCandidatePool || 3));
          emitUpdate('floatingUrl', u.href);
        } else if (payload.id === 'openExternal') {
          emitUpdate('floatingBounds', 'left');
          emitUpdate('floatingBounds', { width: 360, height: 480 });
          const extFile = path.join(__dirname, 'float', 'external.html');
          const u = new URL(url.pathToFileURL(extFile).href);
          u.searchParams.set('channel', state.eventChannel);
          u.searchParams.set('caller', 'rollcall-random');
          emitUpdate('floatingUrl', u.href);
        } else if (payload.id === 'showProbabilities') {
          await ensureStudents();
          const probFile = path.join(__dirname, 'float', 'probabilities.html');
          const u = new URL(url.pathToFileURL(probFile).href);
          u.searchParams.set('channel', state.eventChannel);
          u.searchParams.set('caller', 'rollcall-random');
          emitUpdate('floatingBounds', 'center');
          emitUpdate('floatingBounds', { width: 640, height: 520 });
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

        if (typeof payload.fairMode === 'boolean') state.fairMode = payload.fairMode;
        if (typeof payload.genderBalance === 'boolean') state.genderBalance = payload.genderBalance;
        const csw = Number(payload.coldStartWeight);
        if (Number.isFinite(csw)) state.coldStartWeight = Math.max(1, Math.min(3, csw));
        const mgt = Number(payload.maxGapThreshold);
        if (Number.isFinite(mgt)) state.maxGapThreshold = Math.max(1, Math.min(10, mgt));
        const mcp = Number(payload.minCandidatePool);
        if (Number.isFinite(mcp)) state.minCandidatePool = Math.max(1, Math.min(20, mcp));

        try {
          pluginApi.store.set('fairSettings', {
            fairMode: state.fairMode,
            genderBalance: state.genderBalance,
            coldStartWeight: state.coldStartWeight,
            maxGapThreshold: state.maxGapThreshold,
            minCandidatePool: state.minCandidatePool
          });
        } catch (e) { }
      } else if (payload.type === 'getProbabilities') {
        await ensureStudents();
        const info = getFairPickInfo();
        return { ok: true, ...info };
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
  },
  getFairPickInfo: async () => {
    await ensureStudents();
    return getFairPickInfo();
  }

};

const init = async (api) => {
  pluginApi = api;
  loadHistory();
  loadFairSettings();
};

module.exports = { name: '随机点名', version: '0.2.0', init, functions: { ...functions, getVariable: async (name) => { const k = String(name || ''); if (k === 'currentName') return String(state.currentName || ''); return ''; }, listVariables: () => ['currentName'] } };
