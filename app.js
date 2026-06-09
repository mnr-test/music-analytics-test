
// ─── Constants ───────────────────────────────────────────────
const M_RE = /^M\d+[A-Za-z]*$/i;
const COLORS = {
  commissioned: '#3D7EF5',
  library:      '#F5A623',
  licensed:     '#1DB954',
  pd:           '#A06EF5',
  free:         '#F55E3D',
  other:        '#555'
};
const TYPE_LABELS = {
  commissioned: '書き下ろし',
  library:      'ライブラリー',
  licensed:     'ライセンス・既存曲',
  pd:           'パブリックドメイン',
  free:         '著作権フリー',
  other:        'その他'
};
const LIB_COLORS = ['#F5A623','#3D7EF5','#1DB954','#E50914','#A06EF5','#F55E3D','#5EC4F5','#F5C93D'];

// ─── State ───────────────────────────────────────────────────
let projects = {};       // { id: projectObj }
let currentProjectId = null;
let currentDash = null;  // computed dash data for current project
let delivRows = [];
let newXlsxData = null;
let updateXlsxData = null;

// ─── Persistence ─────────────────────────────────────────────
function saveProjects() {
  try {
    const serializable = {};
    for (const [id, p] of Object.entries(projects)) {
      serializable[id] = { ...p, wavSet: [...(p.wavSet || [])] };
    }
    localStorage.setItem('nma_projects', JSON.stringify(serializable));
  } catch(e) { console.warn('Save failed:', e); }
}

function loadProjects() {
  try {
    const raw = localStorage.getItem('nma_projects');
    if (!raw) return;
    const parsed = JSON.parse(raw);
    for (const [id, p] of Object.entries(parsed)) {
      projects[id] = { ...p, wavSet: new Set(p.wavSet || []) };
    }
  } catch(e) { console.warn('Load failed:', e); }
}

// ─── Parse helpers ───────────────────────────────────────────
function extractM(name) {
  const base = name.replace(/\.[^.]+$/, '');
  const parts = base.split(/[_\s&]+/);
  for (const p of parts) if (M_RE.test(p.trim())) return p.trim().toUpperCase();
  const m = base.match(/\b(M\d+[A-Za-z]?)\b/i);
  return m ? m[1].toUpperCase() : null;
}

// ファイル名からキューシートタイトルとの照合キーを生成
// 例: "KQZ_LiberationA_Final.wav" → "KQZ_LIBERATIONA"
function extractFileKey(filename) {
  // 拡張子・_Final・_ReFinal・(Org)・(Vo)などのサフィックスを除去
  let base = filename.replace(/\.[^.]+$/, '');
  base = base.replace(/[\s_]*(ReFinal|Final|refinal|final)$/i, '');
  base = base.replace(/\s*\((Org|Vo|org|vo)\)/gi, '');
  return base.trim().toUpperCase();
}

// 表記ゆれを正規化（ライブラリ名やスペルミス吸収）
function normalizeTitle(t) {
  return t.toUpperCase()
    .replace(/\s+/g, '_')
    .replace(/[）\)]/g, '')
    .replace(/[（\(]/g, '')
    // 既知の表記ゆれパターン
    .replace('LIBRATION', 'LIBERATION')
    .replace('UNRESET', 'UNREST')
    .replace('WARMS', 'WARMTH')
    .replace('RHYTYM', 'RHYTHM')
    .replace('MARIDOME', 'MARIDOME'); // 統一
}

// キューシートタイトルの正規化キー
function titleKey(t) {
  return normalizeTitle(t.trim());
}

function classifyOrigin(o) {
  if (!o) return 'other';
  const s = String(o).toLowerCase();
  if (s.includes('commission') || s.includes('書き下ろし')) return 'commissioned';
  if (s.includes('library') || s.includes('ライブラリ')) return 'library';
  if (s.includes('licensed') || s.includes('ライセンス') || s.includes('pre-existing') || s.includes('既存曲')) return 'licensed';
  if (s.includes('public domain') || s.includes('パブリックドメイン')) return 'pd';
  if (s.includes('著作権フリー') || s.includes('copyright free') || s.includes('royalty free')) return 'free';
  return 'other';
}

function normPub(p) {
  const l = p.toLowerCase();
  if (l.includes('west one')) return 'West One Music';
  if (l.includes('extreme')) return 'Extreme Music';
  if (l.includes('bmg')) return 'BMG Production Music';
  if (l.includes('scoring')) return 'The Scoring House';
  if (l.includes('fired earth')) return 'Fired Earth Music';
  if (l.includes('percolate')) return 'Percolate Music';
  if (l.includes('osmosis')) return 'Osmosis Music';
  if (l.includes('juice')) return 'Juice Music';
  if (l.includes('evolution')) return 'Evolution Media Music';
  if (l.includes('jingle')) return 'Jingle Punks';
  if (l.includes('pennybank')) return 'Pennybank Tunes';
  if (l.includes('cavendish')) return 'Cavendish Music';
  if (l.includes('bruton')) return 'Bruton Music';
  return p.length > 40 ? p.slice(0, 40) + '…' : p;
}

function guessLibFromTitle(t) {
  if (/west.one/i.test(t)) return 'West One Music';
  if (/extreme/i.test(t)) return 'Extreme Music';
  if (/\bKSD/i.test(t)) return 'KillerTracks';
  if (/\bJCE/i.test(t)) return 'Juice Music';
  if (/BMGPM/i.test(t)) return 'BMG Production Music';
  return null;
}

function parseXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const epSheets = wb.SheetNames.filter(n => /EP/i.test(n.replace(/['\s]/g, '')));
  if (!epSheets.length) throw new Error('EPシートが見つかりません');
  const episodes = {};
  for (const sn of epSheets) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cues = []; let cur = null;
    for (const row of rows) {
      const seq = row[0], title = String(row[1] || '').trim();
      const origin = String(row[4] || '').trim(), role = String(row[14] || '').trim();
      const pub = String(row[17] || '').trim();
      const isSeq = seq !== '' && !isNaN(Number(seq)) && Number(seq) > 0;
      const isTitle = title && title !== '楽曲名をローマ字で入力' && !title.startsWith('キュー') && title !== 'Cue Title' && !/^[\u3000\s]+$/.test(title);
      if (isSeq && isTitle) {
        // Duration: col12=分, col13=秒
        const durMin = Number(row[12]) || 0;
        const durSec2 = Number(row[13]) || 0;
        const durSec = durMin * 60 + durSec2;
        // In/Out (参考用)
        const inH=Number(row[6]||0), inM=Number(row[7]||0), inS=Number(row[8]||0);
        const outH=Number(row[9]||0), outM=Number(row[10]||0), outS=Number(row[11]||0);
        const inSec = inH*3600 + inM*60 + inS;
        const outSec = outH*3600 + outM*60 + outS;
        cur = { seq: Number(seq), title, origin, type: classifyOrigin(origin), isM: M_RE.test(title), publishers: [], inSec, outSec, durSec };
        cues.push(cur);
      } else if (cur && (role.includes('Publisher') || role.includes('出版社'))) {
        if (pub && pub !== 'nan' && pub !== '') cur.publishers.push(pub);
      }
    }
    const key = sn.replace(/['\s]/g, '').replace('GASEP', 'EP');
    episodes[key] = { label: sn.replace(/'/g, '').trim(), cues };
  }
  return episodes;
}

function parseWav(text) {
  const lines = text.trim().split('\n').filter(l => l.trim());
  const keys = new Set();
  for (const l of lines) {
    const fname = l.trim();
    const m = extractM(fname);
    if (m) keys.add(m);
    // RAWキーのみ保存（正規化はcheckDelivery側でやる）
    keys.add(extractFileKey(fname));
  }
  return keys;
}

// 秒数を mm:ss 表示に変換
function fmtDur(sec) {
  if (!sec) return '0:00';
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = Math.floor(sec % 60);
  if (h > 0) return `${h}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
  return `${m}:${String(s).padStart(2,'0')}`;
}

// 書き下ろしタイトルがwavSetに含まれているか判定
// wavSetはRAWファイルキー（正規化なし）を保持
// 戻り値: { delivered, fuzzy, cueKey, fileKey }
//   fuzzy=true → マッチしたがキューとファイル名が違う（要修正）
//   cueKey = キューシート側の正規化前キー
//   fileKey = マッチしたファイル側のキー
function checkDelivery(title, wavSet) {
  const rawCue = title.toUpperCase().trim();  // キューシートの生タイトル（大文字）

  // 完全一致（大文字そのまま）→ fuzzy=false
  if (wavSet.has(rawCue)) return { delivered: true, fuzzy: false, cueKey: rawCue, fileKey: rawCue };

  // Mナンバー完全一致
  const m = extractM(title);
  if (m && wavSet.has(m)) return { delivered: true, fuzzy: false, cueKey: rawCue, fileKey: m };

  // 正規化して一致 → fuzzy=true（rawCue != fileKey なので表記ゆれ）
  const normCue = normalizeTitle(title);
  if (wavSet.has(normCue)) {
    return { delivered: true, fuzzy: normCue !== rawCue, cueKey: rawCue, fileKey: normCue };
  }

  // wavSet内の各キーと正規化比較
  for (const fileKey of wavSet) {
    const normFile = normalizeTitle(fileKey);
    if (normFile === normCue) {
      // どちらかが違う → fuzzy=true
      return { delivered: true, fuzzy: fileKey !== rawCue, cueKey: rawCue, fileKey };
    }
    // プレフィックス違い（KQZ_M9A vs M9A）
    if (normFile.endsWith('_' + normCue) || normCue.endsWith('_' + normFile)) {
      return { delivered: true, fuzzy: true, cueKey: rawCue, fileKey };
    }
  }

  return { delivered: false, fuzzy: false, cueKey: rawCue, fileKey: null };
}

function computeStats(episodes, wavSet, excludeList) {
  // 書き下ろし全タイトルを追跡（Mナンバー以外も含む）
  const commStats = {}; // key: titleKey → { title, count, eps, delivered }
  let total = 0, comm = 0, lib = 0, lic = 0, pd = 0, free = 0, oth = 0;
  const libPubs = {}, epStats = {};

  for (const [k, { label, cues }] of Object.entries(episodes)) {
    epStats[k] = { label, total: 0, commissioned: 0, library: 0, licensed: 0, pd: 0, free: 0, other: 0,
                   durComm: 0, durLib: 0, durLic: 0, durPd: 0, durFree: 0, durOth: 0 };
    for (const c of cues) {
      total++; epStats[k].total++; epStats[k][c.type] = (epStats[k][c.type]||0) + 1;
      const dur = c.durSec || 0;
      if (c.type === 'commissioned') { epStats[k].durComm += dur; }
      else if (c.type === 'library')  { epStats[k].durLib  += dur; }
      else if (c.type === 'licensed') { epStats[k].durLic  += dur; }
      else if (c.type === 'pd')       { epStats[k].durPd   += dur; }
      else if (c.type === 'free')     { epStats[k].durFree += dur; }
      else                            { epStats[k].durOth  += dur; }
      if (c.type === 'commissioned') {
        comm++;
        const tk = titleKey(c.title);
        if (!commStats[tk]) commStats[tk] = (() => {
          const excluded = (excludeList || []).some(ex => ex.toLowerCase() === c.title.toLowerCase());
          if (excluded) return { title: c.title, count: 0, eps: new Set(), delivered: true, fuzzy: false, excluded: true, cueKey: null, fileKey: null };
          const d = checkDelivery(c.title, wavSet);
          return { title: c.title, count: 0, eps: new Set(), delivered: d.delivered, fuzzy: d.fuzzy, excluded: false, cueKey: d.cueKey, fileKey: d.fileKey };
        })();
        commStats[tk].count++;
        commStats[tk].eps.add(label.replace(/GAS|'|\s/g, '').replace(/EP0?/, 'EP').trim());
      } else if (c.type === 'library') {
        lib++;
        let pubs = [...c.publishers];
        if (!pubs.length) { const g = guessLibFromTitle(c.title); if (g) pubs = [g]; }
        for (const p of pubs) { const n = normPub(p); libPubs[n] = (libPubs[n] || 0) + 1; }
      } else if (c.type === 'licensed') lic++;
      else if (c.type === 'pd') pd++;
      else if (c.type === 'free') free++;
      else oth++;
    }
  }

  const totalComm = Object.keys(commStats).length;
  const delivComm = Object.values(commStats).filter(s => s.delivered && !s.excluded).length;
  const excludedCount = Object.values(commStats).filter(s => s.excluded).length;

  // 全EP合計の尺
  let durComm=0, durLib=0, durLic=0, durPd=0, durFree=0, durOth=0;
  for (const s of Object.values(epStats)) {
    durComm += s.durComm; durLib += s.durLib; durLic += s.durLic;
    durPd   += s.durPd;   durFree+= s.durFree; durOth += s.durOth;
  }
  const durTotal = durComm + durLib + durLic + durPd + durFree + durOth;

  return { total, comm, lib, lic, pd, free, oth, totalComm, delivComm, excludedCount, commStats, libPubs, epStats,
           durComm, durLib, durLic, durPd, durFree, durOth, durTotal };
}

// ─── Project CRUD ────────────────────────────────────────────
function createProject() {
  const name = document.getElementById('newProjectName').value.trim();
  if (!name) { toast('作品名を入力してください'); return; }
  if (!newXlsxData) { toast('キューシートをアップしてください'); return; }
  const wavText = document.getElementById('newWavText').value;
  const wavSet = wavText.trim() ? parseWav(wavText) : new Set();
  const id = 'proj_' + Date.now();
  projects[id] = {
    id, name,
    createdAt: new Date().toISOString(),
    episodes: newXlsxData,
    wavSet,
    excludeList: ['Netflix Logo ID']  // 納品対象外タイトル
  };
  saveProjects();
  newXlsxData = null;
  document.getElementById('newProjectName').value = '';
  document.getElementById('newWavText').value = '';
  document.getElementById('newXlsxLabel').textContent = 'クリックまたはドロップ';
  closeModal('newProjectModal');
  renderSidebar();
  renderHome();
  openDashboard(id);
  toast(`「${name}」を作成しました`);
}

function archiveProject(id) {
  const p = projects[id];
  if (!p) return;
  p.archived = !p.archived;
  saveProjects();
  renderSidebar();
  renderHome();
  if (p.archived && currentProjectId === id) showView('home');
  toast(p.archived ? `「${p.name}」をアーカイブしました` : `「${p.name}」を復元しました`);
}

function deleteProject(id) {
  const name = projects[id]?.name;
  if (!confirm(`「${name}」を完全削除しますか？この操作は取り消せません`)) return;
  delete projects[id];
  saveProjects();
  renderSidebar();
  renderHome();
  if (currentProjectId === id) showView('home');
}

// ─── Views ───────────────────────────────────────────────────
function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(`view-${name}`).classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  if (name === 'home') document.querySelectorAll('.nav-btn')[0].classList.add('active');
  if (name === 'compare') document.querySelectorAll('.nav-btn')[1].classList.add('active');
  if (name === 'compare') renderCompare();
}

function openDashboard(id) {
  currentProjectId = id;
  const p = projects[id];
  const stats = computeStats(p.episodes, p.wavSet, p.excludeList || []);
  currentDash = { p, stats };

  // 書き下ろしタイトル一覧を表示用に整列
  delivRows = Object.values(stats.commStats).sort((a, b) => {
    // Mナンバー形式を優先、次にアルファベット順
    const mA = extractM(a.title), mB = extractM(b.title);
    if (mA && mB) {
      const na = parseInt(mA.replace(/\D/g,'')), nb = parseInt(mB.replace(/\D/g,''));
      return na !== nb ? na - nb : mA.localeCompare(mB);
    }
    if (mA) return -1; if (mB) return 1;
    return a.title.localeCompare(b.title);
  }).map(s => ({ m: s.title, ok: s.delivered, fuzzy: s.fuzzy, excluded: s.excluded, cueKey: s.cueKey, fileKey: s.fileKey, count: s.count, epsStr: [...s.eps].sort().join(', ') }));

  document.getElementById('dashTitle').textContent = p.name;
  document.getElementById('dashSub').textContent = `${Object.keys(p.episodes).length} エピソード · 最終更新: ${new Date(p.createdAt).toLocaleDateString('ja-JP')}`;

  // KPIs
  const { total, comm, lib, totalComm, delivComm } = stats;
  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi"><div class="kpi-label">総キュー数</div><div class="kpi-value">${total}</div><div class="kpi-sub">${Object.keys(p.episodes).length} エピソード</div></div>
    <div class="kpi"><div class="kpi-label">書き下ろし (種類)</div><div class="kpi-value">${totalComm}</div><div class="kpi-sub">ユニークタイトル</div></div>
    <div class="kpi"><div class="kpi-label">納品済み</div><div class="kpi-value green">${delivComm}</div><div class="kpi-sub">${pct(delivComm, totalComm)}%</div></div>
    <div class="kpi"><div class="kpi-label">未納品</div><div class="kpi-value ${totalComm - delivComm - (stats.excludedCount||0) > 0 ? 'red' : 'green'}">${totalComm - delivComm - (stats.excludedCount||0)}</div><div class="kpi-sub">${totalComm - delivComm - (stats.excludedCount||0) > 0 ? '要確認' : '完了'}${stats.excludedCount ? ` / 対象外${stats.excludedCount}` : ''}</div></div>
    <div class="kpi"><div class="kpi-label">書き下ろし (使用)</div><div class="kpi-value">${comm}</div><div class="kpi-sub">${pct(comm, total)}% / ${fmtDur(stats.durComm)}</div></div>
    <div class="kpi"><div class="kpi-label">ライブラリー</div><div class="kpi-value amber">${stats.lib}</div><div class="kpi-sub">${pct(stats.lib, total)}% / ${fmtDur(stats.durLib)}</div></div>
  `;

  // Donut
  drawDonut(stats, stats);

  // EP bars
  const eps = Object.keys(stats.epStats);
  document.getElementById('epStackedBars').innerHTML = eps.map(k => {
    const s = stats.epStats[k];
    const short = s.label.replace('GAS', '').replace(/EP0?/, 'EP').trim();
    return `<div class="stacked-row">
      <div class="stacked-label"><span>${short}</span><span style="font-family:var(--mono)">${s.total}</span></div>
      <div class="stacked-track">
        ${['commissioned', 'library', 'licensed', 'other'].map(t => s[t] ? `<div style="width:${pct(s[t], s.total)}%;background:${COLORS[t]}"></div>` : '').join('')}
      </div>
    </div>`;
  }).join('') + `<div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:10px">
    ${Object.entries({ commissioned: '書き下ろし', library: 'ライブラリー', licensed: 'ライセンス', other: 'その他' }).map(([k, l]) =>
      `<span style="font-size:10px;display:flex;align-items:center;gap:5px;color:var(--text2)"><span style="width:8px;height:8px;background:${COLORS[k]};border-radius:2px;display:inline-block"></span>${l}</span>`
    ).join('')}
  </div>`;

  // Duration breakdown panel - 5 categories
  const { durComm, durLib, durLic, durPd, durFree, durOth, durTotal } = stats;
  const durRows = [
    { key: 'commissioned', dur: durComm },
    { key: 'library',      dur: durLib  },
    { key: 'licensed',     dur: durLic  },
    { key: 'pd',           dur: durPd   },
    { key: 'free',         dur: durFree },
    { key: 'other',        dur: durOth  },
  ].filter(d => d.dur > 0);
  document.getElementById('durPanel').innerHTML = durRows.map(d => `
    <div class="bar-row" style="margin-bottom:10px">
      <div class="bar-label">
        <span style="display:flex;align-items:center;gap:6px">
          <span style="width:8px;height:8px;background:${COLORS[d.key]};border-radius:2px;display:inline-block"></span>
          ${TYPE_LABELS[d.key]}
        </span>
        <span style="font-family:var(--mono);font-size:12px">${fmtDur(d.dur)}&nbsp;<span style="color:var(--text3)">(${durTotal?Math.round(d.dur/durTotal*100):0}%)</span></span>
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${durTotal?Math.round(d.dur/durTotal*100):0}%;background:${COLORS[d.key]}"></div>
      </div>
    </div>`).join('') +
    `<div style="margin-top:10px;padding-top:10px;border-top:1px solid var(--border);display:flex;justify-content:space-between;font-size:12px">
      <span style="color:var(--text3)">全体合計尺</span>
      <span style="font-family:var(--mono)">${fmtDur(durTotal)}</span>
    </div>`;

  // EP duration bars
  const epKeys = Object.keys(stats.epStats);
  const maxEpDur = Math.max(...epKeys.map(k => {
    const s = stats.epStats[k];
    return s.durComm + s.durLib + s.durLic + (s.durPd||0) + (s.durFree||0) + s.durOth;
  }), 1);
  document.getElementById('epDurBars').innerHTML = epKeys.map(k => {
    const s = stats.epStats[k];
    const epDur = s.durComm + s.durLib + s.durLic + (s.durPd||0) + (s.durFree||0) + s.durOth;
    const short = s.label.replace('GAS','').replace(/EP0?/,'EP').trim();
    return `<div class="stacked-row">
      <div class="stacked-label">
        <span>${short}</span>
        <span style="font-family:var(--mono);font-size:11px">${fmtDur(epDur)}</span>
      </div>
      <div class="stacked-track" style="height:6px">
        ${s.durComm?`<div style="width:${Math.round(s.durComm/maxEpDur*100)}%;background:${COLORS.commissioned}"></div>`:''}
        ${s.durLib ?`<div style="width:${Math.round(s.durLib /maxEpDur*100)}%;background:${COLORS.library}"></div>`:''}
        ${s.durLic ?`<div style="width:${Math.round(s.durLic /maxEpDur*100)}%;background:${COLORS.licensed}"></div>`:''}
        ${s.durPd  ?`<div style="width:${Math.round(s.durPd  /maxEpDur*100)}%;background:${COLORS.pd}"></div>`:''}
        ${s.durFree?`<div style="width:${Math.round(s.durFree/maxEpDur*100)}%;background:${COLORS.free}"></div>`:''}
        ${s.durOth ?`<div style="width:${Math.round(s.durOth /maxEpDur*100)}%;background:${COLORS.other}"></div>`:''}
      </div>
    </div>`;
  }).join('');

  // Delivery
  renderDelivRows(delivRows);

  // Library
  const libSorted = Object.entries(stats.libPubs).sort((a, b) => b[1] - a[1]);
  const libTotal = libSorted.reduce((s, [, n]) => s + n, 0);
  document.getElementById('libBody').innerHTML = libSorted.map(([pub, n], i) => `
    <tr>
      <td>${pub}</td>
      <td class="mono">${n}</td>
      <td>
        <div style="display:flex;align-items:center;gap:8px">
          <div style="flex:1;height:4px;background:var(--bg4);border-radius:2px;overflow:hidden">
            <div style="width:${pct(n, libTotal)}%;height:100%;background:${LIB_COLORS[i % LIB_COLORS.length]};border-radius:2px"></div>
          </div>
          <span class="mono" style="font-size:11px;color:var(--text2);min-width:28px;text-align:right">${pct(n, libTotal)}%</span>
        </div>
      </td>
    </tr>`).join('');

  // Budget section
  renderBudget();

  // EP select
  const epSel = document.getElementById('epSelect');
  epSel.innerHTML = eps.map(k => `<option value="${k}">${stats.epStats[k].label.replace('GAS', '').trim()}</option>`).join('');

  // Highlight in sidebar
  document.querySelectorAll('.nav-project').forEach(b => b.classList.remove('active'));
  const navBtn = document.getElementById(`nav-${id}`);
  if (navBtn) navBtn.classList.add('active');

  showView('dashboard');
  renderEpDetail();
}

// 2つの文字列で違う部分をハイライトして返す
function diffHighlight(a, b) {
  // 共通プレフィックスの長さ
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  // 共通サフィックスの長さ
  let j = 0;
  while (j < a.length - i && j < b.length - i && a[a.length-1-j] === b[b.length-1-j]) j++;
  const diffA = a.slice(i, a.length - j || undefined);
  const diffB = b.slice(i, b.length - j || undefined);
  const pre = a.slice(0, i), suf = j ? a.slice(a.length - j) : '';
  const preB = b.slice(0, i), sufB = j ? b.slice(b.length - j) : '';
  const hlA = diffA ? `<span style="background:rgba(229,9,20,0.25);border-radius:2px;padding:0 1px">${diffA}</span>` : '';
  const hlB = diffB ? `<span style="background:rgba(29,185,84,0.25);border-radius:2px;padding:0 1px">${diffB}</span>` : '';
  return {
    a: pre + hlA + suf,
    b: preB + hlB + sufB
  };
}

function renderDelivRows(rows) {
  document.getElementById('delivBody').innerHTML = rows.map(r => {
    const rowClass = !r.ok ? 'row-ng' : r.fuzzy ? 'row-warn' : '';

    let titleCell, statusCell;
    if (r.excluded) {
      titleCell = `<span class="mono">${r.m}</span>`;
      statusCell = `<span class="badge b-oth">納品対象外</span>`;
    } else if (r.ok && r.fuzzy && r.cueKey && r.fileKey) {
      // キューシートタイトルとファイル名の差分を表示
      const diff = diffHighlight(r.cueKey, r.fileKey);
      titleCell = `<span class="mono">${r.m}</span>
        <div style="margin-top:4px;font-size:10px;color:var(--text3);line-height:1.6">
          <span style="display:inline-block;width:72px;color:var(--text3)">キューシート:</span>
          <span class="mono" style="font-size:10px">${diff.a}</span><br>
          <span style="display:inline-block;width:72px;color:var(--text3)">音源ファイル:</span>
          <span class="mono" style="font-size:10px">${diff.b}</span>
          <span style="display:block;margin-top:2px;color:var(--amber);font-size:10px">↑ どちらかを統一してください</span>
        </div>`;
      statusCell = `<span class="badge b-ok">✓ 納品済</span>&nbsp;<span class="badge b-warn">⚠ 表記ゆれ</span>`;
    } else {
      titleCell = `<span class="mono">${r.m}</span>`;
      statusCell = r.ok
        ? `<span class="badge b-ok">✓ 納品済</span>`
        : `<span class="badge b-ng">✗ 未納品</span>`;
    }

    return `<tr class="${rowClass}">
      <td>${titleCell}</td>
      <td>${statusCell}</td>
      <td class="mono" style="text-align:center">${r.count}</td>
      <td style="font-size:11px;color:var(--text2)">${r.epsStr}</td>
    </tr>`;
  }).join('');
}

window.filterDeliv = function () {
  const q = document.getElementById('delivFilter').value.toUpperCase();
  const st = document.getElementById('delivStatus').value;
  renderDelivRows(delivRows.filter(r => {
    const matchQ = !q || r.m.toUpperCase().includes(q);
    const matchSt = !st || (
      st === 'ok'   ? (r.ok && !r.excluded) :
      st === 'ng'   ? (!r.ok && !r.excluded) :
      st === 'warn' ? r.fuzzy :
      st === 'excl' ? r.excluded : true
    );
    return matchQ && matchSt;
  }));
};

window.renderEpDetail = function () {
  if (!currentDash) return;
  const k = document.getElementById('epSelect').value;
  const ep = currentDash.p.episodes[k];
  if (!ep) return;
  const wav = currentDash.p.wavSet;
  const tl = { commissioned: '<span class="badge b-comm">書き下ろし</span>', library: '<span class="badge b-lib">ライブラリー</span>', licensed: '<span class="badge b-lic">ライセンス</span>', other: '<span class="badge b-oth">その他</span>' };
  document.getElementById('epDetailBody').innerHTML = ep.cues.map(c => {
    let delivCell = '<span style="color:var(--text3)">—</span>';
    if (c.type === 'commissioned') {
      const d = checkDelivery(c.title, wav);
      if (d.delivered && d.fuzzy) {
        delivCell = '<span class="badge b-ok">✓</span>&nbsp;<span class="badge b-warn">⚠</span>';
      } else if (d.delivered) {
        delivCell = '<span class="badge b-ok">✓</span>';
      } else {
        delivCell = '<span class="badge b-ng">✗</span>';
      }
    }
    return `<tr>
      <td class="mono" style="color:var(--text3)">${c.seq}</td>
      <td>${c.title}</td>
      <td>${tl[c.type] || ''}</td>
      <td>${delivCell}</td>
    </tr>`;
  }).join('');
};

// ─── Donut ───────────────────────────────────────────────────
function drawDonut(stats, fullStats) { fullStats = fullStats || {};
  const cv = document.getElementById('donutCanvas');
  if (!cv) return;
  const ctx = cv.getContext('2d');
  const { total, comm, lib, lic, oth } = stats;
  const data = [
    { l: TYPE_LABELS.commissioned, v: comm, c: COLORS.commissioned },
    { l: TYPE_LABELS.library,      v: lib,  c: COLORS.library },
    { l: TYPE_LABELS.licensed,     v: lic,  c: COLORS.licensed },
    { l: TYPE_LABELS.pd,           v: stats.pd   || 0, c: COLORS.pd },
    { l: TYPE_LABELS.free,         v: stats.free || 0, c: COLORS.free },
    { l: TYPE_LABELS.other,        v: oth,  c: COLORS.other },
  ].filter(d => d.v > 0);
  let ang = -Math.PI / 2;
  ctx.clearRect(0, 0, 100, 100);
  for (const d of data) {
    const sw = (d.v / total) * Math.PI * 2;
    ctx.beginPath(); ctx.moveTo(50, 50); ctx.arc(50, 50, 44, ang, ang + sw); ctx.closePath();
    ctx.fillStyle = d.c; ctx.fill(); ang += sw;
  }
  ctx.beginPath(); ctx.arc(50, 50, 28, 0, Math.PI * 2); ctx.fillStyle = '#141414'; ctx.fill();
  ctx.fillStyle = '#F0EBE0'; ctx.font = '500 12px DM Mono,monospace'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(total, 50, 44); ctx.font = '9px DM Sans,sans-serif'; ctx.fillStyle = '#555'; ctx.fillText('total', 50, 57);
  document.getElementById('donutLegend').innerHTML = data.map(d => {
    const pct = Math.round(d.v / total * 100);
    return `<div class="legend-item">
      <div class="legend-dot" style="background:${d.c}"></div>
      <span style="color:var(--text2)">${d.l}</span>
      <span class="legend-val">${d.v} <span style="color:var(--text3)">(${pct}%)</span></span>
    </div>`;
  }).join('');
}

// ─── Compare view ─────────────────────────────────────────────
function renderCompare() {
  const ids = Object.keys(projects).filter(id => !projects[id].archived);
  if (ids.length < 2) {
    document.getElementById('compareContent').innerHTML = `<div class="empty"><div class="empty-icon">📊</div><div class="empty-title">比較するには2作品以上必要です</div><div class="empty-sub">プロジェクトを追加してください</div></div>`;
    return;
  }
  const all = ids.map(id => {
    const p = projects[id];
    const s = computeStats(p.episodes, p.wavSet, p.excludeList || []);
    const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
    const counts = { commissioned: s.comm, library: s.lib, licensed: s.lic, pd: s.pd||0, free: s.free||0, other: s.oth };
    const durs = { commissioned: s.durComm, library: s.durLib, licensed: s.durLic, pd: s.durPd||0, free: s.durFree||0, other: s.durOth };
    return { id, name: p.name, eps: Object.keys(p.episodes).length, total: s.total,
      comm: s.comm, lib: s.lib, totalComm: s.totalComm, delivered: s.delivComm,
      commPct: pct(s.comm, s.total), libPct: pct(s.lib, s.total), delivPct: pct(s.delivComm, s.totalComm),
      counts, durs, durTotal: s.durTotal, budget: p.budget || 0 };
  });

  const cols = `grid-template-columns: 160px repeat(${all.length}, 1fr)`;
  const hdr = `<div></div>` + all.map(a => `<div style="font-size:12px;font-weight:500;padding:8px 12px;text-align:center">${a.name}</div>`).join('');
  const row = (label, fn, suffix = '') => `
    <div class="compare-row" style="${cols}">
      <div class="compare-metric">${label}</div>
      ${all.map(a => `<div class="compare-val">${fn(a)}${suffix}</div>`).join('')}
    </div>`;

  // 各作品の種別割合（キュー数・尺）を積み上げバーで比較
  const typeOrder = ['commissioned','library','licensed','pd','free','other'];

  const stackedBar = (a, mode) => {
    // mode: 'cue' or 'dur'
    const tot = mode === 'cue' ? a.total : a.durTotal;
    if (!tot) return '<span style="color:var(--text3);font-size:11px">データなし</span>';
    return `<div style="display:flex;height:10px;border-radius:3px;overflow:hidden;gap:1px">
      ${typeOrder.map(k => {
        const v = mode === 'cue' ? (a.counts[k]||0) : (a.durs[k]||0);
        const w = Math.round(v/tot*100);
        return w > 0 ? `<div style="width:${w}%;background:${COLORS[k]}" title="${TYPE_LABELS[k]}: ${w}%"></div>` : '';
      }).join('')}
    </div>
    <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:5px">
      ${typeOrder.map(k => {
        const v = mode === 'cue' ? (a.counts[k]||0) : (a.durs[k]||0);
        if (!v) return '';
        const w = Math.round(v/tot*100);
        const label = mode === 'cue' ? `${v}キュー` : fmtDur(v);
        return `<span style="font-size:10px;display:flex;align-items:center;gap:3px;color:var(--text2)">
          <span style="width:7px;height:7px;background:${COLORS[k]};border-radius:1px;display:inline-block;flex-shrink:0"></span>
          ${TYPE_LABELS[k]} ${label}(${w}%)
        </span>`;
      }).filter(Boolean).join('')}
    </div>`;
  };

  document.getElementById('compareContent').innerHTML = `
    <div style="overflow-x:auto;margin-bottom:24px">
      <div class="compare-header" style="${cols};display:grid">${hdr}</div>
      <div class="compare-grid">
        ${row('エピソード数', a => a.eps)}
        ${row('総キュー数', a => a.total)}
        ${row('書き下ろし (種類)', a => a.totalComm)}
        ${row('納品率', a => `<span style="color:${a.delivPct===100?'var(--green)':'var(--amber)'}">${a.delivPct}%</span> <span style="color:var(--text3);font-size:11px">(${a.delivered}/${a.totalComm})</span>`)}
        ${row('合計尺', a => `<span style="font-family:var(--mono)">${fmtDur(a.durTotal)}</span>`)}
        ${all.length ? `<div class="compare-row" style="${cols}">
          <div class="compare-metric">種別構成<br><span style="color:var(--text3);font-size:10px">キュー数ベース</span></div>
          ${all.map(a => `<div style="padding:8px 12px">${stackedBar(a,'cue')}</div>`).join('')}
        </div>` : ''}
        ${all.length ? `<div class="compare-row" style="${cols}">
          <div class="compare-metric">種別構成<br><span style="color:var(--text3);font-size:10px">尺ベース</span></div>
          ${all.map(a => `<div style="padding:8px 12px">${stackedBar(a,'dur')}</div>`).join('')}
        </div>` : ''}
        ${all.some(a=>a.budget>0) ? row('書き下ろし予算', a => a.budget ? `<span style="font-family:var(--mono)">¥${a.budget.toLocaleString()}</span>` : '—') : ''}
        ${all.some(a=>a.budget>0) ? row('1曲あたり平均', a => a.budget && a.totalComm ? `<span style="font-family:var(--mono)">¥${Math.round(a.budget/a.totalComm).toLocaleString()}</span>` : '—') : ''}
      </div>
    </div>`;
}

// ─── Home ─────────────────────────────────────────────────────
function renderHome() {
  const ids = Object.keys(projects);
  const active = ids.filter(id => !projects[id].archived);
  const archived = ids.filter(id => projects[id].archived);
  document.getElementById('homeEmpty').style.display = active.length ? 'none' : 'block';

  function cardHtml(id, isArchived) {
    const p = projects[id];
    const s = computeStats(p.episodes, p.wavSet, p.excludeList || []);
    const fuzzyCount = Object.values(s.commStats).filter(c => c.fuzzy).length;
    return `<div class="project-card${isArchived ? ' archived-card' : ''}" onclick="${isArchived ? '' : `openDashboard('${id}')`}" style="${isArchived ? 'opacity:0.5;cursor:default' : ''}">
      <div class="card-actions">
        ${!isArchived ? `<button class="icon-btn" onclick="event.stopPropagation();openRenameModal('${id}')" title="名前を変更">✏</button>` : ''}
        <button class="icon-btn" onclick="event.stopPropagation();archiveProject('${id}')" title="${isArchived ? '復元' : 'アーカイブ'}">${isArchived ? '↩' : '⊘'}</button>
        ${isArchived ? `<button class="icon-btn" onclick="event.stopPropagation();deleteProject('${id}')" title="完全削除" style="color:var(--red)">✕</button>` : ''}
      </div>
      <div class="card-name">${p.name}${isArchived ? ' <span style="font-size:10px;color:var(--text3)">[アーカイブ]</span>' : ''}</div>
      <div class="card-meta">${new Date(p.createdAt).toLocaleDateString('ja-JP')} · ${Object.keys(p.episodes).length} EP</div>
      <div class="card-stats">
        <div class="stat"><strong>${s.total}</strong> キュー</div>
        <div class="stat"><strong>${Math.round((s.comm||0)/(s.total||1)*100)}%</strong> 書き下ろし</div>
        <div class="stat"><strong style="color:${s.totalComm - s.delivComm > 0 ? 'var(--amber)' : 'var(--green)'}">${s.delivComm}/${s.totalComm}</strong> 納品</div>
        ${fuzzyCount > 0 ? `<div class="stat"><strong style="color:var(--amber)">⚠ ${fuzzyCount}</strong> 表記ゆれ</div>` : ''}
      </div>
    </div>`;
  }

  document.getElementById('projectGrid').innerHTML = active.map(id => cardHtml(id, false)).join('');

  // archived section
  let archiveSection = document.getElementById('archiveSection');
  if (!archiveSection) {
    archiveSection = document.createElement('div');
    archiveSection.id = 'archiveSection';
    document.getElementById('view-home').appendChild(archiveSection);
  }
  if (archived.length) {
    archiveSection.innerHTML = `
      <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:0.08em;margin:8px 0 12px">アーカイブ済み</div>
      <div class="project-grid">${archived.map(id => cardHtml(id, true)).join('')}</div>`;
  } else {
    archiveSection.innerHTML = '';
  }
}

function renderSidebar() {
  const ids = Object.keys(projects).filter(id => !projects[id].archived);
  document.getElementById('projectNav').innerHTML = ids.map(id => {
    const p = projects[id];
    return `<button class="nav-btn nav-project" id="nav-${id}" onclick="openDashboard('${id}')">
      <div class="dot"></div><span>${p.name}</span>
    </button>`;
  }).join('');
}

// ─── Modals ──────────────────────────────────────────────────
function openNewProjectModal() {
  newXlsxData = null;
  document.getElementById('newProjectModal').classList.add('open');
}

function openUploadModal() {
  updateXlsxData = null;
  if (currentProjectId && projects[currentProjectId]) {
    document.getElementById('updateWavText').value = '';
    document.getElementById('updateXlsxLabel').textContent = 'クリックまたはドロップ';
  }
  document.getElementById('uploadModal').classList.add('open');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('open');
}

function applyUpdate() {
  if (!currentProjectId) return;
  const p = projects[currentProjectId];
  const wavText = document.getElementById('updateWavText').value;
  if (wavText.trim()) p.wavSet = parseWav(wavText);
  if (updateXlsxData) p.episodes = updateXlsxData;
  p.updatedAt = new Date().toISOString();
  saveProjects();
  closeModal('uploadModal');
  openDashboard(currentProjectId);
  toast('データを更新しました');
}

// Xlsx loaders
function loadNewXlsx(file) {
  if (!file) return;
  const r = new FileReader();
  r.onload = e => {
    try {
      newXlsxData = parseXlsx(e.target.result);
      const count = Object.keys(newXlsxData).length;
      document.getElementById('newXlsxLabel').textContent = `✓ ${count} EP 読み込み済み`;
    } catch(err) { toast('読み込みエラー: ' + err.message); }
  };
  r.readAsArrayBuffer(file);
}

function loadUpdateXlsx(file) {
  if (!file) return;
  const r = new FileReader();
  r.onload = e => {
    try {
      updateXlsxData = parseXlsx(e.target.result);
      const count = Object.keys(updateXlsxData).length;
      document.getElementById('updateXlsxLabel').textContent = `✓ ${count} EP 読み込み済み`;
    } catch(err) { toast('読み込みエラー: ' + err.message); }
  };
  r.readAsArrayBuffer(file);
}

// Drag & drop for xlsx
['newXlsxDrop','updateXlsxDrop'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('over'); });
  el.addEventListener('dragleave', () => el.classList.remove('over'));
  el.addEventListener('drop', e => {
    e.preventDefault(); el.classList.remove('over');
    const fn = id === 'newXlsxDrop' ? loadNewXlsx : loadUpdateXlsx;
    fn(e.dataTransfer.files[0]);
  });
});

// ─── Exports ─────────────────────────────────────────────────
function exportCSV() {
  if (!currentDash) return;
  const { p, stats } = currentDash;
  const rows = [['EP', 'Seq#', 'タイトル', '種別', 'M番号', '納品状況', '登場回数']];
  for (const [k, { label, cues }] of Object.entries(p.episodes)) {
    for (const c of cues) {
      const isM = c.isM;
      const dResult = isM ? checkDelivery(c.title, p.wavSet) : null;
      const delivered = isM ? (dResult.delivered ? (dResult.fuzzy ? '納品済(表記ゆれ要確認)' : '納品済') : '未納品') : '—';
      const count = isM ? (stats.mStats[c.title.toUpperCase()]?.count || 0) : '—';
      rows.push([label, c.seq, c.title, { commissioned: '書き下ろし', library: 'ライブラリー', licensed: 'ライセンス', other: 'その他' }[c.type], isM ? c.title.toUpperCase() : '—', delivered, count]);
    }
  }
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${p.name}_music_analysis.csv`;
  a.click();
  toast('CSVをダウンロードしました');
}

function exportPDF() {
  window.print();
}

// ─── Utilities ───────────────────────────────────────────────
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg; el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2500);
}

// ─── Budget ──────────────────────────────────────────────────
function renderBudget() {
  if (!currentDash) return;
  const p = currentDash.p;
  const stats = currentDash.stats;
  const budget = p.budget || {};  // { titleKey: yen }

  // KPI cards
  const totalBudget = Object.values(budget).reduce((s,v)=>s+(v||0),0);
  const commTitles = Object.values(stats.commStats);
  const paidCount = commTitles.filter(c => budget[c.title] > 0).length;
  const avgPerTrack = paidCount ? Math.round(totalBudget / paidCount) : 0;
  document.getElementById('budgetKpis').innerHTML = `
    <div class="kpi"><div class="kpi-label">合計費用</div><div class="kpi-value" style="font-size:18px">¥${totalBudget.toLocaleString()}</div></div>
    <div class="kpi"><div class="kpi-label">入力済み曲数</div><div class="kpi-value">${paidCount} <span style="font-size:12px;color:var(--text2)">/ ${commTitles.length}</span></div></div>
    <div class="kpi"><div class="kpi-label">1曲あたり平均</div><div class="kpi-value" style="font-size:18px">${avgPerTrack ? '¥'+avgPerTrack.toLocaleString() : '—'}</div></div>
  `;

  // Per-track input rows
  const sorted = commTitles.sort((a,b) => {
    const mA = extractM(a.title), mB = extractM(b.title);
    if (mA && mB) { const na=parseInt(mA.replace(/\D/g,'')),nb=parseInt(mB.replace(/\D/g,'')); return na-nb; }
    if (mA) return -1; if (mB) return 1;
    return a.title.localeCompare(b.title);
  });
  document.getElementById('budgetRows').innerHTML = `
    <table>
      <thead><tr><th>楽曲タイトル</th><th>登場回数</th><th style="width:160px">費用（円）</th></tr></thead>
      <tbody>
        ${sorted.map(c => `<tr>
          <td class="mono" style="font-size:12px">${c.title}</td>
          <td style="color:var(--text2);font-size:12px">${c.count}回</td>
          <td><input type="number" id="bgt_${c.title.replace(/[^a-zA-Z0-9]/g,'_')}"
            value="${budget[c.title] || ''}"
            placeholder="0"
            style="width:140px;text-align:right;font-family:var(--mono)"
            oninput="previewBudget()"></td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

function previewBudget() {
  if (!currentDash) return;
  const stats = currentDash.stats;
  const commTitles = Object.values(stats.commStats);
  let total = 0;
  for (const c of commTitles) {
    const id = 'bgt_' + c.title.replace(/[^a-zA-Z0-9]/g,'_');
    const el = document.getElementById(id);
    if (el) total += Number(el.value) || 0;
  }
  // Update total KPI live
  const kpis = document.getElementById('budgetKpis');
  if (kpis) {
    const first = kpis.querySelector('.kpi-value');
    if (first) first.textContent = '¥' + total.toLocaleString();
  }
}

function saveBudget() {
  if (!currentDash || !currentProjectId) return;
  const p = projects[currentProjectId];
  const stats = currentDash.stats;
  const commTitles = Object.values(stats.commStats);
  p.budget = {};
  for (const c of commTitles) {
    const id = 'bgt_' + c.title.replace(/[^a-zA-Z0-9]/g,'_');
    const el = document.getElementById(id);
    const val = Number(el?.value) || 0;
    if (val > 0) p.budget[c.title] = val;
  }
  saveProjects();
  renderBudget();
  toast('費用を保存しました');
}

function clearBudget() {
  if (!currentProjectId) return;
  if (!confirm('費用データをすべてクリアしますか？')) return;
  projects[currentProjectId].budget = {};
  saveProjects();
  renderBudget();
  toast('クリアしました');
}

// ─── Exclude List ────────────────────────────────────────────
function openExcludeModal() {
  if (!currentProjectId) return;
  const p = projects[currentProjectId];
  document.getElementById('excludeInput').value = (p.excludeList || ['Netflix Logo ID']).join('
');
  document.getElementById('excludeModal').classList.add('open');
}
function applyExclude() {
  if (!currentProjectId) return;
  const lines = document.getElementById('excludeInput').value.trim().split('
').map(l=>l.trim()).filter(Boolean);
  projects[currentProjectId].excludeList = lines;
  saveProjects();
  closeModal('excludeModal');
  openDashboard(currentProjectId);
  toast(`対象外: ${lines.length}タイトルを設定しました`);
}

// ─── Rename ──────────────────────────────────────────────────
let renamingId = null;
function openRenameModal(id) {
  renamingId = id;
  document.getElementById('renameInput').value = projects[id]?.name || '';
  document.getElementById('renameModal').classList.add('open');
  setTimeout(() => document.getElementById('renameInput').select(), 50);
}
function applyRename() {
  const name = document.getElementById('renameInput').value.trim();
  if (!name || !renamingId) return;
  projects[renamingId].name = name;
  saveProjects();
  renderSidebar();
  renderHome();
  if (currentProjectId === renamingId) {
    document.getElementById('dashTitle').textContent = name;
  }
  closeModal('renameModal');
  toast(`「${name}」に変更しました`);
  renamingId = null;
}

// ─── Migration ───────────────────────────────────────────────
// 既存プロジェクトのcueにdurSecがなければIn/Outから再計算して補完
function migrateProjects() {
  let migrated = false;
  for (const p of Object.values(projects)) {
    if (!p.excludeList) { p.excludeList = ['Netflix Logo ID']; migrated = true; }
    if (!p.episodes) continue;
    for (const ep of Object.values(p.episodes)) {
      if (!ep.cues) continue;
      for (const c of ep.cues) {
        if (c.durSec !== undefined) continue; // 既にある
        // durSecがない古いデータ → inSec/outSecも当然ない → 0で補完
        // （再アップロードしない限り時間データは取れないので0扱い）
        c.inSec = c.inSec || 0;
        c.outSec = c.outSec || 0;
        c.durSec = c.durSec || 0;
        migrated = true;
      }
    }
    // wavSetがArrayで保存されている場合Setに戻す（旧バージョン互換）
    if (Array.isArray(p.wavSet)) p.wavSet = new Set(p.wavSet);
  }
  if (migrated) {
    saveProjects();
    console.log('[NMA] Migrated old project data');
  }
}

// ─── Init ────────────────────────────────────────────────────
loadProjects();
migrateProjects();
renderSidebar();
renderHome();

// ─── DOM Event Wiring ────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function() {
  // Static action buttons
  const actions = {
    'showHome':       () => showView('home'),
    'showCompare':    () => showView('compare'),
    'newProject':     () => openNewProjectModal(),
    'exportCSV':      () => exportCSV(),
    'exportPDF':      () => exportPDF(),
    'openUpload':     () => openUploadModal(),
    'openRename':     () => openRenameModal(currentProjectId),
    'openExclude':    () => openExcludeModal(),
    'closeNewProject':() => closeModal('newProjectModal'),
    'closeUpload':    () => closeModal('uploadModal'),
    'closeRename':    () => closeModal('renameModal'),
    'closeExclude':   () => closeModal('excludeModal'),
    'createProject':  () => createProject(),
    'applyUpdate':    () => applyUpdate(),
    'applyRename':    () => applyRename(),
    'applyExclude':   () => applyExclude(),
    'saveBudget':     () => saveBudget(),
    'clearBudget':    () => clearBudget(),
    'clickNewXlsx':   () => document.getElementById('newXlsxInput').click(),
    'clickUpdateXlsx':() => document.getElementById('updateXlsxInput').click(),
  };
  // Wire up each button directly
  Object.keys(actions).forEach(action => {
    document.querySelectorAll('[data-action="' + action + '"]').forEach(el => {
      el.addEventListener('click', function(e) {
        e.stopPropagation();
        actions[action]();
      });
    });
  });

  // File input change handlers
  document.getElementById('newXlsxInput').addEventListener('change', e => loadNewXlsx(e.target.files[0]));
  document.getElementById('updateXlsxInput').addEventListener('change', e => loadUpdateXlsx(e.target.files[0]));
  document.getElementById('xlsxIn') && document.getElementById('xlsxIn').addEventListener('change', e => loadXlsx(e.target.files[0]));

  // Rename enter key
  document.getElementById('renameInput').addEventListener('keydown', e => { if(e.key==='Enter') applyRename(); });

  // WAV text input
  const wavText = document.getElementById('wavText');
  if (wavText) wavText.addEventListener('input', onWavInput);
});

