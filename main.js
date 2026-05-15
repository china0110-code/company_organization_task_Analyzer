/* ============================================================
   AI活用業務 優先順位化ツール — main.js
   ============================================================ */

'use strict';

/* ── マスタデータ ─────────────────────────────────────────── */
const Q1_OPTS = [
  { val: 4, lb: '要約・加工',        desc: '議事録・報告書から重要点を抽出し定型フォーマットに落とす' },
  { val: 4, lb: '文章生成',          desc: 'メール回答・報告書・提案書の骨子を毎回作成する' },
  { val: 4, lb: '翻訳・多言語対応',  desc: '文書翻訳や文脈調整' },
  { val: 3, lb: '内容チェック・判定',desc: '申請書が社内ルールや契約に合致しているかOK/NG判断' },
  { val: 3, lb: 'データ検索・特定',  desc: '膨大なフォルダや過去チャットから知見・ファイルを探す' },
  { val: 3, lb: '属人的な相談対応',  desc: '定型的な質問に経験・記憶を頼りに回答する' },
  { val: 1, lb: '情報収集・突き合わせ', desc: '複数ファイル・システムを見比べ差分・整合性を確認' },
  { val: 1, lb: '単純転記・打ち込み',desc: 'システム間のデータ移し替え作業' },
];

const Q2_OPTS = [
  { val: 3, lb: 'テキスト文書',      desc: 'メール・チャット履歴・議事録' },
  { val: 3, lb: '非構造化ファイル',  desc: 'PDF規程集・マニュアル・契約書・仕様書' },
  { val: 3, lb: '音声データ',        desc: '会議録音・コールセンター通話記録' },
  { val: 2, lb: '画像・スキャン文書',desc: '現場写真・図面・スキャン帳票' },
  { val: 2, lb: '暗黙知',            desc: '自分の経験・勘・ベテラン社員へのヒアリング' },
  { val: 1, lb: '構造化データ',      desc: 'DB・Excel・CSV' },
  { val: 1, lb: '社外情報',          desc: 'ニュース・官公庁統計・競合サイト' },
];

const PSI_DEF = {
  p: {
    label: 'P — 拘束時間', min: 0, max: 200, unit: 'h/月', step: 5,
    bands: [
      { v: 10,  l: '10h以下',   lv: 1, c: '#97C459' },
      { v: 40,  l: '11〜40h',   lv: 2, c: '#EF9F27' },
      { v: 80,  l: '41〜80h',   lv: 3, c: '#EF9F27' },
      { v: 160, l: '81〜160h',  lv: 4, c: '#E24B4A' },
      { v: 200, l: '161h以上',  lv: 5, c: '#A32D2D' },
    ],
  },
  s: {
    label: 'S — 属人性（教育日数）', min: 0, max: 30, unit: '日', step: 1,
    bands: [
      { v: 1,  l: '1日未満',  lv: 1, c: '#97C459' },
      { v: 3,  l: '1〜3日',   lv: 2, c: '#EF9F27' },
      { v: 10, l: '4〜10日',  lv: 3, c: '#EF9F27' },
      { v: 20, l: '11〜20日', lv: 4, c: '#E24B4A' },
      { v: 30, l: '21日以上', lv: 5, c: '#A32D2D' },
    ],
  },
  i: {
    label: 'I — 心理的負荷（ミス時工数）', min: 0, max: 50, unit: 'h', step: 1,
    bands: [
      { v: 1,  l: '1h未満',   lv: 1, c: '#97C459' },
      { v: 3,  l: '1〜3h',    lv: 2, c: '#EF9F27' },
      { v: 15, l: '4〜15h',   lv: 3, c: '#EF9F27' },
      { v: 40, l: '16〜40h',  lv: 4, c: '#E24B4A' },
      { v: 50, l: '41h以上',  lv: 5, c: '#A32D2D' },
    ],
  },
};

/* Excel列定義（ヘッダー行に使う名称） */
const EXCEL_COLS = {
  name:   '業務名',
  dept:   '部門',
  desc:   '主な作業内容',
  // Q1
  q1_1:  'Q1_要約・加工',
  q1_2:  'Q1_文章生成',
  q1_3:  'Q1_翻訳・多言語対応',
  q1_4:  'Q1_内容チェック・判定',
  q1_5:  'Q1_データ検索・特定',
  q1_6:  'Q1_属人的な相談対応',
  q1_7:  'Q1_情報収集・突き合わせ',
  q1_8:  'Q1_単純転記・打ち込み',
  // Q2
  q2_1:  'Q2_テキスト文書',
  q2_2:  'Q2_非構造化ファイル',
  q2_3:  'Q2_音声データ',
  q2_4:  'Q2_画像・スキャン文書',
  q2_5:  'Q2_暗黙知',
  q2_6:  'Q2_構造化データ',
  q2_7:  'Q2_社外情報',
  // PSI
  p_raw: 'P_拘束時間(h/月)',
  s_raw: 'S_属人性(教育日数)',
  i_raw: 'I_心理的負荷(ミス時h)',
};

/* ── 状態 ─────────────────────────────────────────────────── */
let bizs = [];
let qChart = null;

/* ── ユーティリティ ──────────────────────────────────────── */
function getLv(key, raw) {
  for (const b of PSI_DEF[key].bands) { if (raw <= b.v) return b.lv; }
  return 5;
}
function getColor(key, raw) {
  for (const b of PSI_DEF[key].bands) { if (raw <= b.v) return b.c; }
  return PSI_DEF[key].bands.at(-1).c;
}
function getRangeLabel(key, raw) {
  for (const b of PSI_DEF[key].bands) { if (raw <= b.v) return b.l; }
  return PSI_DEF[key].bands.at(-1).l;
}
function calcFit(q1mx, q2sm) {
  const sc = q1mx * q2sm;
  return sc >= 12 ? 'high' : sc >= 8 ? 'mid' : 'low';
}
function fitLabel(f) { return f === 'high' ? '高適合' : f === 'mid' ? '中適合' : '低適合'; }
function fitBadge(f) { return f === 'high' ? 'bh' : f === 'mid' ? 'bm' : 'bl'; }
function getQ(p, si) {
  if (p >= 3 && si < 9)  return 'qw';
  if (p >= 3 && si >= 9) return 'st';
  if (p < 3  && si < 9)  return 'pt';
  return 'lp';
}
function qLabel(q) { return { qw: 'Quick Win', st: 'Strategic', pt: 'Potential', lp: 'Low Priority' }[q]; }
function qColor(q) { return { qw: '#639922', st: '#378ADD', pt: '#EF9F27', lp: '#888780' }[q]; }
function esc(s) {
  return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function aiHint(b, q) {
  const map = {
    qw: {
      '要約・加工':      'ChatGPT / Copilot で議事録・レポート要約を即導入。プロンプトテンプレートを整備し1〜2週間でPoC開始可能。',
      '文章生成':        '生成AIでメール下書き・提案書骨子を自動生成。社内文体をFew-shotプロンプトに組み込み品質を担保。',
      '翻訳・多言語対応':'DeepL API または GPT-4o で翻訳ワークフローを自動化。専門用語辞書を整備し高精度を維持。',
      '内容チェック・判定': 'RAGで社内規程を検索し、生成AIがOK/NG判定の根拠を提示するPoC設計が有効。',
      'データ検索・特定': '社内ドキュメントをベクトルDB化しRAGで即時検索できる環境を構築。Microsoft Copilot for M365なら既存環境と親和性が高い。',
      default:           '標準生成AIツールで即導入可能。まずは1〜2週間のPoCを実施し成功体験を積み上げることを推奨。',
    },
    st: { default: 'RAGアーキテクチャや社内データとの連携が必要。専任チームと予算を確保し3〜6ヶ月のPoC計画を策定することを推奨。' },
    pt: { default: '現時点では技術・データの成熟度が課題。6〜12ヶ月後に再評価するタイミングを設定し、データ蓄積を先行して進める。' },
    lp: { default: 'AI化よりも業務プロセス見直し（BPR）や既存システム改修で対応することを推奨。' },
  };
  const qs = map[q] || {};
  for (const lb of (b.q1labels || [])) { if (qs[lb]) return qs[lb]; }
  return qs.default || '';
}

/* ── ステップ遷移 ────────────────────────────────────────── */
function gS(n) {
  if (n === 1 && bizs.filter(b => b.fit !== 'low').length === 0) {
    alert('Step 2に進むには高適合または中適合の業務を1件以上登録してください。');
    return;
  }
  document.querySelectorAll('.sec').forEach((s, i) => s.classList.toggle('vis', i === n));
  document.querySelectorAll('.ps').forEach((s, i) => {
    s.classList.toggle('active', i === n);
    s.classList.toggle('done', i < n);
  });
  if (n === 1) renderPSI();
  if (n === 2) renderQ();
}

/* ── 入力タブ切替 ────────────────────────────────────────── */
function switchInputTab(t) {
  document.getElementById('pane-manual').style.display = t === 'manual' ? '' : 'none';
  document.getElementById('pane-excel').style.display  = t === 'excel'  ? '' : 'none';
  document.getElementById('tab-manual').classList.toggle('act', t === 'manual');
  document.getElementById('tab-excel').classList.toggle('act', t === 'excel');
}

/* ── チェックボックス (手入力) ───────────────────────────── */
function tc(item) {
  const cb = item.querySelector('input[type="checkbox"]');
  cb.checked = !cb.checked;
  item.classList.toggle('chk', cb.checked);
  updSPV();
}

function updSPV() {
  const q1 = [...document.querySelectorAll('#q1g input:checked')];
  const q2 = [...document.querySelectorAll('#q2g input:checked')];
  const el  = document.getElementById('score-preview');
  if (!q1.length || !q2.length) { el.style.display = 'none'; return; }
  const mx = Math.max(...q1.map(c => +c.value));
  const sm = q2.reduce((s, c) => s + +c.value, 0);
  const sc = mx * sm;
  el.style.display = 'block';
  const nEl = el.querySelector('.sp-num');
  const lEl = el.querySelector('.sp-lbl');
  nEl.textContent = sc + '点';
  if      (sc >= 12) { nEl.style.color = '#3B6D11'; lEl.innerHTML = '<span style="color:#3B6D11;font-weight:600">高適合</span> — 最優先でPoC対象'; }
  else if (sc >= 8)  { nEl.style.color = '#BA7517'; lEl.innerHTML = '<span style="color:#BA7517;font-weight:600">中適合</span> — AI化可能。既存ITとの組み合わせを検討'; }
  else               { nEl.style.color = '#A32D2D'; lEl.innerHTML = '<span style="color:#A32D2D;font-weight:600">低適合</span> — AI化の優先度は低い'; }
}

/* ── 業務追加 (手入力) ───────────────────────────────────── */
function addB() {
  const name = document.getElementById('bn').value.trim();
  const dept = document.getElementById('bd').value.trim();
  const desc = document.getElementById('bx').value.trim();
  const q1c  = [...document.querySelectorAll('#q1g input:checked')];
  const q2c  = [...document.querySelectorAll('#q2g input:checked')];
  if (!name)     { alert('業務名を入力してください'); return; }
  if (!q1c.length) { alert('Q1を1つ以上選択してください'); return; }
  if (!q2c.length) { alert('Q2を1つ以上選択してください'); return; }
  const q1mx    = Math.max(...q1c.map(c => +c.value));
  const q2sm    = q2c.reduce((s, c) => s + +c.value, 0);
  const q1labels = q1c.map(c => c.dataset.lb);
  const q2labels = q2c.map(c => c.dataset.lb);
  const score   = q1mx * q2sm;
  bizs.push({ id: Date.now(), name, dept, desc, q1mx, q2sm, q1labels, q2labels, score, fit: calcFit(q1mx, q2sm), pRaw: 40, sRaw: 5, iRaw: 5 });
  renderList();
  resetForm();
}

function resetForm() {
  ['bn', 'bd', 'bx'].forEach(id => document.getElementById(id).value = '');
  document.querySelectorAll('#q1g input, #q2g input').forEach(c => c.checked = false);
  document.querySelectorAll('#q1g .cbi, #q2g .cbi').forEach(i => i.classList.remove('chk'));
  document.getElementById('score-preview').style.display = 'none';
}

/* ── 業務削除 ────────────────────────────────────────────── */
function removeB(id) {
  if (!confirm('この業務を削除しますか？')) return;
  bizs = bizs.filter(b => b.id !== id);
  renderList();
}

/* ── アコーディオン ──────────────────────────────────────── */
function toggleDetail(id) {
  const detail = document.getElementById('bd-' + id);
  const icon   = document.getElementById('bi-icon-' + id);
  const isOpen = detail.classList.contains('open');
  detail.classList.toggle('open', !isOpen);
  if (icon) icon.classList.toggle('open', !isOpen);
}

/* ── 編集フォーム ────────────────────────────────────────── */
function startEdit(id) {
  const b = bizs.find(x => x.id === id); if (!b) return;
  const container = document.getElementById('edit-form-' + id);
  if (container.style.display === 'block') { container.style.display = 'none'; return; }

  const q1Html = Q1_OPTS.map(o => {
    const chk = b.q1labels.includes(o.lb) ? 'checked' : '';
    return `<div class="cbi ${chk ? 'chk' : ''}" onclick="tcEdit(this,${id})">
      <input type="checkbox" value="${o.val}" data-lb="${o.lb}" ${chk}>
      <label>${o.lb}<span class="desc">${o.desc}</span></label></div>`;
  }).join('');
  const q2Html = Q2_OPTS.map(o => {
    const chk = b.q2labels.includes(o.lb) ? 'checked' : '';
    return `<div class="cbi ${chk ? 'chk' : ''}" onclick="tcEdit(this,${id})">
      <input type="checkbox" value="${o.val}" data-lb="${o.lb}" ${chk}>
      <label>${o.lb}<span class="desc">${o.desc}</span></label></div>`;
  }).join('');

  container.innerHTML = `
    <div class="edit-form">
      <div class="row">
        <div><label class="fl">業務名<span class="req">*</span></label><input type="text" id="en-${id}" value="${esc(b.name)}"></div>
        <div><label class="fl">担当部門</label><input type="text" id="ed-${id}" value="${esc(b.dept)}"></div>
      </div>
      <div class="mt10"><label class="fl">主な作業内容</label><textarea id="ex-${id}">${esc(b.desc)}</textarea></div>
      <div class="divider"></div>
      <label class="fl">Q1. 面倒・嫌だと感じる作業<span class="req">*</span></label>
      <div class="cbg" id="eq1-${id}">${q1Html}</div>
      <div class="mt10"><label class="fl">Q2. インプットに使うデータ<span class="req">*</span></label>
      <div class="cbg" id="eq2-${id}">${q2Html}</div></div>
      <div id="edit-spv-${id}" class="score-preview" style="display:none">
        <div class="sp-row"><span style="font-size:13px;color:var(--text-secondary)">スコアプレビュー</span>
        <span class="sp-num"></span></div><div class="sp-lbl"></div>
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px">
        <button class="btn btn-s btn-sm" onclick="cancelEdit(${id})">キャンセル</button>
        <button class="btn btn-p btn-sm" onclick="saveEdit(${id})">✓ 保存</button>
      </div>
    </div>`;
  container.style.display = 'block';
  updEditSPV(id);
}

function tcEdit(item, id) {
  const cb = item.querySelector('input[type="checkbox"]');
  cb.checked = !cb.checked;
  item.classList.toggle('chk', cb.checked);
  updEditSPV(id);
}

function updEditSPV(id) {
  const q1 = [...document.querySelectorAll(`#eq1-${id} input:checked`)];
  const q2 = [...document.querySelectorAll(`#eq2-${id} input:checked`)];
  const el  = document.getElementById(`edit-spv-${id}`);
  if (!q1.length || !q2.length) { el.style.display = 'none'; return; }
  const mx = Math.max(...q1.map(c => +c.value));
  const sm = q2.reduce((s, c) => s + +c.value, 0);
  const sc = mx * sm;
  el.style.display = 'block';
  const nEl = el.querySelector('.sp-num');
  const lEl = el.querySelector('.sp-lbl');
  nEl.textContent = sc + '点';
  if      (sc >= 12) { nEl.style.color = '#3B6D11'; lEl.innerHTML = '<span style="color:#3B6D11;font-weight:600">高適合</span>'; }
  else if (sc >= 8)  { nEl.style.color = '#BA7517'; lEl.innerHTML = '<span style="color:#BA7517;font-weight:600">中適合</span>'; }
  else               { nEl.style.color = '#A32D2D'; lEl.innerHTML = '<span style="color:#A32D2D;font-weight:600">低適合</span>'; }
}

function cancelEdit(id) {
  const c = document.getElementById('edit-form-' + id);
  if (c) c.style.display = 'none';
}

function saveEdit(id) {
  const b = bizs.find(x => x.id === id); if (!b) return;
  const name = document.getElementById('en-' + id).value.trim();
  const dept = document.getElementById('ed-' + id).value.trim();
  const desc = document.getElementById('ex-' + id).value.trim();
  const q1c  = [...document.querySelectorAll(`#eq1-${id} input:checked`)];
  const q2c  = [...document.querySelectorAll(`#eq2-${id} input:checked`)];
  if (!name)     { alert('業務名を入力してください'); return; }
  if (!q1c.length) { alert('Q1を1つ以上選択してください'); return; }
  if (!q2c.length) { alert('Q2を1つ以上選択してください'); return; }
  const q1mx = Math.max(...q1c.map(c => +c.value));
  const q2sm = q2c.reduce((s, c) => s + +c.value, 0);
  Object.assign(b, {
    name, dept, desc,
    q1labels: q1c.map(c => c.dataset.lb),
    q2labels: q2c.map(c => c.dataset.lb),
    q1mx, q2sm, score: q1mx * q2sm, fit: calcFit(q1mx, q2sm),
  });
  renderList();
  setTimeout(() => {
    const el = document.getElementById('bd-' + id); if (el) el.classList.add('open');
    const ic = document.getElementById('bi-icon-' + id); if (ic) ic.classList.add('open');
  }, 50);
}

/* ── 業務リスト描画 ──────────────────────────────────────── */
function renderList() {
  const el = document.getElementById('blist');
  document.getElementById('bcnt').textContent = bizs.length + '件';
  // クリアボタン表示制御
  const btnClear = document.getElementById('btn-clear-all');
  if (btnClear) btnClear.style.display = bizs.length ? 'inline-flex' : 'none';
  if (!bizs.length) {
    el.innerHTML = '<li class="empty">まだ業務が登録されていません</li>';
    return;
  }
  el.innerHTML = bizs.map(b => {
    const q1tags = b.q1labels.map(l => `<span class="tag">${l}</span>`).join('');
    const q2tags = b.q2labels.map(l => `<span class="tag">${l}</span>`).join('');
    return `
    <li class="bi">
      <div class="bi-header" onclick="toggleDetail(${b.id})">
        <div class="bi-left">
          <div class="bi-name">${esc(b.name)}</div>
          <div class="bi-meta">${esc(b.dept) || '—'} ・ 適合スコア ${b.score}点</div>
        </div>
        <div class="bi-actions">
          <span class="sbadge ${fitBadge(b.fit)}">${fitLabel(b.fit)}</span>
          <button class="btn btn-e" onclick="event.stopPropagation();startEdit(${b.id})">✎ 編集</button>
          <button class="btn btn-d" onclick="event.stopPropagation();removeB(${b.id})">🗑</button>
          <span class="bi-chevron" id="bi-icon-${b.id}">▼</span>
        </div>
      </div>
      <div class="bi-detail" id="bd-${b.id}">
        <div class="detail-grid">
          <div class="detail-field"><div class="lbl">業務名</div><div class="val">${esc(b.name)}</div></div>
          <div class="detail-field"><div class="lbl">担当部門</div><div class="val">${esc(b.dept) || '—'}</div></div>
        </div>
        ${b.desc ? `<div class="detail-field" style="margin-bottom:10px"><div class="lbl">主な作業内容</div><div class="val" style="font-weight:400">${esc(b.desc)}</div></div>` : ''}
        <div class="detail-grid">
          <div class="detail-field"><div class="lbl">Q1 選択項目</div><div class="tag-list">${q1tags || '—'}</div></div>
          <div class="detail-field"><div class="lbl">Q2 選択項目</div><div class="tag-list">${q2tags || '—'}</div></div>
        </div>
        <div class="detail-grid" style="margin-top:8px">
          <div class="detail-field"><div class="lbl">Q1最大スコア</div><div class="val">${b.q1mx}点</div></div>
          <div class="detail-field"><div class="lbl">Q2合計スコア</div><div class="val">${b.q2sm}点</div></div>
        </div>
        <div id="edit-form-${b.id}" style="display:none;margin-top:10px"></div>
      </div>
    </li>`;
  }).join('');
}

/* ── PSI 描画 ────────────────────────────────────────────── */
function renderPSI() {
  const eligible = bizs.filter(b => b.fit !== 'low');
  const con = document.getElementById('psi-con');
  if (!eligible.length) {
    con.innerHTML = '<div style="padding:10px 14px;border-radius:8px;background:#E6F1FB;color:#0C447C;border:0.5px solid #B5D4F4;font-size:13px">高適合・中適合の業務がありません。Step 1に戻って業務を追加してください。</div>';
    return;
  }
  con.innerHTML = eligible.map(b => `
    <div class="card psi-card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem">
        <div>
          <div style="font-size:15px;font-weight:600;color:var(--text-primary)">${esc(b.name)}</div>
          <div style="font-size:12px;color:var(--text-secondary)">${esc(b.dept) || ''} ・ 適合スコア: ${b.score}点</div>
        </div>
        <span class="sbadge ${fitBadge(b.fit)}">${fitLabel(b.fit)}</span>
      </div>
      ${mkGauge(b, 'p')}
      ${mkGauge(b, 's')}
      ${mkGauge(b, 'i')}
    </div>`).join('');
}

function mkGauge(b, key) {
  const d   = PSI_DEF[key];
  const raw = b[key + 'Raw'];
  const pct = Math.round((raw - d.min) / (d.max - d.min) * 100);
  const lv  = getLv(key, raw);
  const col = getColor(key, raw);
  const rl  = getRangeLabel(key, raw);
  const dots = [1, 2, 3, 4, 5].map(i =>
    `<div class="lv-dot" style="background:${i <= lv ? col : '#d3d1c7'}"></div>`).join('');
  return `
    <div class="psi-block">
      <div class="psi-hd">
        <span class="psi-label">${d.label}</span>
        <span class="psi-range-lbl" id="${key}-rl-${b.id}">${rl}</span>
      </div>
      <div class="gauge-wrap">
        <div class="gauge-fill" id="${key}-gf-${b.id}" style="width:${pct}%;background:${col}"></div>
        <span class="gauge-raw-val" id="${key}-rv-${b.id}">${raw}${d.unit}</span>
        <span class="gauge-lv-val" id="${key}-lv-${b.id}">Lv.${lv}</span>
      </div>
      <input type="range" min="${d.min}" max="${d.max}" step="${d.step}" value="${raw}"
             oninput="updPSI(${b.id},'${key}',this.value)">
      <div class="lv-dots" id="${key}-dots-${b.id}">${dots}</div>
      <div class="psi-ticks">
        <span>${d.bands[0].l}</span><span>${d.bands[2].l}</span><span>${d.bands[4].l}</span>
      </div>
    </div>`;
}

function updPSI(id, key, val) {
  const b = bizs.find(x => x.id === id); if (!b) return;
  const raw = parseInt(val, 10);
  b[key + 'Raw'] = raw;
  const d   = PSI_DEF[key];
  const pct = Math.round((raw - d.min) / (d.max - d.min) * 100);
  const lv  = getLv(key, raw);
  const col = getColor(key, raw);
  const rl  = getRangeLabel(key, raw);
  const gf  = document.getElementById(`${key}-gf-${id}`);
  if (gf) { gf.style.width = pct + '%'; gf.style.background = col; }
  const rv = document.getElementById(`${key}-rv-${id}`); if (rv) rv.textContent = raw + d.unit;
  const lv2 = document.getElementById(`${key}-lv-${id}`); if (lv2) lv2.textContent = 'Lv.' + lv;
  const rl2 = document.getElementById(`${key}-rl-${id}`); if (rl2) rl2.textContent = rl;
  const dots = document.getElementById(`${key}-dots-${id}`);
  if (dots) dots.innerHTML = [1, 2, 3, 4, 5].map(i =>
    `<div class="lv-dot" style="background:${i <= lv ? col : '#d3d1c7'}"></div>`).join('');
  if (qChart) renderQ();
}

/* ── 4象限 描画 ──────────────────────────────────────────── */
function renderQ() {
  const eligible = bizs.filter(b => b.fit !== 'low');
  const ctx = document.getElementById('qc').getContext('2d');
  if (qChart) { qChart.destroy(); qChart = null; }

  const datasets = eligible.map(b => {
    const si  = getLv('s', b.sRaw) * getLv('i', b.iRaw);
    const pLv = getLv('p', b.pRaw);
    const q   = getQ(pLv, si);
    return {
      label: b.name,
      data: [{ x: si, y: pLv }],
      backgroundColor: qColor(q) + 'CC',
      borderColor: qColor(q),
      borderWidth: 2,
      pointRadius: 10,
      pointHoverRadius: 13,
    };
  });

  qChart = new Chart(ctx, {
    type: 'bubble',
    data: { datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 40, right: 40, bottom: 30, left: 20 } },
      scales: {
        x: { min: 0, max: 28, title: { display: true, text: '実現難易度 (S×I)', color: '#888780', font: { size: 12 } }, grid: { color: 'rgba(136,135,128,0.15)' }, ticks: { color: '#888780', font: { size: 11 } } },
        y: { min: 0, max: 6,  title: { display: true, text: 'ビジネスインパクト (P)', color: '#888780', font: { size: 12 } }, grid: { color: 'rgba(136,135,128,0.15)' }, ticks: { stepSize: 1, color: '#888780', font: { size: 11 } } },
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const b = eligible[ctx.datasetIndex];
              const si = getLv('s', b.sRaw) * getLv('i', b.iRaw);
              const q  = getQ(getLv('p', b.pRaw), si);
              return [b.name, `分類: ${qLabel(q)}`, `P:Lv.${getLv('p', b.pRaw)} S:Lv.${getLv('s', b.sRaw)} I:Lv.${getLv('i', b.iRaw)}`];
            },
          },
        },
      },
    },
    plugins: [{
      id: 'qbg',
      beforeDraw(chart) {
        const { ctx, chartArea: { left, right, top, bottom, width, height } } = chart;
        const xMid = left + width * (9 / 28);
        const yMid = top + height * (1 - 3 / 6);
        ctx.save();
        [
          { x: left, y: top,  w: xMid - left, h: yMid - top,    col: 'rgba(99,153,34,0.07)',    lbl: 'Quick Win',   tc: '#3B6D11' },
          { x: xMid, y: top,  w: right - xMid,h: yMid - top,    col: 'rgba(55,138,221,0.07)',   lbl: 'Strategic',   tc: '#0C447C' },
          { x: left, y: yMid, w: xMid - left, h: bottom - yMid, col: 'rgba(239,159,39,0.07)',   lbl: 'Potential',   tc: '#854F0B' },
          { x: xMid, y: yMid, w: right - xMid,h: bottom - yMid, col: 'rgba(136,135,128,0.07)', lbl: 'Low Priority', tc: '#5F5E5A' },
        ].forEach(z => {
          ctx.fillStyle = z.col;
          ctx.fillRect(z.x, z.y, z.w, z.h);
          ctx.fillStyle = z.tc + 'CC';
          ctx.font = '600 12px sans-serif';
          ctx.textAlign = 'left';
          ctx.fillText(z.lbl, z.x + 8, z.y + 16);
        });
        ctx.strokeStyle = 'rgba(136,135,128,0.5)'; ctx.lineWidth = 1.5; ctx.setLineDash([6, 4]);
        ctx.beginPath(); ctx.moveTo(xMid, top); ctx.lineTo(xMid, bottom); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(left, yMid); ctx.lineTo(right, yMid); ctx.stroke();
        ctx.setLineDash([]);
        eligible.forEach((b, i) => {
          const pt = chart.getDatasetMeta(i).data[0]; if (!pt) return;
          const si = getLv('s', b.sRaw) * getLv('i', b.iRaw);
          ctx.fillStyle = qColor(getQ(getLv('p', b.pRaw), si));
          ctx.font = '12px sans-serif'; ctx.textAlign = 'center';
          const name = b.name.length > 8 ? b.name.slice(0, 8) + '…' : b.name;
          ctx.fillText(name, pt.x, pt.y - 14);
        });
        ctx.restore();
      },
    }],
  });
  renderTable(eligible);
}

function renderTable(eligible) {
  const tbody   = document.getElementById('rtbody');
  const actions = { qw: '標準ツールで即導入・PoC開始', st: '予算確保・専任チーム編成', pt: '技術進化・データ蓄積後に再評価', lp: '現行維持・AI化見送り' };
  const tagCls  = { qw: 'tqw', st: 'tst', pt: 'tpt', lp: 'tlp' };
  tbody.innerHTML = eligible.map(b => {
    const si = getLv('s', b.sRaw) * getLv('i', b.iRaw);
    const q  = getQ(getLv('p', b.pRaw), si);
    return `<tr>
      <td style="font-weight:600">${esc(b.name)}</td>
      <td>${esc(b.dept) || '—'}</td>
      <td><span class="sbadge ${fitBadge(b.fit)}">${fitLabel(b.fit)}</span></td>
      <td style="text-align:center">Lv.${getLv('p', b.pRaw)}</td>
      <td style="text-align:center">Lv.${getLv('s', b.sRaw)}</td>
      <td style="text-align:center">Lv.${getLv('i', b.iRaw)}</td>
      <td><span class="qtag ${tagCls[q]}">${qLabel(q)}</span>
          <div style="font-size:11px;color:var(--text-secondary);margin-top:3px">${actions[q]}</div></td>
      <td><div class="ai-hint">${aiHint(b, q)}</div></td>
    </tr>`;
  }).join('');
}

/* ── タブ切替 (Step3) ────────────────────────────────────── */
function swTab(t) {
  document.getElementById('tb-c').classList.toggle('act', t === 'chart');
  document.getElementById('tb-t').classList.toggle('act', t === 'table');
  document.getElementById('pane-chart').style.display = t === 'chart' ? '' : 'none';
  document.getElementById('pane-table').style.display = t === 'table' ? '' : 'none';
}

/* ── Excelアップロード ───────────────────────────────────── */
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.remove('drag');
  const f = e.dataTransfer.files[0]; if (f) onFileSelected(f);
}

function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    const r = new FileReader();
    r.onload = e => parseExcelCSV(e.target.result);
    r.readAsText(file, 'UTF-8');
  } else if (ext === 'xlsx' || ext === 'xls') {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      parseExcelCSV(XLSX.utils.sheet_to_csv(ws));
    };
    r.readAsBinaryString(file);
  } else {
    showIR('err', '対応していないファイル形式です（.xlsx / .csv のみ）');
  }
}

/**
 * Excel/CSVパース
 * ヘッダー行の列名でフィールドを特定する。
 * Q1/Q2列は "○" "1" "true" "yes" の場合にチェック済みとみなす。
 * PSI値が指定されている場合はそのまま数値として使う。
 * PSI値が全列欠落している行は pRaw=40, sRaw=5, iRaw=5 をデフォルト値とする。
 */
function parseExcelCSV(text) {
  const lines = text.trim().split('\n').filter(l => l.trim());
  if (lines.length < 2) { showIR('err', 'データが見つかりません'); return; }

  // ヘッダー解析
  const headers = splitCSVLine(lines[0]);

  // 列インデックス取得ヘルパー
  const colIdx = name => headers.findIndex(h => h.trim() === name);

  // 列インデックスマップ
  const idx = {};
  for (const [k, v] of Object.entries(EXCEL_COLS)) { idx[k] = colIdx(v); }

  if (idx.name < 0) { showIR('err', `「${EXCEL_COLS.name}」列が見つかりません。テンプレートのヘッダーを確認してください。`); return; }

  let added = 0, skipped = 0;

  for (let i = 1; i < lines.length; i++) {
    const cols = splitCSVLine(lines[i]);
    const get  = key => (idx[key] >= 0 ? (cols[idx[key]] || '').trim() : '');
    const chk  = key => { const v = get(key); return v === '○' || v === '1' || v.toLowerCase() === 'true' || v.toLowerCase() === 'yes'; };

    const name = get('name');
    if (!name) { skipped++; continue; }

    const dept = get('dept');
    const desc = get('desc');

    // Q1
    const q1Checked = Q1_OPTS.map((o, j) => chk(`q1_${j + 1}`) ? o : null).filter(Boolean);
    // Q2
    const q2Checked = Q2_OPTS.map((o, j) => chk(`q2_${j + 1}`) ? o : null).filter(Boolean);

    if (!q1Checked.length || !q2Checked.length) { skipped++; continue; }

    const q1mx    = Math.max(...q1Checked.map(o => o.val));
    const q2sm    = q2Checked.reduce((s, o) => s + o.val, 0);
    const q1labels = q1Checked.map(o => o.lb);
    const q2labels = q2Checked.map(o => o.lb);
    const score   = q1mx * q2sm;

    // PSI
    const pRaw = idx.p_raw >= 0 && get('p_raw') !== '' ? Math.min(200, Math.max(0, parseInt(get('p_raw'), 10) || 40)) : 40;
    const sRaw = idx.s_raw >= 0 && get('s_raw') !== '' ? Math.min(30,  Math.max(0, parseInt(get('s_raw'), 10) || 5))  : 5;
    const iRaw = idx.i_raw >= 0 && get('i_raw') !== '' ? Math.min(50,  Math.max(0, parseInt(get('i_raw'), 10) || 5))  : 5;

    bizs.push({ id: Date.now() + i, name, dept, desc, q1mx, q2sm, q1labels, q2labels, score, fit: calcFit(q1mx, q2sm), pRaw, sRaw, iRaw });
    added++;
  }

  renderList();

  if (added > 0) {
    const hasPSI = idx.p_raw >= 0;
    const msg = `${added}件の業務を登録しました${skipped > 0 ? `（${skipped}件スキップ）` : ''}。`
      + (hasPSI ? ' PSI値も取り込み済みです。Step 3の結果をご確認ください。' : ' Step 2でPSI評価を入力してください。');
    showIR('ok', msg);
    // PSI値が含まれていれば Step 3 まで一気にジャンプ
    if (hasPSI && bizs.some(b => b.fit !== 'low')) {
      setTimeout(() => jumpToResult(), 300);
    }
  } else {
    showIR('err', '登録できる業務がありませんでした。Q1・Q2の少なくとも1列に「○」を入力しているか確認してください。');
  }
}

function jumpToResult() {
  document.querySelectorAll('.sec').forEach((s, i) => s.classList.toggle('vis', i === 2));
  document.querySelectorAll('.ps').forEach((s, i) => {
    s.classList.toggle('active', i === 2);
    s.classList.toggle('done', i < 2);
  });
  renderQ();
  // ページ先頭へスクロール
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/** CSV行をカンマ分割（ダブルクォート対応） */
function splitCSVLine(line) {
  const result = []; let cur = ''; let inQ = false;
  for (let ci = 0; ci < line.length; ci++) {
    const ch = line[ci];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { result.push(cur); cur = ''; }
    else { cur += ch; }
  }
  result.push(cur);
  return result.map(s => s.replace(/^"|"$/g, ''));
}

function showIR(type, msg) {
  const el = document.getElementById('import-result');
  el.style.display = 'block';
  el.className = `import-result ${type}`;
  el.textContent = msg;
}

/* ── テンプレートダウンロード ────────────────────────────── */
function downloadTemplate() {
  const cols = Object.values(EXCEL_COLS);
  // サンプル行 1: 議事録要約（Q1=要約・加工 ○、Q2=音声データ・テキスト文書 ○、PSI値あり）
  const sample1 = [
    '議事録要約', '総務部', '会議録音を文字起こしし要点・アクションを抽出してメール配布する',
    '○','','','','','','','',   // Q1 (8列)
    '○','○','○','','','','',   // Q2 (7列)
    80, 5, 3,                   // PSI
  ];
  // サンプル行 2: メール下書き（Q1=文章生成 ○、Q2=テキスト文書 ○）
  const sample2 = [
    'メール下書き生成', '営業部', '問い合わせ対応メールを毎回ゼロから作成している',
    '','○','','','','','','',
    '○','','','','','','',
    40, 3, 2,
  ];
  const rows = [cols, sample1, sample2];
  const csv  = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = 'ai_usecase_template.csv';
  a.click();
}

/* ══════════════════════════════════════════════════════════
   PDF エクスポート  — 全ページ A4 横向き統一
   p.1 : 4象限マップ（縮小）＋ 象限説明カード4枚（横並び・下段）
   p.2〜: 業務一覧（Q1作業種別・主な作業内容）
   p.末〜: 優先順位 結果一覧（AI活用の可能性付き）
   ══════════════════════════════════════════════════════════ */
async function exportPDF() {
  const eligible = bizs.filter(b => b.fit !== 'low');
  if (!eligible.length) { alert('出力対象の業務がありません。'); return; }

  const overlay = document.getElementById('pdf-overlay');
  overlay.classList.add('show');
  await waitForjsPDF();

  const { jsPDF } = window.jspdf;
  const now   = new Date();
  const stamp = `${now.getFullYear()}/${String(now.getMonth()+1).padStart(2,'0')}/${String(now.getDate()).padStart(2,'0')}`;
  const TITLE = 'AI活用業務 優先順位化レポート';

  const PW = 297, PH = 210, PM = 14;
  const CW = PW - PM * 2;   // 269 mm
  const HDR_H = 20;
  const FTR_H = 16;
  const LH = 3.6;

  const Q_COLOR = {
    qw:{ text:[59,109,17],  bg:[234,243,222], label:'Quick Win'    },
    st:{ text:[24,95,165],  bg:[230,241,251], label:'Strategic'    },
    pt:{ text:[133,79,11],  bg:[250,238,218], label:'Potential'    },
    lp:{ text:[95,94,90],   bg:[241,239,232], label:'Low Priority' },
  };
  const FIT_COLOR = {
    high:{ text:[59,109,17],  bg:[234,243,222], label:'高適合' },
    mid :{ text:[133,79,11],  bg:[250,238,218], label:'中適合' },
    low :{ text:[163,45,45],  bg:[252,235,235], label:'低適合' },
  };

  // テーブルヘッダー（doc・cols・y を引数で受け取る）
  const drawTblHdr = (doc, cols, y) => {
    doc.setFillColor(241,239,232);
    doc.rect(PM, y, CW, 7, 'F');
    doc.setDrawColor(200,198,190); doc.setLineWidth(0.3);
    doc.rect(PM, y, CW, 7, 'S');
    setJP(doc, 6.5, 'bold'); doc.setTextColor(95,94,90);
    let cx = PM;
    cols.forEach(c => {
      const tx = c.al === 'c' ? cx + c.w/2 : cx + 1.5;
      doc.text(c.l, tx, y + 4.8, { align: c.al === 'c' ? 'center' : 'left', maxWidth: c.w - 1 });
      cx += c.w;
    });
    return y + 7;
  };

  // 行背景＋下線（doc を引数で受け取る）
  const drawRowBg = (doc, y, h, ri) => {
    doc.setFillColor(ri%2===0?255:249, ri%2===0?255:248, ri%2===0?255:244);
    doc.rect(PM, y, CW, h, 'F');
    doc.setDrawColor(220,218,210); doc.setLineWidth(0.22);
    doc.line(PM, y+h, PM+CW, y+h);
  };

  // バッジ描画ヘルパー
  const drawBadge = (doc, x, midY, w, colorObj, labelStr) => {
    const bh = 4.8;
    doc.setFillColor(...colorObj.bg); doc.setDrawColor(...colorObj.text); doc.setLineWidth(0.28);
    doc.roundedRect(x+1.5, midY-bh/2, w-3, bh, 1.2, 1.2, 'FD');
    setJP(doc, 5.8, 'bold'); doc.setTextColor(...colorObj.text);
    doc.text(labelStr, x+w/2, midY+0.6, { align:'center' });
  };

  try {
    const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
    embedJPFont(doc);
    let pageNum = 1;

    /* ─── p.1: 4象限マップ ＋ 下段に象限説明カード4枚横並び ─── */
    drawPageHeader(doc, TITLE, '4象限マップ', stamp, PW, PM);

    // カード高さを先に確定してチャート領域を逆算
    const CARD_H   = 38;
    const CARD_GAP = 4;
    const CARD_Y   = PH - FTR_H - CARD_H;
    const chartTop = HDR_H + 2;
    const chartAreaH = CARD_Y - chartTop - CARD_GAP;

    // チャート画像（アスペクト比保持・中央揃え）
    const canvas  = document.getElementById('qc');
    const imgData = canvas.toDataURL('image/png', 1.0);
    const aspect  = canvas.width / canvas.height;
    let chartW = Math.min(CW, chartAreaH * aspect);
    let chartH = chartW / aspect;
    if (chartH > chartAreaH) { chartH = chartAreaH; chartW = chartH * aspect; }
    const chartX = PM + (CW - chartW) / 2;
    doc.addImage(imgData, 'PNG', chartX, chartTop, chartW, chartH);

    // ① 象限説明カード4枚を横並び（Webのq-legend-gridと同じ構成）
    const Q_DEFS = [
      { label:'Quick Win（左上）',
        color:[59,109,17],  bg:[234,243,222], border:[99,153,34],
        desc:'標準ツールで即導入し成功体験を積み上げる領域。議事録要約・翻訳・メール下書きなど、AIが即時に価値を発揮できる業務。' },
      { label:'Strategic（右上）',
        color:[24,95,165],  bg:[230,241,251], border:[55,138,221],
        desc:'予算・リソースを集中投下し独自の知能資産を築く領域。技術承継・特許解析・高度な見積自動化など複雑な業務。' },
      { label:'Potential（左下）',
        color:[133,79,11],  bg:[250,238,218], border:[239,159,39],
        desc:'将来の種となる領域。技術進化やデータ蓄積を待ってから判断する。高度な需要予測・未整備データの分析など。' },
      { label:'Low Priority（右下）',
        color:[95,94,90],   bg:[241,239,232], border:[180,178,169],
        desc:'AI化しない領域。無理にAIを当てはめず現行通りとする。対面接遇・身体的業務・低頻度タスクなど。' },
    ];
    const cardW = (CW - CARD_GAP * 3) / 4;
    let cardX = PM;
    Q_DEFS.forEach(qd => {
      doc.setFillColor(...qd.bg); doc.setDrawColor(...qd.border); doc.setLineWidth(0.4);
      doc.roundedRect(cardX, CARD_Y, cardW, CARD_H, 2, 2, 'FD');
      doc.setFillColor(...qd.border);
      doc.roundedRect(cardX, CARD_Y, 2.5, CARD_H, 1, 1, 'F');
      setJP(doc, 7, 'bold'); doc.setTextColor(...qd.color);
      doc.text(qd.label, cardX + 5, CARD_Y + 6);
      // ① setJP の後に splitTextToSize を呼ぶ（フォント設定→計算の順序）
      setJP(doc, 6, 'normal'); doc.setTextColor(50,50,48);
      const dLines = doc.splitTextToSize(qd.desc, cardW - 7);
      doc.text(dLines, cardX + 5, CARD_Y + 11.5);
      cardX += cardW + CARD_GAP;
    });

    drawFooter(doc, PW, PH, PM, stamp, pageNum++);

    /* ─── p.2〜: 業務一覧 ─────────────────────────────────── */
    doc.addPage([PW, PH], 'landscape');
    drawPageHeader(doc, TITLE, '業務一覧', stamp, PW, PM);

    // ② 列幅の合計が CW(269) に厳密に合うよう変数で管理
    const W = { nm:35, dp:20, ft:16, p:8, s:8, i:8, qt:26, q1:46 };
    W.ds = CW - W.nm - W.dp - W.ft - W.p - W.s - W.i - W.qt - W.q1;
    const D_COL = [
      { l:'業務名',       w:W.nm, al:'l' },
      { l:'部門',         w:W.dp, al:'l' },
      { l:'適合性',       w:W.ft, al:'c' },
      { l:'P',            w:W.p,  al:'c' },
      { l:'S',            w:W.s,  al:'c' },
      { l:'I',            w:W.i,  al:'c' },
      { l:'分類',         w:W.qt, al:'c' },
      { l:'Q1 作業種別',  w:W.q1, al:'l' },
      { l:'主な作業内容', w:W.ds, al:'l' },
    ];

    let dy = drawTblHdr(doc, D_COL, HDR_H + 2);

    eligible.forEach((b, ri) => {
      const si  = getLv('s', b.sRaw) * getLv('i', b.iRaw);
      const q   = getQ(getLv('p', b.pRaw), si);
      const qc  = Q_COLOR[q], fc = FIT_COLOR[b.fit];

      // ② フォントを先に設定してから splitTextToSize（崩れの根本原因を修正）
      setJP(doc, 6.5, 'normal');
      const q1Text    = (b.q1labels||[]).join('、') || '—';
      const q1Lines   = doc.splitTextToSize(q1Text,  W.q1 - 2);
      const descText  = b.desc || '—';
      const descLines = doc.splitTextToSize(descText, W.ds  - 4);
      const rowH = Math.max(9, Math.max(q1Lines.length, descLines.length) * LH + 4);

      if (dy + rowH > PH - FTR_H) {
        drawFooter(doc, PW, PH, PM, stamp, pageNum++);
        doc.addPage([PW, PH], 'landscape');
        drawPageHeader(doc, TITLE, '業務一覧（続き）', stamp, PW, PM);
        dy = drawTblHdr(doc, D_COL, HDR_H + 2);
      }

      drawRowBg(doc, dy, rowH, ri);
      const MID = dy + rowH / 2;
      const TOP = dy + 4.5;
      let cx = PM;

      // 業務名
      setJP(doc, 6.8, 'bold'); doc.setTextColor(26,26,24);
      const nmL = doc.splitTextToSize(b.name, W.nm - 2);
      doc.text(nmL, cx+1.5, TOP); cx += W.nm;

      // 部門
      setJP(doc, 6.5, 'normal'); doc.setTextColor(90,89,85);
      doc.text((b.dept||'—').slice(0,10), cx+1.5, MID); cx += W.dp;

      // 適合性バッジ
      drawBadge(doc, cx, MID, W.ft, fc, fc.label); cx += W.ft;

      // P / S / I
      doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(26,26,24);
      [W.p, W.s, W.i].forEach((w, ki) => {
        const k = ['p','s','i'][ki];
        doc.text('Lv.'+getLv(k,b[k+'Raw']), cx+w/2, MID, {align:'center'});
        cx += w;
      });

      // 分類バッジ
      drawBadge(doc, cx, MID, W.qt, qc, qc.label); cx += W.qt;

      // Q1 作業種別
      setJP(doc, 6.5, 'normal'); doc.setTextColor(60,59,55);
      doc.text(q1Lines, cx+1.5, TOP); cx += W.q1;

      // 主な作業内容（左アクセントライン付き）
      doc.setDrawColor(...qc.text); doc.setLineWidth(0.6);
      doc.line(cx+1, dy+2, cx+1, dy+rowH-2); doc.setLineWidth(0.22);
      setJP(doc, 6.5, 'normal'); doc.setTextColor(70,69,65);
      doc.text(descLines, cx+3, TOP);

      dy += rowH;
    });

    drawFooter(doc, PW, PH, PM, stamp, pageNum++);

    /* ─── p.末〜: 優先順位 結果一覧 ─────────────────────────── */
    doc.addPage([PW, PH], 'landscape');
    drawPageHeader(doc, TITLE, '優先順位 結果一覧', stamp, PW, PM);

    const RW = { nm:34, dp:20, ft:15, p:8, s:8, i:8, qt:26 };
    RW.ht = CW - RW.nm - RW.dp - RW.ft - RW.p - RW.s - RW.i - RW.qt;
    const R_COL = [
      { l:'業務名',         w:RW.nm, al:'l' },
      { l:'部門',           w:RW.dp, al:'l' },
      { l:'適合性',         w:RW.ft, al:'c' },
      { l:'P',              w:RW.p,  al:'c' },
      { l:'S',              w:RW.s,  al:'c' },
      { l:'I',              w:RW.i,  al:'c' },
      { l:'分類',           w:RW.qt, al:'c' },
      { l:'AI活用の可能性', w:RW.ht, al:'l' },
    ];

    let ry = drawTblHdr(doc, R_COL, HDR_H + 2);

    eligible.forEach((b, ri) => {
      const si   = getLv('s', b.sRaw) * getLv('i', b.iRaw);
      const q    = getQ(getLv('p', b.pRaw), si);
      const qc   = Q_COLOR[q], fc = FIT_COLOR[b.fit];
      const hint = aiHint(b, q);

      setJP(doc, 6.5, 'normal');
      const hintLines = doc.splitTextToSize(hint, RW.ht - 4);
      const rowH = Math.max(9, hintLines.length * LH + 4);

      if (ry + rowH > PH - FTR_H) {
        drawFooter(doc, PW, PH, PM, stamp, pageNum++);
        doc.addPage([PW, PH], 'landscape');
        drawPageHeader(doc, TITLE, '優先順位 結果一覧（続き）', stamp, PW, PM);
        ry = drawTblHdr(doc, R_COL, HDR_H + 2);
      }

      drawRowBg(doc, ry, rowH, ri);
      const MID = ry + rowH / 2;
      const TOP = ry + 4.5;
      let cx = PM;

      setJP(doc, 6.8, 'bold'); doc.setTextColor(26,26,24);
      const nmL = doc.splitTextToSize(b.name, RW.nm - 2);
      doc.text(nmL, cx+1.5, TOP); cx += RW.nm;

      setJP(doc, 6.5, 'normal'); doc.setTextColor(90,89,85);
      doc.text((b.dept||'—').slice(0,10), cx+1.5, MID); cx += RW.dp;

      drawBadge(doc, cx, MID, RW.ft, fc, fc.label); cx += RW.ft;

      doc.setFont('helvetica','normal'); doc.setFontSize(7); doc.setTextColor(26,26,24);
      [RW.p, RW.s, RW.i].forEach((w, ki) => {
        const k = ['p','s','i'][ki];
        doc.text('Lv.'+getLv(k,b[k+'Raw']), cx+w/2, MID, {align:'center'});
        cx += w;
      });

      drawBadge(doc, cx, MID, RW.qt, qc, qc.label); cx += RW.qt;

      doc.setDrawColor(...qc.text); doc.setLineWidth(0.6);
      doc.line(cx+1, ry+2, cx+1, ry+rowH-2); doc.setLineWidth(0.22);
      setJP(doc, 6.5, 'normal'); doc.setTextColor(70,69,65);
      doc.text(hintLines, cx+3, TOP);

      ry += rowH;
    });

    drawFooter(doc, PW, PH, PM, stamp, pageNum);
    doc.save('AI_ユースケース_優先順位化レポート.pdf');

  } finally {
    overlay.classList.remove('show');
  }
}

function waitForjsPDF() {
  return new Promise(resolve => {
    const check = () => (window.jspdf ? resolve() : setTimeout(check, 100));
    check();
  });
}

/* ── 日本語フォント埋め込みヘルパー ─────────────────────── */
function embedJPFont(doc) {
  if (typeof IPAG_B64 === 'undefined') return;
  doc.addFileToVFS('ipag.ttf', IPAG_B64);
  doc.addFont('ipag.ttf', 'IPAGothic', 'normal');
  doc.addFont('ipag.ttf', 'IPAGothic', 'bold');
}
function setJP(doc, size, style) {
  doc.setFont(typeof IPAG_B64 !== 'undefined' ? 'IPAGothic' : 'helvetica',
              style === 'bold' ? 'bold' : 'normal');
  doc.setFontSize(size);
}

/* ── 共通ヘッダー描画 ────────────────────────────────────── */
function drawPageHeader(doc, mainTitle, subTitle, stamp, PW, MARGIN) {
  doc.setFillColor(24, 95, 165);
  doc.rect(0, 0, PW, 20, 'F');
  setJP(doc, 12, 'bold'); doc.setTextColor(255,255,255);
  doc.text(mainTitle, MARGIN, 9);
  setJP(doc, 9, 'normal');
  doc.text(subTitle, MARGIN, 16);
  doc.text(stamp, PW - MARGIN, 13, { align: 'right' });
}

/* ── フッター描画 ─────────────────────────────────────────── */
function drawFooter(doc, PW, PH, MARGIN, stamp, pageNum) {
  doc.setDrawColor(200,198,190); doc.setLineWidth(0.3);
  doc.line(MARGIN, PH - 12, PW - MARGIN, PH - 12);
  setJP(doc, 7.5, 'normal'); doc.setTextColor(140,138,130);
  doc.text('AI活用業務 優先順位化ツール', MARGIN, PH - 7);
  doc.text(`${pageNum} ページ  |  ${stamp}`, PW - MARGIN, PH - 7, { align:'right' });
}

/* ══════════════════════════════════════════════════════════
   ④ 全件クリア
   ══════════════════════════════════════════════════════════ */
function clearAll() {
  if (!bizs.length) return;
  if (!confirm(`登録済みの ${bizs.length} 件の業務をすべて削除します。よろしいですか？`)) return;
  bizs = [];
  if (qChart) { qChart.destroy(); qChart = null; }
  renderList();
}

/* ══════════════════════════════════════════════════════════
   ⑤ Excel ファイル選択 → 確認 → 実行 の 2 ステップフロー
   ══════════════════════════════════════════════════════════ */
let pendingFile = null; // 選択済みだが未実行のファイル

function onFileSelected(file) {
  if (!file) return;
  pendingFile = file;

  // ドロップゾーン表示を更新
  document.getElementById('uz-icon').textContent = '📋';
  document.getElementById('uz-title').textContent = '別のファイルを選択する場合はここをクリック';
  document.getElementById('uz-sub').textContent   = '';

  // 実行ボタン行を表示
  const row = document.getElementById('import-exec-row');
  row.style.display = 'flex';
  document.getElementById('import-filename').textContent = `選択中: ${file.name}`;

  // 前回の結果をリセット
  const ir = document.getElementById('import-result');
  ir.style.display = 'none'; ir.className = 'import-result';
}

function clearFileSelection() {
  pendingFile = null;
  document.getElementById('xlfile').value = '';
  document.getElementById('uz-icon').textContent  = '📤';
  document.getElementById('uz-title').textContent = 'Excelファイルをドロップ、またはクリックして選択';
  document.getElementById('uz-sub').textContent   = '.xlsx / .csv 形式対応';
  document.getElementById('import-exec-row').style.display = 'none';
  const ir = document.getElementById('import-result');
  ir.style.display = 'none'; ir.className = 'import-result';
}

function executeImport() {
  if (!pendingFile) return;
  handleFile(pendingFile);
  clearFileSelection();
}
