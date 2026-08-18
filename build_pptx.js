// 小兒創傷 permissive hypotension 教學簡報 → .pptx
// node build_pptx.js   (pptxgenjs 在 /Users/mac/Downloads/claude working/node_modules)
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';               // 13.333 x 7.5
pres.author = '林自強';
pres.title = '小兒創傷 Permissive Hypotension：實證與 EMT 教學建議';

const F = 'Arial';
const C = {
  bg:'0C0E13', panel:'161B24', panel2:'1E2530', line:'2E3745',
  ink:'E9EDF3', muted:'A6B0BE', dim:'838C99',
  red:'E2404C', amber:'F0A836', teal:'3EC6B4', white:'FFFFFF',
  gold:'FFC000', darkred:'C00000', green:'5FBF6A',
};
const M = 0.55;                             // 左右邊界
const W = 13.333 - M * 2;                   // 12.233 可用寬

// ---------- helpers -------------------------------------------------------
function base(kicker, title, sub) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addText(kicker, { x:M, y:0.30, w:W, h:0.28, fontSize:11, bold:true, color:C.red,
                      charSpacing:2.2, fontFace:F, margin:0 });
  s.addText(title, { x:M, y:0.60, w:W, h:0.60, fontSize:30, bold:true, color:C.ink,
                     fontFace:F, margin:0 });
  if (sub) s.addText(sub, { x:M, y:1.21, w:W, h:0.32, fontSize:13, color:C.muted,
                            fontFace:F, margin:0 });
  return s;
}
// 卡片：圓角矩形 + 標題 + 條列
function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x:o.x, y:o.y, w:o.w, h:o.h, fill:{ color:o.fill || C.panel },
    line:{ color:o.accent || C.line, width:o.accent ? 1.5 : 1 }, rectRadius:0.10,
  });
  let ty = o.y + 0.22;
  if (o.label) {
    s.addText(o.label, { x:o.x+0.28, y:ty, w:o.w-0.56, h:0.24, fontSize:10.5, bold:true,
                         color:C.dim, charSpacing:1.6, fontFace:F, margin:0 });
    ty += 0.28;
  }
  if (o.title) {
    s.addText(o.title, { x:o.x+0.28, y:ty, w:o.w-0.56, h:0.34, fontSize:17, bold:true,
                         color:o.titleColor || C.ink, fontFace:F, margin:0 });
    ty += 0.44;
  }
  if (o.body) {
    s.addText(o.body, { x:o.x+0.28, y:ty, w:o.w-0.56, h:o.y+o.h-ty-0.18, fontSize:o.size||13.5,
                        color:C.ink, fontFace:F, margin:0, lineSpacingMultiple:1.25, valign:'top' });
  }
  if (o.bullets) {
    s.addText(o.bullets.map((t,i) => ({ text:t, options:{ bullet:{ code:'2022' },
      breakLine:i !== o.bullets.length-1 } })),
      { x:o.x+0.28, y:ty, w:o.w-0.56, h:o.y+o.h-ty-0.18, fontSize:o.size||13.5, color:C.ink,
        fontFace:F, margin:0, paraSpaceAfter:5, valign:'top' });
  }
}
// 數字大字報
function stat(s, x, y, w, n, k, col) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h:1.28, fill:{ color:C.panel2 },
    line:{ color:C.line, width:1 }, rectRadius:0.09 });
  s.addText(n, { x, y:y+0.14, w, h:0.60, fontSize:30, bold:true, color:col||C.red,
                 align:'center', fontFace:F, margin:0 });
  s.addText(k, { x:x+0.10, y:y+0.74, w:w-0.20, h:0.46, fontSize:10.5, color:C.muted,
                 align:'center', fontFace:F, margin:0, lineSpacingMultiple:1.1 });
}
// 房子招牌：紅底金框 🔥 金句（PHPLS 系列自訂樣式）
// emoji 自己一個 run 且不指定 typeface，否則 Mac PowerPoint 開檔會 crash
function punch(s, text, y, h) {
  s.addText(
    [ { text:'🔥 ', options:{ bold:true, fontSize:15, color:C.white } },
      { text, options:{ bold:true, fontSize:15, color:C.white, fontFace:F } } ],
    { shape:pres.ShapeType.roundRect, rectRadius:0.10, x:M, y, w:W, h:h||0.52,
      fill:{ color:C.darkred }, line:{ color:C.gold, width:1.5 },
      align:'left', valign:'middle', margin:[6,10,6,10] });
}
// 比喻框
function analogy(s, emo, text, y, h) {
  s.addShape(pres.ShapeType.roundRect, { x:M, y, w:W, h:h||0.72, fill:{ color:'241D0F' },
    line:{ color:C.amber, width:1 }, rectRadius:0.08 });
  s.addText(emo, { x:M+0.16, y:y+0.10, w:0.5, h:(h||0.72)-0.2, fontSize:20, align:'center',
                   valign:'middle', margin:0 });
  s.addText([ { text:'比喻　', options:{ bold:true, color:C.amber, fontFace:F } },
              { text, options:{ color:C.ink, fontFace:F } } ],
    { x:M+0.72, y:y+0.06, w:W-0.92, h:(h||0.72)-0.12, fontSize:12.5, valign:'middle',
      margin:0, lineSpacingMultiple:1.2 });
}
// 表格
function table(s, rows, o) {
  const head = rows[0].map(t => ({ text:t, options:{ bold:true, color:C.muted, fontSize:11,
    fill:{ color:C.panel2 } } }));
  const body = rows.slice(1).map((r, i) => r.map(cell => {
    const isKey = o.keyRows && o.keyRows.includes(i);
    if (typeof cell === 'object') return Object.assign({}, cell,
      { options:Object.assign({ fontSize:o.size||11.5, color:C.ink,
        fill:{ color:isKey ? '2A171B' : C.panel } }, cell.options||{}) });
    return { text:cell, options:{ fontSize:o.size||11.5, color:C.ink,
      fill:{ color:isKey ? '2A171B' : C.panel } } };
  }));
  s.addTable([head, ...body], {
    x:M, y:o.y, w:W, colW:o.colW, fontFace:F, border:{ type:'solid', color:C.line, pt:1 },
    valign:'middle', margin:o.cellMargin || 5, rowH:o.rowH,
  });
}
// 分段條（代償曲線 / 輸液量帶）
function band(s, y, segs, ticks) {
  let x = M;
  segs.forEach(sg => {
    const w = W * sg.f;
    s.addShape(pres.ShapeType.rect, { x, y, w, h:0.78, fill:{ color:sg.c },
      line:{ color:C.line, width:1 } });
    s.addText(sg.t, { x:x+0.04, y, w:w-0.08, h:0.78, fontSize:11.5, bold:true, color:C.ink,
                      align:'center', valign:'middle', fontFace:F, margin:0,
                      lineSpacingMultiple:1.05 });
    x += w;
  });
  ticks.forEach(tk => {
    const cx = M + W * tk.f;
    s.addText(tk.t, { x:cx - 0.85, y:y+0.82, w:1.7, h:0.26, fontSize:10.5, color:C.dim,
                      align:tk.a || 'center', fontFace:F, margin:0 });
  });
}

/* =========================================================================
   1 — 封面
   ====================================================================== */
{
  const s = pres.addSlide();
  s.background = { color:'0C0E13' };
  s.addShape(pres.ShapeType.roundRect, { x:-1.2, y:-1.6, w:8.4, h:4.6, rectRadius:0.5,
    fill:{ color:C.darkred, transparency:78 }, line:{ color:C.bg, width:0 } });
  s.addText('PHPLS · 小兒到院前急救', { x:M, y:1.95, w:W, h:0.32, fontSize:12.5, bold:true,
    color:C.red, charSpacing:2.4, fontFace:F, margin:0 });
  s.addText('小兒創傷該不該「刻意壓低血壓」？', { x:M, y:2.35, w:W, h:0.72, fontSize:36,
    bold:true, color:C.white, fontFace:F, margin:0 });
  s.addText('Permissive Hypotension 的實證與 EMT 教學建議', { x:M, y:3.05, w:W, h:0.62,
    fontSize:30, bold:true, color:C.red, fontFace:F, margin:0 });
  s.addText('從成人 RCT、兒童共識、進行中的英國 PRESSURE 試驗，\n回到現場能記得住、做得出來的四個動作',
    { x:M, y:3.90, w:W, h:0.80, fontSize:15, color:C.muted, fontFace:F, margin:0,
      lineSpacingMultiple:1.3 });
  s.addShape(pres.ShapeType.line, { x:M, y:5.05, w:4.2, h:0, line:{ color:C.line, width:1 } });
  s.addText('EMT／EMT-P 繼續教育　·　2026-08-17　·　林自強', { x:M, y:5.18, w:W, h:0.34,
    fontSize:13, color:C.dim, fontFace:F, margin:0 });
  s.addNotes('這堂課只有一個核心：把「少給水」和「讓他低血壓」分開。');
}

/* 2 — 開場情境 */
{
  const s = base('CASE', '先想一個場景', '你在現場，接下來 10 分鐘做什麼？');
  card(s, { x:M, y:1.75, w:W, h:1.60, accent:C.red,
    body:'6 歲男童，機車後座乘客，被撞飛約 5 公尺。\n意識躁動、臉色蒼白、四肢冰冷。HR 160、RR 34、BP 96/70、CRT 4 秒，腹部脹且壓痛，無明顯外出血。',
    size:15 });
  card(s, { x:M, y:3.55, w:5.95, h:1.30, label:'選項 A', size:13,
    body:'血壓還有 96，看起來還好 —— 先不打點滴，快點送。' });
  card(s, { x:M+6.28, y:3.55, w:5.95, h:1.30, label:'選項 B', size:13,
    body:'休克了 —— 兩條大針，生理食鹽水全開，一路灌到血壓上來。' });
  punch(s, '兩個都錯。今天要回答的是：中間那條路長什麼樣子。', 5.10);
  s.addNotes('讓學員先舉手選 A 或 B，再進入結論頁。多數人會選 B。');
}

/* 3 — 一句話結論 */
{
  const s = base('BOTTOM LINE', '一句話結論', '如果今天只帶走一句話，就是這句');
  s.addShape(pres.ShapeType.rect, { x:M, y:1.80, w:0.06, h:1.05, fill:{ color:C.red },
    line:{ color:C.red, width:0 } });
  s.addText('兒童創傷不做 permissive hypotension（刻意維持低血壓），\n但要做限制性輸液（不要灌爆）。',
    { x:M+0.28, y:1.80, w:W-0.28, h:1.05, fontSize:21, bold:true, color:C.white, fontFace:F,
      margin:0, lineSpacingMultiple:1.35 });
  card(s, { x:M, y:3.10, w:5.95, h:1.62, accent:C.red, title:'✕ 不是這個', titleColor:C.red,
    size:13, body:'把血壓「放著低」當成治療目標，等到手術室才拉回來。\n這是成人的策略，兒童沒有證據支持。' });
  card(s, { x:M+6.28, y:3.10, w:5.95, h:1.62, accent:C.teal, title:'✓ 是這個', titleColor:C.teal,
    size:13, body:'小量、多次、看反應。目標不是某個數字，而是意識、脈搏、皮膚回溫有沒有改善。' });
  punch(s, '學員最常把「少給水」聽成「讓他低血壓」——這是本堂最大的陷阱。', 5.00);
  s.addNotes('這兩件事在成人教材裡常被寫在同一段，所以要主動拆開講。');
}

/* 4 — 名詞界定 */
{
  const s = base('DEFINITIONS', '兩個名詞，不要混在一起',
    'Permissive hypotension vs. restrictive fluid strategy');
  card(s, { x:M, y:1.75, w:5.95, h:2.35, accent:C.amber, label:'PERMISSIVE HYPOTENSION',
    title:'容許性低血壓', size:13,
    body:'目標是血壓：止血前刻意把 SBP 維持在 50–70 或 80–90 mmHg，不去把它拉到正常。\n理由：血壓拉高會沖掉剛形成的血栓、加速失血。' });
  card(s, { x:M+6.28, y:1.75, w:5.95, h:2.35, accent:C.teal, label:'RESTRICTIVE FLUID STRATEGY',
    title:'限制性輸液', size:13,
    body:'目標是輸液量：不給不必要的晶體液，每次少量後重新評估，需要時儘早改給血品。\n理由：大量生理食鹽水本身有害。' });
  analogy(s, '🩹', '傷口剛結了一層薄痂。前者是「別用高壓水槍去沖它」，後者是「別在旁邊一直加水稀釋膠水」。一個管水壓、一個管水量——小孩的水壓不能亂降，水量卻要省著用。',
    4.30, 0.86);
  punch(s, '教學上的講法：我們省的是「水」，不是「血壓」。', 5.35);
  s.addNotes('這一頁是全場的定義基準，後面所有內容都回頭指這裡。');
}

/* 5 — 成人證據 */
{
  const s = base('ADULT EVIDENCE', '成人：permissive hypotension 是有證據的',
    '這是整個概念的來源，也是被誤植到兒童的起點');
  const w4 = (W - 0.36 * 3) / 4;
  stat(s, M,                  1.72, w4, '6',        '篇 RCT 統合分析');
  stat(s, M+(w4+0.36),        1.72, w4, '1,158',    '受試者總數', C.white);
  stat(s, M+(w4+0.36)*2,      1.72, w4, 'OR 0.70',  '死亡（0.53–0.92）', C.teal);
  stat(s, M+(w4+0.36)*3,      1.72, w4, '50–70',    '介入組目標 SBP（mmHg）', C.amber);
  card(s, { x:M, y:3.20, w:5.95, h:1.72, title:'試驗內容', size:12.5,
    body:'介入組目標 SBP 50–70 或 MAP ≥50；對照組 SBP 65–100／MAP ≥65，直到出血控制。輸血量與失血量都較少。\n但各試驗 power 不足、異質性高。（Tran 2018）' });
  card(s, { x:M+6.28, y:3.20, w:5.95, h:1.72, title:'指引怎麼改的', size:12.5, bullets:[
    'ATLS 9th：初始晶體液 2 L → 1 L，更早給血漿與血小板',
    '歐洲創傷指引：止血前 SBP 80–90 mmHg（1C），明文排除腦傷者' ] });
  punch(s, '成人的前提是：有明確出血源、而且很快進手術室——小孩兩個都不一定成立。', 5.10);
  s.addNotes('Tran 2018：6 篇 RCT、n=1,158，OR 0.70（0.53–0.92）。');
}

/* 6 — 兒童證據地圖 */
{
  const s = base('PEDIATRIC EVIDENCE', '兒童：沒有 RCT，只有共識',
    '2023 Pediatric Traumatic Hemorrhagic Shock Consensus Conference');
  const w4 = (W - 0.36 * 3) / 4;
  stat(s, M,             1.72, w4, '21', '條聲明', C.white);
  stat(s, M+(w4+0.36),   1.72, w4, '2',  'clinical recommendation', C.teal);
  stat(s, M+(w4+0.36)*2, 1.72, w4, '14', 'expert consensus', C.amber);
  stat(s, M+(w4+0.36)*3, 1.72, w4, '5',  'good practice statement', C.amber);
  card(s, { x:M, y:3.20, w:5.95, h:1.72, accent:C.red, title:'全文沒有設定的東西', size:13,
    titleColor:C.red, bullets:[ '兒童 permissive hypotension 的血壓目標',
      '兒童專屬的低血壓復甦流程', 'ATLS／PALS／PECARN／EAST 也都沒有' ] });
  card(s, { x:M+6.28, y:3.20, w:5.95, h:1.72, accent:C.teal, title:'它實際涵蓋的', size:13,
    titleColor:C.teal, bullets:[ '血品與輸液復甦', '院前血品使用', '止血輔助劑、止血帶',
      '院前呼吸道與血壓處置' ] });
  punch(s, '作者自己寫：兒童的出血控制與復甦「幾乎沒有高品質證據可循」。', 5.10);
  s.addNotes('Russell 2023, J Trauma Acute Care Surg 94(1S):S2-S10.');
}

/* 7 — 生理 */
{
  const s = base('PHYSIOLOGY', '為什麼小孩不能照抄大人',
    '兒童的血壓是「最後才掉」的生命徵象');
  band(s, 1.78,
    [ { f:0.62, c:'1E4A34', t:'代償期：血壓正常（心跳快、血管收縮、CRT 變長）' },
      { f:0.22, c:'4A3A16', t:'失代償：血壓開始掉' },
      { f:0.16, c:'4A1D22', t:'瀕臨停止' } ],
    [ { f:0.0, t:'失血 0%', a:'left' }, { f:0.62, t:'約 30%' },
      { f:0.84, t:'約 40%' }, { f:1.0, t:'45%+', a:'right' } ]);
  card(s, { x:M, y:3.05, w:5.95, h:1.60, accent:C.red, title:'臨床意義', size:13,
    body:'量到低血壓時代償儲備已幾乎用完。那不是「還可以觀察的中間狀態」，而是接近心跳停止的訊號。' });
  card(s, { x:M+6.28, y:3.05, w:5.95, h:1.60, title:'所以現場要看什麼', size:13, bullets:[
    '意識反應', '橈動脈摸不摸得到、脈搏強不強', 'CRT > 2 秒、皮膚濕冷蒼白' ] });
  analogy(s, '🔋', '小孩的血壓像手機的電量顯示：一路顯示 100%、100%、100%……然後直接跳到關機。你不能等它顯示 5% 才開始找充電器。',
    4.82, 0.78);
  s.addNotes('大人的電量條是慢慢降的，所以大人可以「開到 20% 再說」。');
}

/* 8 — 血壓門檻不可靠 */
{
  const s = base('THRESHOLDS', '而且那個「血壓門檻」本身就不可靠',
    '同一個孩子，換一本教科書就換一個答案');
  table(s, [
    ['來源', '說法', '問題'],
    ['常背的公式', '1–10 歲 SBP < 70 + 2×年齡', '創傷教材沿用，未針對預後驗證'],
    ['Sarganas 2019\n（德國 KiGGS 族群）', '3–9 歲第 5 百分位其實約 82 + 年齡；10–17 歲非線性',
     '比常用公式高，代表公式會低估休克'],
    ['Hagedoorn 2019\n（系統性回顧）', 'PALS 定義 <12 歲尚可；>12 歲會低估低血壓；APLS 定義則高估',
     '不同來源門檻可差 15–30 mmHg'],
  ], { y:1.75, colW:[3.0, 5.4, 3.833], keyRows:[2], rowH:0.62 });
  analogy(s, '⚖️', '三台體重計，站上去差 3 公斤。你不會拿它來決定要不要開刀——你會去看這個人有沒有變瘦。血壓數字只能當警報，不能當方向盤。',
    4.32, 0.78);
  punch(s, '血壓是「紅線」（跌破就是很糟），不是「目標」（滴定到剛好）。', 5.30);
  s.antNotes = null;
  s.addNotes('兩篇分析的共同結論：沒有任何一條公式對預後有良好驗證。');
}

/* 9 — 輸液怎麼給 */
{
  const s = base('FLUIDS', '那輸液到底怎麼給？',
    '2025 PALS 與 ERC 的共同方向：小量、分次、重評');
  card(s, { x:M, y:1.75, w:5.95, h:1.95, accent:C.teal, label:'2025 AHA／AAP PALS', size:12.5,
    body:'等張晶體液以 10 mL/kg 或 20 mL/kg 分次給予，每次給完就重新評估「有沒有反應」與「有沒有容積過載」。\n重點不在選 10 還是 20，在於打完要重評。' });
  card(s, { x:M+6.28, y:1.75, w:5.95, h:1.95, accent:C.amber, label:'ERC 兒童生命支持', size:12.5,
    body:'單次上限壓到 10 mL/kg（取代舊的 20），理由是減少容積過載造成的呼吸衰竭與稀釋性凝血病。\n出血性／創傷性休克：晶體液要限制，儘早給血品。' });
  const bw = (W - 0.9) / 4, ay = 4.00;
  ['評估休克', '10 mL/kg 快速給', '立刻重評\n意識／脈搏／CRT', '有改善→停\n沒改善→再一次']
    .forEach((t, i) => {
      const x = M + i * (bw + 0.30);
      card(s, { x, y:ay, w:bw, h:0.92, size:12, body:t });
      if (i < 3) s.addText('▶', { x:x+bw+0.02, y:ay+0.28, w:0.26, h:0.36, fontSize:14,
        bold:true, color:C.red, align:'center', fontFace:F, margin:0 });
    });
  punch(s, '輸液不是「掛上去就走」，是一次一次給、一次一次看。', 5.20);
  s.addNotes('ERC 把單次上限降到 10 mL/kg，是這幾年最重要的方向改變。');
}

/* 10 — 過量的傷害 */
{
  const s = base('HARM', '灌太多，是真的會死人的',
    '兒童創傷的觀察性資料一致指向同一個方向');
  const w4 = (W - 0.36 * 3) / 4;
  stat(s, M,             1.72, w4, '2.96',  '急診第 1 小時 20–40 mL/kg\n死亡 adjusted OR', C.amber);
  stat(s, M+(w4+0.36),   1.72, w4, '6.26',  '急診第 1 小時 ≥40 mL/kg\n死亡 adjusted OR');
  stat(s, M+(w4+0.36)*2, 1.72, w4, '>60',   'mL/kg／24–48 小時\n存活出院率下降', C.amber);
  stat(s, M+(w4+0.36)*3, 1.72, w4, '5 / 6', '研究顯示\nICU／住院天數延長', C.white);
  card(s, { x:M, y:3.20, w:5.95, h:1.62, accent:C.red, title:'怎麼死的', size:13, bullets:[
    '稀釋性凝血病 —— 凝血因子被沖淡', '低體溫 —— 冷輸液讓凝血酶失效',
    '組織與肺水腫 —— 插管率上升' ] });
  card(s, { x:M+6.28, y:3.20, w:5.95, h:1.62, title:'資料的限制（要誠實說）', size:13,
    body:'全部是觀察性研究。「傷得重的孩子本來就會被灌比較多水」這個混淆無法完全排除，雖然研究已校正創傷嚴重度（ISS）。' });
  punch(s, '灌水不是「多做一點總沒錯」——它有劑量、有毒性、有上限。', 5.00);
  s.addNotes('Mbadiwe 2021, J Surg Res 262:93-100（單中心世代，已校正 ISS）。');
}

/* 11 — 40 mL/kg 那條線 */
{
  const s = base('THE LINE', '現場心裡要有的那條線：40 mL/kg',
    '到了這條線還在惡化，代表他需要的不是水');
  band(s, 1.78,
    [ { f:0.33, c:'1E4A34', t:'0–20 mL/kg\n參考組' },
      { f:0.34, c:'4A3A16', t:'20–40 mL/kg\n死亡 aOR 2.96' },
      { f:0.33, c:'4A1D22', t:'≥40 mL/kg\n死亡 aOR 6.26' } ],
    [ { f:0.0, t:'0', a:'left' }, { f:0.33, t:'20 mL/kg' },
      { f:0.67, t:'40 mL/kg' }, { f:1.0, t:'更多', a:'right' } ]);
  card(s, { x:M, y:3.05, w:5.95, h:1.55, title:'換算成現場的動作', size:13,
    body:'40 mL/kg ≒ 10 mL/kg 打 4 次，或 20 mL/kg 打 2 次（＝台北市 SOP 的上限）。\n20 kg 的孩子 ≒ 800 mL。' });
  card(s, { x:M+6.28, y:3.05, w:5.95, h:1.55, accent:C.red, title:'越線之後該做的', size:13,
    bullets:[ '通報，指定創傷中心', '加壓止血、骨盆固定帶再確認', '保溫', '快送——不要在現場繼續加碼' ],
    size:12 });
  analogy(s, '🚰', '牆壁裡的水管破了，你在客廳一直拖地。拖得再勤，牆裡還在漏。到了 40 mL/kg，該叫的是水電工（開刀房），不是再拿一條抹布。',
    4.78, 0.78);
  s.addNotes('腹腔與骨盆的出血在現場沒有辦法止——這是這一頁的重點。');
}

/* 12 — TBI 數據 */
{
  const s = base('CONTRAINDICATION', '頭部外傷：低血壓 = 二次腦傷',
    'PEGASUS 多中心研究，5 家兒童創傷中心，n=234');
  const w4 = (W - 0.36 * 3) / 4;
  stat(s, M,             1.72, w4, '26%',      '早期照護中曾低血壓\n（60/234）', C.amber);
  stat(s, M+(w4+0.36),   1.72, w4, '23.3%',    '院內死亡（曾低血壓）\nvs 8.6%，p=0.01');
  stat(s, M+(w4+0.36)*2, 1.72, w4, 'aRR 0.46', '30 分鐘內矯正\n死亡（0.24–0.90）', C.teal);
  stat(s, M+(w4+0.36)*3, 1.72, w4, 'aRR 0.54', '30 分鐘內矯正\n不良 GOS（0.39–0.76）', C.teal);
  card(s, { x:M, y:3.20, w:5.95, h:1.62, title:'定義', size:12.5,
    body:'低血壓＝ SBP < 70 + 2×年齡（BTF 定義）；缺氧＝ SpO₂ < 90% 或 PaO₂ < 60。\n「及時處置」＝紀錄到該事件後 30 分鐘內以輸液、血品或升壓劑矯正。' });
  card(s, { x:M+6.28, y:3.20, w:5.95, h:1.62, accent:C.red, title:'另一個佐證', size:12.5,
    body:'Luerssen 系列（成人 6,908／兒童 1,906）：低血壓對死亡的不良影響，在兒童比成人更大。' });
  punch(s, '腦傷的孩子，低血壓不是「還撐得住」，是正在製造第二次傷害。', 5.00);
  s.addNotes('Kannan 2018, Pediatr Emerg Care 34(5):325-329；經 BTF 2023 指引採用。');
}

/* 13 — 禁忌整理 */
{
  const s = base('RULE', '哪些情況「一分鐘都不能低」', '這一頁請學員抄下來');
  card(s, { x:M, y:1.75, w:5.95, h:2.05, accent:C.red, title:'絕對不容許低血壓',
    titleColor:C.red, size:13, bullets:[ '頭部外傷（懷疑就算）', '脊髓損傷',
    '成人指引在合併重度 TBI（GCS ≤8）時目標是 MAP ≥80 mmHg，而不是 permissive' ] });
  card(s, { x:M+6.28, y:1.75, w:5.95, h:2.05, accent:C.teal, title:'現場怎麼判斷',
    titleColor:C.teal, size:13, bullets:[ '意識改變、瞳孔不等大', '頭皮血腫、耳鼻出血',
    '受傷機轉：高處墜落、車禍拋飛', '不確定的時候，當成有' ] });
  card(s, { x:M, y:3.95, w:W, h:1.00, accent:C.amber, size:13,
    body:'BTF 指引特別提醒：小兒 TBI 即使血壓「正常」，只要有灌流不良徵象（CRT 長、脈搏弱、意識差），仍然可以給輸液。這是唯一會叫你「積極一點」的情境。' });
  analogy(s, '🧠', '受傷的腦子像已經跳電一半的房子：電壓再掉一點點，剩下還亮著的燈也會跟著熄。省電（低血壓）在這裡不叫節約，叫加速停電。',
    5.10, 0.74);
  s.addNotes('這一頁是全場唯一的「硬規則」，其餘都是判斷。');
}

/* 14 — PRESSURE 設計 */
{
  const s = base('ONGOING TRIAL', '英國 PRESSURE 試驗',
    '兒童血壓目標領域的第一個大型 RCT — 正在進行中');
  table(s, [
    ['項目', '內容'],
    ['全名', 'PRotocolised Evaluation of permiSSive blood pressure targets versus USual caRE'],
    ['註冊／資助', 'ISRCTN 20609635；NIHR HTA（128895），ICNARC 主持'],
    ['規模', '英國 18 家 PICU，目標 1,900 人；2021 年 11 月第一位個案'],
    ['介入組', '維持 MAP > 該年齡第 5 百分位（permissive target band）'],
    ['對照組', 'MAP 目標由臨床團隊自行決定（usual care）'],
    ['主要結果', '30 天內死亡 ＋ 侵入性呼吸器天數的複合排序結果；含經濟評估'],
    ['納入', '矯正妊娠 37 週–<16 歲，入 PICU、使用呼吸器、預期血管活性藥 ≥6 小時'],
    ['排除', '急性腦損傷、心臟術後、心肌病變、依賴動脈導管循環、肺高壓、惡性高血壓、瀕死'],
  ], { y:1.72, colW:[2.4, 9.833], keyRows:[3, 7], size:11.5, rowH:0.42 });
  punch(s, '截至 2026 年 8 月，PRESSURE 尚未有結果發表。', 5.55);
  s.addNotes('PCCM 2024 protocol；50% 收案時做期中分析（Peto–Haybittle）。');
}

/* 15 — PRESSURE MAP 帶 */
{
  const s = base('ONGOING TRIAL', '它用的 MAP 目標帶', '可以直接當「地板值」的教學參考');
  s.addTable([
    [ { text:'年齡', options:{ bold:true, color:C.muted, fontSize:12, fill:{ color:C.panel2 } } },
      { text:'MAP 目標帶（mmHg）', options:{ bold:true, color:C.muted, fontSize:12,
        fill:{ color:C.panel2 }, align:'right' } } ],
    ...[['37 週–6 個月','40–43'],['>6 個月–<1 歲','40–45'],['1–3 歲','45–50'],
        ['4–9 歲','50–55'],['≥10 歲','55–60']].map(r => [
      { text:r[0], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel } } },
      { text:r[1], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel }, align:'right',
        bold:true } } ]),
  ], { x:M, y:1.75, w:5.4, colW:[3.2, 2.2], fontFace:F,
       border:{ type:'solid', color:C.line, pt:1 }, valign:'middle', margin:6, rowH:0.44 });
  card(s, { x:M+5.85, y:1.75, w:6.38, h:1.90, accent:C.amber,
    title:'注意這裡的 “permissive” 是什麼意思', size:13,
    body:'它不是「允許血壓掉下去」，而是「不必用升壓藥把 MAP 拱得比第 5 百分位更高」。\n第 5 百分位是地板，不是目標。' });
  analogy(s, '🏦', '銀行帳戶的最低留存餘額 1,000 元：意思是「不准低於 1,000」，不是「你的目標是存 1,000」。有人把 permissive 讀成後者——那是完全相反的做法。',
    3.85, 0.90);
  s.addNotes('用這一頁把 permissive 這個字的方向講清楚。');
}

/* 16 — PRESSURE 不能回答什麼 */
{
  const s = base('INTERPRETATION', 'PRESSURE 不能拿來改現場流程', '被問到的時候，用這張表回答');
  table(s, [
    ['面向', 'PRESSURE 在做的', '到院前小兒創傷'],
    ['場域', 'PICU 內，已插管', '現場／救護車上'],
    ['工具', '用升壓劑調整 MAP', '用輸液處理失血'],
    ['病人', '多為敗血症等分佈性休克', '出血性休克'],
    ['腦傷', '明文排除急性腦損傷', '現場常無法排除'],
    ['問題', '「用藥把血壓拉到多高才夠？」', '「輸液可以少給多少？」'],
  ], { y:1.75, colW:[2.0, 5.4, 4.833], keyRows:[1], size:13, rowH:0.52 });
  punch(s, 'PRESSURE 問的是「拉到多高才夠」，不是「可以放著多低」。', 5.05, 0.56);
  s.addText('在它公布結果前——即使公布之後——到院前小兒創傷的建議完全不變。',
    { x:M, y:5.72, w:W, h:0.34, fontSize:13, color:C.muted, fontFace:F, margin:0 });
  s.addNotes('這一頁是為了應付「我看到英國在做 permissive」這個提問。');
}

/* 17 — 台灣 SOP 對照 */
{
  const s = base('LOCAL PROTOCOL', '回到台灣的 SOP',
    '好消息：限制性輸液其實已經內建在裡面了');
  table(s, [
    ['項目', '臺北市 EMTP SOP（T5 創傷性休克）', '新北市 2024 SOP（兒童流程）'],
    ['兒童輸液', '<40 kg：20 mL/kg，得重複一次（上限 40）', '休克者考慮 10–20 mL/kg'],
    ['兒童休克血壓', '未特別列（成人 SBP<90）', '<1 歲 <70；1–10 歲 <70+2×歲；>10 歲 <90'],
    ['IV 嘗試次數', '休克患者現場最多 2 次（或 5 分鐘）', '同左'],
    ['保溫', '鋁箔／毛毯、車內暖氣、加溫輸液 39°C', '同左'],
    ['TXA', '成人指標：SBP<90 或 HR≥120、3 小時內、1 g／10 分鐘', '兒童劑量未載'],
  ], { y:1.72, colW:[2.2, 5.6, 4.433], keyRows:[0], size:11.5, rowH:0.50 });
  card(s, { x:M, y:4.68, w:W, h:0.92, accent:C.green, size:13,
    body:'不必新增任何規則。現行 SOP 的「20 mL/kg × 2 為限」＝ 40 mL/kg，剛好停在文獻的傷害訊號之前——教學只要把「為什麼有這個上限」講清楚。' });
  s.addText('2026-08-17 核對本機 SOP 全文。兒童 TXA 劑量請依線上醫療指導，勿由成人劑量外推。',
    { x:M, y:5.72, w:W, h:0.30, fontSize:10.5, color:C.dim, fontFace:F, margin:0 });
  s.addNotes('這一頁讓學員知道現行 protocol 本來就對——不是要他們改規則。');
}

/* 18 — 口訣 */
{
  const s = base('MNEMONIC', '帶回現場的口訣', '四句、七言，走路的時候都背得起來');
  s.addShape(pres.ShapeType.roundRect, { x:M, y:1.80, w:W, h:1.75, rectRadius:0.10,
    fill:{ color:C.darkred }, line:{ color:C.gold, width:2 } });
  s.addText('先止血、再給水；一次十、打完看；\n摸到脈、就停手；撞到頭、不能低。',
    { x:M+0.3, y:1.80, w:W-0.6, h:1.75, fontSize:28, bold:true, color:C.white, align:'center',
      valign:'middle', fontFace:F, margin:0, lineSpacingMultiple:1.5 });
  const w3 = (W - 0.36 * 2) / 3;
  card(s, { x:M, y:3.80, w:w3, h:1.55, label:'一次十', size:12.5,
    body:'10 mL/kg。貴縣市寫 20 就教「一次二十」——重點在打完必須重評，不是一路掛到底。' });
  card(s, { x:M+w3+0.36, y:3.80, w:w3, h:1.55, label:'摸到脈', size:12.5,
    body:'終點是灌流不是數字：意識、橈動脈、皮膚回溫。' });
  card(s, { x:M+(w3+0.36)*2, y:3.80, w:w3, h:1.55, label:'撞到頭', size:12.5,
    body:'唯一的例外條款。懷疑腦傷就不容許任何低血壓。' });
  s.addNotes('讓全班一起唸兩次，再進到四步流程。');
}

/* 19 — 四步流程 */
{
  const s = base('FIELD FLOW', '現場只有四步', '流程越短，壓力下越不會出錯');
  card(s, { x:M, y:1.75, w:5.95, h:1.42, accent:C.red, title:'① 止血優先', size:13,
    titleColor:C.red, body:'直接加壓 → 止血帶 → 骨盆固定帶。任何輸液都不會止血。' });
  card(s, { x:M+6.28, y:1.75, w:5.95, h:1.42, accent:C.amber, title:'② 判斷休克，別只看血壓',
    size:13, titleColor:C.amber, body:'意識、橈動脈、CRT>2 秒、皮膚濕冷、心跳快。血壓只當紅線。' });
  card(s, { x:M, y:3.32, w:5.95, h:1.42, accent:C.teal, title:'③ 小量、多次、看反應', size:13,
    titleColor:C.teal, body:'10（或 20）mL/kg → 立刻重評 → 有改善就停。心裡那條線：40 mL/kg。' });
  card(s, { x:M+6.28, y:3.32, w:5.95, h:1.42, title:'④ 保溫 ＋ 快送', size:13,
    body:'脫濕衣、鋁毯、暖氣、加溫輸液。腹腔／骨盆出血 → scoop and run。' });
  analogy(s, '🍲', '為什麼灌太多會讓血止不住？把濃湯一直加水，量變多了、味道卻沒了。血液被稀釋後凝血因子濃度不夠，血栓就結不起來。',
    4.90, 0.78);
  s.addNotes('四步順序不可調換：止血永遠在輸液之前。');
}

/* 20 — 口袋卡 */
{
  const s = base('POCKET CARD', '口袋卡：紅線與劑量', '貼在血壓計上的那一張');
  s.addText('血壓紅線（跌破＝很糟）', { x:M, y:1.72, w:5.95, h:0.32, fontSize:15, bold:true,
    color:C.red, fontFace:F, margin:0 });
  s.addTable([
    [ { text:'年齡', options:{ bold:true, color:C.muted, fontSize:11, fill:{ color:C.panel2 } } },
      { text:'收縮壓', options:{ bold:true, color:C.muted, fontSize:11, fill:{ color:C.panel2 },
        align:'right' } } ],
    ...[['< 1 歲','SBP < 70'],['1–10 歲','SBP < 70 + 2×歲'],['> 10 歲','SBP < 90']].map(r => [
      { text:r[0], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel } } },
      { text:r[1], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel }, align:'right',
        bold:true } } ]),
  ], { x:M, y:2.10, w:5.95, colW:[3.0, 2.95], fontFace:F,
       border:{ type:'solid', color:C.line, pt:1 }, valign:'middle', margin:6, rowH:0.44 });
  s.addText('口算：5 歲 = 80、8 歲 = 86、10 歲 = 90', { x:M, y:3.95, w:5.95, h:0.30,
    fontSize:11.5, color:C.dim, fontFace:F, margin:0 });

  s.addText('一次的輸液量（生理食鹽水）', { x:M+6.28, y:1.72, w:5.95, h:0.32, fontSize:15,
    bold:true, color:C.teal, fontFace:F, margin:0 });
  s.addTable([
    [ { text:'體重', options:{ bold:true, color:C.muted, fontSize:11, fill:{ color:C.panel2 } } },
      { text:'10–20 mL/kg', options:{ bold:true, color:C.muted, fontSize:11,
        fill:{ color:C.panel2 }, align:'right' } } ],
    ...[['10 kg（約 1 歲）','100–200 mL'],['20 kg（約 6 歲）','200–400 mL'],
        ['30 kg（約 10 歲）','300–600 mL']].map(r => [
      { text:r[0], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel } } },
      { text:r[1], options:{ fontSize:13, color:C.ink, fill:{ color:C.panel }, align:'right',
        bold:true } } ]),
  ], { x:M+6.28, y:2.10, w:5.95, colW:[3.0, 2.95], fontFace:F,
       border:{ type:'solid', color:C.line, pt:1 }, valign:'middle', margin:6, rowH:0.44 });
  s.addText('體重未知：Broselow 帶或 (年齡+4)×2　｜　現場總量上限 40 mL/kg（20 kg ≒ 800 mL）',
    { x:M+6.28, y:3.95, w:5.95, h:0.30, fontSize:11.5, color:C.dim, fontFace:F, margin:0 });
  punch(s, '交班一定要講：總共給了幾 mL、幾點給的、給完之後脈搏和意識有沒有變。', 4.55);
  s.addText('沒講量，醫院就會從頭再灌一次。', { x:M, y:5.22, w:W, h:0.32, fontSize:13,
    color:C.muted, fontFace:F, margin:0 });
  s.addNotes('這一頁可以印成 A6 護貝，發給每一位學員。');
}

/* 21 — 誤解 */
{
  const s = base('MISCONCEPTIONS', '要主動打掉的四個誤解', '這些是課後最容易被記錯的版本');
  table(s, [
    ['常見誤解', '正確版本'],
    ['「現在流行 permissive hypotension，小孩也少給水就好」',
     '兒童沒有這個建議。少給水是為了避免過量傷害，不是追求低血壓'],
    ['「血壓還可以，那就沒休克」', '兒童代償強，血壓最後才掉。看意識、脈搏強度、CRT'],
    ['「休克就先掛兩袋跑」', '每次 bolus 後都要重評；>40 mL/kg 與死亡率上升相關（觀察性）'],
    ['「英國已經在試 permissive 了，所以小孩可以低血壓」',
     'PRESSURE 是 PICU 重症、用升壓藥、排除腦傷，且尚無結果；它的 permissive 是「MAP 不低於第 5 百分位」'],
  ], { y:1.75, colW:[5.2, 7.033], keyRows:[3], size:12.5, rowH:0.72 });
  punch(s, '學員記不住否定句，只記得住動作——糾正時一定要給替代版本。', 5.25);
  s.addNotes('每一條都先問學員「這句話哪裡怪怪的？」再給正確版本。');
}

/* 22 — OSCE */
{
  const s = base('OSCE', '情境檢核', '四題，每題只問一個決策點');
  table(s, [
    ['情境', '該做什麼', '考點'],
    ['A　6 歲、車禍、腹脹壓痛、HR 160、CRT 4 秒、BP 96/70',
     '當休克處理：確認止血 → 10–20 mL/kg → 重評 → 快送', '血壓正常 ≠ 沒休克'],
    ['B　4 歲、墜落、意識混亂、瞳孔不等大、SBP 72',
     '疑似 TBI：積極矯正低血壓、避免缺氧、送創傷中心', '腦傷是禁忌'],
    ['C　8 歲約 30 kg，已給 1,200 mL 仍蒼白脈弱',
     '已達 40 mL/kg → 停止加碼、通報、保溫、快送', '需要的是血與手術'],
    ['D　2 歲、下肢大量外出血、HR 180',
     '先止血帶／加壓止血，再考慮輸液', '止血優先於輸液'],
  ], { y:1.75, colW:[4.6, 5.0, 2.633], keyRows:[1], size:12, rowH:0.66 });
  punch(s, '不要問「該給多少」，要問「打完這一次，你看什麼決定要不要打第二次」。', 5.20);
  s.addNotes('C 題可請學員現場算 40 mL/kg 是多少，強化換算。');
}

/* 23 — 證據等級 */
{
  const s = base('PROVENANCE', '本簡報數字的來源與等級', '教學材料也要能被查核');
  table(s, [
    ['關鍵數字', '來源', '等級'],
    ['成人 permissive OR 0.70（0.53–0.92）', 'Tran 2018 統合分析（6 RCT, n=1,158）', '✓ 同儕審查'],
    ['≥40 mL/kg 死亡 aOR 6.26', 'Mbadiwe 2021（單中心回溯，已校正 ISS）', '○ 單一中心觀察性'],
    ['>60 mL/kg 存活下降', '轉引自 Russell 2023 共識之文獻回顧', '○ 二手引用'],
    ['PEGASUS 23.3% vs 8.6%、aRR 0.46', 'Kannan 2018, Pediatr Emerg Care（已核對原文）', '✓ 同儕審查'],
    ['PRESSURE 設計與 MAP 帶', 'PCCM 2024 protocol ＋ ISRCTN 20609635', '⚠ 進行中，無結果'],
    ['臺北／新北 SOP 條文', '本機 SOP 全文，2026-08-17 核對', '✓ 官方文件'],
    ['「40 mL/kg ≒ 20 kg 童 800 mL」', '本簡報換算', '△ 概算，僅供教學'],
  ], { y:1.72, colW:[4.6, 5.2, 2.433], size:11.5, rowH:0.44 });
  s.addText('✓ 官方／同儕審查　　⚠ 進行中或草案　　○ 單一來源／二手　　△ 自行推算',
    { x:M, y:5.55, w:W, h:0.30, fontSize:11, color:C.dim, fontFace:F, margin:0 });
  s.addNotes('被問到「這個數字哪裡來的」時翻這一頁。');
}

/* 24 — 參考文獻 */
{
  const s = base('REFERENCES', '參考文獻', '完整可點版本見線上全文閱讀頁');
  const refs = [
    'Russell RT, et al. Pediatric Traumatic Hemorrhagic Shock Consensus Conference Recommendations. J Trauma Acute Care Surg. 2023;94(1S):S2-S10.',
    'Tran A, et al. Permissive hypotension versus conventional resuscitation in adult trauma. J Trauma Acute Care Surg. 2018;84(5):802-808.',
    'Lasa JJ, et al. Part 8: Pediatric Advanced Life Support: 2025 AHA/AAP Guidelines. Pediatrics. 2026;157(1):e2025074351.',
    'Buis ML, Turner NM. New ERC guidelines for pediatric life support. Paediatr Anaesth. 2022;32(4):497-503.',
    'Mbadiwe N, et al. Higher Crystalloid Volume During Initial Pediatric Trauma Resuscitation Is Associated With Mortality. J Surg Res. 2021;262:93-100.',
    'Sarganas G, et al. An unambiguous definition of pediatric hypotension is still lacking. J Trauma Acute Care Surg. 2019;86(3):448-453.',
  ];
  const refs2 = [
    'Hagedoorn NN, et al. A comparison of clinical paediatric guidelines for hypotension with population-based lower centiles. Crit Care. 2019;23:380.',
    'Lulla A, et al. Prehospital Guidelines for the Management of TBI – 3rd Edition. Prehosp Emerg Care. 2023;27(5):507-538.',
    'Kannan N, et al. Timely Hemodynamic Resuscitation and Outcomes in Severe Pediatric TBI. Pediatr Emerg Care. 2018;34(5):325-329.',
    'Peters MJ, et al. PRESSURE trial protocol. Pediatr Crit Care Med. 2024;25(7):629-637.',
    'PATCH-Trauma Investigators. Prehospital Tranexamic Acid for Severe Trauma. N Engl J Med. 2023;389(2):127-136.',
    'Spinella PC, et al. Use of Antifibrinolytics in Pediatric Life-Threatening Hemorrhage. Crit Care Med. 2022;50(4):e382-e392.',
  ];
  const mk = (arr, start) => arr.map((t, i) => ({ text:`${start+i}. ${t}`,
    options:{ breakLine:i !== arr.length-1 } }));
  s.addText(mk(refs, 1), { x:M, y:1.75, w:5.95, h:3.6, fontSize:10.5, color:C.muted,
    fontFace:F, margin:0, paraSpaceAfter:7, lineSpacingMultiple:1.12, valign:'top' });
  s.addText(mk(refs2, 7), { x:M+6.28, y:1.75, w:5.95, h:3.6, fontSize:10.5, color:C.muted,
    fontFace:F, margin:0, paraSpaceAfter:7, lineSpacingMultiple:1.12, valign:'top' });
  s.addShape(pres.ShapeType.line, { x:M, y:5.50, w:W, h:0, line:{ color:C.line, width:1 } });
  s.addText('線上簡報：limzijiang.github.io/ped-permissive-hypotension　｜　全文閱讀版：同網址 /read.html',
    { x:M, y:5.62, w:W, h:0.32, fontSize:12, color:C.ink, fontFace:F, margin:0 });
  s.addText('實證檢索：OpenEvidence，2026-08-17（2 次查詢）；SOP 條文核對自臺北市與新北市 EMTP SOP 全文。',
    { x:M, y:5.96, w:W, h:0.30, fontSize:10.5, color:C.dim, fontFace:F, margin:0 });
  s.addNotes('把網址念一次，請學員課後自己點文獻。');
}

/* 25 — 收尾 */
{
  const s = pres.addSlide();
  s.background = { color:C.bg };
  s.addShape(pres.ShapeType.roundRect, { x:6.6, y:4.4, w:8.4, h:4.6, rectRadius:0.5,
    fill:{ color:C.darkred, transparency:80 }, line:{ color:C.bg, width:0 } });
  s.addText('帶走這三句', { x:M, y:1.55, w:W, h:0.5, fontSize:14, bold:true, color:C.red,
    charSpacing:2.4, fontFace:F, margin:0 });
  const lines = [
    '兒童創傷休克三件事：壓住出血、保暖、快送。',
    '輸液是小口喝，不是灌水；打完一定重評。',
    '撞到頭，血壓一低就要救回來。',
  ];
  lines.forEach((t, i) => {
    s.addText(`${i+1}`, { x:M, y:2.20 + i*1.05, w:0.55, h:0.55, fontSize:22, bold:true,
      color:C.red, align:'center', fontFace:F, margin:0 });
    s.addText(t, { x:M+0.75, y:2.20 + i*1.05, w:W-0.75, h:0.55, fontSize:22, bold:true,
      color:C.white, fontFace:F, margin:0, valign:'middle' });
  });
  s.addShape(pres.ShapeType.line, { x:M, y:5.55, w:4.2, h:0, line:{ color:C.line, width:1 } });
  s.addText('林自強　·　2026-08-17　·　limzijiang.github.io/ped-permissive-hypotension',
    { x:M, y:5.70, w:W, h:0.34, fontSize:12.5, color:C.dim, fontFace:F, margin:0 });
  s.addNotes('結束前再問一次開場的案例：現在你會怎麼做？');
}

pres.writeFile({ fileName: '/Users/mac/Downloads/claude working/ped-permissive-hypotension/小兒創傷_permissive_hypotension_教學.pptx' })
  .then(f => console.log('WROTE', f));
