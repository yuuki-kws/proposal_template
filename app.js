// ─── Initialization ──────────────────────────────────────────────────────────

function init() {
  // Year select
  const currentYear = new Date().getFullYear();
  const perfYearSel = document.getElementById('perfYear');
  for (let y = currentYear; y <= currentYear + 15; y++) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    perfYearSel.appendChild(opt);
  }

  // Total-minutes select (5-min increments, 0–360 min)
  const perfMinSel = document.getElementById('perfTotalMinutes');
  for (let m = 0; m <= 360; m += 5) {
    const h = Math.floor(m / 60), r = m % 60;
    let label;
    if (m === 0) label = '0分';
    else if (h === 0)  label = `${m}分`;
    else if (r === 0)  label = `${m}分（${h}時間）`;
    else               label = `${m}分（${h}時間${r}分）`;
    const opt = document.createElement('option');
    opt.value = m; opt.textContent = label;
    perfMinSel.appendChild(opt);
  }

  // Cast auto-total
  const maleSel   = document.getElementById('maleCast');
  const femaleSel = document.getElementById('femaleCast');
  const totalInp  = document.getElementById('totalCast');
  function updateTotal() {
    const m = parseInt(maleSel.value)   || 0;
    const f = parseInt(femaleSel.value) || 0;
    totalInp.value = m + f;
  }
  maleSel.addEventListener('change', updateTotal);
  femaleSel.addEventListener('change', updateTotal);

  // その他 toggle
  const otherInput = document.getElementById('perfRightsOther');
  document.querySelectorAll('input[name="perfRights"]').forEach(r => {
    r.addEventListener('change', () => {
      otherInput.disabled =
        document.querySelector('input[name="perfRights"]:checked')?.value !== 'その他';
    });
  });

  document.getElementById('generateBtn').addEventListener('click', generateWord);
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}

// ─── Data collection ─────────────────────────────────────────────────────────

function collectData() {
  const v     = id   => (document.getElementById(id)?.value || '').trim();
  const radio = name => document.querySelector(`input[name="${name}"]:checked`)?.value || '';
  return {
    submitterName:    v('submitterName'),
    perfYear:         v('perfYear'),
    perfMonth:        v('perfMonth'),
    theater:          v('theater'),
    workType:         radio('workType'),
    workNameFuri:     v('workNameFuri'),
    workName:         v('workName'),
    origNameFuri:     v('origNameFuri'),
    origName:         v('origName'),
    authorFuri:       v('authorFuri'),
    authorName:       v('authorName'),
    translator:       v('translator'),
    scriptAvail:      radio('scriptAvail'),
    maleCast:         v('maleCast'),
    femaleCast:       v('femaleCast'),
    totalCast:        v('totalCast'),
    perfTotalMinutes: v('perfTotalMinutes'),
    perfRights:       radio('perfRights'),
    perfRightsOther:  v('perfRightsOther'),
    perfHistory:      v('perfHistory'),
    authorIntro:      v('authorIntro'),
    proposalReason:   v('proposalReason'),
    completionDate:   v('completionDate'),
    specialNotes:     v('specialNotes'),
  };
}

// ─── Word generation ─────────────────────────────────────────────────────────

async function generateWord() {
  const btn = document.getElementById('generateBtn');
  btn.disabled = true; btn.textContent = '生成中...';
  try {
    const data = collectData();
    const blob = await docx.Packer.toBlob(buildDocument(data));
    const safe = (data.workName || '無題').replace(/[/\\?%*:|"<>]/g, '_');
    downloadBlob(blob, `企画書_${safe}.docx`);
  } catch (err) {
    console.error(err);
    alert('Wordファイルの生成中にエラーが発生しました。\n' + err.message);
  } finally {
    btn.disabled = false; btn.textContent = 'Word ファイルを生成・ダウンロード';
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── Document builder ────────────────────────────────────────────────────────

function buildDocument(d) {
  const {
    Document, Paragraph, Table, TableRow, TableCell, TextRun,
    AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, HeightRule,
  } = docx;

  // ── Page metrics ─────────────────────────────────────────────────────────
  // A4: 11906 × 16838 twips. Left/right margin 720, top/bottom 720.
  // Available width = 11906 - 1440 = 10466 twips
  // 12-column grid: each unit = 872 twips (12 × 872 = 10464 ≈ 10466)
  const U    = 872;
  const COLS = 12;
  const TW   = U * COLS;          // 10464 twips — table fills page width

  const FONT = 'MS 明朝';
  const FS   = 20;  // 10pt
  const FSS  = 16;  //  8pt – ふりがな / small notes
  const FSL  = 32;  // 16pt – title

  // ── Borders ──────────────────────────────────────────────────────────────
  const bOuter = { style: BorderStyle.SINGLE, size: 8,  color: '000000' };
  const bInner = { style: BorderStyle.SINGLE, size: 2,  color: '444444' };
  const bNone  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

  const OUTER_B = { top: bOuter, bottom: bOuter, left: bOuter, right: bOuter,
                    insideH: bInner, insideV: bInner };
  const BOX_B   = { top: bInner, bottom: bInner, left: bInner, right: bInner,
                    insideH: bNone,  insideV: bNone  };

  // ── Primitive helpers ─────────────────────────────────────────────────────

  const mkRun = (text, opts = {}) =>
    new TextRun({ text: text || ' ', font: FONT, size: opts.size || FS, ...opts });

  const mkPara = (runs, align = AlignmentType.LEFT) =>
    new Paragraph({
      alignment: align,
      spacing: { before: 0, after: 0 },
      children: Array.isArray(runs) ? runs : [runs],
    });

  // Core cell: paragraphs[], columnSpan, optional gray shade
  function mkCell(paras, span, shade) {
    return new TableCell({
      columnSpan: span,
      width:   { size: span * U, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      shading: shade
        ? { type: ShadingType.CLEAR, fill: shade, color: 'auto' }
        : undefined,
      children: paras,
    });
  }

  // Label cell: gray bg, bold, supports \n
  const lc = (text, span) =>
    mkCell(
      (text || ' ').split('\n').map(l => mkPara(mkRun(l, { bold: true }))),
      span, 'EFEFEF'
    );

  // Value cell: white bg, supports \n, empty → non-breaking space
  const vc = (text, span) =>
    mkCell(
      (text || '').split('\n').map(l => mkPara(mkRun(l || ' '))),
      span
    );

  // Furi-label cell: small "(ふりがな)" row + bold field name row
  const flc = (fieldName, span) =>
    mkCell([
      mkPara(mkRun('（ふりがな）', { size: FSS })),
      mkPara(mkRun(fieldName, { bold: true })),
    ], span, 'EFEFEF');

  // Furi-value cell: small furigana + normal name
  const fvc = (furi, name, span) =>
    mkCell([
      mkPara(mkRun(furi || ' ', { size: FSS })),
      mkPara(mkRun(name || ' ')),
    ], span);

  // ── Format values ─────────────────────────────────────────────────────────

  const today   = new Date();
  const recDate = `${today.getFullYear()}年　${today.getMonth() + 1}月　${today.getDate()}日`;

  const perfYM = [
    d.perfYear  ? `${d.perfYear}年`  : '',
    d.perfMonth ? `${d.perfMonth}月` : '',
  ].filter(Boolean).join('　') || ' ';

  // Minutes → hours + minutes conversion
  const totalMins = parseInt(d.perfTotalMinutes) || 0;
  const ph = Math.floor(totalMins / 60);
  const pm = totalMins % 60;
  const perfTimeStr = totalMins === 0 ? ' '
    : ph > 0 && pm > 0 ? `${ph}時間${pm}分`
    : ph > 0 ? `${ph}時間`
    : `${pm}分`;
  const perfTimeText = perfTimeStr;

  const chk = (sel, val) => sel === val ? '■' : '□';

  const workTypeText =
    `${chk(d.workType, '書き下ろし作品')} 1. 書き下ろし作品　　` +
    `${chk(d.workType, '海外の既存作品')} 2. 海外の既存作品`;

  const scriptText =
    `${chk(d.scriptAvail, '有')} 1. 有\n${chk(d.scriptAvail, '無')} 2. 無`;

  const rOpts = ['取得済み', '調査済み', '請査中', '未調査'];
  const rightsLine1 = rOpts.map((o, i) => `${chk(d.perfRights, o)} ${i + 1}. ${o}`).join('　');
  const rightsLine2 = `${chk(d.perfRights, 'その他')} 5. その他（${
    d.perfRights === 'その他' && d.perfRightsOther
      ? d.perfRightsOther : '　　　　　　　　　　'
  }）`;

  // ── Structured info table (12-column grid, fixed columnWidths) ────────────

  const COL_WIDTHS = Array(COLS).fill(U);   // [872, 872, …] × 12

  function tRow(cells, minH) {
    return new TableRow({
      height: minH ? { value: minH, rule: HeightRule.AT_LEAST } : undefined,
      children: cells,
    });
  }

  const infoTable = new Table({
    width:        { size: 5000, type: WidthType.PERCENTAGE },
    columnWidths: COL_WIDTHS,
    borders:      OUTER_B,
    rows: [
      // 企画書制作日 [2][7][3] = 12
      tRow([ lc('企画書制作日', 2), vc(recDate, 7), vc('（2020年度以降）', 3) ]),

      // 企画提出者名 | 上演希望年月 | 希望劇場 [2][2][2][2][2][2] = 12
      tRow([
        lc('企画提出者名', 2), vc(d.submitterName, 2),
        lc('上演希望年月', 2), vc(perfYM, 2),
        lc('希望劇場',    2), vc(d.theater, 2),
      ]),

      // 作品区分 [2][10] = 12
      tRow([ lc('作品区分', 2), vc(workTypeText, 10) ]),

      // 作品名 / 原作名 [2][4][2][4] = 12
      tRow([
        flc('作品名', 2), fvc(d.workNameFuri, d.workName, 4),
        flc('原作名', 2), fvc(d.origNameFuri, d.origName, 4),
      ]),

      // 作者名 / 翻訳者 [2][4][2][4] = 12
      tRow([
        flc('作者名', 2), fvc(d.authorFuri, d.authorName, 4),
        lc('翻訳・脚色・翻案者名', 3), vc(d.translator, 3),
      ]),

      // 台本の有無 | 登場人物 [2][2][2][2][2][2] = 12
      tRow([
        lc('台本の有無', 2), vc(scriptText, 2),
        lc('登場人物の人数', 2),
        vc(`男　${d.maleCast  || ' '}　名`, 2),
        vc(`女　${d.femaleCast || ' '}　名`, 2),
        vc(`合計　${d.totalCast || ' '}　名`, 2),
      ]),

      // 上演予定時間 [2][10] = 12
      tRow([ lc('上演予定時間', 2), vc(perfTimeText, 10) ]),

      // 上演権 [2][10] = 12
      tRow([ lc('上演権', 2), vc(rightsLine1 + '\n' + rightsLine2, 10) ]),
    ],
  });

  // ── Free text box helper ──────────────────────────────────────────────────

  function freeBox(heading, content, minHeight) {
    const lines    = content ? content.split('\n') : [];
    const cellParas = lines.length
      ? lines.map(l => mkPara(mkRun(l || ' ')))
      : [mkPara(mkRun(' '))];

    return [
      new Paragraph({
        spacing: { before: 180, after: 80 },
        children: [mkRun(heading, { bold: true })],
      }),
      new Table({
        width:        { size: 5000, type: WidthType.PERCENTAGE },
        columnWidths: [TW],
        borders:      BOX_B,
        rows: [
          new TableRow({
            height: { value: minHeight, rule: HeightRule.AT_LEAST },
            children: [
              new TableCell({
                width:   { size: 5000, type: WidthType.PERCENTAGE },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: cellParas,
              }),
            ],
          }),
        ],
      }),
    ];
  }

  // ── Assemble pages ────────────────────────────────────────────────────────

  const children = [
    // Title
    mkPara(mkRun('企画提出書（本公演）', { bold: true, size: FSL }),
           AlignmentType.CENTER),
    new Paragraph({ spacing: { before: 0, after: 140 }, children: [] }),

    infoTable,

    // ─ Page 1: free sections ─
    ...freeBox('■ 上演歴', d.perfHistory, 2600),
    ...freeBox('■ 作者紹介・あらすじ等', d.authorIntro, 3200),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing:   { before: 80, after: 60 },
      children:  [mkRun('裏面に続く', { size: FSS })],
    }),

    // Page break → 裏面
    new Paragraph({ children: [new PageBreak()] }),

    // ─ Page 2: back side ─
    ...freeBox('■ 企画趣旨（推薦理由等）', d.proposalReason, 2600),
    ...freeBox('■ 未翻訳又は書き下ろしの場合、いつ出来上がりますか？', d.completionDate, 2800),
    ...freeBox('■ 特記事項：希望する演出者・出演者があれば記入して下さい。', d.specialNotes, 3600),

    // 注記
    new Paragraph({
      spacing: { before: 180, after: 60 },
      children: [mkRun('注', { bold: true })],
    }),
    ...[
      '１．企画作品の原本コピーを添えて提出して下さい。',
      '２．未翻訳の作品は、粗訳又は粗筋を添えて提出して下さい。',
      '３．企画提出書に書ききれない場合は、別紙を添付して下さい。',
    ].map(t =>
      new Paragraph({ spacing: { before: 0, after: 0 }, children: [mkRun(t, { size: FSS })] })
    ),

    // 検討結果（blank – theater use）
    ...freeBox('■ 検討結果（検討日　　　年　　月　　日）', '', 2800),
  ];

  return new Document({
    sections: [{
      properties: {
        page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } },
      },
      children,
    }],
  });
}
