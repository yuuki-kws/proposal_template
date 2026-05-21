// ─── Initialization ──────────────────────────────────────────────────────────

function init() {
  const currentYear = new Date().getFullYear();
  const perfYearSel = document.getElementById('perfYear');
  for (let y = currentYear; y <= currentYear + 15; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    perfYearSel.appendChild(opt);
  }

  const maleCast  = document.getElementById('maleCast');
  const femaleCast = document.getElementById('femaleCast');
  const totalCast  = document.getElementById('totalCast');
  function updateTotal() {
    totalCast.value = (parseInt(maleCast.value) || 0) + (parseInt(femaleCast.value) || 0);
  }
  maleCast.addEventListener('input', updateTotal);
  femaleCast.addEventListener('input', updateTotal);

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
  const v    = id   => (document.getElementById(id)?.value || '').trim();
  const radio = name => document.querySelector(`input[name="${name}"]:checked`)?.value || '';
  return {
    submitterName:  v('submitterName'),
    perfYear:       v('perfYear'),
    perfMonth:      v('perfMonth'),
    theater:        v('theater'),
    workType:       radio('workType'),
    workNameFuri:   v('workNameFuri'),
    workName:       v('workName'),
    origNameFuri:   v('origNameFuri'),
    origName:       v('origName'),
    authorFuri:     v('authorFuri'),
    authorName:     v('authorName'),
    translator:     v('translator'),
    scriptAvail:    radio('scriptAvail'),
    maleCast:       v('maleCast'),
    femaleCast:     v('femaleCast'),
    totalCast:      v('totalCast'),
    perfHours:      v('perfHours'),
    perfMinutes:    v('perfMinutes'),
    perfActs:       v('perfActs'),
    perfRights:     radio('perfRights'),
    perfRightsOther: v('perfRightsOther'),
    perfHistory:    v('perfHistory'),
    authorIntro:    v('authorIntro'),
    proposalReason: v('proposalReason'),
    completionDate: v('completionDate'),
    specialNotes:   v('specialNotes'),
  };
}

// ─── Word generation ─────────────────────────────────────────────────────────

async function generateWord() {
  const btn = document.getElementById('generateBtn');
  btn.disabled = true;
  btn.textContent = '生成中...';
  try {
    const data = collectData();
    const doc  = buildDocument(data);
    const blob = await docx.Packer.toBlob(doc);
    const safe = (data.workName || '無題').replace(/[/\\?%*:|"<>]/g, '_');
    downloadBlob(blob, `企画書_${safe}.docx`);
  } catch (err) {
    console.error(err);
    alert('Wordファイルの生成中にエラーが発生しました。\n' + err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Word ファイルを生成・ダウンロード';
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── Document builder ────────────────────────────────────────────────────────

function buildDocument(d) {
  const {
    Document, Paragraph, Table, TableRow, TableCell, TextRun,
    AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, HeightRule,
  } = docx;

  // A4: 11906 twips wide. Margins 1080 each → available = 9746 twips
  // 10-column grid: each unit U = 974 twips (10 × 974 = 9740 ≈ 9746)
  const U    = 974;
  const TW   = U * 10;           // total table width
  const FONT = 'MS 明朝';
  const FS   = 20;               // 10pt (half-points)
  const FSS  = 16;               // 8pt  – ふりがな
  const FST  = 36;               // 18pt – title

  // Borders
  const bThick = { style: BorderStyle.SINGLE, size: 8,  color: '000000' };
  const bThin  = { style: BorderStyle.SINGLE, size: 2,  color: '000000' };
  const bNone  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };
  const OUTER  = { top: bThick, bottom: bThick, left: bThick, right: bThick, insideH: bThin,  insideV: bThin  };
  const BOX    = { top: bThin,  bottom: bThin,  left: bThin,  right: bThin,  insideH: bNone,  insideV: bNone  };

  // ── Primitives ────────────────────────────────────────────────────────────

  const run = (text, opts = {}) =>
    new TextRun({ text: text || '', font: FONT, size: opts.size || FS, ...opts });

  const para = (runs, align = AlignmentType.LEFT) =>
    new Paragraph({
      alignment: align,
      spacing: { before: 0, after: 0 },
      children: Array.isArray(runs) ? runs : [runs],
    });

  // Generic cell builder
  function mkCell(paragraphs, span, shade) {
    return new TableCell({
      columnSpan: span,
      width:   { size: span * U, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: 'auto' } : undefined,
      children: Array.isArray(paragraphs) ? paragraphs : [paragraphs],
    });
  }

  // Label cell (gray background, bold text, supports \n)
  const lc = (text, span) =>
    mkCell(
      (text || '').split('\n').map(l => para(run(l, { bold: true }))),
      span, 'F0F0F0'
    );

  // Value cell (white, supports \n)
  const vc = (text, span) =>
    mkCell(
      (text || '').split('\n').map(l => para(run(l))),
      span
    );

  // Furi-label cell: small "(ふりがな)" line + bold field name
  const flc = (fieldName, span) =>
    mkCell([
      para(run('（ふりがな）', { size: FSS })),
      para(run(fieldName, { bold: true })),
    ], span, 'F0F0F0');

  // Furi-value cell: small furigana value + normal name value
  const fvc = (furi, name, span) =>
    mkCell([
      para(run(furi || '', { size: FSS })),
      para(run(name || '')),
    ], span);

  // ── Format field values ───────────────────────────────────────────────────

  const today   = new Date();
  const recDate = `${today.getFullYear()} 年　${today.getMonth() + 1} 月　${today.getDate()} 日`;

  const perfYM = [
    d.perfYear  ? `${d.perfYear}年`  : '',
    d.perfMonth ? `${d.perfMonth}月` : '',
  ].filter(Boolean).join('　');

  const chk = (selected, val) => (selected === val ? '■' : '□');

  const workTypeText =
    `${chk(d.workType, '書き下ろし作品')} 1. 書き下ろし作品　　` +
    `${chk(d.workType, '海外の既存作品')} 2. 海外の既存作品`;

  const scriptText =
    `${chk(d.scriptAvail, '有')} 1. 有\n${chk(d.scriptAvail, '無')} 2. 無`;

  const perfTimeText =
    `${d.perfHours   || '　'} 時間　` +
    `${d.perfMinutes || '　'} 分　（` +
    `${d.perfActs    || '　'} 幕中場）`;

  const rightsRow1 =
    `${chk(d.perfRights, '取得済み')} 1. 取得済み　` +
    `${chk(d.perfRights, '調査済み')} 2. 調査済み　` +
    `${chk(d.perfRights, '請査中')}   3. 請査中　` +
    `${chk(d.perfRights, '未調査')}   4. 未調査`;
  const rightsRow2 =
    `${chk(d.perfRights, 'その他')} 5. その他（` +
    (d.perfRights === 'その他' && d.perfRightsOther
      ? d.perfRightsOther
      : '　　　　　　　　　　') + `）`;

  // ── Structured info table (10-column grid) ────────────────────────────────

  function tRow(cells, minHeight) {
    return new TableRow({
      height: minHeight ? { value: minHeight, rule: HeightRule.AT_LEAST } : undefined,
      children: cells,
    });
  }

  const infoTable = new Table({
    width:   { size: TW, type: WidthType.DXA },
    borders: OUTER,
    rows: [
      // 受付日  [2][5][3]
      tRow([
        lc('受付日', 2),
        vc(recDate, 5),
        vc('（2020年度以降）', 3),
      ]),

      // 企画提出者名 | 上演希望年月 | 希望劇場  [1][3][1][2][1][2]
      tRow([
        lc('企画\n提出者名', 1),
        vc(d.submitterName, 3),
        lc('上演\n希望年月', 1),
        vc(perfYM, 2),
        lc('希望\n劇場', 1),
        vc(d.theater, 2),
      ]),

      // 作品区分  [2][8]
      tRow([
        lc('作品区分', 2),
        vc(workTypeText, 8),
      ]),

      // 作品名 / 原作名  [2][3][2][3]
      tRow([
        flc('作品名', 2),
        fvc(d.workNameFuri, d.workName, 3),
        flc('原作名', 2),
        fvc(d.origNameFuri, d.origName, 3),
      ]),

      // 作者名 / 翻訳者  [2][3][2][3]
      tRow([
        flc('作者名', 2),
        fvc(d.authorFuri, d.authorName, 3),
        lc('（翻訳・脚色・翻案）\n者名', 2),
        vc(d.translator, 3),
      ]),

      // 台本の有無 | 登場人物  [1][2][2][2][1][2]
      tRow([
        lc('台本の\n有無', 1),
        vc(scriptText, 2),
        lc('登場人物\nの人数', 2),
        vc(`男　${d.maleCast  || '　'}　名`, 2),
        vc(`女　${d.femaleCast || '　'}　名`, 1),
        vc(`合計　${d.totalCast || '　'}　名`, 2),
      ]),

      // 上演予定時間  [2][8]
      tRow([
        lc('上演予定時間', 2),
        vc(perfTimeText, 8),
      ]),

      // 上演権  [2][8]
      tRow([
        lc('上演権', 2),
        vc(rightsRow1 + '\n' + rightsRow2, 8),
      ]),
    ],
  });

  // ── Free text box helper ──────────────────────────────────────────────────

  function freeBox(heading, content, minHeight = 2400) {
    const lines = (content || '').split('\n');
    const cellParas = lines.length && content
      ? lines.map(l => para(run(l)))
      : [para(run(''))];

    return [
      new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [run(heading, { bold: true })],
      }),
      new Table({
        width:   { size: TW, type: WidthType.DXA },
        borders: BOX,
        rows: [
          new TableRow({
            height: { value: minHeight, rule: HeightRule.AT_LEAST },
            children: [
              new TableCell({
                width:   { size: TW, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: cellParas,
              }),
            ],
          }),
        ],
      }),
    ];
  }

  // ── Assemble document ─────────────────────────────────────────────────────

  const children = [
    // Title
    para(run('企画提出書（本公演）', { bold: true, size: FST }), AlignmentType.CENTER),
    new Paragraph({ spacing: { before: 0, after: 160 }, children: [] }),

    // Info table
    infoTable,

    // Page 1 free sections
    ...freeBox('■ 上演歴', d.perfHistory, 2800),
    ...freeBox('■ 作者紹介・あらすじ等', d.authorIntro, 3200),

    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing:   { before: 100, after: 60 },
      children:  [run('裏面に続く', { size: FSS })],
    }),

    ...freeBox('■ 企画趣旨（推薦理由等）', d.proposalReason, 2800),

    // Page break
    new Paragraph({ children: [new PageBreak()] }),

    // Page 2 free sections
    ...freeBox('■ 未翻訳又は書き下ろしの場合、いつ出来上がりますか？', d.completionDate, 1600),
    ...freeBox('■ 特記事項：希望する演出者・出演者があれば記入して下さい。', d.specialNotes, 2400),

    // 注記
    new Paragraph({ spacing: { before: 200, after: 60 }, children: [run('注', { bold: true })] }),
    ...[
      '１．企画作品の原本コピーを添えて提出して下さい。',
      '２．未翻訳の作品は、粗訳又は粗筋を添えて提出して下さい。',
      '３．企画提出書に書ききれない場合は、別紙を添付して下さい。',
    ].map(t => new Paragraph({ spacing: { before: 0, after: 0 }, children: [run(t, { size: FSS })] })),

    // 検討結果（blank – for theater use）
    ...freeBox('■ 検討結果（検討日　　　年　　月　　日）', '', 1600),
  ];

  return new Document({
    sections: [{
      properties: {
        page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } },
      },
      children,
    }],
  });
}
