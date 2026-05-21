document.addEventListener('DOMContentLoaded', () => {
  // Auto-calculate total cast
  const maleCast = document.getElementById('maleCast');
  const femaleCast = document.getElementById('femaleCast');
  const totalCast = document.getElementById('totalCast');

  function updateTotal() {
    const m = parseInt(maleCast.value) || 0;
    const f = parseInt(femaleCast.value) || 0;
    totalCast.value = m + f;
  }
  maleCast.addEventListener('input', updateTotal);
  femaleCast.addEventListener('input', updateTotal);

  // Enable/disable その他 text input
  const otherInput = document.getElementById('perfRightsOther');
  document.querySelectorAll('input[name="perfRights"]').forEach(radio => {
    radio.addEventListener('change', () => {
      const isOther = document.querySelector('input[name="perfRights"]:checked')?.value === 'その他';
      otherInput.disabled = !isOther;
    });
  });

  document.getElementById('generateBtn').addEventListener('click', generateWord);
});

function collectData() {
  const v = id => (document.getElementById(id)?.value || '').trim();
  const radio = name => document.querySelector(`input[name="${name}"]:checked`)?.value || '';

  return {
    recYear: v('recYear'),
    recMonth: v('recMonth'),
    recDay: v('recDay'),
    submitterName: v('submitterName'),
    perfYear: v('perfYear'),
    theater: v('theater'),
    workType: radio('workType'),
    workNameFuri: v('workNameFuri'),
    workName: v('workName'),
    origNameFuri: v('origNameFuri'),
    origName: v('origName'),
    authorFuri: v('authorFuri'),
    authorName: v('authorName'),
    translator: v('translator'),
    scriptAvail: radio('scriptAvail'),
    maleCast: v('maleCast'),
    femaleCast: v('femaleCast'),
    totalCast: v('totalCast'),
    perfHours: v('perfHours'),
    perfMinutes: v('perfMinutes'),
    perfActs: v('perfActs'),
    perfRights: radio('perfRights'),
    perfRightsOther: v('perfRightsOther'),
    perfHistory: v('perfHistory'),
    authorIntro: v('authorIntro'),
    proposalReason: v('proposalReason'),
    completionDate: v('completionDate'),
    specialNotes: v('specialNotes'),
  };
}

async function generateWord() {
  const btn = document.getElementById('generateBtn');
  btn.disabled = true;
  btn.textContent = '生成中...';

  try {
    const data = collectData();
    const doc = buildDocument(data);
    const blob = await docx.Packer.toBlob(doc);
    const safeName = (data.workName || '無題').replace(/[/\\?%*:|"<>]/g, '_');
    downloadBlob(blob, `企画書_${safeName}.docx`);
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
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── Document builder ───────────────────────────────────────────────────────

function buildDocument(d) {
  const {
    Document, Paragraph, Table, TableRow, TableCell, TextRun,
    AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak,
  } = docx;

  const LABEL_WIDTH = 2800; // twips
  const FONT_SIZE = 20;     // half-points = 10pt
  const FONT_SIZE_SM = 18;

  // ── Border presets ──────────────────────────────────────────────────────
  const borderThin = { style: BorderStyle.SINGLE, size: 1, color: 'AAAAAA' };
  const borderMid  = { style: BorderStyle.SINGLE, size: 4, color: '2C5F8A' };
  const borderNone = { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' };

  const tableBorders = {
    top: borderMid, bottom: borderMid, left: borderMid, right: borderMid,
    insideH: borderThin, insideV: borderThin,
  };

  const boxBorders = {
    top: borderThin, bottom: borderThin, left: borderThin, right: borderThin,
    insideH: borderNone, insideV: borderNone,
  };

  // ── Helpers ─────────────────────────────────────────────────────────────

  function run(text, opts = {}) {
    return new TextRun({ text, size: FONT_SIZE, font: 'メイリオ', ...opts });
  }

  function para(children, opts = {}) {
    return new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
  }

  // Row: section header (full-width, dark background)
  function sectionRow(title) {
    return new TableRow({
      children: [
        new TableCell({
          columnSpan: 2,
          shading: { type: ShadingType.CLEAR, fill: '2C5F8A', color: 'auto' },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [para(run(title, { bold: true, color: 'FFFFFF', size: FONT_SIZE }))],
        }),
      ],
    });
  }

  // Row: label + value
  function infoRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: LABEL_WIDTH, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: 'E8E8E8', color: 'auto' },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [para(run(label, { bold: true }))],
        }),
        new TableCell({
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [para(run(value || ''))],
        }),
      ],
    });
  }

  // Checkbox representation: ■ selected  □ unselected
  function checkboxLine(options, selected) {
    return options.map(opt => (opt === selected ? '■ ' : '□ ') + opt).join('　');
  }

  // Multi-line text in a bordered box
  function freeTextBlock(heading, text) {
    const lines = (text || '').split('\n');
    const cellChildren = lines.length > 0 && text
      ? lines.map(line => para(run(line || ' ', { size: FONT_SIZE }), { spacing: { after: 40 } }))
      : [para(run('　'))];

    return [
      para(run(heading, { bold: true, size: FONT_SIZE }), { spacing: { before: 180, after: 80 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: boxBorders,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: cellChildren,
              }),
            ],
          }),
        ],
      }),
    ];
  }

  // ── Format values ────────────────────────────────────────────────────────

  const recDate = (d.recYear || d.recMonth || d.recDay)
    ? `${d.recYear || '　　'}年　${d.recMonth || '　'}月　${d.recDay || '　'}日`
    : '';

  const perfTime = [
    d.perfHours ? `${d.perfHours}時間` : '',
    d.perfMinutes ? `${d.perfMinutes}分` : '',
    d.perfActs ? `（${d.perfActs}幕中場）` : '',
  ].filter(Boolean).join('　');

  const rightsValue = d.perfRights === 'その他' && d.perfRightsOther
    ? checkboxLine(['取得済み', '調査済み', '請査中', '未調査', 'その他'], 'その他') + `（${d.perfRightsOther}）`
    : checkboxLine(['取得済み', '調査済み', '請査中', '未調査', 'その他'], d.perfRights);

  const castValue = [
    `男　${d.maleCast || '　'}名`,
    `女　${d.femaleCast || '　'}名`,
    `合計　${d.totalCast || '　'}名`,
  ].join('　　');

  // ── Main info table ──────────────────────────────────────────────────────

  const mainTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [
      sectionRow('基本情報'),
      infoRow('受付日', recDate),
      infoRow('企画提出者名', d.submitterName),
      infoRow('上演希望年', d.perfYear ? `${d.perfYear}年` : ''),
      infoRow('希望劇場', d.theater),

      sectionRow('作品情報'),
      infoRow('作品区分', checkboxLine(['書き下ろし作品', '海外の既存作品'], d.workType)),
      infoRow('作品名（ふりがな）', d.workNameFuri),
      infoRow('作品名', d.workName),
      infoRow('原作名（ふりがな）', d.origNameFuri),
      infoRow('原作名', d.origName),
      infoRow('作者名（ふりがな）', d.authorFuri),
      infoRow('作者名', d.authorName),
      infoRow('翻訳・脚色・翻案者', d.translator),

      sectionRow('台本・出演情報'),
      infoRow('台本の有無', checkboxLine(['有', '無'], d.scriptAvail)),
      infoRow('登場人物', castValue),

      sectionRow('上演情報'),
      infoRow('上演予定時間', perfTime),
      infoRow('上演権', rightsValue),
    ],
  });

  // ── Divider ──────────────────────────────────────────────────────────────

  const divider = new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB', space: 4 } },
    spacing: { before: 160, after: 160 },
    children: [],
  });

  // ── Document children ────────────────────────────────────────────────────

  const children = [
    // Title
    para(run('企　画　書', { bold: true, size: 36 }), {
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
    }),

    // Main table
    mainTable,

    divider,

    // 記述欄 heading
    para(run('記述欄', { bold: true, size: 24, color: '2C5F8A' }), {
      spacing: { before: 0, after: 100 },
    }),

    ...freeTextBlock('■ 上演歴', d.perfHistory),
    ...freeTextBlock('■ 作者紹介・あらすじ等', d.authorIntro),
    ...freeTextBlock('■ 企画趣旨（推薦理由等）', d.proposalReason),

    para(run('※ 裏面に続く', { size: FONT_SIZE_SM, italics: true, color: '888888' }), {
      alignment: AlignmentType.RIGHT,
      spacing: { before: 120, after: 0 },
    }),

    // Page break
    para(new PageBreak()),

    // 裏面 title
    para(run('企　画　書（裏面）', { bold: true, size: 32 }), {
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),

    ...freeTextBlock('■ 未翻訳又は書き下ろしの場合、いつ出来上がりますか？', d.completionDate),
    ...freeTextBlock('■ 特記事項（希望する演出者・出演者があれば記入して下さい。）', d.specialNotes),

    divider,

    // 注記
    para(run('注記', { bold: true, size: FONT_SIZE }), { spacing: { before: 0, after: 80 } }),
    para(run('１．企画作品の原本コピーを添えて提出して下さい。', { size: FONT_SIZE_SM })),
    para(run('２．未翻訳の作品は、粗訳又は粗筋を添えて提出して下さい。', { size: FONT_SIZE_SM })),
    para(run('３．企画提出に至れない場合は、別紙を添付して下さい。', { size: FONT_SIZE_SM })),
  ];

  return new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 720, bottom: 720, left: 1080, right: 1080 },
        },
      },
      children,
    }],
  });
}
