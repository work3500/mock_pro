const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, HeadingLevel, TabStopType,
  TabStopPosition, TableOfContents, PageBreak, UnderlineType
} = require('docx');
const fs = require('fs');

// ══════════════════════════════════════════════════════
//  KTNET 사내 정책 문서 표준 서식 v2
//  컨셉: 모던 다크그레이 + 포인트 컬러 (앰버/골드)
//  A4 · 맑은 고딕 · 스타일 갤러리 · TOC · 줄무늬 표 · 강조 블록
// ══════════════════════════════════════════════════════

const C = {
  // 기본 팔레트
  charcoal:    '2D2D2D',   // 제목 텍스트
  darkGray:    '404040',   // 본문 텍스트
  midGray:     '6B6B6B',   // 보조 텍스트
  lightGray:   'F2F2F2',   // 줄무늬 배경 (홀수)
  white:       'FFFFFF',
  // 포인트 컬러 (앰버/골드)
  amber:       'E8A020',   // 메인 포인트
  amberDark:   'C47E00',   // 진한 포인트
  amberLight:  'FDF3DC',   // 연한 강조 배경
  // 구조용
  ruleDark:    '2D2D2D',   // 굵은 구분선
  ruleLight:   'DDDDDD',   // 얇은 구분선
  tableBorder: 'CCCCCC',
  tableHeader: '2D2D2D',   // 표 헤더 배경
  tableStripe: 'F7F7F7',   // 짝수 행 배경
  infoBox:     'FDF3DC',   // 정보/주의 박스
  warnBox:     'FFF0F0',   // 경고 박스
};

const PAGE = {
  w: 11906, h: 16838,
  mTop: 1800, mBottom: 1700, mLeft: 1800, mRight: 1700,
  get cw() { return this.w - this.mLeft - this.mRight; }, // 8406
};

// ── 헬퍼 ──────────────────────────────────────────────

const gap = (n = 120) => new Paragraph({ spacing: { before: n, after: n }, children: [] });

const rule = (color = C.ruleLight, size = 4) =>
  new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    children: [],
  });

const thickRule = () =>
  new Paragraph({
    spacing: { before: 0, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 20, color: C.amber, space: 1 } },
    children: [],
  });

// ── 강조 블록 (텍스트 박스 대용, 단락 테두리) ─────────
const callout = (label, text, bgColor = C.amberLight, lineColor = C.amber) =>
  new Paragraph({
    spacing: { before: 160, after: 160 },
    indent: { left: 200, right: 200 },
    shading: { fill: bgColor, type: ShadingType.CLEAR },
    border: {
      top:    { style: BorderStyle.SINGLE, size: 12, color: lineColor, space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 4,  color: lineColor, space: 4 },
      left:   { style: BorderStyle.THICK,  size: 24, color: lineColor, space: 8 },
      right:  { style: BorderStyle.SINGLE, size: 4,  color: lineColor, space: 4 },
    },
    children: [
      new TextRun({ text: `${label}  `, bold: true, size: 19, font: '맑은 고딕', color: C.amberDark }),
      new TextRun({ text, size: 19, font: '맑은 고딕', color: C.darkGray }),
    ],
  });

const warnCallout = (text) => callout('⚠ 주의', text, C.warnBox, 'CC3333');
const noteCallout = (text) => callout('ℹ 참고', text, C.amberLight, C.amber);

// ── 표 공통 ───────────────────────────────────────────
const bord = (color = C.tableBorder) => ({
  top:    { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left:   { style: BorderStyle.SINGLE, size: 1, color },
  right:  { style: BorderStyle.SINGLE, size: 1, color },
});

const cell = (text, w, opts = {}) => new TableCell({
  borders: bord(opts.borderColor || C.tableBorder),
  width: { size: w, type: WidthType.DXA },
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  verticalAlign: VerticalAlign.CENTER,
  columnSpan: opts.span,
  children: [new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({
      text,
      bold: opts.bold || false,
      size: opts.size || 18,
      font: '맑은 고딕',
      color: opts.color || C.darkGray,
    })],
  })],
});

// ── 문서 정보 표 ──────────────────────────────────────
function metaTable() {
  const lw = 1500, vw = 2703; // label / value width (×2 = 8406)
  const lbl = (t) => cell(t, lw, { fill: C.tableHeader, bold: true, color: C.white, center: true, size: 18 });
  const val = (t, span) => cell(t, vw, { size: 18, span });

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [lw, vw, lw, vw],
    rows: [
      new TableRow({ children: [lbl('문서 번호'), val('KTNET-POL-YYYY-NNN'), lbl('문서 등급'), val('□ 일반  □ 대외비  □ 기밀')] }),
      new TableRow({ children: [lbl('제 정 일'), val('YYYY. MM. DD.'), lbl('개정 번호'), val('v1.0')] }),
      new TableRow({ children: [lbl('작성 부서'), val('AX추진실'), lbl('승  인  자'), val('')] }),
      new TableRow({ children: [lbl('관련 근거'), val('', 3)] }),   // span=3 → 나머지 3열 합병
    ],
  });
}

// ── 개정 이력 표 (줄무늬) ─────────────────────────────
function historyTable() {
  const cols = [900, 1300, 1500, 4706];
  const hdrs = ['버전', '일자', '작성자', '변경 내용'];

  const makeRow = (vals, stripe) => new TableRow({
    children: vals.map((v, i) => cell(v, cols[i], {
      fill: stripe ? C.tableStripe : C.white,
      center: i < 3,
      size: 18,
    })),
  });

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      // 헤더 행
      new TableRow({ tableHeader: true, children: hdrs.map((h, i) =>
        cell(h, cols[i], { fill: C.tableHeader, bold: true, color: C.white, center: true, size: 18 })
      )}),
      makeRow(['v1.0', 'YYYY.MM.DD', '홍길동', '최초 제정'], false),
      makeRow(['',     '',           '',       ''],          true),
      makeRow(['',     '',           '',       ''],          false),
    ],
  });
}

// ── 헤더 ──────────────────────────────────────────────
const pageHeader = new Header({ children: [
  new Paragraph({
    spacing: { before: 0, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.amber, space: 2 } },
    children: [
      new TextRun({ text: 'KTNET 한국무역정보통신', bold: true, size: 18, font: '맑은 고딕', color: C.charcoal }),
      new TextRun({ text: '\t', }),
      new TextRun({ text: '내부 정책 문서 (Policy Document)', size: 17, font: '맑은 고딕', color: C.midGray }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  }),
]});

// ── 푸터 ──────────────────────────────────────────────
const pageFooter = new Footer({ children: [
  new Paragraph({
    spacing: { before: 100, after: 0 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
    children: [
      new TextRun({ text: '본 문서는 KTNET 내부 정책 문서입니다. 허가 없는 외부 배포를 금합니다.', size: 15, font: '맑은 고딕', color: C.midGray }),
      new TextRun({ text: '\t', }),
      new TextRun({ text: 'Page ', size: 15, font: '맑은 고딕', color: C.midGray }),
      new TextRun({ children: [PageNumber.CURRENT], size: 15, font: '맑은 고딕', color: C.amber }),
      new TextRun({ text: ' / ', size: 15, font: '맑은 고딕', color: C.midGray }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, font: '맑은 고딕', color: C.midGray }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  }),
]});

// ── 본문 헬퍼 ──────────────────────────────────────────
const articleTitle = (no, title) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [
    new TextRun({ text: `제${no}조`, bold: true, size: 24, font: '맑은 고딕', color: C.amber }),
    new TextRun({ text: `  (${title})`, bold: true, size: 24, font: '맑은 고딕', color: C.charcoal }),
  ],
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 200, after: 80 },
  children: [new TextRun({ text, bold: true, size: 21, font: '맑은 고딕', color: C.charcoal })],
});

const body = (text) => new Paragraph({
  spacing: { before: 80, after: 80 },
  children: [new TextRun({ text, size: 20, font: '맑은 고딕', color: C.darkGray })],
});

const clause = (no, text) => new Paragraph({
  spacing: { before: 60, after: 60 },
  indent: { left: 400, hanging: 400 },
  children: [
    new TextRun({ text: `${no}  `, bold: true, size: 20, font: '맑은 고딕', color: C.amber }),
    new TextRun({ text, size: 20, font: '맑은 고딕', color: C.darkGray }),
  ],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  spacing: { before: 40, after: 40 },
  children: [new TextRun({ text, size: 20, font: '맑은 고딕', color: C.darkGray })],
});

// ══ 문서 조립 ═════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '–',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 560, hanging: 280 } } },
      }],
    }],
  },
  styles: {
    default: {
      document: { run: { font: '맑은 고딕', size: 20, color: C.darkGray } },
    },
    paragraphStyles: [
      // Heading 1 — 조항 번호 스타일 (포인트 컬러 왼쪽 테두리)
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: '맑은 고딕', color: C.charcoal },
        paragraph: {
          spacing: { before: 360, after: 120 },
          outlineLevel: 0,
          border: { left: { style: BorderStyle.THICK, size: 20, color: C.amber, space: 12 } },
        },
      },
      // Heading 2 — 소항목
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 21, bold: true, font: '맑은 고딕', color: C.charcoal },
        paragraph: {
          spacing: { before: 200, after: 80 },
          outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
        },
      },
      // Heading 3 — 세부항목
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 20, bold: true, font: '맑은 고딕', color: C.midGray },
        paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2 },
      },
      // 강조 본문 (별도 캐릭터 스타일)
      {
        id: 'IntenseQuote', name: 'Intense Quote', basedOn: 'Normal',
        run: { size: 20, italics: true, color: C.amberDark, font: '맑은 고딕' },
        paragraph: {
          spacing: { before: 160, after: 160 },
          indent: { left: 600, right: 600 },
          shading: { fill: C.amberLight, type: ShadingType.CLEAR },
        },
      },
    ],
  },

  sections: [
    // ──────────────────────────────────────────────────
    //  섹션 1: 표지 (헤더/푸터 없음)
    // ──────────────────────────────────────────────────
    {
      properties: {
        page: {
          size:   { width: PAGE.w, height: PAGE.h },
          margin: { top: 2400, bottom: 1800, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      children: [
        // 상단 앰버 굵은 선
        new Paragraph({
          spacing: { before: 0, after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 36, color: C.amber, space: 1 } },
          children: [],
        }),
        gap(600),
        // 기관명
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: 'KTNET  한국무역정보통신', size: 20, font: '맑은 고딕', color: C.midGray, characterSpacing: 60 })],
        }),
        // 문서 유형 배지
        new Paragraph({
          spacing: { before: 0, after: 280 },
          children: [
            new TextRun({ text: ' POLICY DOCUMENT ', bold: true, size: 16, font: '맑은 고딕', color: C.white,
              shading: { fill: C.amber, type: ShadingType.CLEAR } }),
          ],
        }),
        // 메인 제목
        new Paragraph({
          spacing: { before: 0, after: 80 },
          children: [new TextRun({ text: '[정책 문서 제목]', bold: true, size: 52, font: '맑은 고딕', color: C.charcoal })],
        }),
        // 부제목
        new Paragraph({
          spacing: { before: 0, after: 480 },
          children: [new TextRun({ text: '부제목 또는 정책 구분 (예: AI 서비스 이용 정책)', size: 24, font: '맑은 고딕', color: C.midGray })],
        }),
        thickRule(),
        gap(240),
        // 문서 정보 표
        metaTable(),
        gap(500),
        // 하단 기관 서명
        new Paragraph({
          spacing: { before: 0, after: 40 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 8 } },
          children: [new TextRun({ text: '한국무역정보통신 (KTNET)', bold: true, size: 22, font: '맑은 고딕', color: C.charcoal })],
        }),
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: 'AX추진실  |  YYYY년 MM월 DD일', size: 18, font: '맑은 고딕', color: C.midGray })],
        }),
        // 표지 끝 → 다음 페이지
        new Paragraph({ children: [new PageBreak()] }),
        // ── 개정 이력 (표지 다음 페이지) ─────────────
        gap(200),
        new Paragraph({
          spacing: { before: 0, after: 120 },
          children: [new TextRun({ text: '개정 이력', bold: true, size: 21, font: '맑은 고딕', color: C.charcoal })],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 4 } },
        }),
        historyTable(),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ──────────────────────────────────────────────────
    //  섹션 2: 목차 페이지
    // ──────────────────────────────────────────────────
    {
      properties: {
        page: {
          size:   { width: PAGE.w, height: PAGE.h },
          margin: { top: PAGE.mTop, bottom: PAGE.mBottom, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      headers: { default: pageHeader },
      footers: { default: pageFooter },
      children: [
        gap(200),
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: '목  차', bold: true, size: 32, font: '맑은 고딕', color: C.charcoal })],
        }),
        new Paragraph({
          spacing: { before: 0, after: 280 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.amber, space: 2 } },
          children: [new TextRun({ text: 'Table of Contents', size: 18, font: '맑은 고딕', color: C.midGray })],
        }),
        new TableOfContents('목차', { hyperlink: true, headingStyleRange: '1-3' }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ──────────────────────────────────────────────────
    //  섹션 3: 본문
    // ──────────────────────────────────────────────────
    {
      properties: {
        page: {
          size:   { width: PAGE.w, height: PAGE.h },
          margin: { top: PAGE.mTop, bottom: PAGE.mBottom, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      headers: { default: pageHeader },
      footers: { default: pageFooter },
      children: [
        gap(200),

        // ── 제1조 목적 ───────────────────────────────
        articleTitle('1', '목적'),
        body('본 정책은 [정책 대상 범위 및 목적을 기재합니다]. KTNET의 업무 효율성 향상과 정보 보호를 위해 다음과 같이 정합니다.'),
        gap(80),
        noteCallout('본 정책은 관련 법령 및 상위 지침에 우선하지 않으며, 법령 개정 시 자동으로 반영됩니다.'),
        gap(120),

        // ── 제2조 적용 범위 ──────────────────────────
        articleTitle('2', '적용 범위'),
        body('본 정책은 다음에 해당하는 모든 대상에 적용됩니다.'),
        bullet('[적용 대상 1] — 임직원 및 계약직'),
        bullet('[적용 대상 2] — 협력사 직원 (KTNET 시스템 접근자)'),
        bullet('[적용 대상 3] — 외주 개발자 및 파견 인력'),
        gap(80),

        // ── 제3조 용어의 정의 ────────────────────────
        articleTitle('3', '용어의 정의'),
        body('본 정책에서 사용하는 주요 용어의 정의는 다음과 같습니다.'),
        clause('①', '"[용어1]"이란 [정의]를 의미한다.'),
        clause('②', '"[용어2]"이란 [정의]를 의미한다.'),
        clause('③', '"[용어3]"이란 [정의]를 의미한다.'),
        gap(80),

        // ── 제4조 역할 및 책임 ──────────────────────
        articleTitle('4', '역할 및 책임'),
        h2('① 담당 부서 (AX추진실)'),
        bullet('[역할 내용 1] — 정책 수립 및 개정'),
        bullet('[역할 내용 2] — 교육 및 가이드 배포'),
        bullet('[역할 내용 3] — 위반 사항 모니터링'),
        h2('② 이용자'),
        bullet('[이용자 책임 1] — 본 정책 숙지 및 준수'),
        bullet('[이용자 책임 2] — 위반 사항 발견 시 즉시 신고'),
        gap(80),

        // ── 제5조 세부 정책 내용 ─────────────────────
        articleTitle('5', '세부 정책 내용'),
        h2('5.1 [소항목 제목]'),
        body('[세부 정책 본문 내용을 작성합니다. 항목별로 번호를 붙여 명확하게 기재하세요.]'),
        clause('①', '[세부 내용 1]'),
        clause('②', '[세부 내용 2]'),
        h2('5.2 [소항목 제목]'),
        body('[소항목 내용을 작성합니다.]'),
        gap(80),
        warnCallout('아래 사항은 정책 위반으로 간주되며, 위반 시 제6조에 따른 조치가 취해집니다.'),
        gap(80),

        // ── 제6조 위반 시 조치 ──────────────────────
        articleTitle('6', '위반 시 조치'),
        body('본 정책을 위반할 경우 다음의 절차에 따라 조치합니다.'),
        clause('①', '[경미한 위반] — 서면 경고 및 재교육 이수'),
        clause('②', '[중대한 위반] — 관련 규정에 따른 징계 처분'),
        clause('③', '[법령 위반] — 관계 기관 신고 및 법적 조치'),
        gap(80),

        // ── 개정 이력 표 ─────────────────────────────
        articleTitle('7', '시행일 및 개정 이력'),
        body('본 정책은 YYYY년 MM월 DD일부터 시행합니다. 이전에 제정된 동일 목적의 정책은 본 정책으로 대체됩니다.'),
        gap(240),

        // ── 서명란 ───────────────────────────────────
        rule(C.ruleLight, 4),
        gap(120),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 60 },
          children: [new TextRun({ text: 'YYYY년 MM월 DD일', size: 18, font: '맑은 고딕', color: C.midGray })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 40 },
          children: [new TextRun({ text: '한국무역정보통신 (KTNET)', bold: true, size: 22, font: '맑은 고딕', color: C.charcoal })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: 'AX추진실장  ___________________  (인)', size: 20, font: '맑은 고딕', color: C.darkGray })],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/KTNET_정책문서_표준양식_v3.docx', buf);
  console.log('✅ KTNET_정책문서_표준양식_v3.docx 생성 완료');
});
