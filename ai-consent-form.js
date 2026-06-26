const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, HeadingLevel, TabStopType,
  TabStopPosition, PageBreak,
} = require('docx');
const fs = require('fs');

// ══════════════════════════════════════════════════════
//  AI 도구 활용 동의서
//  · 내용/구성: AI_활용_동의서_v0_1_임직원용_5.docx 원문 유지
//  · 서식/표/디자인/색상: ktnet-policy-template.js (모던 다크그레이 + 앰버) 참고
//  · 표지·목차 등 템플릿 고유 구성은 적용하지 않음
// ══════════════════════════════════════════════════════

const C = {
  charcoal:    '2D2D2D',
  darkGray:    '404040',
  midGray:     '6B6B6B',
  lightGray:   'F2F2F2',
  white:       'FFFFFF',
  amber:       'E8A020',
  amberDark:   'C47E00',
  amberLight:  'FDF3DC',
  ruleDark:    '2D2D2D',
  ruleLight:   'DDDDDD',
  tableBorder: 'CCCCCC',
  tableHeader: '2D2D2D',
  tableStripe: 'F7F7F7',
  infoBox:     'FDF3DC',
  warnBox:     'FFF0F0',
  consentBox:  'FBF6EA',
};

const PAGE = {
  w: 11906, h: 16838,
  mTop: 1800, mBottom: 1700, mLeft: 1800, mRight: 1700,
  get cw() { return this.w - this.mLeft - this.mRight; }, // 8406
};

const FONT = '맑은 고딕';

// ── 기본 헬퍼 ─────────────────────────────────────────
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

// ── 강조 블록 (단락 테두리 텍스트 박스) ────────────────
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
      ...(label ? [new TextRun({ text: `${label}  `, bold: true, size: 19, font: FONT, color: C.amberDark })] : []),
      new TextRun({ text, size: 19, font: FONT, color: C.darkGray }),
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
  margins: { top: 90, bottom: 90, left: 130, right: 130 },
  verticalAlign: VerticalAlign.CENTER,
  columnSpan: opts.span,
  children: [new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({
      text,
      bold: opts.bold || false,
      size: opts.size || 17,
      font: FONT,
      color: opts.color || C.darkGray,
    })],
  })],
});

// ── 데이터 표: 다크 헤더 + 줄무늬 본문 ─────────────────
// headers: [..], rows: [[..],..], cols: [..] (DXA), centerCols: 가운데 정렬할 열 인덱스 배열
function dataTable(headers, rows, cols, opts = {}) {
  const centerCols = opts.centerCols || [];
  const boldFirst = opts.boldFirst !== false;

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cell(h, cols[i], { fill: C.tableHeader, bold: true, color: C.white, center: true, size: 17 })),
  });

  const bodyRows = rows.map((vals, r) => new TableRow({
    children: vals.map((v, i) => cell(v, cols[i], {
      fill: r % 2 === 1 ? C.tableStripe : C.white,
      center: centerCols.includes(i),
      bold: boldFirst && i === 0,
      color: boldFirst && i === 0 ? C.charcoal : C.darkGray,
      size: 17,
    })),
  }));

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: cols,
    rows: [headerRow, ...bodyRows],
  });
}

// ── 헤더 ──────────────────────────────────────────────
const pageHeader = new Header({ children: [
  new Paragraph({
    spacing: { before: 0, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.amber, space: 2 } },
    children: [
      new TextRun({ text: 'KTNET 한국무역정보통신', bold: true, size: 18, font: FONT, color: C.charcoal }),
      new TextRun({ text: '\t' }),
      new TextRun({ text: 'AI 도구 활용 동의서 (Consent Form)', size: 17, font: FONT, color: C.midGray }),
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
      new TextRun({ text: '본 문서는 개인정보가 포함된 KTNET 내부 문서입니다. 허가 없는 외부 배포를 금합니다.', size: 15, font: FONT, color: C.midGray }),
      new TextRun({ text: '\t' }),
      new TextRun({ text: 'Page ', size: 15, font: FONT, color: C.midGray }),
      new TextRun({ children: [PageNumber.CURRENT], size: 15, font: FONT, color: C.amber }),
      new TextRun({ text: ' / ', size: 15, font: FONT, color: C.midGray }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, font: FONT, color: C.midGray }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  }),
]});

// ── 본문 헬퍼 ──────────────────────────────────────────
const articleTitle = (no, title) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [
    new TextRun({ text: `제${no}조`, bold: true, size: 23, font: FONT, color: C.amber }),
    new TextRun({ text: `  ${title}`, bold: true, size: 23, font: FONT, color: C.charcoal }),
  ],
});

const sectionTitle = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [new TextRun({ text, bold: true, size: 23, font: FONT, color: C.charcoal })],
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 220, after: 80 },
  children: [new TextRun({ text, bold: true, size: 20, font: FONT, color: C.charcoal })],
});

const body = (text) => new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.JUSTIFIED,
  children: [new TextRun({ text, size: 19, font: FONT, color: C.darkGray })],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  spacing: { before: 30, after: 30 },
  children: [new TextRun({ text, size: 19, font: FONT, color: C.darkGray })],
});

// ※ 보조 안내 문구
const note = (text) => new Paragraph({
  spacing: { before: 30, after: 30 },
  indent: { left: 200, hanging: 200 },
  children: [new TextRun({ text, size: 16, font: FONT, color: C.midGray })],
});

// ── 안내(읽어주세요) 박스 ─────────────────────────────
function noticeBox(title, lines) {
  const paras = [];
  paras.push(new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: title, bold: true, size: 19, font: FONT, color: C.amberDark })],
  }));
  lines.forEach((ln, i) => paras.push(new Paragraph({
    spacing: { before: 40, after: i === lines.length - 1 ? 0 : 40 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text: ln, size: 18, font: FONT, color: C.darkGray })],
  })));

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [PAGE.cw],
    rows: [new TableRow({ children: [new TableCell({
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 4,  color: C.amber },
        bottom: { style: BorderStyle.SINGLE, size: 4,  color: C.amber },
        left:   { style: BorderStyle.THICK,  size: 28, color: C.amber },
        right:  { style: BorderStyle.SINGLE, size: 4,  color: C.amber },
      },
      width: { size: PAGE.cw, type: WidthType.DXA },
      shading: { fill: C.amberLight, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 220, right: 200 },
      children: paras,
    })] })],
  });
}

// ── 동의 여부 박스 (체크) ─────────────────────────────
function checkRuns(options) {
  const parts = options.split('\n');
  const runs = [];
  parts.forEach((p, i) => {
    if (i > 0) runs.push(new TextRun({ text: p, bold: true, size: 18, font: FONT, color: C.amberDark, break: 1 }));
    else runs.push(new TextRun({ text: p, bold: true, size: 18, font: FONT, color: C.amberDark }));
  });
  return runs;
}

function consentRow(statement, options) {
  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [6006, 2400],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 6, color: C.amberDark },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: C.amberDark },
          left:   { style: BorderStyle.THICK,  size: 22, color: C.amberDark },
          right:  { style: BorderStyle.SINGLE, size: 4, color: C.amberDark },
        },
        width: { size: 6006, type: WidthType.DXA },
        shading: { fill: C.consentBox, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: statement, size: 18, font: FONT, color: C.charcoal })],
        })],
      }),
      new TableCell({
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 6, color: C.amberDark },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: C.amberDark },
          left:   { style: BorderStyle.SINGLE, size: 4, color: C.amberDark },
          right:  { style: BorderStyle.SINGLE, size: 6, color: C.amberDark },
        },
        width: { size: 2400, type: WidthType.DXA },
        shading: { fill: C.white, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 0 },
          children: checkRuns(options),
        })],
      }),
    ] })],
  });
}

// ── 서명란 표 ─────────────────────────────────────────
function signatureTable() {
  const lw = 1500, vw = 2703;
  const lbl = (t) => cell(t, lw, { fill: C.tableHeader, bold: true, color: C.white, center: true, size: 17 });
  const val = (t, span) => cell(t || ' ', span ? PAGE.cw - lw : vw, { size: 17, span });

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [lw, vw, lw, vw],
    rows: [
      new TableRow({ children: [lbl('성   명'), val(''), lbl('사   번'), val('')] }),
      new TableRow({ children: [lbl('부 서 명'), val(''), lbl('직급/직책'), val('')] }),
      new TableRow({ children: [lbl('업무용 이메일'), val('', 3)] }),
      new TableRow({ children: [lbl('동의 일자'), val('        년       월       일'), lbl('서   명'), val('(서명 또는 날인)')] }),
    ],
  });
}

// ── 접수처 박스 ───────────────────────────────────────
function receiverBox() {
  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [PAGE.cw],
    rows: [new TableRow({ children: [new TableCell({
      borders: bord(C.amberDark),
      width: { size: PAGE.cw, type: WidthType.DXA },
      shading: { fill: C.charcoal, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 180, right: 180 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({ text: '㈜한국무역정보통신 귀중', bold: true, size: 19, font: FONT, color: C.white }),
          new TextRun({ text: '      수령 부서: AX추진실   |   보관 담당: AX추진실 및 CPO실   |   이의사항 문의: cpo@ktnet.com', size: 16, font: FONT, color: C.amberLight }),
        ],
      })],
    })] })],
  });
}

// ══ 문서 조립 ═════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '●',
        alignment: AlignmentType.LEFT,
        style: { run: { color: C.amber }, paragraph: { indent: { left: 460, hanging: 240 } } },
      }],
    }],
  },
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: C.darkGray } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 23, bold: true, font: FONT, color: C.charcoal },
        paragraph: {
          spacing: { before: 360, after: 120 },
          outlineLevel: 0,
          border: { left: { style: BorderStyle.THICK, size: 20, color: C.amber, space: 12 } },
        },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 20, bold: true, font: FONT, color: C.charcoal },
        paragraph: {
          spacing: { before: 220, after: 80 },
          outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
        },
      },
    ],
  },

  sections: [{
    properties: {
      page: {
        size:   { width: PAGE.w, height: PAGE.h },
        margin: { top: PAGE.mTop, bottom: PAGE.mBottom, left: PAGE.mLeft, right: PAGE.mRight },
      },
    },
    headers: { default: pageHeader },
    footers: { default: pageFooter },
    children: [
      // ── 제목 영역 (표지 아님, 본문 상단 타이틀) ──────
      thickRule(),
      gap(120),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: ' CONSENT FORM ', bold: true, size: 15, font: FONT, color: C.white,
          shading: { fill: C.amber, type: ShadingType.CLEAR } })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: 'AI 도구 활용 동의서', bold: true, size: 40, font: FONT, color: C.charcoal })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: '인공지능(AI) 도구 업무 활용 및 개인정보 처리에 관한 임직원 동의서', size: 20, font: FONT, color: C.midGray })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '「인공지능 활용 및 관리 요령」 제34조 / AI-GOV-003-001 5.2장, 5.3장 근거', size: 16, font: FONT, color: C.midGray })],
      }),
      thickRule(),
      gap(160),

      // ── 안내 박스 ────────────────────────────────────
      noticeBox('[동의서 작성 전 반드시 읽어 주시기 바랍니다]', [
        "한국무역정보통신(이하 '회사')은 임직원이 업무 목적으로 생성형 AI 도구를 사용함에 따라 발생하는 개인정보 처리 및 사용 모니터링에 대한 사전 동의를 확보하기 위해 본 동의서를 수집합니다.",
        '본 동의서는 AI 게이트웨이·DLP 기술 통제 시스템이 구축되기 전까지 임시 운영되는 관리적 통제 수단으로, AI 도구 활용 정책(AI-GOV-003-001) 5.2장 및 5.3장의 운영 전제 조건입니다.',
        '동의하지 않으실 경우 AI 도구를 업무에 활용할 수 없습니다. 이의사항은 AX추진실로 문의하시기 바랍니다.',
      ]),
      gap(120),

      // ── 제1조 ────────────────────────────────────────
      articleTitle('1', '개인정보 수집·이용에 관한 동의'),
      body('회사는 아래와 같이 임직원의 개인정보를 수집·이용합니다. 「개인정보 보호법」 제15조 제1항 제1호(정보주체의 동의) 및 「KTNET 개인정보보호세칙」 제15조에 따라 다음 사항에 동의를 구합니다.'),
      h2('가. 수집하는 개인정보 항목'),
      dataTable(
        ['구분', '항목', '비고'],
        [
          ['필수 항목', '성명, 사번, 부서명, 직급/직책, 업무용 이메일 주소', 'AI 도구 계정 관리 및 사용자 식별 목적'],
          ['사용 관련 정보', 'AI 도구 접속 일시, 도구명, 사용 목적(업무 분류), 망 구분(업무망/인터넷망)', '모니터링 및 감사 목적 (요령 제40조)'],
          ['동의 관리 정보', '동의서 제출일, 동의 여부, 서명', '동의 이력 관리 목적'],
        ],
        [1500, 4406, 2500],
      ),
      h2('나. 수집·이용 목적'),
      bullet('AI 도구 업무 활용 승인 및 계정 관리'),
      bullet('AI 도구 사용 현황 모니터링 및 정책 준수 여부 확인'),
      bullet('법인카드 전표 기반 AI 도구 결제 확인 및 분기 감사 대응'),
      bullet('개인정보 침해 사고 발생 시 원인 조사 및 사고 대응'),
      bullet('AI 거버넌스 KPI 산출 및 경영 보고'),
      bullet('「인공지능 활용 및 관리 요령」 제48조에 따른 반기 자체 점검 지원'),
      h2('다. 보유 및 이용 기간'),
      dataTable(
        ['구분', '보유 기간', '근거'],
        [
          ['동의서 원본', '재직 기간 + 퇴직 후 3년', '개인정보보호세칙 제17조 / 내부 관리계획'],
          ['AI 도구 접속 로그', '최소 2년 (사고 조사 중에는 종결 시까지 연장)', '요령 제40조 · 개인정보보호세칙 제49조'],
          ['감사 기록', '5년', 'AI 거버넌스 위원회 의결 사항 보관 기준'],
        ],
        [1800, 3306, 3300],
      ),
      note('※ 보유 기간 종료 후에는 지체 없이 파기합니다. 다만 관련 법령에 의해 보존이 필요한 경우에는 해당 기간까지 보관 후 파기합니다.'),
      h2('라. 제3자 제공 및 처리위탁'),
      body('수집된 개인정보는 원칙적으로 제3자에게 제공하지 않습니다. AI 도구 서비스 제공사(Anthropic, Inc. 등 승인 목록 등재 업체)에 대한 처리위탁은 「KTNET 개인정보보호세칙」 제34조 및 「인공지능 활용 및 관리 요령」 제35조에 따라 Zero Data Retention(최종 데이터 미저장·미학습) 또는 이에 준하는 계약 조건이 적용됩니다. 수탁 업체 현황은 회사 개인정보 처리방침(사내 게시)에 공개됩니다.'),
      gap(120),
      consentRow('[제1조 동의 여부]  본인은 위 개인정보 수집·이용에 관한 사항을 충분히 읽고 이해하였으며 이에 동의합니다.', '□ 동의함\n□ 동의하지 않음'),
      gap(120),

      // ── 제2조 ────────────────────────────────────────
      articleTitle('2', '개인정보 국외 이전에 관한 동의'),
      body('Tier A 등급으로 승인된 AI 도구(예: Claude Team, Cursor Team, JetBrains AI Business 등)는 서비스 제공사의 서버가 해외에 위치하여 개인정보가 국외로 이전될 수 있습니다. 「개인정보 보호법」 제28조의 8 및 「KTNET 개인정보보호세칙」에 따라 아래와 같이 고지하고 동의를 구합니다.'),
      gap(60),
      dataTable(
        ['항목', '내용', '비고'],
        [
          ['이전받는 자', 'Anthropic, Inc. / JetBrains s.r.o. / Anysphere Inc. (Cursor) 등 승인 AI 도구 제공사', 'AI 도구 승인 목록(v1.6 이상) 등재 업체 한정'],
          ['이전 국가', '미국, EU 등 (도구별 Data Residency 정책 참조)', '4.3장 부합성 검증 ⑥항 확인 완료'],
          ['이전 항목', '업무용 이메일, 사용자 식별 정보, 입력 프롬프트 메타데이터', 'ZDR 적용 시 입력 데이터 저장·학습 금지'],
          ['이전 목적', 'AI 서비스 제공 및 계정 관리', '목적 외 이용 계약으로 금지'],
          ['보유·이용 기간', '서비스 계약 기간 내. 계약 종료 시 파기(ZDR 조항 적용)', '요령 제35조 ²-4 준수'],
          ['보호 수단', 'DPA 체결, Zero Data Retention, TLS 암호화, SOC 2 Type II / ISO 27001 인증 보유', 'AI 도구 활용 정책 4.3장 참조'],
        ],
        [1600, 4406, 2400],
      ),
      note('※ 국외 이전에 동의하지 않을 경우 해당 AI 도구를 업무에 사용할 수 없습니다.'),
      note("※ 승인 AI 도구 목록 및 각 도구의 국외 이전 상세 현황은 AX추진실이 관리하는 '관리 AI 도구 승인 목록'(사내 공유)에서 확인 가능합니다."),
      gap(120),
      consentRow('[제2조 동의 여부]  본인은 위 개인정보 국외 이전에 관한 사항을 충분히 읽고 이해하였으며 이에 동의합니다.', '□ 동의함\n□ 동의하지 않음'),
      gap(120),

      // ── 제3조 ────────────────────────────────────────
      articleTitle('3', 'AI 도구 사용 모니터링 및 통제에 관한 동의'),
      body('회사는 「인공지능 활용 및 관리 요령」 제40조 및 「AI 도구 활용 정책」 6.2장에 따라 임직원의 AI 도구 사용 현황을 모니터링합니다. 이는 정보보호 및 개인정보 침해 방지를 위한 업무상 목적에 해당하며, 아래 범위 내에서 수행됩니다.'),
      h2('가. 모니터링 범위 및 방법'),
      dataTable(
        ['모니터링 항목', '방법', '주기', '담당 부서'],
        [
          ['도구별 사용량·비용', 'Admin Console 조회', '분기', 'AX추진실'],
          ['법인카드 AI 도구 결제', '법인카드 전표 점검 (허용 화이트리스트 외 결제 선별)', '분기', '감사실'],
          ['미승인 도구 접근 시도', '내부 신고 + 향후 AI 게이트웨이 도입 시 로그 기반 전환', '월 / 분기', '인프라운영실'],
          ['Level 2 이상 데이터 입력 의심', '내부 신고 + 향후 DLP 도입 시 이벤트 로그 전환', '월 / 분기', '인프라운영실'],
          ['정책 위반 적발 현황', '감사실 적발·제재 대장', '분기', '감사실'],
        ],
        [2100, 3306, 1200, 1800],
        { centerCols: [2, 3] },
      ),
      h2('나. 모니터링의 목적 및 한계'),
      bullet('모니터링은 정보보호 및 개인정보 보호 목적으로만 활용되며, 인사 평가에 직접 활용되지 않습니다.'),
      bullet('현재 AI 게이트웨이·DLP 미구축 상태로, 개인 카드 결제·무료 플랜 사용은 기술적으로 완전 통제되지 않는 한계가 있습니다.'),
      bullet('향후 AI 게이트웨이·DLP 시스템 구축 시 본 동의서는 갱신됩니다.'),
      gap(120),
      consentRow('[제3조 동의 여부]  본인은 위 모니터링 및 사용 통제에 관한 사항을 충분히 읽고 이해하였으며 이에 동의합니다.', '□ 동의함\n□ 동의하지 않음'),
      gap(120),

      // ── 제4조 ────────────────────────────────────────
      articleTitle('4', 'AI 도구 활용 준수 의무 확인'),
      body('본인은 AI 도구를 업무에 활용함에 있어 다음 사항을 준수할 것을 확인합니다.'),
      h2('가. 데이터 입력 수준(Level) 준수'),
      dataTable(
        ['Level', '데이터 유형', '입력 가능 AI 도구 범위'],
        [
          ['Level 1 (공개)', '공개 가능한 정보, 일반 지식 조회', 'Tier A / B / C 모두 허용 (망별 매트릭스 준수)'],
          ['Level 2 (내부)', '사내 업무 정보, 미공개 프로젝트, 회원사 일반 정보', 'Tier A 도구만 허용 (업무망·인터넷망, 본인확인망 금지)'],
          ['Level 3 (기밀)', '개인식별정보, 고객 개인정보, 금융·계좌 정보, 직원 민감정보', '원칙적 금지. AI 거버넌스 위원회 심의·승인 후 온프레미스 도구에 한해 조건부 허용'],
        ],
        [1700, 2706, 4000],
      ),
      h2('나. 금지 행위 준수'),
      bullet('미승인 AI 도구 업무 사용 금지 (「AI 도구 활용 정책」 8.2장)'),
      bullet('개인 계정(Free/Pro)으로 Level 2 이상 데이터 입력 금지'),
      bullet('본인확인망에서 외부 AI 도구 접근 일체 금지'),
      bullet('AI 생성 결과물의 무검증 업무 활용 금지 (요령 제25조, 제26조)'),
      bullet('타인 계정 공유 또는 대리 사용 금지 (요령 제39조 ² 2호)'),
      bullet('AI 도구 사용 중 개인정보 침해 의심 상황 발견 시 즉시 AX추진실 및 CPO에 신고'),
      h2('다. 교육 이수 의무'),
      bullet('신규 입사자: 입사 후 30일 이내 AI 활용 교육 이수 (요령 제47조 ¹)'),
      bullet('전 임직원: 연 1회 이상 AI 활용 교육 이수'),
      bullet('협력업체 상주 인력: 업무 수행 전 준수 서약서 별도 제출'),
      gap(120),
      consentRow('[제4조 확인]  본인은 위 AI 도구 활용 준수 의무를 충분히 이해하였으며 이를 준수할 것을 확인합니다.', '□ 확인함'),
      gap(120),

      // ── 제5조 ────────────────────────────────────────
      articleTitle('5', '정보주체의 권리 안내'),
      body('임직원(정보주체)은 「개인정보 보호법」 제35조~제37조의 2 및 「KTNET 개인정보보호세칙」에 따라 다음 권리를 행사할 수 있습니다.'),
      gap(60),
      dataTable(
        ['권리', '내용', '행사 방법'],
        [
          ['열람 요구권', '본인에 관한 개인정보 처리 현황 및 내용 확인', 'CPO실 서면 요청 cpo@ktnet.com'],
          ['정정·삭제 요구권', '잘못된 정보의 정정 또는 불필요한 정보의 삭제 요청 (법령에 의해 수집된 정보 제외)', 'CPO실 서면 요청'],
          ['처리 정지 요구권', '개인정보 처리 정지 요청 (단, 업무 수행상 필수 항목은 처리 정지 불가)', 'CPO실 서면 요청'],
          ['동의 철회권', '수집·이용 동의 철회 가능. 단 철회 시 AI 도구 업무 활용이 제한될 수 있음', 'AX추진실 서면 통보'],
          ['자동화 결정 거부권', '완전 자동화된 AI 의사결정으로 인한 중대한 영향에 대해 설명 요구 및 거부 가능 (법 제37조의 2)', 'CPO실 서면 요청'],
        ],
        [1700, 3906, 2800],
      ),
      note('※ 권리 행사는 개인정보 보호책임자(CPO) 앞으로 서면·이메일로 신청하며, 접수 후 10일 이내 처리합니다.'),
      note('※ 개인정보 침해 관련 민원은 개인정보 분쟁조정위원회(www.kopico.go.kr) 또는 개인정보 침해신고센터(privacy.kisa.or.kr, 국번 없이 118)에 신고하실 수 있습니다.'),
      gap(160),

      // ── 최종 확인 및 서명 ────────────────────────────
      sectionTitle('동의자 서명 — 최종 확인 및 서명'),
      body('본인은 위 모든 사항(제1조~제5조)을 충분히 읽고 이해하였으며, 자유로운 의사에 따라 동의합니다.'),
      gap(120),
      signatureTable(),
      gap(160),
      receiverBox(),

      // ── [참고] 관련 규정 및 근거 ─────────────────────
      sectionTitle('[참고] 관련 규정 및 근거'),
      bullet('개인정보 보호법 제15조(개인정보의 수집·이용), 제28조의 8(개인정보의 국외 이전), 제35조~제37조의 2(정보주체 권리)'),
      bullet('인공지능 활용 및 관리 요령 v1.0-r6 제34조(개인정보 보호), 제35조(외부 AI 서비스 처리위탁), 제39조(접근통제), 제40조(로그 및 모니터링), 제47조(교육)'),
      bullet('KTNET 개인정보보호세칙 제15조(수집·이용), 제17조(보유기간), 제34조(처리위탁), 제49조(접속기록 보관)'),
      bullet('KTNET 개인정보 내부 관리계획 제1조(목적), 제15조(접근권한관리), 제17조(암호화), 제18조(접속기록), 제24조(사고대응), 제25조(위험관리)'),
      bullet('AI 도구 활용 정책 AI-GOV-003-001 v2.3 5.2장(업무망 운영), 5.3장(인터넷망 운영), 6.2장(모니터링 지표), 8.2장(금지사항)'),
      bullet('생성형 인공지능(AI) 개발·활용을 위한 개인정보 처리 안내서 (개인정보보호위원회, 2025.8.)'),

      // ── 붙임: Tier 안내 ──────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      sectionTitle('붙임  AI 도구 Tier(보안 등급) 안내'),
      body('본 붙임은 동의서 제2조와 제4조에서 언급하는 AI 도구 Tier(보안 등급)의 의미와 분류 기준을 설명합니다.'),
      note('근거: 인공지능 활용 및 관리 요령 v1.0-r6 제11조 / AI 도구 활용 정책 AI-GOV-003-001 v2.3 제2장'),
      h2('1.  Tier란 무엇인가?'),
      body("Tier는 KTNET이 AI 도구의 보안 강도 및 계약 요건 충족 수준에 따라 부여하는 도구별 보안 등급입니다. 동일한 AI 서비스라도 플랜(무료/유료/기업용)에 따라 Tier가 달라지며, Tier가 높을수록 회사가 제공사와 체결한 보안 청약 수준이 높고 개인정보 보호 장치가 강화됩니다. 본 동의서에서 언급하는 ‘Tier A 도구’ 역시 이 기준으로 승인된 기업용 AI 도구를 의미합니다."),
      h2('2.  Tier 분류 기준 및 각 Tier의 의미'),
      body('Tier 분류는 요령 제11조의 등급 체계와 정합하며, AI 도구 활용 정책 2장에서 다음과 같이 정의합니다.'),
      gap(40),
      dataTable(
        ['Tier', '등급명', '해당 플랜 예시', 'DPA 체결', '학습 OFF', '요령 등급', '허용 데이터'],
        [
          ['A', '허용 (Approved)', 'Claude Team/Enterprise, Cursor Business, JetBrains AI Business 등 기업용 계약', '✓ 체결', '✓ 기본적용', '등급 2', 'Level 1(공개)부터 Level 2(내부) 허용'],
          ['B', '제한적 허용 (Restricted)', 'Claude Pro, ChatGPT Plus 등 개인 명의 유료 플랜', '△ 미체결', '△ 선택적용', '등급 1', 'Level 1(공개)만 허용 + 자기신고 필수'],
          ['C', '주의 (Caution)', 'Claude Free, ChatGPT Free 등 무료 플랜', '✗ 미체결', '✗ 미지원', '등급 1', 'Level 1(공개)만 허용, 가이드라인 준수 필수'],
          ['S', '사내 전용 (On-Premise)', 'Ollama, 사내 LLM 등 회사 폐쇄망/전용 클라우드 배포', '해당 없음', '해당 없음', '등급 3', '망 구분 문제없음. Level 1부터 Level 2 허용'],
        ],
        [600, 1300, 2106, 900, 900, 900, 1700],
        { centerCols: [0, 3, 4, 5] },
      ),
      note('※ DPA(Data Processing Addendum): 개인정보 처리 위탁 계약 부속서. 회사 명의로 체결하여 입력 데이터 학습에 사용 금지를 법적으로 보장하는 계약 조항입니다.'),
      note('※ 학습 OFF: AI 서비스 제공사가 사용자 입력 데이터를 모델 학습에 사용하지 않도록 하는 설정. Tier A는 계약상 기본 보장(ZDR), Tier B는 선택적 사용자 설정.'),
      note('※ Tier S는 외부 클라우드 AI 도구가 아니며 회사 내부 망에 배포된 도구로 본 문서의 동의 대상은 아닙니다.'),
    ],
  }],
});

const OUT = process.env.OUT || '/workspace/AI_활용_동의서_v0_1_임직원용_5_KTNET서식.docx';
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('✅ 생성 완료:', OUT);
});
