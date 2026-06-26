const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  ShadingType,
  TabStopPosition,
  TabStopType,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} = require('docx');
const fs = require('fs');
const path = require('path');

// KTNET 사내 문서 표준 서식 기반 AI 활용 가이드북 생성기
// - 원본: AI_활용_가이드북_v0_1.docx
// - 서식: ktnet-policy-template.js의 차콜/앰버 테마, A4, 맑은 고딕, TOC, 헤더/푸터, 표 스타일

const C = {
  charcoal: '2D2D2D',
  darkGray: '404040',
  midGray: '6B6B6B',
  lightGray: 'F2F2F2',
  white: 'FFFFFF',
  amber: 'E8A020',
  amberDark: 'C47E00',
  amberLight: 'FDF3DC',
  ruleDark: '2D2D2D',
  ruleLight: 'DDDDDD',
  tableBorder: 'CCCCCC',
  tableHeader: '2D2D2D',
  tableStripe: 'F7F7F7',
  infoBox: 'FDF3DC',
  warnBox: 'FFF0F0',
};

const PAGE = {
  w: 11906,
  h: 16838,
  mTop: 1800,
  mBottom: 1700,
  mLeft: 1800,
  mRight: 1700,
  get cw() {
    return this.w - this.mLeft - this.mRight;
  },
};

const FONT = '맑은 고딕';

const gap = (n = 120) =>
  new Paragraph({ spacing: { before: n, after: n }, children: [] });

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const thickRule = () =>
  new Paragraph({
    spacing: { before: 0, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 20, color: C.amber, space: 1 } },
    children: [],
  });

const rule = (color = C.ruleLight, size = 4) =>
  new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 1 } },
    children: [],
  });

const run = (text, opts = {}) =>
  new TextRun({
    text,
    bold: opts.bold,
    italics: opts.italics,
    size: opts.size || 20,
    font: FONT,
    color: opts.color || C.darkGray,
    characterSpacing: opts.characterSpacing,
    shading: opts.shading,
  });

const paragraph = (children, opts = {}) =>
  new Paragraph({
    heading: opts.heading,
    alignment: opts.alignment,
    spacing: opts.spacing || { before: 80, after: 80 },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
    numbering: opts.numbering,
    tabStops: opts.tabStops,
    children,
  });

const body = (text) => paragraph([run(text)]);

const noteCallout = (text) =>
  paragraph(
    [
      run('참고  ', { bold: true, size: 19, color: C.amberDark }),
      run(text, { size: 19 }),
    ],
    {
      spacing: { before: 160, after: 160 },
      indent: { left: 200, right: 200 },
      shading: { fill: C.infoBox, type: ShadingType.CLEAR },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: C.amber, space: 4 },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: C.amber, space: 4 },
        left: { style: BorderStyle.THICK, size: 24, color: C.amber, space: 8 },
        right: { style: BorderStyle.SINGLE, size: 4, color: C.amber, space: 4 },
      },
    },
  );

const warnCallout = (text) =>
  paragraph(
    [
      run('주의  ', { bold: true, size: 19, color: 'B00020' }),
      run(text, { size: 19 }),
    ],
    {
      spacing: { before: 160, after: 160 },
      indent: { left: 200, right: 200 },
      shading: { fill: C.warnBox, type: ShadingType.CLEAR },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: 'CC3333', space: 4 },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CC3333', space: 4 },
        left: { style: BorderStyle.THICK, size: 24, color: 'CC3333', space: 8 },
        right: { style: BorderStyle.SINGLE, size: 4, color: 'CC3333', space: 4 },
      },
    },
  );

const bord = (color = C.tableBorder) => ({
  top: { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left: { style: BorderStyle.SINGLE, size: 1, color },
  right: { style: BorderStyle.SINGLE, size: 1, color },
});

function cell(text, w, opts = {}) {
  const textRuns = String(text ?? '')
    .split('\n')
    .flatMap((line, index) => [
      ...(index === 0 ? [] : [new TextRun({ break: 1 })]),
      run(line, {
        bold: opts.bold || false,
        size: opts.size || 17,
        color: opts.color || C.darkGray,
      }),
    ]);

  return new TableCell({
    borders: bord(opts.borderColor || C.tableBorder),
    width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span,
    children: [
      new Paragraph({
        alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { before: 20, after: 20 },
        children: textRuns,
      }),
    ],
  });
}

function makeTable(rows, options = {}) {
  const colCount = Math.max(...rows.map((r) => r.length));
  const widths = options.widths || Array.from({ length: colCount }, () => Math.floor(PAGE.cw / colCount));

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(
      (row, rowIndex) =>
        new TableRow({
          tableHeader: rowIndex === 0,
          children: row.map((value, colIndex) =>
            cell(value, widths[colIndex] || widths[widths.length - 1], {
              fill: rowIndex === 0 ? C.tableHeader : rowIndex % 2 === 0 ? C.tableStripe : C.white,
              bold: rowIndex === 0,
              color: rowIndex === 0 ? C.white : C.darkGray,
              center: rowIndex === 0 || options.centerCols?.includes(colIndex),
            }),
          ),
        }),
    ),
  });
}

function metaTable() {
  const lw = 1500;
  const vw = 2703;
  const lbl = (t) => cell(t, lw, { fill: C.tableHeader, bold: true, color: C.white, center: true });
  const val = (t, span) => cell(t, vw, { span });

  return new Table({
    width: { size: PAGE.cw, type: WidthType.DXA },
    columnWidths: [lw, vw, lw, vw],
    rows: [
      new TableRow({ children: [lbl('문서 번호'), val('AI-GUIDEBOOK-001'), lbl('문서 등급'), val('일반')] }),
      new TableRow({ children: [lbl('제 정 일'), val('2026. 06.'), lbl('개정 번호'), val('v0.1')] }),
      new TableRow({ children: [lbl('작성 부서'), val('AX추진실'), lbl('적용 대상'), val('KTNET 전 임직원')] }),
      new TableRow({ children: [lbl('관련 문서'), val('인공지능 활용 및 관리 요령 v1.0\nAI 도구 활용 정책 (AI-GOV-003-001)', 3)] }),
      new TableRow({ children: [lbl('문의처'), val('kiwhan@ktnet.co.kr / 내선 x2759 / smartflow 메신저(김기환)', 3)] }),
    ],
  });
}

function historyTable() {
  return makeTable(
    [
      ['버전', '일자', '작성자', '변경 내용'],
      ['v0.1', '2026.06', 'AX추진실', '초안 작성 및 KTNET 표준 문서 양식 적용'],
      ['', '', '', ''],
    ],
    { widths: [900, 1300, 1500, 4706], centerCols: [0, 1, 2] },
  );
}

const pageHeader = new Header({
  children: [
    new Paragraph({
      spacing: { before: 0, after: 100 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.amber, space: 2 } },
      children: [
        run('KTNET 한국무역정보통신', { bold: true, size: 18, color: C.charcoal }),
        new TextRun({ text: '\t' }),
        run('AI 활용 가이드북 (AI Utilization Guidebook)', { size: 17, color: C.midGray }),
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    }),
  ],
});

const pageFooter = new Footer({
  children: [
    new Paragraph({
      spacing: { before: 100, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
      children: [
        run('본 문서는 KTNET 내부 가이드 문서입니다. 허가 없는 외부 배포를 금합니다.', { size: 15, color: C.midGray }),
        new TextRun({ text: '\t' }),
        run('Page ', { size: 15, color: C.midGray }),
        new TextRun({ children: [PageNumber.CURRENT], size: 15, font: FONT, color: C.amber }),
        run(' / ', { size: 15, color: C.midGray }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, font: FONT, color: C.midGray }),
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    }),
  ],
});

const chapterTitle = (text) =>
  paragraph([run(text, { bold: true, size: 25, color: C.charcoal })], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 140 },
    border: { left: { style: BorderStyle.THICK, size: 20, color: C.amber, space: 12 } },
  });

const sectionTitle = (text) =>
  paragraph([run(text, { bold: true, size: 21, color: C.charcoal })], {
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 90 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
  });

const subTitle = (text) =>
  paragraph([run(text, { bold: true, size: 20, color: C.midGray })], {
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
  });

const bullet = (text) =>
  paragraph([run(text)], {
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 40 },
  });

function addTable(children, rows, options) {
  children.push(makeTable(rows, options), gap(120));
}

function addBullets(children, items) {
  items.forEach((item) => children.push(bullet(item)));
}

const bodyChildren = [
  gap(200),
  chapterTitle('들어가며'),
  sectionTitle('이 가이드북의 목적과 읽는 법'),
  body('AI는 이제 업무의 일부입니다. 문서 작성, 코드 리뷰, 번역, 데이터 분석 등 다양한 영역에서 AI 도구가 활용되고 있으며, 앞으로 그 범위는 더 넓어질 것입니다.'),
  body('이 가이드북은 KTNET 임직원이 AI 도구를 안전하고 올바르게 사용할 수 있도록 돕기 위해 만들어졌습니다. 어떤 정보를 AI에 입력할 수 있는지, 결과물을 어떻게 검증해야 하는지, 문제가 생겼을 때 어떻게 대응해야 하는지를 실무 중심으로 설명합니다.'),
  noteCallout('이 가이드북은 법적 구속력이 있는 규정 문서가 아닙니다. 규정의 핵심 내용을 임직원이 쉽게 이해하고 실무에 적용할 수 있도록 정리한 참고 자료입니다. 규정과 가이드북이 충돌하는 경우 규정이 우선합니다.'),
  subTitle('읽는 방법'),
];
addBullets(bodyChildren, [
  '처음 AI 도구를 사용하는 분은 처음부터 순서대로 읽으세요.',
  '특정 상황에서 궁금한 것이 있는 분은 해당 장을 찾아 바로 참고하세요.',
]);
bodyChildren.push(sectionTitle('빠른 참조 카드'), body('아래 표를 통해 상황별로 어디를 확인하면 되는지 바로 찾을 수 있습니다.'));
addTable(bodyChildren, [
  ['상황', '확인할 곳'],
  ['어떤 AI 도구를 쓸 수 있는지 모를 때', 'AI 도구 활용 정책 승인 목록 또는 AX추진실 문의'],
  ['이 정보를 AI에 입력해도 될지 모를 때', '제1장 1.2 입력 금지 정보, 1.4 판단 플로차트'],
  ['망이 달라 어떤 도구를 써야 할지 모를 때', '제1장 1.3 망별 활용 가이드'],
  ['프롬프트를 어떻게 써야 할지 모를 때', '제2장 2.1 프롬프트 작성법'],
  ['AI 결과물을 그냥 써도 될지 모를 때', '제2장 2.2 결과물 검증'],
  ['사고가 발생했을 때', '제4장 4.2 신고 절차 Step by Step'],
  ['전문가에게 문의하고 싶을 때', '제5장 AI 헬프데스크 이용 안내'],
], { widths: [3000, 5406] });

bodyChildren.push(
  chapterTitle('제1장  AI에 어떤 정보를 입력할 수 있나요?'),
  body('AI 도구를 사용할 때 가장 먼저 확인해야 할 것은 입력하려는 정보가 해당 AI 도구에 입력해도 되는 정보인지 여부입니다. 이 판단을 위해 두 가지를 알아야 합니다. 하나는 입력하려는 데이터의 민감도이고, 다른 하나는 사용하려는 AI 도구의 보안 등급입니다.'),
  sectionTitle('1.1  데이터 민감도 3단계 이해하기'),
);
addTable(bodyChildren, [
  ['등급', '명칭', '정의 및 예시'],
  ['Level 1', '일반정보', '외부에 공개되어 있거나 공개해도 무방한 정보. 공개된 회사 소개 자료, 법령 원문, 공개 통계 등.'],
  ['Level 2', '대외비', '외부에는 공개하지 않으나 사내에서는 공유하는 정보. 내부 보고서, 기획 문서, 가명처리된 자료 등.'],
  ['Level 3', '비밀', '유출 시 회사·회원사·개인에게 심각한 피해를 줄 수 있는 정보. 개인정보, 회원사 거래 데이터, 미공개 재무 자료 등.'],
], { widths: [1300, 1500, 5606], centerCols: [0, 1] });
bodyChildren.push(
  noteCallout('판단 원칙: 어떤 등급인지 확신하기 어려울 때는 상향 분류 원칙을 적용합니다. Level 1인지 2인지 모르겠다면 Level 2로 처리하세요.'),
  sectionTitle('1.2  절대 입력 금지 정보 9가지'),
  body('아래 9가지 정보는 Tier A 도구를 포함하여 어떤 외부 AI 도구에도 원칙적으로 입력할 수 없습니다.'),
);
addTable(bodyChildren, [
  ['번호', '절대 입력 금지 정보'],
  ['①', '회원사의 수출입 신고 원문, 통관 내역, 원산지 정보, 가격 자료'],
  ['②', '회원사·거래처와의 계약서 원본 및 미공개 계약 조건'],
  ['③', '개인정보 (성명, 주민등록번호, 여권번호, 연락처, 계좌번호 등)'],
  ['④', '임직원 인사·평가·급여 자료'],
  ['⑤', '미공개 재무·경영·M&A 관련 자료'],
  ['⑥', '회사 정보시스템의 소스코드 중 인증, 암호화, 핵심 비즈니스 로직'],
  ['⑦', 'API 키, 비밀번호, 인증서, 접속 크리덴셜'],
  ['⑧', '법무·감사·징계·소송 관련 진행 중인 자료'],
  ['⑨', '국가보안 관련 정보 및 법령상 비밀로 규정된 정보'],
], { widths: [900, 7506], centerCols: [0] });
bodyChildren.push(
  warnCallout('위 정보를 실수로 입력했다면 즉시 AX추진실에 보고해야 합니다. 스스로 처리하려다 늦어지면 대응이 더 어려워집니다.'),
  body('근거: 인공지능 활용 및 관리 요령 제16조'),
  sectionTitle('1.3  망별 활용 가이드 한눈에 보기'),
  body('KTNET의 네트워크는 세 가지 망으로 구성되어 있으며, 망에 따라 AI 도구 사용 가능 범위가 달라집니다.'),
);
addTable(bodyChildren, [
  ['망', '통제 수준', 'AI 도구 사용 기준'],
  ['본인확인망', '최상', '외부 클라우드 AI 도구 전면 금지. 온프레미스 LLM만 CISO 정책 결정 후 제한적 허용.'],
  ['업무망', '양호', '§3.2 매트릭스 적용. 회사 워크스페이스 가입 필수. 관리자 콘솔 기반 사용자 권한·사용량 통제(Tier A). AI 활용 동의서 사전 제출. 분기별 로그 감시. AI 게이트웨이·DLP는 도입 준비 중이며 도입 후 실시간 통제로 전환 예정.'],
  ['인터넷망', '강화 필요', '§3.2 매트릭스 적용. 회사 워크스페이스 가입 필수. 부서장 서면 승인 필수. AI 활용 동의서 사전 제출. 월 1회 로그 감시 및 분기 감사실 검토. DLP·EPP는 도입 준비 중이며 도입 후 실시간 감시로 전환 예정.'],
], { widths: [1200, 1300, 5906], centerCols: [0, 1] });
bodyChildren.push(subTitle('Tier × Level 매트릭스 요약'), body('본인확인망은 모든 조합에서 금지이므로 업무망과 인터넷망만 표시합니다.'));
addTable(bodyChildren, [
  ['망 / Tier', '업무망 Tier A (허용)', '업무망 Tier B (제한)', '업무망 Tier C (주의)', '인터넷망 Tier A (허용)', '인터넷망 Tier B (제한)', '인터넷망 Tier C (주의)'],
  ['Level 1 (일반)', '허용', '자기신고 후 허용', '허용(AI 도구 활용 정책 5.4)', '허용', '자기신고 후 허용', '허용(AI 도구 활용 정책 5.4)'],
  ['Level 2 (대외비)', '워크스페이스 통제 허용\n(현: 콘솔+동의서 / 향후: 게이트웨이·DLP)', '금지', '금지', '부서장 승인 허용\n(현: 콘솔+동의서 / 향후: DLP)', '금지', '금지'],
  ['Level 3 (비밀)', '금지', '금지', '금지', '금지', '금지', '금지'],
], { widths: [1100, 1350, 1100, 1100, 1350, 1100, 1306], centerCols: [0, 1, 2, 3, 4, 5, 6] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제10조·제11조 / AI 도구 활용 정책 3.1~3.2'),
  subTitle('현 시점 통제 체계 — AI 활용 동의서'),
  body('AI 게이트웨이·DLP가 도입되기 전까지, 업무망 및 인터넷망에서 AI 도구를 사용하는 임직원은 사전에 「AI 활용 동의서」를 AX추진실에 제출해야 합니다.'),
  body('동의서는 회사가 관리자 콘솔·분기 감사를 통해 AI 사용 현황을 모니터링하는 것에 대한 사전 동의를 확인하는 문서입니다.'),
  warnCallout('동의서 미제출 상태에서 AI 도구를 사용하면 AI 도구 활용 정책 8.4 제재 절차의 대상이 됩니다.'),
  body('동의서 양식은 AX추진실(kiwhan@ktnet.co.kr)이 배포합니다. 근거: AI 도구 활용 정책 5.2, 5.3'),
  subTitle('Tier B 도구 자기신고 방법'),
  body('Tier B(개인 명의 유료 플랜) 도구를 업무 목적으로 사용하려는 임직원은 사용 개시 전 반드시 AX추진실에 자기신고를 완료해야 합니다.'),
);
addTable(bodyChildren, [
  ['항목', '내용'],
  ['신고 대상', 'Tier B로 분류된 AI 도구를 업무 목적으로 사용하려는 전 임직원'],
  ['신고 항목', '① 도구명  ② 사용 목적  ③ 입력 예정 데이터 Level  ④ 사용 빈도'],
  ['신고 방법', 'kiwhan@ktnet.co.kr  |  내선 x2759  |  smartflow 메신저(김기환)'],
  ['처리 기한', '신고 접수 후 5 영업일 이내 심사 결과 통보'],
  ['미신고 사용', 'AI 도구 활용 정책 §8.4 제재 대상 (위반 사항으로 처리)'],
], { widths: [1800, 6606] });
bodyChildren.push(
  body('근거: AI 도구 활용 정책 8.3'),
  sectionTitle('1.4  판단이 어려울 때 — 플로차트'),
  body('아래 4단계를 순서대로 확인하세요. 판단이 어렵거나 예외적인 상황은 AX추진실(kiwhan@ktnet.co.kr)에 문의하세요.'),
);
addTable(bodyChildren, [
  ['단계', '판단 흐름'],
  ['Step 1', '입력 정보가 1.2의 절대 금지 9가지에 해당하는가?\n→ YES: 입력 금지. 가명처리 후 재검토하거나 AX추진실 문의.\n→ NO: Step 2로 이동.'],
  ['Step 2', '데이터 등급은 무엇인가? (확신 없으면 상향 분류)\n→ Level 3 (비밀): 외부 AI 도구 입력 금지.\n→ Level 1·2: Step 3으로 이동.'],
  ['Step 3', '사용 AI 도구의 Tier는? (승인 목록에서 확인)\n→ Tier A + Level 1·2: 허용.\n→ Tier B + Level 2: 가명·마스킹·AX추진실 승인 필요.\n→ Tier C + Level 2: 금지. Tier C + Level 1: 회사명·맥락 제외 후 허용.'],
  ['Step 4', '현재 망에서 해당 도구 사용이 허용되는가? (1.3장 참고)\n→ YES: 입력 가능. 프롬프트 작성 원칙(제2장) 준수.\n→ NO: AX추진실에 별도 승인 문의.'],
], { widths: [1200, 7206], centerCols: [0] });
bodyChildren.push(
  sectionTitle('1.5  업무별 데이터 분류 예시'),
  body('1.1의 3단계 정의만으로 판단이 어려울 때 아래 예시를 참고하세요. KTNET의 주요 업무 영역별로 구체적인 사례를 정리했습니다.'),
  body('이 예시집은 AX추진실이 주기적으로 갱신합니다. 업무 특성상 Level 판단이 반복적으로 어렵거나 새로운 데이터 유형이 생겼다면 kiwhan@ktnet.co.kr로 제보해 주세요.'),
  body('근거: AI 도구 활용 정책 3.3'),
  subTitle('무역·통관·물류 업무'),
);
addTable(bodyChildren, [
  ['등급', '예시', '비고'],
  ['Level 1', '관세청·무역협회 등이 공표한 수출입 통계', '공개 데이터'],
  ['Level 1', 'HS코드 체계 설명, 품목 분류 일반 기준', '공개 정보'],
  ['Level 1', '관세법·FTA 협정문 원문 및 해설', '공개 법령'],
  ['Level 1', 'uTradeHub 공개 서비스 매뉴얼', '외부 공개 자료'],
  ['Level 2', '내부 통관 프로세스 절차서 (회원사 정보 미포함)', '사내 업무 문서'],
  ['Level 2', 'HS코드 분류 검토 결과 (특정 거래·회원사 정보 제거 후)', '개인정보·기밀 미포함 시'],
  ['Level 2', '물류 서비스 기획 내부 검토 자료 (미공개, 거래 정보 미포함)', '사내 기획 문서'],
  ['Level 3', '회원사별 수출입 신고 원문, 신고번호, 통관 내역', '절대 입력 금지'],
  ['Level 3', '특정 회원사의 원산지 정보, 가격 자료, 물량 데이터', '절대 입력 금지'],
  ['Level 3', '화주·수하인·선사 등 거래 당사자 개인정보가 포함된 B/L, AWB', '절대 입력 금지'],
  ['Level 3', '특정 회원사와의 미공개 계약 조건, 단가 협의 내역', '절대 입력 금지'],
], { widths: [1200, 5200, 2006], centerCols: [0] });
bodyChildren.push(subTitle('서비스 개발·운영 업무'));
addTable(bodyChildren, [
  ['등급', '예시', '비고'],
  ['Level 1', '공개 오픈소스 라이브러리 코드 질문', '공개 기술 정보'],
  ['Level 1', '일반 알고리즘·디자인 패턴 설명 요청', '공개 기술 지식'],
  ['Level 1', '공개 API 명세서 분석 요청', '외부 공개 자료'],
  ['Level 2', '사내 시스템의 일반 업무 로직 코드 (인증·암호화·핵심 로직 제외)', '미공개이나 기밀 아님'],
  ['Level 2', '내부 API 설계 문서 (실제 데이터 미포함)', '사내 기술 문서'],
  ['Level 2', '장애 로그 분석 요청 (개인정보·크리덴셜 제거 후)', '마스킹 처리 필수'],
  ['Level 2', '테스트 환경의 가명·합성 데이터', '실제 개인정보 미포함 시'],
  ['Level 3', '인증·암호화·핵심 비즈니스 로직이 포함된 소스코드', '절대 입력 금지'],
  ['Level 3', 'DB 접속 정보, API 키, 비밀번호, 인증서', '절대 입력 금지'],
  ['Level 3', '실 운영 서버 로그 (회원사 거래 기록, 개인정보 포함)', '절대 입력 금지'],
  ['Level 3', 'TradeSign, eTradeBill 등 서비스의 실 사용자 데이터', '절대 입력 금지'],
], { widths: [1200, 5200, 2006], centerCols: [0] });
bodyChildren.push(subTitle('경영·행정 업무'));
addTable(bodyChildren, [
  ['등급', '예시', '비고'],
  ['Level 1', '공개된 산업 동향 보고서, 경쟁사 공시 자료 분석', '외부 공개 자료'],
  ['Level 1', '일반 비즈니스 문서 작성 요령, 기획서 구조 질문', '회사 정보 미포함'],
  ['Level 2', '내부 사업 기획서 초안 (미공개, 재무 수치 미포함)', '사내 기획 문서'],
  ['Level 2', '부서별 업무 보고서 (개인정보·재무 수치 미포함)', '사내 내부 보고'],
  ['Level 2', '내부 회의록 (인사·재무·법무 관련 내용 제외)', '사내 행정 문서'],
  ['Level 3', '미공개 재무제표, 예산 계획, 손익 자료', '절대 입력 금지'],
  ['Level 3', '임직원 인사·평가·급여 자료', '절대 입력 금지'],
  ['Level 3', 'M&A, 투자, 사업 매각 관련 검토 자료', '절대 입력 금지'],
  ['Level 3', '법무·감사·징계·소송 관련 진행 중인 자료', '절대 입력 금지'],
], { widths: [1200, 5200, 2006], centerCols: [0] });
bodyChildren.push(subTitle('고객(회원사) 대응 업무'));
addTable(bodyChildren, [
  ['등급', '예시', '비고'],
  ['Level 1', '서비스 이용 방법 일반 안내문 작성', '공개 정보'],
  ['Level 1', 'FAQ 답변 초안 (특정 고객 정보 미포함)', '공개 서비스 정보'],
  ['Level 2', '고객 문의 유형 분석 보고서 (개인 식별 정보 제거 후)', '집계·통계 형태'],
  ['Level 2', '서비스 장애 공지 초안 (특정 고객 정보 미포함)', '사내 공지 문서'],
  ['Level 3', '특정 회원사 담당자 성명, 연락처, 이메일', '절대 입력 금지'],
  ['Level 3', '특정 회원사 민원·분쟁 내용, CS 이력', '절대 입력 금지'],
  ['Level 3', '회원사 사업자번호, 계좌 정보 등', '절대 입력 금지'],
], { widths: [1200, 5200, 2006], centerCols: [0] });
bodyChildren.push(subTitle('헷갈리기 쉬운 경계 사례'), body('실무에서 자주 발생하는 판단 어려운 사례입니다. 모르겠으면 상향 분류가 원칙입니다.'));
addTable(bodyChildren, [
  ['사례', '판정', '이유'],
  ['고객사 이름 바꾼 계약서 ("XX사"로 대체)', 'Level 3', '계약 구조·조건이 특정 거래를 식별 가능. 실제 계약 내용 기반이면 금지.'],
  ['이름 없고 사번·직급·평가 점수만 있는 자료', 'Level 3', '사번은 개인 식별 정보. 인사 정보와 결합 시 개인정보에 해당.'],
  ['사용자 이름·이메일 제거 후 IP 주소만 남긴 로그', 'Level 3', 'IP 주소는 다른 정보와 결합 시 개인 식별 가능. 완전 익명화 필요.'],
  ['특정 회원사명 포함 장애 로그 (회사명 마스킹)', 'Level 2', '마스킹 후 기술 패턴만 남은 경우 Level 2 검토 가능. 확신 없으면 상향.'],
  ['관세청 통계 기반 내부 가공 분석 보고서', 'Level 2', '가공·분석 결과는 내부 자료. 재무 수치·회원사 정보 포함 시 Level 3 상향.'],
], { widths: [3000, 1200, 4206], centerCols: [1] });

bodyChildren.push(
  chapterTitle('제2장  AI를 제대로 활용하는 실전 방법'),
  body('데이터 입력 조건을 충족했다면, 이제 AI를 효과적으로 활용하는 방법을 알아봅니다.'),
  sectionTitle('2.1  프롬프트 잘 쓰는 법'),
  body('프롬프트는 AI에게 주는 지시문입니다. 잘 작성된 프롬프트는 정확한 결과물을 만들고 검증도 쉽게 해줍니다.'),
);
addTable(bodyChildren, [
  ['원칙', '설명'],
  ['① 민감정보 제외', '민감정보를 포함하지 않는 것이 기본입니다. 불가피한 경우 가명·익명화 후 입력합니다.\n나쁜 예: "홍길동(사번1234)의 평가 결과 요약해줘"\n좋은 예: "직원 A의 평가 항목 3가지를 요약하는 형식 만들어줘"'],
  ['② 맥락·목적 명시', '구체적인 맥락·목적·원하는 결과물 형식을 명시할수록 유용한 결과를 얻습니다.\n나쁜 예: "이메일 써줘"\n좋은 예: "거래처에 프로젝트 일정 변경 안내 이메일. 기존 12/1→12/15, 공식적이나 친근한 톤, 5문장 이내."'],
  ['③ 검증 기준 포함', '"법령에 근거하여", "출처를 명시하여" 등의 표현을 추가하면 결과물 검증이 쉬워집니다.'],
  ['④ 시스템 프롬프트 보호', '시스템 프롬프트를 설계하는 임직원은 사용자 입력에 의해 덮어써지지 않도록 기술적 보호조치를 적용해야 합니다.'],
], { widths: [1800, 6606] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제17조'),
  sectionTitle('2.2  AI 결과물은 반드시 검증하세요'),
  body('AI가 생성한 내용은 틀릴 수 있습니다. AI 결과물을 업무에 사용하는 임직원은 스스로 검증할 책임이 있습니다.'),
);
addTable(bodyChildren, [
  ['항목', '검증 체크리스트'],
  ['1', '사실 관계가 정확한가? (특히 수치, 날짜, 통계, 법조문)'],
  ['2', '출처가 실제로 존재하는가? (AI가 만들어낸 가짜 출처인지 확인)'],
  ['3', '편향이나 불공정한 서술이 없는가?'],
  ['4', '개인정보나 민감정보가 결과물에 포함되어 있지 않은가?'],
  ['5', '저작권 침해 가능성이 있는 내용이 포함되어 있지 않은가?'],
  ['6', '업무 맥락에 맞는 내용인가? (AI가 문맥을 잘못 이해했을 가능성)'],
], { widths: [900, 7506], centerCols: [0] });
bodyChildren.push(
  warnCallout('AI가 생성한 출처·인용·법조문·통계는 반드시 원문을 직접 확인한 후에 인용하세요. AI는 그럴듯하지만 실제로 존재하지 않는 출처를 만들어내는 경우가 있습니다.'),
  body('근거: 인공지능 활용 및 관리 요령 제18조'),
  sectionTitle('2.3  AI가 만든 문서에 표시해야 할 것들'),
  body('생성형 AI가 실질적으로 관여한 결과물을 대외적으로 공개·배포할 때는 반드시 그 사실을 표시해야 합니다.'),
);
addTable(bodyChildren, [
  ['대상', '표시 방법'],
  ['문서·보고서·기사 등', '"본 문서는 생성형 AI의 지원을 받아 작성되었으며 담당자가 검토하였습니다." 또는 이에 상당하는 문구를 문서 내 또는 파일 속성에 표시.'],
  ['이미지·영상·음성', '비가시적 워터마크 또는 C2PA 규격 메타데이터 삽입 원칙. 기술적으로 어려운 경우 배포 설명문에 AI 활용 사실 명시.'],
], { widths: [2200, 6206] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제19조'),
  sectionTitle('2.4  AI 에이전트 사용 시 특별 주의사항'),
  body('AI 에이전트는 외부 시스템·도구를 자율적으로 호출하여 일련의 작업을 수행하는 AI입니다. 실제 시스템에 영향을 미치는 행동을 취할 수 있어 별도의 주의가 필요합니다.'),
  body('다음 5가지 작업은 에이전트 실행 직전 사람의 명시적 승인이 필요합니다.'),
);
addTable(bodyChildren, [
  ['번호', '사람 승인이 필요한 작업'],
  ['①', '결제·송금·계약 체결'],
  ['②', '고객·회원사에 대한 대외 통신(이메일, 문자 등) 발송'],
  ['③', '파일 시스템의 영구적 변경 및 삭제'],
  ['④', '외부 네트워크로의 데이터 전송'],
  ['⑤', '사내 인사·재무 시스템에 대한 쓰기 작업'],
], { widths: [900, 7506], centerCols: [0] });
bodyChildren.push(body('근거: 인공지능 활용 및 관리 요령 제42조'), sectionTitle('2.5  업무 유형별 활용 예시'));
addTable(bodyChildren, [
  ['업무 유형', '활용 예시', '주의사항'],
  ['문서 작성', '보고서 초안, 이메일 문구, 회의록 정리, 외국어 자료 번역', '회원사 정보·개인정보 포함 내용은 입력 전 마스킹 필수'],
  ['코딩 지원', '반복 코드 자동완성, 리팩터링, 버그 원인 분석, 문서 주석 생성', '인증·암호화 로직 및 API 키·비밀번호 절대 입력 금지'],
  ['번역·다국어', '공개 자료 번역, 외국어 문서 요약', '계약서·법적 문서는 AI 번역 결과 그대로 사용 금지, 전문가 검토 필요'],
  ['데이터 분석', '공개 데이터·Level 1 대상 분석 방법론 탐색, 시각화 아이디어', '고객·거래 데이터 등 Level 2 이상은 입력 전 승인 절차 확인'],
], { widths: [1400, 3800, 3206] });

bodyChildren.push(
  chapterTitle('제3장  AI 윤리와 책임 있는 사용'),
  body('이 장은 AI를 단순히 사용하는 것을 넘어 윤리적으로 올바르게 사용하기 위한 내용을 다룹니다. KTNET의 AI 윤리 교육 보조 자료로도 활용됩니다.'),
  sectionTitle('3.1  AI도 틀릴 수 있다 — 편향이란 무엇인가?'),
  body('AI는 학습 데이터를 기반으로 동작합니다. 학습 데이터에 편향이 있다면 AI의 결과물에도 편향이 나타날 수 있습니다.'),
);
addTable(bodyChildren, [
  ['편향 유형', '설명'],
  ['데이터 편향', '특정 집단이 학습 데이터에 과소 또는 과대 대표된 경우. 특정 국가·언어 데이터만 많이 학습한 AI는 다른 문화권의 맥락을 잘못 이해할 수 있습니다.'],
  ['알고리즘 편향', '모델 설계 과정에서 의도치 않게 특정 패턴을 더 중요하게 학습하는 경우.'],
  ['피드백 편향', 'AI 결과물을 사용자가 검토·수정하는 과정에서 특정 방향으로 AI가 강화되는 경우.'],
], { widths: [1800, 6606] });
bodyChildren.push(
  warnCallout('인사 평가·채용·신용 심사 등 사람에게 직접 영향을 미치는 의사결정에 AI를 사용할 때는 편향 가능성에 특히 주의하고, AI 판단을 그대로 따르지 말고 사람이 반드시 검토해야 합니다.'),
  body('근거: 인공지능 활용 및 관리 요령 제24조의2'),
  sectionTitle('3.2  KTNET AI 거버넌스 7대 원칙'),
);
addTable(bodyChildren, [
  ['원칙', '설명'],
  ['① 사람 중심의 인공지능', 'AI는 인간의 존엄·기본권을 존중. 최종 의사결정의 책임은 항상 사람에게 있다.'],
  ['② 안전성과 신뢰성', '예상 가능한 위험에 선제적으로 대응. 검증 없이 중요한 결정에 AI를 사용하지 않는다.'],
  ['③ 투명성과 설명가능성', 'AI의 작동 원리·사용 데이터·결과 근거를 합리적 수준에서 설명할 수 있어야 한다.'],
  ['④ 공정성과 비차별', '성별·연령·출신·장애·종교에 따른 부당한 차별을 발생시키지 않아야 한다.'],
  ['⑤ 데이터 보호와 프라이버시', '개인정보·민감정보·영업비밀·회원사 정보를 법령과 규정에 따라 보호한다.'],
  ['⑥ 책임성과 감독', 'AI 활용·운영 책임 소재를 명확히 하고, 사람이 실질적으로 감독할 수 있게 한다.'],
  ['⑦ 지속 가능성', 'AI 활용이 환경·사회·인사에 미치는 영향을 종합적으로 고려한다.'],
], { widths: [2400, 6006] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제5조 (AI 기본법, NIST AI RMF, 금융위 AI 가이드라인 종합)'),
  sectionTitle('3.3  저작권·지식재산 주의사항'),
);
addBullets(bodyChildren, [
  'AI 결과물의 저작권 귀속은 아직 명확하게 정립되지 않은 상태입니다. 각 AI 서비스의 이용약관을 확인하고 상업적 사용 가능 여부를 파악하세요.',
  'AI는 학습 데이터에 포함된 저작물의 내용을 그대로 출력할 수 있습니다. 외부 공개 전 제3자 저작물과의 유사 여부를 확인하세요.',
  '회사 고유의 핵심 기술·사업 전략·알고리즘을 AI에 입력하는 행위는 「무체재산권 보호에 관한 요령」 위반에 해당할 수 있습니다.',
]);
bodyChildren.push(
  sectionTitle('3.4  AI 판단에 이의가 있다면 — 재검토 요청 절차'),
  body('AI가 내린 판단 또는 추천에 이의가 있는 경우 다음 절차에 따라 재검토를 요청할 수 있습니다. 이 권리를 행사한다는 이유로 불이익을 받지 않습니다.'),
);
addTable(bodyChildren, [
  ['단계', '절차'],
  ['Step 1', 'AI 헬프데스크(kiwhan@ktnet.co.kr 또는 x2759)에 재검토 요청 제출'],
  ['Step 2', '이의 내용, 해당 AI 도구명, 결과물 내용을 함께 기재'],
  ['Step 3', '회사는 합리적인 기간 내에 사람에 의한 재검토 결과를 회신'],
], { widths: [1200, 7206], centerCols: [0] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제24조의2 ⑤항'),
  sectionTitle('3.5  책임 소재와 면책 요건'),
  body('AI 결과물을 업무에 사용한 임직원은 그 결과물로 인한 손실에 대해 책임을 집니다. 다만 다음 조건을 모두 충족한 경우에는 면책이 적용될 수 있습니다.'),
);
addTable(bodyChildren, [
  ['면책 적용 조건 (모두 충족 시)', '면책 불가 사유'],
  ['✓ 승인된 AI 도구만 사용했을 것\n✓ 데이터 입력 기준(1장) 준수했을 것\n✓ 결과물 검증 절차(2.2) 합리적으로 수행했을 것\n✓ 규정 위반 사항이 없을 것', '✗ 승인되지 않은 AI 도구 사용\n✗ 입력 금지 정보를 의도적으로 입력\n✗ 검증 없이 AI 결과물 그대로 사용\n✗ 사고 발생 사실을 숨기고 늦게 보고'],
], { widths: [4203, 4203] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제18조 ②항, 제42조'),
  sectionTitle('3.6  취약계층·다양성 배려'),
);
addBullets(bodyChildren, [
  '고령자, 장애인, 외국인 등 디지털 접근성이 제한될 수 있는 이용자를 대상으로 AI 기반 서비스를 제공할 때는 대체 수단을 함께 제공해야 합니다.',
  '비숙련 임직원·신규 입사자에게 AI 결과물을 일방적으로 제시하거나 검증 없이 따르도록 강요하는 것은 바람직하지 않습니다.',
  'AI 도입이 특정 직무에 중대한 변화를 가져올 경우, 회사는 해당 임직원에게 재교육 기회를 우선 제공하며 AI 도입만을 이유로 한 해고·구조조정은 배제합니다.',
]);
bodyChildren.push(body('근거: 인공지능 활용 및 관리 요령 제24조의2 ④항, 제56조'));

bodyChildren.push(
  chapterTitle('제4장  문제가 생겼을 때'),
  body('AI를 사용하다 보면 예상치 못한 상황이 발생할 수 있습니다. 이 장에서는 문제 상황별 대처 방법과 신고 절차를 안내합니다.'),
  sectionTitle('4.1  AI 사고 유형'),
);
addTable(bodyChildren, [
  ['유형', '설명', '예시'],
  ['① 민감정보 유출', '입력 금지 정보가 AI에 입력되거나, AI 결과물을 통해 외부 유출', '개인정보 포함 파일 AI 업로드 / 다른 사용자 정보가 응답에 포함'],
  ['② 오류 결과물 피해', '잘못된 AI 결과물 무검증 사용으로 업무·회원사 피해 발생', '틀린 법조문 인용 / 오류 수치 포함 보고서로 의사결정'],
  ['③ 보안 침해', '악의적 프롬프트로 AI 시스템 조작 또는 내부 시스템 접근', '프롬프트 인젝션 공격 / AI가 악성 명령 실행'],
  ['④ 규정 위반', '승인되지 않은 도구 사용, 입력 금지 정보 입력 등 규정 위반', '미승인 도구 사용 / 입력 금지 정보 의도적 입력'],
], { widths: [1800, 3500, 3106] });
bodyChildren.push(
  body('근거: 인공지능 활용 및 관리 요령 제2조 제12호'),
  sectionTitle('4.2  즉시 해야 할 일 — 신고 절차 Step by Step'),
  body('AI 사고가 발생했거나 의심될 때는 먼저 숨기지 마세요. 사고를 숨기고 늦게 보고하면 피해가 커지고 책임도 무거워집니다.'),
);
addTable(bodyChildren, [
  ['단계', '조치 내용'],
  ['Step 1 즉시 중단', '사고가 의심되는 순간 해당 AI 도구의 사용을 즉시 중단합니다.'],
  ['Step 2 헬프데스크 신고', '이메일(kiwhan@ktnet.co.kr) 또는 내선 x2759로 즉시 신고합니다. 신고 내용: 발생 일시, 사용 도구명, 입력한 정보의 종류, 의심 정황.'],
  ['Step 3 부서장 보고', '헬프데스크 신고와 별도로 소속 부서장에게도 즉시 보고합니다.'],
  ['Step 4 안내에 따라 대응', 'AX추진실·인프라운영실의 안내에 따라 추가 조치를 취합니다. 사고 내용을 혼자 처리하거나 증거를 삭제하지 마세요.'],
], { widths: [2200, 6206] });
bodyChildren.push(
  warnCallout('개인정보 유출 사고의 경우 CPO에게도 즉시 보고해야 합니다. 회원사 정보가 유출된 경우 해당 회원사에 72시간 이내 통지해야 합니다.'),
  body('근거: 인공지능 활용 및 관리 요령 제38조~제50조'),
  sectionTitle('4.3  위반 시 처리 기준 개요'),
  body('AI 관련 규정 위반은 「인사관리규정」 및 「취업규칙」에 따라 처리됩니다.'),
);
addTable(bodyChildren, [
  ['위반 유형', '처리 방향', '비고'],
  ['승인되지 않은 AI 도구 사용', '경고 ~ 시말서 (고의·반복 시 중징계)', ''],
  ['입력 금지 정보 입력', '유출 피해 규모에 따라 경고 ~ 징계', ''],
  ['사고 미신고·은폐', '중징계 (사고 자체보다 엄중 처리)', '자진신고 시 감경'],
  ['결과물 무검증 사용으로 피해 발생', '피해 규모에 따라 손해배상 청구 가능', '합리적 검증 시 면책'],
], { widths: [3000, 3600, 1806] });
bodyChildren.push(body('근거: 인공지능 활용 및 관리 요령 제48·49조 / 별표 12 징계 양정 가이드'));

bodyChildren.push(
  chapterTitle('제5장  AI 헬프데스크 이용 안내'),
  body('AI 사용 중 궁금한 점이 생기거나 판단이 어려운 상황이 발생하면 언제든지 AI 헬프데스크를 이용하세요. 질문하는 것이 규정 위반보다 훨씬 낫습니다.'),
  sectionTitle('5.1  문의 채널'),
);
addTable(bodyChildren, [
  ['채널', '연락처', '운영 시간', '응답 목표'],
  ['전용 이메일', 'kiwhan@ktnet.co.kr', '업무시간', '1영업일 이내'],
  ['내선 전화', 'x2759', '업무시간', '즉시'],
  ['사내 메신저 (smartflow)', '김기환', '업무시간', '즉시'],
], { widths: [2200, 2400, 1800, 2006], centerCols: [2, 3] });
bodyChildren.push(sectionTitle('5.2  문의 유형별 적합 채널 안내'));
addTable(bodyChildren, [
  ['상황', '적합 채널'],
  ['사고 의심, 긴급 문의', '내선 x2759 또는 kiwhan@ktnet.co.kr 직접 연락'],
  ['일반 정책 문의 (입력 여부, 도구 선택 등)', 'smartflow 메신저(김기환) 또는 이메일'],
  ['신규 AI 도구 도입 검토 요청', 'AI 도구 활용 정책 1.4 도입 신청 절차에 따라 AX추진실에 도입 검토 요청서 제출'],
  ['교육·가이드북 내용 문의', '이메일(kiwhan@ktnet.co.kr) 또는 smartflow 메신저(김기환)'],
], { widths: [3400, 5006] });
bodyChildren.push(body('근거: 인공지능 활용 및 관리 요령 제8조 ④항'));

bodyChildren.push(chapterTitle('부록'), sectionTitle('A.  용어 해설'));
addTable(bodyChildren, [
  ['용어', '정의'],
  ['AI 에이전트', '사용자의 지시를 받아 외부 시스템·도구를 자율적으로 호출하여 일련의 작업을 수행하는 AI.'],
  ['가명처리', '개인정보의 일부를 삭제·대체하여 추가 정보 없이는 특정 개인을 알아볼 수 없도록 처리하는 것.'],
  ['대외비', '외부에는 공개하지 않으나 사내에서는 공유되는 정보. 데이터 민감도 Level 2에 해당.'],
  ['마스킹', "정보의 일부를 '*' 등으로 대체하여 식별 불가능하게 처리하는 것. 예: 주민번호 뒷자리 처리."],
  ['생성형 AI', '이용자의 입력에 따라 텍스트·이미지·음성·영상·소프트웨어 코드 등을 새로 생성하는 AI.'],
  ['Tier', 'AI 도구의 보안 등급. Tier A(높은 보안)·Tier B·Tier C(낮은 보안) 3단계 분류.'],
  ['프롬프트', 'AI에 입력하는 지시문. 시스템 프롬프트(기본 동작 설정)와 사용자 프롬프트(실제 질문)로 구분.'],
  ['프롬프트 인젝션', '악의적인 사용자가 프롬프트에 특수 지시를 삽입하여 AI 동작을 조작하는 공격 기법.'],
  ['LLM', '대규모 텍스트 말뭉치로 학습된 생성형 AI로서 자연어 이해·생성을 주된 기능으로 하는 모델.'],
], { widths: [1800, 6606] });
bodyChildren.push(sectionTitle('B.  관련 사규·문서 및 담당 연락처'), subTitle('관련 사규'));
addTable(bodyChildren, [
  ['사규명', '주요 내용'],
  ['인공지능 활용 및 관리 요령', 'AI 활용·관리 최상위 요령 (이 가이드북의 근거 규정)'],
  ['AI 도구 활용 정책 (AI-GOV-003-001)', '승인 도구 목록, Tier 분류, 매트릭스, 망별 활용 신규 도구 도입 신청 및 폐기 절차'],
  ['AI 코딩어시스턴트 구독 정책', '개발 직군 AI 코딩 도구 구독·비용 기준'],
  ['KTNET 윤리강령', 'AI 활용에도 적용되는 최상위 윤리 기준'],
  ['정보보호관리세칙', '정보 등급 분류, 접근 통제, 사고 대응 기준'],
  ['개인정보보호세칙', '개인정보 처리 기준'],
], { widths: [3100, 5306] });
bodyChildren.push(subTitle('주요 연락처'));
addTable(bodyChildren, [
  ['담당 부서', '연락처', '문의 내용'],
  ['AX추진실 (AI 담당)', 'kiwhan@ktnet.co.kr\n내선번호 x2759\nsmartflow(김기환)', 'AI 정책 문의·도입 신청·사고 신고'],
  ['인프라운영실', '인프라운영실 김종기 실장', '망 접근 통제·보안 사고 대응'],
  ['CPO', '전자물류사업본부장 고병권 상무', '개인정보 관련 문의'],
], { widths: [2300, 3100, 3006] });
bodyChildren.push(
  rule(C.ruleLight, 4),
  body('이 가이드북은 「인공지능 활용 및 관리 요령」 제4조 ⑤항에 따른 운영 지원 안내자료로서, AI 거버넌스 위원회 검토 후 AX추진실이 발행합니다.'),
  body('내용의 정확성을 위해 주기적으로 업데이트되며, 최신 버전은 AX추진실로 문의하세요.'),
);

const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: '–',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 560, hanging: 280 } } },
          },
        ],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: 20, color: C.darkGray } },
    },
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 25, bold: true, font: FONT, color: C.charcoal },
        paragraph: {
          spacing: { before: 360, after: 140 },
          outlineLevel: 0,
          border: { left: { style: BorderStyle.THICK, size: 20, color: C.amber, space: 12 } },
        },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 21, bold: true, font: FONT, color: C.charcoal },
        paragraph: {
          spacing: { before: 220, after: 90 },
          outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 2 } },
        },
      },
      {
        id: 'Heading3',
        name: 'Heading 3',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 20, bold: true, font: FONT, color: C.midGray },
        paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE.w, height: PAGE.h },
          margin: { top: 2400, bottom: 1800, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 36, color: C.amber, space: 1 } },
          children: [],
        }),
        gap(600),
        paragraph([run('KTNET  한국무역정보통신', { size: 20, color: C.midGray, characterSpacing: 60 })], {
          spacing: { before: 0, after: 60 },
        }),
        paragraph(
          [
            run(' GUIDEBOOK ', {
              bold: true,
              size: 16,
              color: C.white,
              shading: { fill: C.amber, type: ShadingType.CLEAR },
            }),
          ],
          { spacing: { before: 0, after: 280 } },
        ),
        paragraph([run('AI 활용 가이드북', { bold: true, size: 52, color: C.charcoal })], {
          spacing: { before: 0, after: 80 },
        }),
        paragraph([run('AI Utilization Guidebook', { size: 24, color: C.midGray })], {
          spacing: { before: 0, after: 480 },
        }),
        thickRule(),
        gap(240),
        metaTable(),
        gap(340),
        noteCallout('이 가이드북은 「인공지능 활용 및 관리 요령」 제4조 ⑤항에 따른 운영 지원 안내자료로서 법적 구속력이 있는 규정 문서가 아닙니다. 규정과 가이드북이 충돌하는 경우 규정이 우선합니다.'),
        gap(220),
        paragraph([run('한국무역정보통신 (KTNET)', { bold: true, size: 22, color: C.charcoal })], {
          spacing: { before: 0, after: 40 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 8 } },
        }),
        paragraph([run('AX추진실  |  2026년 06월', { size: 18, color: C.midGray })], {
          spacing: { before: 0, after: 0 },
        }),
        pageBreak(),
        gap(200),
        paragraph([run('개정 이력', { bold: true, size: 21, color: C.charcoal })], {
          spacing: { before: 0, after: 120 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.ruleLight, space: 4 } },
        }),
        historyTable(),
        pageBreak(),
      ],
    },
    {
      properties: {
        page: {
          size: { width: PAGE.w, height: PAGE.h },
          margin: { top: PAGE.mTop, bottom: PAGE.mBottom, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      headers: { default: pageHeader },
      footers: { default: pageFooter },
      children: [
        gap(200),
        paragraph([run('목  차', { bold: true, size: 32, color: C.charcoal })], {
          spacing: { before: 0, after: 60 },
        }),
        paragraph([run('Table of Contents', { size: 18, color: C.midGray })], {
          spacing: { before: 0, after: 280 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.amber, space: 2 } },
        }),
        new TableOfContents('목차', { hyperlink: true, headingStyleRange: '1-3' }),
        pageBreak(),
      ],
    },
    {
      properties: {
        page: {
          size: { width: PAGE.w, height: PAGE.h },
          margin: { top: PAGE.mTop, bottom: PAGE.mBottom, left: PAGE.mLeft, right: PAGE.mRight },
        },
      },
      headers: { default: pageHeader },
      footers: { default: pageFooter },
      children: bodyChildren,
    },
  ],
});

const outputPath = path.join(__dirname, 'AI_활용_가이드북_v0_1_KTNET_표준양식.docx');

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outputPath, buf);
  console.log(`Generated ${outputPath}`);
});
