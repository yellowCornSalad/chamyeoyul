/**
 * 특허심의위원회 설치·운영(안) — 대표이사 보고 자료
 * 사내 표준 양식(엔키화이트햇 회사소개서) 서식 적용
 * 실행: NODE_PATH=<pptxgenjs 경로> node build_deck.js
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "10173A", BLUE = "1C67C7", BLUE_DK = "163A78", GRAY_HD = "7C8BA1";
const BG = "F7F7F7", WHITE = "FFFFFF", LINE = "C9D8EC", TBLHD = "E8F0FA";
const SOFT = "F3F7FB", PILLLN = "A9C4E6", TXT = "333333", MUTED = "7B8492", SEP = "C9CDD6";
const F = "맑은 고딕";
const MX = 1.0, CW = 11.33, SW = 13.333, SH = 7.5;
const A = (f) => path.join(__dirname, "assets", f);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "㈜엔키화이트햇 연구기획팀";
pres.title = "특허심의위원회 설치·운영(안)";

const txt = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontFace: F, color: TXT }, o));
const rect = (s, o) => s.addShape(pres.ShapeType.rect, o);
const rrect = (s, o) => s.addShape(pres.ShapeType.roundRect, o);

/** 본문 슬라이드 골격: 배경 + "대분류 │ 소분류" 제목 + 리드문 + 페이지 번호 */
function page(cat, sub, lead, no) {
  const s = pres.addSlide();
  s.background = { color: BG };
  txt(s, [
    { text: cat, options: { bold: true, fontSize: 21, color: NAVY } },
    { text: "   │   ", options: { fontSize: 18, color: SEP } },
    { text: sub, options: { bold: true, fontSize: 21, color: NAVY } },
  ], { x: MX, y: 0.76, w: CW, h: 0.42, valign: "middle" });
  if (lead) txt(s, lead, { x: MX, y: 1.33, w: CW, h: 0.62, fontSize: 11.5, color: "3A3A3A", lineSpacing: 19 });
  txt(s, String(no), { x: 0, y: 6.98, w: SW, h: 0.24, fontSize: 9, color: "A9AEB8", align: "center" });
  return s;
}
/** 연한 파랑 알약형 구분 띠 */
function pill(s, y, label, x, w) {
  x = x === undefined ? MX : x; w = w === undefined ? CW : w;
  rrect(s, { x, y, w, h: 0.38, rectRadius: 0.19, fill: { color: SOFT }, line: { color: PILLLN, width: 1 } });
  txt(s, label, { x, y, w, h: 0.38, fontSize: 12.5, bold: true, color: BLUE_DK, align: "center", valign: "middle" });
}
/** 번호 원이 있는 진한 네이비 띠 */
function band(s, y, num, label, h) {
  h = h || 0.42;
  rrect(s, { x: MX, y, w: CW, h, rectRadius: h / 2, fill: { color: NAVY }, line: { color: NAVY } });
  if (num) {
    s.addShape(pres.ShapeType.ellipse, { x: MX + 0.16, y: y + (h - 0.28) / 2, w: 0.28, h: 0.28, fill: { color: NAVY }, line: { color: WHITE, width: 1 } });
    txt(s, num, { x: MX + 0.16, y: y + (h - 0.28) / 2, w: 0.28, h: 0.28, fontSize: 9, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  txt(s, label, { x: MX, y, w: CW, h, fontSize: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
}
/** 파란 헤더 바 */
function hbar(s, o) {
  rect(s, { x: o.x, y: o.y, w: o.w, h: o.h || 0.42, fill: { color: o.fill || BLUE }, line: { color: o.fill || BLUE } });
  txt(s, o.t, { x: o.x + 0.08, y: o.y, w: o.w - 0.16, h: o.h || 0.42, fontSize: o.fs || 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
}
/** 흰 본문 박스 */
function wbox(s, o) {
  rect(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill || WHITE }, line: { color: o.line || LINE, width: 1 } });
}
const note = (s, y, t, h) => {
  rect(s, { x: MX, y, w: CW, h: h || 0.62, fill: { color: SOFT }, line: { color: PILLLN, width: 1 } });
  txt(s, t, { x: MX + 0.24, y, w: CW - 0.48, h: h || 0.62, fontSize: 10, color: "44506A", valign: "middle", lineSpacing: 15 });
};
const cell = (t, o) => ({ text: t, options: Object.assign({ fontFace: F, fontSize: 11, color: TXT, align: "center", valign: "middle" }, o || {}) });
const TB = { type: "solid", color: LINE, pt: 1 };

/* ── 1. 표지 ─────────────────────────────────────────── */
{
  const s = pres.addSlide();
  s.addImage({ path: A("bg_cover.jpg"), x: 0, y: 0, w: SW, h: SH });
  txt(s, "특허심의위원회 설치·운영(안)", { x: MX, y: 2.02, w: 8.6, h: 0.72, fontSize: 34, bold: true, color: WHITE });
  txt(s, "「발명진흥법」에 따른 직무발명 관리체계 정비", { x: MX + 0.02, y: 2.86, w: 8.6, h: 0.36, fontSize: 15, color: "AEBBD8" });
  txt(s, [
    { text: "2026. 9. 4.", options: { fontSize: 12, color: "D6DDEC", breakLine: true } },
    { text: "연구기획팀", options: { fontSize: 12, color: "D6DDEC" } },
  ], { x: MX + 0.02, y: 5.92, w: 5, h: 0.62, lineSpacing: 20 });
}

/* ── 2. CONTENTS ─────────────────────────────────────── */
{
  const s = pres.addSlide();
  s.addImage({ path: A("bg_contents.jpg"), x: 0, y: 0, w: SW, h: SH });
  txt(s, "CONTENTS", { x: 1.5, y: 2.12, w: 5, h: 0.7, fontSize: 33, bold: true, color: WHITE });
  txt(s, "특허 출원·등록 프로세스 재정립을 위한\n특허심의위원회 설치·운영 및 직무발명 보상 기준 보고", { x: 1.52, y: 2.94, w: 5.6, h: 0.7, fontSize: 12.5, color: "BFC9DE", lineSpacing: 22 });
  const items = [
    ["1. 추진 배경 및 법적 근거", "특허 관리 방향의 전환 · 발명진흥법 이행 사항"],
    ["2. 위원회 구성 및 운영", "위원 6인 구성 · 개선된 출원·등록 프로세스"],
    ["3. 직무발명 보상 기준", "보상규정(안) · 보상 등급 결정 기준 · 서식 4종"],
    ["4. 제1회 위원회 안건 및 승인 요청", "안건 4건 · 운영상 관리 포인트"],
  ];
  items.forEach((it, i) => {
    const y = 2.28 + i * 1.02;
    rect(s, { x: 8.62, y: y + 0.02, w: 0.028, h: 0.62, fill: { color: "9FB0D0" }, line: { color: "9FB0D0" } });
    txt(s, it[0], { x: 8.84, y: y, w: 4.0, h: 0.32, fontSize: 15, bold: true, color: WHITE });
    txt(s, it[1], { x: 8.84, y: y + 0.34, w: 4.0, h: 0.3, fontSize: 10, color: "A9B4CC" });
  });
}

/* ── 3. 보고 개요 ────────────────────────────────────── */
{
  const s = page("보고 개요", "지시 사항 및 승인 요청",
    "2026. 8. 28.(금) 전성학 부사장의 지시에 따라 특허 출원·등록 프로세스 재정립을 위한 특허심의위원회 구성(안)을 보고드립니다.\n위원회는 「발명진흥법」에 따른 권리 승계·보상 절차를 사내에서 이행하고, 특허 관리 체계를 상장 기업 수준으로 정비하는 것을 목적으로 합니다.", 2);
  const rows = [
    ["지시 사항", "2026. 8. 28.(금) 전성학 부사장 — 자사의 특허 출원·등록 프로세스 재정립을 위한 특허심의위원회 구성 지시"],
    ["설치 목적", "「발명진흥법」 제13조·제15조에 따른 권리 승계 및 보상 절차의 사내 이행, 사업 기여도 중심의 발명 선별 체계 확립"],
    ["보고 범위", "위원회 구성·운영(안), 개선된 출원·등록 프로세스, 직무발명 보상규정(안) 및 서식 4종, 제1회 위원회 안건"],
  ];
  rows.forEach((r, i) => {
    const y = 2.16 + i * 0.60;
    rect(s, { x: MX, y, w: 1.55, h: 0.60, fill: { color: BLUE }, line: { color: BLUE } });
    txt(s, r[0], { x: MX, y, w: 1.55, h: 0.60, fontSize: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    wbox(s, { x: MX + 1.55, y, w: CW - 1.55, h: 0.60 });
    txt(s, r[1], { x: MX + 1.78, y, w: CW - 2.0, h: 0.60, fontSize: 11, color: TXT, valign: "middle" });
  });
  band(s, 4.24, null, "승 인 요 청 사 항");
  const asks = [
    ["특허심의위원회 설치·운영(안)", "위원 6인 구성\n안건 발생 시 수시 개최\n출석위원 과반수 찬성으로 의결"],
    ["직무발명 보상규정(안) 및 서식 4종", "등급별 보상금 지급 기준\n발명신고서 · 등급 평가표\n특허권 양도계약서"],
    ["제1회 특허심의위원회 개최", "운영 방안 및 보상 기준 심의\n기 보유 지식재산권 기한 관리\n제품개발팀 아이디어 심의"],
  ];
  const bw = (CW - 2 * 0.25) / 3;
  asks.forEach((a, i) => {
    const x = MX + i * (bw + 0.25);
    hbar(s, { x, y: 4.92, w: bw, t: a[0], fs: 11 });
    wbox(s, { x, y: 5.34, w: bw, h: 1.02 });
    txt(s, a[1], { x: x + 0.18, y: 5.34, w: bw - 0.36, h: 1.02, fontSize: 10.5, color: TXT, valign: "middle", lineSpacing: 16 });
  });
  txt(s, "※ 직무발명 보상규정(안)은 「발명진흥법」 제15조제3항에 따라 확정 전 임직원 협의 절차를 거칩니다.", { x: MX, y: 6.55, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
  s.addNotes("8/28 부사장님 지시로 위원회 구성을 준비했습니다. 오늘은 구성·운영안, 보상규정(안), 제1회 개최 세 건의 승인을 요청드립니다.");
}

/* ── 4. 추진 배경 ────────────────────────────────────── */
{
  const s = page("추진 배경", "특허 관리 방향의 전환",
    "기존에는 OFFen의 상품성 향상 및 기술평가를 위해 다수의 특허 등록을 추진해 왔으나, 발명 선별 기준과 보상 기준이 정립되어 있지 않았습니다.\n앞으로는 사업 기여도가 높은 발명을 선별하여 집중 지원하고, 보상·연차료·해외 출원까지 일관된 기준으로 관리하고자 합니다.", 3);
  const colW = (CW - 0.33) / 2;
  const cols = [
    { x: MX, fill: GRAY_HD, t: "기존 (AS-IS)   양(量) 중심의 출원",
      rows: ["OFFen 상품성 향상 및 기술평가 목적의 다수 출원 추진", "발명 선별 기준 부재 — 사업 기여도와 무관한 출원 혼재", "직무발명보상금 지급 기준 미정립", "특허 연차료 및 해외 출원 관리 체계 부재"] },
    { x: MX + colW + 0.33, fill: BLUE, t: "개선 (TO-BE)   가치 중심의 선별·보상",
      rows: ["사업 기여도가 높은 발명을 선별하여 집중 지원", "직무발명보상금 지급 기준 정립 (A·B·C 등급제)", "특허 연차료 관리 — 등급별 유지·포기 판단", "해외 출원 검토 — 조약우선권 기한 내 의사결정"] },
  ];
  cols.forEach((c) => {
    hbar(s, { x: c.x, y: 2.16, w: colW, t: c.t, fill: c.fill });
    c.rows.forEach((r, i) => {
      const y = 2.70 + i * 0.68;
      wbox(s, { x: c.x, y, w: colW, h: 0.58 });
      txt(s, r, { x: c.x + 0.22, y, w: colW - 0.44, h: 0.58, fontSize: 11, color: TXT, valign: "middle" });
    });
  });
  band(s, 5.62, null, "발명 신고부터 보상·해외 출원까지 문서로 증빙되는 절차 확립 → 상장 심사·기술평가·정부 R&D 대응이 가능한 지식재산 관리 체계", 0.52);
  txt(s, "※ 특허심의위원회는 발명의 승계·출원 여부와 보상 등급을 함께 의결하여, 선별과 보상이 하나의 절차로 연결되도록 합니다.", { x: MX, y: 6.38, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 5. 법적 근거 ────────────────────────────────────── */
{
  const s = page("추진 배경", "법적 근거 및 사내 이행 사항",
    "특허심의위원회는 「발명진흥법」이 사용자에게 요구하는 권리 승계 통지, 정당한 보상, 비밀유지 의무를 사내 절차로 이행하기 위한 기구입니다.\n각 조문별 이행 사항은 다음과 같으며, 관련 서식은 모두 초안 작성을 완료하였습니다.", 4);
  const hd = { fill: TBLHD, bold: true, color: NAVY, fontSize: 11.5 };
  const rows = [
    [cell("조   문", hd), cell("주요 내용", hd), cell("사내 이행 사항", hd)],
    [cell("제12조\n직무발명 완성사실의 통지", { bold: true, color: NAVY, fontSize: 10.5 }), cell("종업원은 직무발명을 완성한 경우 지체 없이 사용자에게 문서로 통지", { align: "left" }), cell("직무발명 신고서 제출", { color: BLUE, bold: true })],
    [cell("제13조\n승계 여부의 통지", { bold: true, color: NAVY, fontSize: 10.5 }), cell("승계 규정이 있는 경우 발명 완성 시 권리 승계, 미승계 시 4개월 이내 서면 통지 (시행령 제7조)", { align: "left" }), cell("위원회 심의 및 결과 통지", { color: BLUE, bold: true })],
    [cell("제15조\n직무발명에 대한 보상", { bold: true, color: NAVY, fontSize: 10.5 }), cell("권리 승계 시 정당한 보상, 보상 기준은 종업원과 협의하여 규정으로 정함", { align: "left" }), cell("직무발명 보상규정(안) 제정", { color: BLUE, bold: true })],
    [cell("제19조\n비밀유지의 의무", { bold: true, color: NAVY, fontSize: 10.5 }), cell("사용자가 출원할 때까지 발명 내용의 비밀유지 (위반 시 제58조 벌칙)", { align: "left" }), cell("신고서 서약 · 양도계약 제7조", { color: BLUE, bold: true })],
  ];
  s.addTable(rows, { x: MX, y: 2.14, w: CW, colW: [2.35, 5.68, 3.30], rowH: 0.66, border: TB, fill: { color: WHITE }, fontFace: F });
  note(s, 5.70, "※ 사전 승계 규정이 있는 경우 권리는 발명을 완성한 때부터 회사에 승계되며, 접수일부터 4개월 이내에 미승계 사실을 통지하지 않으면 승계한 것으로 처리됩니다.\n※ 보상에 이의가 있는 경우에는 「발명진흥법」 제17조에 따른 직무발명심의위원회를 별도로 구성하여 심의합니다. (특허심의위원회와 별개 기구)", 0.78);
}

/* ── 6. 위원회 구성 및 운영 ──────────────────────────── */
{
  const s = page("위원회 운영", "구성 및 운영 방안",
    "위원회는 대표이사를 포함한 6인으로 구성하며, 연구기획팀이 간사를 맡아 안건 접수·심의 준비·결과 통지를 담당합니다.\n안건 발생 시 수시로 개최하여 발명의 승계·출원 여부와 보상 등급을 함께 의결합니다.", 5);
  pill(s, 2.14, "특허심의위원회 위원 구성 (총 6인)");
  const mem = [["이성권", "대표이사"], ["전성학", "부사장"], ["최창진", "이사"], ["이정민", "실장"], ["이태우", "이사"], ["배준호", "연구원 (간사)"]];
  const mw = (CW - 5 * 0.16) / 6;
  mem.forEach((m, i) => {
    const x = MX + i * (mw + 0.16);
    const dark = i === 0, blue = i === 1;
    wbox(s, { x, y: 2.74, w: mw, h: 0.86, fill: dark ? NAVY : blue ? BLUE : WHITE, line: dark ? NAVY : blue ? BLUE : LINE });
    txt(s, m[0], { x, y: 2.86, w: mw, h: 0.3, fontSize: 13, bold: true, color: dark || blue ? WHITE : NAVY, align: "center" });
    txt(s, m[1], { x, y: 3.16, w: mw, h: 0.26, fontSize: 9.5, color: dark || blue ? "C6D2EA" : MUTED, align: "center" });
  });
  const ops = [
    ["개   최", "안건 발생 시 수시 개최\n연구기획팀이 안건을 접수·상정"],
    ["진   행", "아이디어 발표 15분\n질의응답 10분"],
    ["의   결", "재적위원 과반수 출석, 출석위원 과반수 찬성\n가부동수인 경우 위원장이 결정"],
  ];
  const ow = (CW - 2 * 0.25) / 3;
  ops.forEach((o, i) => {
    const x = MX + i * (ow + 0.25);
    hbar(s, { x, y: 3.94, w: ow, t: o[0] });
    wbox(s, { x, y: 4.36, w: ow, h: 0.86 });
    txt(s, o[1], { x: x + 0.18, y: 4.36, w: ow - 0.36, h: 0.86, fontSize: 10.5, color: TXT, valign: "middle", lineSpacing: 16 });
  });
  note(s, 5.52, [
    { text: "위원회 역할   ", options: { bold: true, color: NAVY } },
    { text: "직무발명의 승계·출원 여부 의결   ·   보상 등급(A·B·C) 결정   ·   기 보유 지식재산권의 연차료 및 유지·포기 판단   ·   해외 출원 대상 선정" },
  ], 0.56);
  txt(s, "※ 본 특허심의위원회는 「발명진흥법」 제17조의 직무발명심의위원회(보상 이의 제기 심의 기구)와는 별개의 사내 의사결정 기구입니다.", { x: MX, y: 6.30, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 7. 개선 프로세스 ────────────────────────────────── */
{
  const s = page("위원회 운영", "개선된 특허 출원·등록 프로세스",
    "발명 신고에서 해외 출원 결정까지 6단계로 표준화하고, 각 단계의 산출물과 기한을 연구기획팀이 관리합니다.\n특히 양도계약서 날인 단계를 신설하여, 퇴사 이후의 권리 분쟁 소지를 사전에 차단합니다.", 6);
  band(s, 2.14, "1", "발명 신고 → 심의 → 출원·등록 → 양도계약 → 보상 → 해외 출원 (6단계)");
  const st = [
    ["① 발명신고서 제출", "발명자 → 연구기획팀\n발명진흥법 제12조"],
    ["② 위원회 심의", "대면 발표 15분\n가결 시 등급 결정"],
    ["③ 출원 · 등록", "특허법률사무소\n(주)차원에 의뢰"],
    ["④ 양도계약 날인", "권리 일체를\n회사로 이전"],
    ["⑤ 보상금 지급", "등급별 출원 ·\n등록보상금 지급"],
    ["⑥ 해외 출원 결정", "국내 출원일부터\n12개월 이내"],
  ];
  const sw = (CW - 5 * 0.22) / 6;
  st.forEach((v, i) => {
    const x = MX + i * (sw + 0.22);
    hbar(s, { x, y: 2.86, w: sw, h: 0.44, t: v[0], fs: 10.5 });
    wbox(s, { x, y: 3.30, w: sw, h: 1.02 });
    txt(s, v[1], { x: x + 0.10, y: 3.30, w: sw - 0.20, h: 1.02, fontSize: 9.5, color: TXT, align: "center", valign: "middle", lineSpacing: 15 });
    if (i < 5) txt(s, "▶", { x: x + sw, y: 2.86, w: 0.22, h: 0.44, fontSize: 10, color: BLUE, align: "center", valign: "middle" });
  });
  const kv = [
    ["산출물", "발명신고서 · 등급 평가표 · 심의 결과 통지서 · 특허권 양도계약서"],
    ["기   한", "승계 여부 통지 4개월 이내 (제13조) · 해외 출원 12개월 이내 (조약우선권)"],
    ["유의사항", "출원 전까지 발명 내용은 비밀유지 대상이며, 논문·학회·전시 등 사전 공개 시 신규성 상실로 등록이 거절될 수 있습니다. (제19조)"],
  ];
  kv.forEach((r, i) => {
    const y = 4.66 + i * 0.56;
    rect(s, { x: MX, y, w: 1.35, h: 0.56, fill: { color: TBLHD }, line: { color: LINE, width: 1 } });
    txt(s, r[0], { x: MX, y, w: 1.35, h: 0.56, fontSize: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    wbox(s, { x: MX + 1.35, y, w: CW - 1.35, h: 0.56 });
    txt(s, r[1], { x: MX + 1.55, y, w: CW - 1.75, h: 0.56, fontSize: 10.5, color: TXT, valign: "middle" });
  });
  txt(s, "※ 양도계약서는 발명자가 퇴사 후 해당 발명에 대해 권리를 주장하는 것을 방지하고, 출원·등록 절차에 필요한 협력 의무를 확보하기 위한 것입니다.", { x: MX, y: 6.44, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 8. 보상 체계 ────────────────────────────────────── */
{
  const s = page("보상 기준", "직무발명 보상규정(안)",
    "「발명진흥법」 제15조에 따라 보상의 종류·금액·지급 방법을 규정으로 정하고, 심의 시 결정된 등급(A·B·C)에 따라 차등 지급합니다.\n보상은 국내 출원 시와 국내 등록 시 두 차례 지급하며, 아래 금액은 특허 1건 기준입니다.", 7);
  pill(s, 2.10, "보상의 종류 및 금액 (특허 1건 기준)");
  const hd = { fill: BLUE, bold: true, color: WHITE, fontSize: 11.5 };
  const nm = { bold: true, color: NAVY, align: "left", fontSize: 11 };
  const sm = { fill: TBLHD, bold: true, color: NAVY, fontSize: 11 };
  const rows = [
    [cell("보상 종류", hd), cell("지급 시기", hd), cell("A등급", hd), cell("B등급", hd), cell("C등급", hd)],
    [cell("출원보상금", nm), cell("국내 출원일"), cell("150만원", { bold: true, color: BLUE }), cell("100만원"), cell("50만원")],
    [cell("등록보상금", nm), cell("국내 등록일"), cell("150만원", { bold: true, color: BLUE }), cell("100만원"), cell("50만원")],
    [cell("합       계", Object.assign({ align: "left" }, sm)), cell("특허 1건 기준", sm), cell("300만원", Object.assign({}, sm, { color: BLUE })), cell("200만원", sm), cell("100만원", sm)],
  ];
  s.addTable(rows, { x: MX, y: 2.62, w: CW, colW: [2.35, 3.10, 1.96, 1.96, 1.96], rowH: 0.54, border: TB, fill: { color: WHITE }, fontFace: F });
  const rule = [
    ["공동발명 분배 (제7조)", "발명자가 2명 이상인 경우 발명신고서에 기재된 기여율에 따라 분배하며, 합계액은 건당 보상액을 초과하지 않습니다."],
    ["출원보상금의 성격 (제5조)", "등록 여부와 무관하게 지급하며, 등록 거절·취하 시에도 회수하지 않습니다. 분할·계속출원은 중복 지급하지 않습니다."],
    ["지급 시기 및 통지 (제9조)", "지급 사유가 발생한 날이 속하는 달의 다음 달 급여 지급일에 지급하고, 등급과 항목별 점수를 서면 통지합니다."],
    ["감액 · 부지급 (제8조)", "허위 기재, 무단 공개로 인한 신규성 상실, 절차 협력 거부의 경우 위원회 의결로 감액하거나 지급하지 않을 수 있습니다."],
  ];
  const rw = (CW - 3 * 0.24) / 4;
  rule.forEach((r, i) => {
    const x = MX + i * (rw + 0.24);
    hbar(s, { x, y: 5.04, w: rw, h: 0.42, t: r[0], fs: 11 });
    wbox(s, { x, y: 5.46, w: rw, h: 0.96 });
    txt(s, r[1], { x: x + 0.16, y: 5.46, w: rw - 0.32, h: 0.96, fontSize: 10, color: TXT, valign: "middle", lineSpacing: 15 });
  });
  txt(s, "※ 보상금은 「소득세법」 제12조제3호 어목에 따라 연간 700만원까지 비과세되며, 재직 중 지급분은 근로소득, 퇴직 후 지급분은 기타소득으로 처리합니다.", { x: MX, y: 6.54, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 9. 등급 결정 기준 ───────────────────────────────── */
{
  const s = page("보상 기준", "보상 등급 결정 기준",
    "각 위원이 4개 항목을 1~5점으로 채점하고, 위원 채점의 평균값(소수점 첫째 자리 반올림)을 합계 점수로 하여 등급을 결정합니다.\n등급 결정 결과와 항목별 점수는 심의 결과 통지서에 기재하여 발명자에게 서면으로 통지합니다.", 8);
  const LW = 6.70, RX = MX + LW + 0.33, RW = CW - LW - 0.33;
  pill(s, 2.14, "평가 항목 (각 1~5점 · 20점 만점)", MX, LW);
  const crit = [
    ["① 사업 연관성", "자사 제품·서비스(OFFen 등)에 이미 적용되었거나\n적용이 확정되어 있는가"],
    ["② 권리 활용성", "회피설계가 어렵고, 제3자의 침해를 외부에서\n탐지·입증할 수 있는가"],
    ["③ 기술 독창성", "선행기술 대비 차별성이 뚜렷하고\n등록 가능성이 높은가"],
    ["④ 대외 활용 가치", "기술평가·인증·정부사업 가점·국가 R&D 성과로\n활용할 수 있는가"],
  ];
  crit.forEach((c, i) => {
    const y = 2.74 + i * 0.78;
    rect(s, { x: MX, y, w: 1.62, h: 0.70, fill: { color: TBLHD }, line: { color: LINE, width: 1 } });
    txt(s, c[0], { x: MX, y, w: 1.62, h: 0.70, fontSize: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    wbox(s, { x: MX + 1.62, y, w: LW - 1.62 - 0.72, h: 0.70 });
    txt(s, c[1], { x: MX + 1.80, y, w: LW - 2.60, h: 0.70, fontSize: 10, color: TXT, valign: "middle", lineSpacing: 14 });
    rect(s, { x: MX + LW - 0.72, y, w: 0.72, h: 0.70, fill: { color: WHITE }, line: { color: LINE, width: 1 } });
    txt(s, "5점", { x: MX + LW - 0.72, y, w: 0.72, h: 0.70, fontSize: 10.5, bold: true, color: BLUE, align: "center", valign: "middle" });
  });
  pill(s, 2.14, "등급 환산", RX, RW);
  const gr = [
    ["A", "16점 이상", "핵심 특허", "해외 출원 및 사업화 연계 우선 검토", NAVY, WHITE],
    ["B", "11 ~ 15점", "사업 연계 특허", "자사 제품·서비스와 직접 연관", BLUE, WHITE],
    ["C", "10점 이하", "방어·포트폴리오 특허", "기술평가 및 권리 방어 목적", WHITE, NAVY],
  ];
  gr.forEach((g, i) => {
    const y = 2.74 + i * 1.04;
    const solid = g[4] !== WHITE;
    wbox(s, { x: RX, y, w: RW, h: 0.94, fill: g[4], line: solid ? g[4] : LINE });
    txt(s, g[0], { x: RX + 0.16, y: y + 0.20, w: 0.54, h: 0.54, fontSize: 22, bold: true, color: solid ? WHITE : BLUE, align: "center", valign: "middle" });
    txt(s, g[1] + "   " + g[2], { x: RX + 0.84, y: y + 0.18, w: RW - 1.0, h: 0.28, fontSize: 12, bold: true, color: solid ? WHITE : NAVY });
    txt(s, g[3], { x: RX + 0.84, y: y + 0.48, w: RW - 1.0, h: 0.28, fontSize: 9.5, color: solid ? "C6D2EA" : MUTED });
  });
  note(s, 6.00, "※ 2점·4점도 부여할 수 있으며, 위 문구는 판단 기준선입니다.   ※ 등급 판정에 위원 간 이견이 있는 경우 위원장이 최종 결정하며, 등록 시점에 사업 적용 상황이 현저히 변경된 경우 의결로 1개 등급 범위에서 조정할 수 있습니다.", 0.56);
}

/* ── 10. 서식 정비 ───────────────────────────────────── */
{
  const s = page("보상 기준", "규정 및 서식 4종 정비", 
    "위원회 운영에 필요한 규정과 서식은 모두 초안 작성을 완료하였으며, 승인 후 사내 규정으로 확정하고자 합니다.\n각 서식은 「발명진흥법」의 요구 사항과 사내 절차 단계에 1:1로 대응하도록 구성하였습니다.", 9);
  const docs = [
    ["직무발명 신고서", "발명진흥법 제12조", "발명 완성사실 통지\n발명자 및 기여율 확정\n외부 공개 이력 확인\n국가 R&D 연계 여부 확인"],
    ["직무발명 보상규정(안)", "발명진흥법 제15조", "보상의 종류·금액 규정\n지급 시기 및 방법\n등급(A·B·C) 기준\n세무 처리 및 퇴직자 보상"],
    ["직무발명 등급 평가표", "보상규정(안) 제6조", "위원별 채점지\n항목별 점수 기준 제시\n가결·부결·보류 의견\n부결 사유 체크리스트"],
    ["특허권 양도계약서", "발명진흥법 제13조", "권리 승계 사실 확인\n퇴사 후 권리 주장 방지\n출원·등록 협력 의무\n비밀유지 의무 (퇴직 후 유효)"],
  ];
  const dw = (CW - 3 * 0.25) / 4;
  docs.forEach((d, i) => {
    const x = MX + i * (dw + 0.25);
    hbar(s, { x, y: 2.16, w: dw, h: 0.46, t: d[0], fs: 12 });
    wbox(s, { x, y: 2.62, w: dw, h: 1.94 });
    rect(s, { x: x + 0.20, y: 2.82, w: dw - 0.40, h: 0.34, fill: { color: TBLHD }, line: { color: LINE, width: 1 } });
    txt(s, d[1], { x: x + 0.20, y: 2.82, w: dw - 0.40, h: 0.34, fontSize: 9.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, d[2], { x: x + 0.24, y: 3.30, w: dw - 0.48, h: 1.12, fontSize: 10, color: TXT, lineSpacing: 17 });
  });
  band(s, 4.90, null, "4종 모두 초안 작성 완료 — 승인 후 사내 규정으로 확정하고, 보상규정(안)은 임직원 협의를 거쳐 시행", 0.52);
  note(s, 5.72, [
    { text: "규정 제정 절차   ", options: { bold: true, color: NAVY } },
    { text: "「발명진흥법」 제15조제3항에 따라 보상 규정의 작성·변경은 임직원과 협의하여야 하며, 임직원에게 불리하게 변경하는 경우에는 적용 대상 임직원 과반수의 동의가 필요합니다.", options: { breakLine: true } },
    { text: "시행 시점   ", options: { bold: true, color: NAVY } },
    { text: "시행일 이후 신고된 직무발명부터 적용하되, 시행일 현재 출원 중이거나 등록된 직무발명에 대하여는 위원회 의결로 적용할 수 있습니다." },
  ], 0.80);
}

/* ── 11. 제1회 안건 ──────────────────────────────────── */
{
  const s = page("제1회 위원회", "안건 및 특이사항",
    "제1회 위원회에서는 운영 방안과 보상 기준을 확정하고, 기 보유 지식재산권의 기한을 공유한 뒤 첫 발명 심의를 진행합니다.\n제품개발팀 강충현 연구원의 아이디어는 대표님 구두 컨펌이 완료된 건으로, 정식 심의 절차를 거쳐 승계·출원 여부를 의결합니다.", 10);
  band(s, 2.14, "1", "제1회 특허심의위원회 안건 (4건)");
  const ag = [
    ["01", "특허심의위원회 추진 목적 및 운영 방안", "설치 배경 · 위원 구성 · 의결 방식 확정"],
    ["02", "기 보유 지식재산권 마감일 알림", "특허 / 출원 건별 연차료 및 기한 공유"],
    ["03", "직무발명보상금 기준(안) 검토", "보상규정(안) 및 등급 평가표 심의"],
    ["04", "제품개발팀 강충현 연구원 아이디어 발표", "발표 15분 + 질의응답 10분 · 승계 여부 의결"],
  ];
  const LW2 = 7.30;
  ag.forEach((a, i) => {
    const y = 2.80 + i * 0.86;
    rect(s, { x: MX, y, w: 0.72, h: 0.74, fill: { color: BLUE }, line: { color: BLUE } });
    txt(s, a[0], { x: MX, y, w: 0.72, h: 0.74, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    wbox(s, { x: MX + 0.72, y, w: LW2 - 0.72, h: 0.74 });
    txt(s, a[1], { x: MX + 0.94, y: y + 0.10, w: LW2 - 1.16, h: 0.30, fontSize: 12.5, bold: true, color: NAVY });
    txt(s, a[2], { x: MX + 0.94, y: y + 0.40, w: LW2 - 1.16, h: 0.26, fontSize: 10, color: MUTED });
  });
  const RX2 = MX + LW2 + 0.33, RW2 = CW - LW2 - 0.33;
  hbar(s, { x: RX2, y: 2.80, w: RW2, t: "특 이 사 항" });
  wbox(s, { x: RX2, y: 3.22, w: RW2, h: 1.14, fill: SOFT, line: PILLLN });
  txt(s, "제품개발팀 강충현 연구원의 아이디어는\n대표님 구두 컨펌이 완료된 건입니다.\n제1회 위원회에서 정식 심의 절차를 거쳐\n승계·출원 및 보상 등급을 의결합니다.", { x: RX2 + 0.20, y: 3.22, w: RW2 - 0.40, h: 1.14, fontSize: 10.5, color: "44506A", valign: "middle", lineSpacing: 16 });
  hbar(s, { x: RX2, y: 4.52, w: RW2, t: "진 행 방 식" });
  wbox(s, { x: RX2, y: 4.94, w: RW2, h: 1.32 });
  txt(s, "아이디어 발표 15분\n질의응답 10분\n위원별 평가표 채점 후 등급 결정\n가결 시 (주)차원에 출원 의뢰", { x: RX2 + 0.24, y: 4.94, w: RW2 - 0.48, h: 1.32, fontSize: 10.5, color: TXT, valign: "middle", lineSpacing: 19 });
  txt(s, "※ 부결 의결 시 신고일부터 4개월 이내에 미승계 사실을 서면으로 통지해야 합니다. (「발명진흥법」 제13조, 시행령 제7조)", { x: MX, y: 6.42, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 12. 관리 포인트 ─────────────────────────────────── */
{
  const s = page("운영 관리", "주요 관리 포인트",
    "위원회 운영 과정에서 기한을 놓치면 권리 자체를 잃거나 보상 절차에 하자가 생길 수 있습니다.\n아래 항목은 연구기획팀이 발명 신고 접수 시점부터 관리하며, 위원회 개최 시 마감 임박 건을 안건으로 보고합니다.", 11);
  const pts = [
    ["승계 통지 기한 (4개월)", "신고 접수일부터 4개월 이내에 승계 여부를 서면 통지해야 하며, 미통지 시 승계를 포기한 것으로 봅니다.\n근거 : 제13조 · 시행령 제7조"],
    ["조약우선권 기한 (12개월)", "국내 출원일부터 12개월 이내에 해외 출원 여부를 결정해야 우선권 주장이 가능합니다.\n근거 : 파리협약 우선권"],
    ["보상금 비과세 한도 (700만원)", "연간 700만원까지 비과세되므로 동일 발명자의 연간 누적 지급액을 지급 전 확인합니다.\n근거 : 「소득세법」 제12조제3호 어목"],
    ["출원 전 비밀유지", "논문·학회·전시 등 사전 공개 시 신규성 상실로 등록이 거절될 수 있으며, 위반 시 처벌 대상입니다.\n근거 : 제19조 · 제58조"],
    ["규정 제정 시 협의 절차", "보상 규정은 임직원과 협의하여 작성하며, 불리하게 변경할 경우 과반수 동의가 필요합니다.\n근거 : 제15조제3항"],
    ["보상 이의 제기 대응", "보상에 이의가 있는 경우 사용자위원·종업원위원 동수의 직무발명심의위원회를 별도로 구성합니다.\n근거 : 제17조 · 제18조"],
  ];
  const pw = (CW - 2 * 0.25) / 3;
  pts.forEach((p, i) => {
    const x = MX + (i % 3) * (pw + 0.25), y = 2.16 + Math.floor(i / 3) * 2.10;
    hbar(s, { x, y, w: pw, h: 0.44, t: p[0], fs: 11.5 });
    wbox(s, { x, y: y + 0.44, w: pw, h: 1.28 });
    txt(s, p[1], { x: x + 0.20, y: y + 0.44, w: pw - 0.40, h: 1.28, fontSize: 10.5, color: TXT, valign: "middle", lineSpacing: 16 });
  });
  txt(s, "※ 마감 임박 건은 위원회 개최 시 별도 안건으로 보고하며, 기한 경과가 우려되는 건은 서면 심의로 대체할 수 있도록 운영 규칙에 반영하겠습니다.", { x: MX, y: 6.44, w: CW, h: 0.28, fontSize: 9.5, color: MUTED });
}

/* ── 13. 승인 요청 (표지 서식) ───────────────────────── */
{
  const s = pres.addSlide();
  s.addImage({ path: A("bg_cover.jpg"), x: 0, y: 0, w: SW, h: SH });
  txt(s, "승인 요청 사항", { x: MX, y: 1.42, w: 7, h: 0.6, fontSize: 30, bold: true, color: WHITE });
  const asks = [
    ["01", "특허심의위원회 설치·운영(안) 승인", "위원 6인 구성 · 안건 발생 시 수시 개최 · 출석위원 과반수 찬성으로 의결"],
    ["02", "직무발명 보상규정(안) 및 서식 4종 승인", "등급별 보상금 기준 · 발명신고서 · 등급 평가표 · 특허권 양도계약서"],
    ["03", "제1회 특허심의위원회 개최 승인", "안건 4건 — 운영 방안 · 기한 관리 · 보상 기준 · 아이디어 심의"],
  ];
  asks.forEach((a, i) => {
    const y = 2.42 + i * 0.86;
    txt(s, a[0], { x: MX, y: y + 0.02, w: 0.5, h: 0.32, fontSize: 15, bold: true, color: "8FA8D8" });
    txt(s, a[1], { x: MX + 0.56, y: y, w: 7.6, h: 0.34, fontSize: 15, bold: true, color: WHITE });
    txt(s, a[2], { x: MX + 0.56, y: y + 0.36, w: 7.6, h: 0.3, fontSize: 10.5, color: "A9B7D4" });
  });
  txt(s, "후속 조치", { x: MX, y: 5.28, w: 3, h: 0.28, fontSize: 11, bold: true, color: "8FA8D8" });
  txt(s, "① 보상규정(안) 임직원 협의    ②  제1회 위원회 개최    ③  규정 제정 및 시행    ④  발명 신고 상시 접수", { x: MX, y: 5.62, w: 8.6, h: 0.3, fontSize: 12, color: "E3E8F4" });
  txt(s, "㈜엔키화이트햇  연구기획팀  ·  2026. 9. 4.", { x: MX, y: 6.48, w: 6, h: 0.28, fontSize: 10.5, color: "9BA9C8" });
  s.addNotes("승인해 주시면 보상규정(안)의 임직원 협의 절차를 진행하고, 제1회 특허심의위원회 일정을 확정하여 안내드리겠습니다.");
}

const out = path.join(__dirname, "특허심의위원회_설치운영안_보고.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("saved:", out));
