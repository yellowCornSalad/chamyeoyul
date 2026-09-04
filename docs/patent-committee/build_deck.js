/**
 * 특허심의위원회 설치·운영(안) — 대표이사 보고 자료 생성 스크립트
 * 실행: node build_deck.js  (pptxgenjs 필요)
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const NAVY = "13294B", NAVY2 = "1F3E63", GOLD = "C8A96A", GOLD_D = "9A7C3E";
const BG = "F5F7FA", WHITE = "FFFFFF", TEXT = "263248", MUTED = "6B7A99", LINE = "DDE3ED";
const F = "맑은 고딕";
const SW = 13.333, SH = 7.5, M = 0.62;
const CW = SW - M * 2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "㈜엔키화이트햇 연구기획팀";
pres.title = "특허심의위원회 설치·운영(안)";

const sh = () => ({ type: "outer", color: "9AABC4", blur: 10, offset: 1, angle: 90, opacity: 0.22 });
let pageNo = 0;

function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.06,
    fill: { color: o.fill || WHITE }, line: { color: o.line || LINE, width: 1 },
    shadow: o.noShadow ? undefined : sh(),
  });
}
function txt(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontFace: F, color: TEXT }, o)); }

function base(title, eyebrow) {
  const s = pres.addSlide();
  pageNo += 1;
  s.background = { color: BG };
  txt(s, eyebrow, { x: M, y: 0.34, w: 8, h: 0.24, fontSize: 10.5, bold: true, color: GOLD_D, charSpacing: 1.6 });
  txt(s, title, { x: M, y: 0.60, w: 10.6, h: 0.52, fontSize: 27, bold: true, color: NAVY });
  s.addShape(pres.ShapeType.roundRect, { x: SW - M - 0.42, y: 0.36, w: 0.42, h: 0.42, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
  txt(s, String(pageNo).padStart(2, "0"), { x: SW - M - 0.42, y: 0.36, w: 0.42, h: 0.42, fontSize: 12, bold: true, color: GOLD, align: "center", valign: "middle" });
  txt(s, "㈜엔키화이트햇  |  연구기획팀", { x: M, y: SH - 0.5, w: 6, h: 0.22, fontSize: 8.5, color: MUTED });
  return s;
}

/* ─── 01 표지 ─────────────────────────────────────────── */
{
  const s = pres.addSlide(); pageNo = 0;
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: 9.9, y: -1.5, w: 5.4, h: 5.4, fill: { color: NAVY2 }, line: { color: NAVY2 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.6, y: 4.7, w: 3.2, h: 3.2, fill: { color: NAVY2 }, line: { color: NAVY2 } });
  txt(s, "연구기획팀 보고  |  2026. 9. 4.", { x: M, y: 1.42, w: 8, h: 0.28, fontSize: 12, bold: true, color: GOLD, charSpacing: 1.6 });
  txt(s, "특허심의위원회 설치·운영(안)", { x: M, y: 1.92, w: 10.4, h: 0.86, fontSize: 42, bold: true, color: WHITE });
  txt(s, "「발명진흥법」에 따른 직무발명 관리체계 정비", { x: M, y: 2.86, w: 10.4, h: 0.42, fontSize: 18, color: "AEC0DC" });
  const chips = [["위원", "6인"], ["정비 서식", "4종"], ["개선 절차", "6단계"], ["1회차 안건", "4건"]];
  chips.forEach((c, i) => {
    const x = M + i * 2.34;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.02, w: 2.1, h: 0.92, rectRadius: 0.07, fill: { color: NAVY2 }, line: { color: "2E5177", width: 1 } });
    txt(s, c[0], { x: x + 0.18, y: 4.16, w: 1.8, h: 0.22, fontSize: 10, color: "AEC0DC" });
    txt(s, c[1], { x: x + 0.18, y: 4.40, w: 1.8, h: 0.38, fontSize: 20, bold: true, color: GOLD });
  });
  txt(s, [
    { text: "지시  ", options: { color: "8FA4C4", fontSize: 11 } },
    { text: "2026. 8. 28.(금) 전성학 부사장", options: { color: WHITE, fontSize: 11, bold: true, breakLine: true } },
    { text: "보고  ", options: { color: "8FA4C4", fontSize: 11 } },
    { text: "연구기획팀 배준호 연구원 (위원회 간사)", options: { color: WHITE, fontSize: 11, bold: true } },
  ], { x: M, y: 5.42, w: 8, h: 0.72, lineSpacing: 20 });
}

/* ─── 02 보고 개요 ─────────────────────────────────────── */
{
  const s = base("보고 개요", "REPORT SUMMARY");
  const items = [
    ["01", "지시 사항", "2026. 8. 28.(금) 전성학 부사장께서 자사의 특허 출원·등록 프로세스\n재정립을 위한 특허심의위원회 구성을 지시"],
    ["02", "설치 목적", "「발명진흥법」 제13조·제15조에 따른 권리 승계·보상 절차를 사내에서\n이행하고, 특허 관리 체계를 상장 기업 수준으로 정비"],
    ["03", "보고 범위", "위원회 구성·운영(안), 개선된 출원·등록 프로세스,\n직무발명 보상규정(안) 및 서식 4종, 제1회 안건"],
  ];
  items.forEach((it, i) => {
    const y = 1.42 + i * 1.42;
    card(s, { x: M, y, w: 7.72, h: 1.22 });
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.26, y: y + 0.32, w: 0.56, h: 0.56, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, it[0], { x: M + 0.26, y: y + 0.32, w: 0.56, h: 0.56, fontSize: 14, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, it[1], { x: M + 1.02, y: y + 0.20, w: 6.4, h: 0.3, fontSize: 15, bold: true, color: NAVY });
    txt(s, it[2], { x: M + 1.02, y: y + 0.55, w: 6.4, h: 0.56, fontSize: 11.5, color: TEXT, lineSpacing: 17 });
  });
  const rx = M + 8.02, rw = CW - 8.02;
  card(s, { x: rx, y: 1.42, w: rw, h: 4.28, fill: NAVY, line: NAVY });
  txt(s, "승인 요청 사항", { x: rx + 0.32, y: 1.72, w: rw - 0.64, h: 0.34, fontSize: 16, bold: true, color: WHITE });
  const asks = [
    ["특허심의위원회 설치·운영(안)", "위원 6인 구성 및 의결 방식"],
    ["직무발명 보상규정(안)", "등급별 보상금 기준 및 서식 4종"],
    ["제1회 위원회 개최 및 안건", "운영 방안·보상 기준·아이디어 발표"],
  ];
  asks.forEach((a, i) => {
    const y = 2.32 + i * 1.05;
    s.addShape(pres.ShapeType.roundRect, { x: rx + 0.32, y: y + 0.04, w: 0.24, h: 0.24, rectRadius: 0.06, fill: { color: GOLD }, line: { color: GOLD } });
    txt(s, a[0], { x: rx + 0.70, y: y, w: rw - 1.02, h: 0.3, fontSize: 12.5, bold: true, color: WHITE });
    txt(s, a[1], { x: rx + 0.70, y: y + 0.32, w: rw - 1.02, h: 0.44, fontSize: 10.5, color: "AEC0DC", lineSpacing: 15 });
  });
  txt(s, "※ 보상규정(안)은 「발명진흥법」 제15조제3항에 따라 확정 전 임직원 협의 절차를 거칩니다.", { x: rx + 0.32, y: 5.16, w: rw - 0.64, h: 0.42, fontSize: 9.5, color: "8FA4C4", lineSpacing: 14 });
  s.addNotes("8/28 전성학 부사장 지시로 특허심의위원회 구성을 준비했습니다. 오늘 보고는 위원회 구성·운영안, 개선 프로세스, 보상규정(안) 승인을 요청드리는 자리입니다.");
}

/* ─── 03 추진 배경 ─────────────────────────────────────── */
{
  const s = base("추진 배경 — 특허 관리 방향의 전환", "BACKGROUND");
  const colW = 5.86;
  const cols = [
    { x: M, title: "기존 (AS-IS)", sub: "양(量) 중심의 출원", accent: MUTED, fill: WHITE,
      rows: ["OFFen의 상품성 향상 및 기술평가를 위한 다수의 특허 등록 추진", "발명 선별 기준 부재 — 사업 기여도와 무관한 출원 혼재", "직무발명보상금 지급 기준 미정립", "특허 연차료·해외 출원에 대한 관리 체계 부재"] },
    { x: M + colW + 0.373, title: "개선 (TO-BE)", sub: "가치 중심의 선별·보상", accent: GOLD_D, fill: WHITE,
      rows: ["사업 기여도가 높은 발명을 선별하여 집중 지원", "직무발명보상금 지급 기준 정립 (등급제 운영)", "특허 연차료 관리 — 등급별 유지·포기 판단", "해외 출원 검토 — 조약우선권 기한 내 의사결정"] },
  ];
  cols.forEach((c) => {
    card(s, { x: c.x, y: 1.42, w: colW, h: 3.72 });
    txt(s, c.title, { x: c.x + 0.3, y: 1.66, w: colW - 0.6, h: 0.32, fontSize: 16, bold: true, color: c.accent === MUTED ? MUTED : NAVY });
    txt(s, c.sub, { x: c.x + 0.3, y: 2.00, w: colW - 0.6, h: 0.26, fontSize: 11, color: c.accent });
    c.rows.forEach((r, i) => {
      const y = 2.46 + i * 0.66;
      s.addShape(pres.ShapeType.roundRect, { x: c.x + 0.3, y: y + 0.07, w: 0.16, h: 0.16, rectRadius: 0.04, fill: { color: c.accent }, line: { color: c.accent } });
      txt(s, r, { x: c.x + 0.62, y: y, w: colW - 0.92, h: 0.56, fontSize: 11.5, color: TEXT, lineSpacing: 16 });
    });
  });
  card(s, { x: M, y: 5.36, w: CW, h: 0.92, fill: NAVY, line: NAVY });
  txt(s, [
    { text: "기대 효과   ", options: { fontSize: 11, color: GOLD, bold: true } },
    { text: "발명 신고부터 보상·해외 출원까지 문서로 증빙되는 절차 확립 → 상장 심사·기술평가·정부 R&D 대응이 가능한 지식재산 관리 체계", options: { fontSize: 13, color: WHITE, bold: true } },
  ], { x: M + 0.34, y: 5.36, w: CW - 0.68, h: 0.92, valign: "middle", lineSpacing: 19 });
}

/* ─── 04 법적 근거 ─────────────────────────────────────── */
{
  const s = base("법적 근거 — 「발명진흥법」 조문별 사내 이행 사항", "LEGAL BASIS");
  const rows = [
    ["제12조", "직무발명 완성사실의 통지", "발명자는 발명 완성 시 지체 없이 회사에 서면 신고", "직무발명 신고서 제출"],
    ["제13조", "승계 여부의 통지", "신고 접수일부터 4개월 이내 승계 여부를 서면 통지 (시행령 제7조)", "특허심의위원회 심의·결과 통지"],
    ["제15조", "직무발명에 대한 보상", "권리 승계 시 정당한 보상, 기준·절차를 규정으로 정하고 협의", "직무발명 보상규정(안) 제정"],
    ["제16조", "출원 유보 시의 보상", "승계 후 출원하지 않거나 포기한 경우에도 보상", "출원유보보상금 지급"],
    ["제19조", "비밀유지의 의무", "출원 시까지 발명 내용 비밀유지 (위반 시 제58조 벌칙)", "신고서 서약 · 양도계약 제7조"],
  ];
  rows.forEach((r, i) => {
    const y = 1.44 + i * 0.94;
    card(s, { x: M, y, w: CW, h: 0.80 });
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.22, y: y + 0.19, w: 0.92, h: 0.42, rectRadius: 0.08, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, r[0], { x: M + 0.22, y: y + 0.19, w: 0.92, h: 0.42, fontSize: 12, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, r[1], { x: M + 1.28, y: y + 0.16, w: 2.5, h: 0.48, fontSize: 12.5, bold: true, color: NAVY, valign: "middle" });
    txt(s, r[2], { x: M + 3.92, y: y + 0.16, w: 4.86, h: 0.48, fontSize: 11, color: TEXT, valign: "middle", lineSpacing: 15 });
    s.addShape(pres.ShapeType.roundRect, { x: M + 8.96, y: y + 0.17, w: 3.14, h: 0.46, rectRadius: 0.08, fill: { color: "EEF2F8" }, line: { color: "D7DFEC", width: 1 } });
    txt(s, r[3], { x: M + 9.08, y: y + 0.17, w: 2.9, h: 0.46, fontSize: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  txt(s, "※ 「발명진흥법」 제13조에 따라 사전 승계 규정이 있는 경우 권리는 발명을 완성한 때부터 회사에 승계되며, 4개월 내 미승계를 통지하지 않으면 승계한 것으로 처리됩니다.", { x: M, y: 6.24, w: CW, h: 0.3, fontSize: 9.5, color: MUTED });
  s.addNotes("제13조의 정식 조문 제목은 '승계 여부의 통지'입니다. 4개월 기한은 발명진흥법 시행령 제7조에 규정되어 있습니다.");
}

/* ─── 05 위원회 구성 및 운영 ────────────────────────────── */
{
  const s = base("위원회 구성 및 운영(안)", "COMMITTEE");
  const mem = [
    ["이성권", "대표이사"], ["전성학", "부사장"], ["최창진", "이사"],
    ["이정민", "실장"], ["이태우", "이사"], ["배준호", "연구원 · 간사"],
  ];
  txt(s, "위원 구성 (총 6인)", { x: M, y: 1.30, w: 5, h: 0.28, fontSize: 13, bold: true, color: NAVY });
  mem.forEach((m, i) => {
    const x = M + i * 2.036;
    card(s, { x, y: 1.64, w: 1.913, h: 1.56 });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.677, y: 1.86, w: 0.56, h: 0.56, fill: { color: i === 5 ? GOLD : NAVY }, line: { color: i === 5 ? GOLD : NAVY } });
    txt(s, m[0].charAt(0), { x: x + 0.677, y: 1.86, w: 0.56, h: 0.56, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, m[0], { x: x + 0.11, y: 2.54, w: 1.693, h: 0.3, fontSize: 14, bold: true, color: NAVY, align: "center" });
    txt(s, m[1], { x: x + 0.11, y: 2.85, w: 1.693, h: 0.26, fontSize: 10.5, color: MUTED, align: "center" });
  });
  txt(s, "운영 방식", { x: M, y: 3.44, w: 5, h: 0.28, fontSize: 13, bold: true, color: NAVY });
  const ops = [
    ["개최", "안건 발생 시 수시 개최"],
    ["진행", "아이디어 발표 15분 + 질의응답 10분"],
    ["의결", "재적위원 과반수 출석, 출석위원 과반수 찬성\n(가부동수 시 위원장 결정)"],
  ];
  ops.forEach((o, i) => {
    const x = M + i * 4.131;
    card(s, { x, y: 3.78, w: 3.83, h: 1.30 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: 4.00, w: 0.72, h: 0.34, rectRadius: 0.08, fill: { color: "EEF2F8" }, line: { color: "D7DFEC", width: 1 } });
    txt(s, o[0], { x: x + 0.28, y: 4.00, w: 0.72, h: 0.34, fontSize: 11, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, o[1], { x: x + 0.28, y: 4.46, w: 3.27, h: 0.56, fontSize: 11.5, color: TEXT, lineSpacing: 16 });
  });
  card(s, { x: M, y: 5.28, w: CW, h: 0.92, fill: "EEF2F8", line: "D7DFEC" });
  txt(s, [
    { text: "역할  ", options: { fontSize: 11, bold: true, color: GOLD_D } },
    { text: "직무발명의 승계·출원 여부 의결  ·  보상 등급(A/B/C) 결정  ·  기존 지식재산권의 연차료 및 유지·포기 판단  ·  해외 출원 대상 선정", options: { fontSize: 12, color: NAVY } },
  ], { x: M + 0.34, y: 5.28, w: CW - 0.68, h: 0.92, valign: "middle", lineSpacing: 18 });
  txt(s, "※ 본 특허심의위원회는 「발명진흥법」 제17조의 직무발명심의위원회(이의 제기 심의 기구)와는 별개의 사내 의사결정 기구입니다.", { x: M, y: 6.36, w: CW, h: 0.3, fontSize: 9.5, color: MUTED });
}

/* ─── 06 개선 프로세스 ─────────────────────────────────── */
{
  const s = base("개선된 특허 출원·등록 프로세스", "PROCESS");
  const steps = [
    ["1", "발명신고서 제출", "발명자 → 연구기획팀\n「발명진흥법」 제12조"],
    ["2", "특허심의위원회 심의", "대면 발표 15분 + Q&A 10분\n가결 시 보상 등급 동시 결정"],
    ["3", "출원 · 등록", "가결 건은 특허법률사무소\n(주)차원에 의뢰"],
    ["4", "양도계약서 날인", "특허를 받을 수 있는 권리 및\n특허권 일체를 회사로 이전"],
    ["5", "직무발명 보상금 지급", "보상규정(안)에 따라\n등급별 출원·등록보상금 지급"],
    ["6", "해외 출원 여부 결정", "국내 출원일부터 12개월 이내\n(조약우선권 주장 기한)"],
  ];
  steps.forEach((st, i) => {
    const x = M + (i % 3) * 4.131, y = 1.40 + Math.floor(i / 3) * 1.86;
    card(s, { x, y, w: 3.83, h: 1.62 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.26, y: y + 0.26, w: 0.5, h: 0.5, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, st[0], { x: x + 0.26, y: y + 0.26, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, st[1], { x: x + 0.88, y: y + 0.30, w: 2.73, h: 0.42, fontSize: 13.5, bold: true, color: NAVY, valign: "middle" });
    txt(s, st[2], { x: x + 0.26, y: y + 0.92, w: 3.35, h: 0.56, fontSize: 11, color: TEXT, lineSpacing: 16 });
    if (i % 3 !== 2) txt(s, "▶", { x: x + 3.83, y: y + 0.62, w: 0.24, h: 0.34, fontSize: 11, color: GOLD, align: "center", valign: "middle" });
  });
  card(s, { x: M, y: 5.20, w: CW, h: 1.02, fill: "EEF2F8", line: "D7DFEC" });
  txt(s, [
    { text: "※ 양도계약서는 발명자가 퇴사 후 해당 발명에 대해 권리를 주장하는 것을 방지하고, 출원·등록 절차에 필요한 협력 의무를 확보하기 위한 것입니다.", options: { fontSize: 10.5, color: NAVY, breakLine: true } },
    { text: "※ 출원 전까지 발명 내용은 비밀유지 대상입니다. (「발명진흥법」 제19조) — 논문·학회·전시 등 사전 공개 시 신규성 상실로 등록이 거절될 수 있습니다.", options: { fontSize: 10.5, color: NAVY } },
  ], { x: M + 0.34, y: 5.20, w: CW - 0.68, h: 1.02, valign: "middle", lineSpacing: 18 });
}

/* ─── 07 보상 체계 ─────────────────────────────────────── */
{
  const s = base("직무발명 보상규정(안) — 보상 체계", "COMPENSATION");
  const hdr = { fill: NAVY, color: WHITE, bold: true, fontSize: 12, align: "center", valign: "middle" };
  const c = (t, o) => Object.assign({ text: t, options: Object.assign({ fontFace: F, fontSize: 11.5, color: TEXT, valign: "middle", align: "center" }, o || {}) });
  const rows = [
    [c("보상 종류", hdr), c("지급 시기", hdr), c("A등급", hdr), c("B등급", hdr), c("C등급", hdr)],
    [c("출원보상금", { bold: true, color: NAVY, align: "left" }), c("국내 출원일"), c("150만원", { bold: true, color: GOLD_D }), c("100만원"), c("50만원")],
    [c("등록보상금", { bold: true, color: NAVY, align: "left" }), c("국내 등록일"), c("150만원", { bold: true, color: GOLD_D }), c("100만원"), c("50만원")],
    [c("출원유보보상금", { bold: true, color: NAVY, align: "left" }), c("유보·포기 결정일"), c("75만원", { bold: true, color: GOLD_D }), c("50만원"), c("25만원")],
    [c("해외출원보상금", { bold: true, color: NAVY, align: "left" }), c("해외 출원일"), c("50만원 / 국가   (등급 무관)", { colspan: 3, fill: "EEF2F8", bold: true, color: NAVY })],
    [c("처분·실시보상금", { bold: true, color: NAVY, align: "left" }), c("수익 발생 연도 종료 후 3개월 내"), c("순수익의 10%   (등급 무관)", { colspan: 3, fill: "EEF2F8", bold: true, color: NAVY })],
  ];
  s.addTable(rows, {
    x: M, y: 1.40, w: CW, colW: [2.55, 3.15, 2.13, 2.13, 2.13], rowH: 0.52,
    border: { type: "solid", color: LINE, pt: 1 }, fill: { color: WHITE }, fontFace: F,
  });
  const notes = [
    ["1건 기준", "발명자가 2명 이상인 경우 발명신고서의 기여율에 따라 분배하며, 합계액은 건당 보상액을 초과하지 않습니다."],
    ["출원보상금", "등록 여부와 무관하게 지급하며, 등록 거절·취하 시에도 회수하지 않습니다."],
    ["해외출원보상금", "발명자 1인당 연간 2개국 한도, 원칙적으로 A등급 발명을 우선 검토합니다."],
    ["세무 처리", "「소득세법」 제12조제3호 어목에 따라 연간 700만원까지 비과세 (지배주주 등 특수관계인 임원 제외)"],
  ];
  notes.forEach((n, i) => {
    const y = 4.66 + i * 0.44;
    s.addShape(pres.ShapeType.roundRect, { x: M, y: y + 0.02, w: 1.62, h: 0.32, rectRadius: 0.07, fill: { color: "EEF2F8" }, line: { color: "D7DFEC", width: 1 } });
    txt(s, n[0], { x: M, y: y + 0.02, w: 1.62, h: 0.32, fontSize: 10, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, n[1], { x: M + 1.78, y: y + 0.02, w: CW - 1.78, h: 0.32, fontSize: 10.5, color: TEXT, valign: "middle" });
  });
}

/* ─── 08 등급 결정 기준 ────────────────────────────────── */
{
  const s = base("보상 등급 결정 기준 (규정(안) 제6조)", "GRADING");
  txt(s, "평가 항목 — 각 항목 1~5점, 20점 만점", { x: M, y: 1.34, w: 7, h: 0.28, fontSize: 13, bold: true, color: NAVY });
  const crit = [
    ["①", "사업 연관성", "자사 제품·서비스(OFFen 등)에 이미 적용되었거나 적용이 확정되어 있는가"],
    ["②", "권리 활용성", "회피설계가 어렵고, 제3자의 침해를 외부에서 탐지·입증할 수 있는가"],
    ["③", "기술 독창성", "선행기술 대비 차별성이 뚜렷하고 등록 가능성이 높은가"],
    ["④", "대외 활용 가치", "기술평가·인증·정부사업 가점·국가 R&D 성과 활용 가치가 있는가"],
  ];
  crit.forEach((cr, i) => {
    const y = 1.72 + i * 0.98;
    card(s, { x: M, y, w: 7.5, h: 0.84 });
    txt(s, cr[0], { x: M + 0.24, y: y + 0.22, w: 0.4, h: 0.4, fontSize: 15, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, cr[1], { x: M + 0.72, y: y + 0.14, w: 2.1, h: 0.3, fontSize: 13, bold: true, color: NAVY });
    txt(s, cr[2], { x: M + 0.72, y: y + 0.45, w: 5.5, h: 0.3, fontSize: 10.5, color: TEXT });
    s.addShape(pres.ShapeType.roundRect, { x: M + 6.62, y: y + 0.22, w: 0.64, h: 0.4, rectRadius: 0.08, fill: { color: "EEF2F8" }, line: { color: "D7DFEC", width: 1 } });
    txt(s, "5점", { x: M + 6.62, y: y + 0.22, w: 0.64, h: 0.4, fontSize: 10.5, bold: true, color: NAVY, align: "center", valign: "middle" });
  });
  const rx = M + 7.8, rw = CW - 7.8;
  txt(s, "등급 환산", { x: rx, y: 1.34, w: 4, h: 0.28, fontSize: 13, bold: true, color: NAVY });
  const grades = [
    ["A", "16점 이상", "핵심 특허", "해외 출원 및 사업화 연계 우선 검토", GOLD],
    ["B", "11 ~ 15점", "사업 연계 특허", "자사 제품·서비스와 직접 연관", "2C5282"],
    ["C", "10점 이하", "방어·포트폴리오 특허", "기술평가 및 권리 방어 목적", MUTED],
  ];
  grades.forEach((g, i) => {
    const y = 1.72 + i * 1.31;
    card(s, { x: rx, y, w: rw, h: 1.17 });
    s.addShape(pres.ShapeType.roundRect, { x: rx + 0.26, y: y + 0.31, w: 0.56, h: 0.56, rectRadius: 0.1, fill: { color: g[4] }, line: { color: g[4] } });
    txt(s, g[0], { x: rx + 0.26, y: y + 0.31, w: 0.56, h: 0.56, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, g[1], { x: rx + 0.96, y: y + 0.20, w: 2.6, h: 0.28, fontSize: 12.5, bold: true, color: NAVY });
    txt(s, g[2], { x: rx + 0.96, y: y + 0.50, w: 2.6, h: 0.26, fontSize: 11, bold: true, color: g[4] === GOLD ? GOLD_D : g[4] });
    txt(s, g[3], { x: rx + 0.96, y: y + 0.76, w: 3.1, h: 0.3, fontSize: 10, color: MUTED });
  });
  txt(s, "※ 각 위원이 항목별로 채점하고 위원 채점의 평균값(소수점 첫째 자리 반올림)으로 등급을 결정하며, 등급 결정 결과와 항목별 점수는 발명자에게 서면 통지합니다.", { x: M, y: 5.72, w: CW, h: 0.3, fontSize: 9.5, color: MUTED });
}

/* ─── 09 정비 서식 4종 ─────────────────────────────────── */
{
  const s = base("정비 완료 서식 4종", "DOCUMENTS");
  const docs = [
    ["01", "직무발명 신고서", "발명진흥법 제12조", "발명 완성사실 통지 · 발명자와 기여율 확정 · 외부 공개 이력 및 국가 R&D 연계 확인"],
    ["02", "직무발명 보상규정(안)", "발명진흥법 제15조", "보상의 종류·금액·지급 방법 규정 · 등급(A/B/C) 및 세무 처리 기준 명문화"],
    ["03", "직무발명 등급 평가표", "보상규정(안) 제6조", "위원별 채점지 · 항목별 점수 기준 · 가결/부결/보류 의견 및 사유 기재"],
    ["04", "특허권 양도계약서", "발명진흥법 제13조", "권리 승계 확인 · 퇴사 후 권리 주장 방지 · 절차 협력 및 비밀유지 의무 확보"],
  ];
  docs.forEach((d, i) => {
    const x = M + i * 3.084;
    card(s, { x, y: 1.44, w: 2.84, h: 3.46 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: 1.74, w: 0.54, h: 0.54, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, d[0], { x: x + 0.28, y: 1.74, w: 0.54, h: 0.54, fontSize: 13, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, d[1], { x: x + 0.28, y: 2.48, w: 2.28, h: 0.6, fontSize: 14, bold: true, color: NAVY, lineSpacing: 20 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: 3.14, w: 1.86, h: 0.3, rectRadius: 0.07, fill: { color: "EEF2F8" }, line: { color: "D7DFEC", width: 1 } });
    txt(s, d[2], { x: x + 0.28, y: 3.14, w: 1.86, h: 0.3, fontSize: 9.5, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, d[3], { x: x + 0.28, y: 3.60, w: 2.28, h: 1.1, fontSize: 10.5, color: TEXT, lineSpacing: 16 });
  });
  card(s, { x: M, y: 5.16, w: CW, h: 1.0, fill: NAVY, line: NAVY });
  txt(s, [
    { text: "검토 필요   ", options: { fontSize: 11, bold: true, color: GOLD } },
    { text: "4종 모두 초안 작성 완료 — 승인 후 사내 규정으로 확정하고, 보상규정(안)은 임직원 협의를 거쳐 시행합니다.", options: { fontSize: 12.5, color: WHITE } },
  ], { x: M + 0.34, y: 5.16, w: CW - 0.68, h: 1.0, valign: "middle", lineSpacing: 18 });
}

/* ─── 10 제1회 안건 ───────────────────────────────────── */
{
  const s = base("제1회 특허심의위원회 안건", "AGENDA (1st)");
  const ag = [
    ["01", "특허심의위원회 추진 목적 및 운영 방안", "설치 배경 · 위원 구성 · 의결 방식 확정"],
    ["02", "기 보유 지식재산권 마감일 알림", "특허 / 출원 건별 연차료 및 기한 공유"],
    ["03", "직무발명보상금 기준(안) 검토", "보상규정(안) 및 등급 평가표 심의"],
    ["04", "제품개발팀 강충현 연구원 아이디어 발표", "발표 15분 + 질의응답 10분 · 승계 여부 의결"],
  ];
  ag.forEach((a, i) => {
    const y = 1.42 + i * 1.06;
    card(s, { x: M, y, w: 8.3, h: 0.92 });
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.24, y: y + 0.21, w: 0.5, h: 0.5, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, a[0], { x: M + 0.24, y: y + 0.21, w: 0.5, h: 0.5, fontSize: 13, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, a[1], { x: M + 0.88, y: y + 0.17, w: 7.2, h: 0.3, fontSize: 13.5, bold: true, color: NAVY });
    txt(s, a[2], { x: M + 0.88, y: y + 0.50, w: 7.2, h: 0.28, fontSize: 11, color: MUTED });
  });
  const rx = M + 8.6, rw = CW - 8.6;
  card(s, { x: rx, y: 1.42, w: rw, h: 2.06, fill: NAVY, line: NAVY });
  txt(s, "특이 사항", { x: rx + 0.3, y: 1.68, w: rw - 0.6, h: 0.3, fontSize: 14, bold: true, color: GOLD });
  txt(s, "제품개발팀 강충현 연구원의 아이디어는\n대표님 구두 컨펌이 완료된 건으로,\n제1회 위원회에서 정식 심의 절차를 거쳐\n승계·출원 및 보상 등급을 의결합니다.", { x: rx + 0.3, y: 2.08, w: rw - 0.6, h: 1.24, fontSize: 11.5, color: WHITE, lineSpacing: 19 });
  card(s, { x: rx, y: 3.68, w: rw, h: 2.02 });
  txt(s, "진행 방식", { x: rx + 0.3, y: 3.92, w: rw - 0.6, h: 0.3, fontSize: 14, bold: true, color: NAVY });
  txt(s, [
    { text: "아이디어 발표 15분", options: { fontSize: 11.5, color: TEXT, bullet: true, breakLine: true } },
    { text: "질의응답 10분", options: { fontSize: 11.5, color: TEXT, bullet: true, breakLine: true } },
    { text: "위원별 평가표 채점 후 등급 결정", options: { fontSize: 11.5, color: TEXT, bullet: true, breakLine: true } },
    { text: "가결 시 (주)차원에 출원 의뢰", options: { fontSize: 11.5, color: TEXT, bullet: true } },
  ], { x: rx + 0.3, y: 4.32, w: rw - 0.6, h: 1.24, paraSpaceAfter: 6 });
}

/* ─── 11 관리 포인트 ──────────────────────────────────── */
{
  const s = base("운영상 관리 포인트", "KEY CONTROLS");
  const pts = [
    ["4개월", "승계 통지 기한", "신고 접수일부터 4개월 이내 승계 여부를 서면 통지해야 하며, 미통지 시 권리 승계를 포기한 것으로 봅니다. (제13조·시행령 제7조)"],
    ["12개월", "조약우선권 기한", "국내 출원일부터 12개월 이내에 해외 출원 여부를 결정해야 우선권 주장이 가능합니다."],
    ["700만원", "보상금 비과세 한도", "연간 700만원까지 비과세되므로, 동일 발명자의 연간 누적 지급액을 지급 전 확인합니다. (「소득세법」)"],
    ["제19조", "출원 전 비밀유지", "논문·학회·전시 등 사전 공개 시 신규성 상실로 거절될 수 있으며, 위반 시 제58조에 따른 처벌 대상입니다."],
    ["제15조", "규정 제정 절차", "보상규정은 임직원과 협의하여 작성하며, 불리하게 변경할 경우 과반수 동의가 필요합니다."],
    ["별도 기구", "이의 제기 대응", "보상에 이의가 있는 경우 「발명진흥법」 제17조의 직무발명심의위원회를 별도로 구성하여 심의합니다."],
  ];
  pts.forEach((p, i) => {
    const x = M + (i % 3) * 4.131, y = 1.44 + Math.floor(i / 3) * 2.28;
    card(s, { x, y, w: 3.83, h: 2.04 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.28, y: y + 0.26, w: 1.5, h: 0.42, rectRadius: 0.08, fill: { color: NAVY }, line: { color: NAVY } });
    txt(s, p[0], { x: x + 0.28, y: y + 0.26, w: 1.5, h: 0.42, fontSize: 12, bold: true, color: GOLD, align: "center", valign: "middle" });
    txt(s, p[1], { x: x + 0.28, y: y + 0.82, w: 3.27, h: 0.3, fontSize: 13.5, bold: true, color: NAVY });
    txt(s, p[2], { x: x + 0.28, y: y + 1.18, w: 3.27, h: 0.72, fontSize: 10.5, color: TEXT, lineSpacing: 15 });
  });
  txt(s, "※ 위 기한은 모두 연구기획팀이 발명 신고 접수 시점부터 관리하며, 위원회 개최 시 마감 임박 건을 안건으로 보고합니다.", { x: M, y: 6.10, w: CW, h: 0.3, fontSize: 9.5, color: MUTED });
}

/* ─── 12 승인 요청 및 후속 조치 ─────────────────────────── */
{
  const s = pres.addSlide(); pageNo += 1;
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x: 10.6, y: -1.8, w: 5.0, h: 5.0, fill: { color: NAVY2 }, line: { color: NAVY2 } });
  txt(s, "REQUEST FOR APPROVAL", { x: M, y: 0.62, w: 8, h: 0.26, fontSize: 10.5, bold: true, color: GOLD, charSpacing: 1.6 });
  txt(s, "승인 요청 및 후속 조치", { x: M, y: 0.90, w: 10, h: 0.54, fontSize: 30, bold: true, color: WHITE });
  const asks = [
    ["01", "특허심의위원회 설치·운영(안) 승인", "위원 6인 구성 · 수시 개최 · 과반수 의결"],
    ["02", "직무발명 보상규정(안) 및 서식 4종 승인", "등급별 보상금 기준 · 신고서 · 평가표 · 양도계약서"],
    ["03", "제1회 특허심의위원회 개최 승인", "안건 4건 (운영 방안 · 기한 관리 · 보상 기준 · 아이디어 심의)"],
  ];
  asks.forEach((a, i) => {
    const y = 1.72 + i * 1.14;
    s.addShape(pres.ShapeType.roundRect, { x: M, y, w: CW, h: 1.0, rectRadius: 0.06, fill: { color: NAVY2 }, line: { color: "2E5177", width: 1 } });
    s.addShape(pres.ShapeType.roundRect, { x: M + 0.3, y: y + 0.26, w: 0.5, h: 0.5, rectRadius: 0.1, fill: { color: GOLD }, line: { color: GOLD } });
    txt(s, a[0], { x: M + 0.3, y: y + 0.26, w: 0.5, h: 0.5, fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, a[1], { x: M + 0.98, y: y + 0.22, w: 10.6, h: 0.32, fontSize: 15, bold: true, color: WHITE });
    txt(s, a[2], { x: M + 0.98, y: y + 0.58, w: 10.6, h: 0.28, fontSize: 11, color: "AEC0DC" });
  });
  txt(s, "후속 조치", { x: M, y: 5.32, w: 4, h: 0.28, fontSize: 12, bold: true, color: GOLD });
  const nx = ["규정(안) 임직원 협의", "제1회 위원회 개최", "규정 제정 · 시행", "발명 신고 상시 접수"];
  nx.forEach((n, i) => {
    const x = M + i * 3.084;
    s.addShape(pres.ShapeType.roundRect, { x, y: 5.68, w: 2.84, h: 0.62, rectRadius: 0.08, fill: { color: NAVY2 }, line: { color: "2E5177", width: 1 } });
    txt(s, `${i + 1}. ${n}`, { x: x + 0.14, y: 5.68, w: 2.56, h: 0.62, fontSize: 11.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    if (i < 3) txt(s, "▶", { x: x + 2.84, y: 5.68, w: 0.22, h: 0.62, fontSize: 10, color: GOLD, align: "center", valign: "middle" });
  });
  txt(s, "㈜엔키화이트햇  |  연구기획팀  |  2026. 9. 4.", { x: M, y: SH - 0.56, w: 8, h: 0.24, fontSize: 9, color: "8FA4C4" });
  s.addNotes("승인해 주시면 보상규정(안)의 임직원 협의 절차를 진행하고, 제1회 특허심의위원회 일정을 확정하여 안내드리겠습니다.");
}

const out = path.join(__dirname, "특허심의위원회_설치운영안_보고.pptx");
pres.writeFile({ fileName: out }).then(() => console.log("saved:", out));
