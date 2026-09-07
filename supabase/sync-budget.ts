// Supabase Edge Function: sync-budget
// 구글 시트의 Apps Script 가 파싱한 예산/집행 데이터를 받아 암호화해 저장한다.
// Apps Script 는 평문 JSON 만 보내고, 암호화는 서버에서 한다(시트 쪽에 키를 두지 않기 위해).

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, apikey, x-client-info, content-type, x-sync-token",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });

const toB64 = (b: Uint8Array) => {           // 큰 배열 spread 는 위험해 루프로
  let s = "";
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
  return btoa(s);
};

/* 과제 메타 — 시트에 없는 값(발주처·사업명·기간)은 여기서 붙인다.
   과제가 추가되면 이 표만 고치면 된다. */
const META: Record<string, any> = {
  "암호화트래픽": { sheet:"암호화트래픽", agency:"IITP", dept:"과학기술정보통신부", biz:"암호화사이버위협대응기술연구개발", from:"2023-04-01", to:"2027-12-31", y2026:533334000, total:2139668000, title:"ICT융합 공공 서비스·인프라 암호화 사이버위협에 대한 네트워크 행위기반 보안관제 기술 개발", period:"2023.04.01 ~ 2027.12.31 (4년 9개월)", thisYear:"2026.01.01 ~ 2026.12.31", lead:"이성권", pm:"김주원" },
  "복원력": { sheet:"복원력", agency:"KISA", dept:"과학기술정보통신부", biz:"정보보호핵심원천기술개발", from:"2024-06-01", to:"2027-12-31", y2026:430000000, total:1453667000, title:"사이버 위기 대응 능력 및 복원력 시험평가 도구 개발", period:"2024.06.01 ~ 2027.12.31", thisYear:"2026.01.01 ~ 2026.12.31", mgmtNo:"RS-2024-00439139", lead:"정시헌", pm:"윤영서", notes:["신규채용인력(김종성) 퇴사로 신규 1명 추가 투입 필요 (성예린)", "연구지원인력 변경 필요 (성예린)"] },
  "에어갭": { sheet:"에어갭", agency:"IITP", dept:"과학기술정보통신부", biz:"정보보호핵심원천기술개발", from:"2024-06-01", to:"2026-12-31", y2026:286000000, total:762667000, note:"올해 종료", title:"이동통신 및 AirGap 환경에서 스니핑 방지 기술 개발", period:"2024.06.01 ~ 2026.12.31", thisYear:"2026.01.01 ~ 2026.12.31", lead:"조정현", pm:"정수환", notes:["박상준 미지급인건비로 투입되어 있으나, 현물로 변경하여 인건비 재계산 필요 (성예린)", "퇴사예정자(김동준) 제외 협약 변경 필요 (성예린)", "연구지원인력 변경 필요"] },
  "시뮬레이터": { sheet:"시뮬레이터", agency:"IITP", dept:"과학기술정보통신부", biz:"정보보호핵심원천기술개발사업", from:"2026-04-01", to:"2029-12-31", y2026:333334000, total:1920002000, title:"사이버공격 실시간 방어수준 지표기반평가 및 전주기(탐지·분석·대응) 자율보안을 위한 시뮬레이터 개발", period:"2026.04 ~ 2029.12", thisYear:"2026.04 ~ 2026.12", mgmtNo:"2026-IITP-99304", lead:"최창진", pm:"홍성현 → 배준호", notes:["1차년도 정량 성과목표 변경 승인성 협약 변경 진행 예정 (배준호)", "쏘마 인수합병으로 승인성 협약 변경 진행 필요 (7~8월 예정) // 서류 안내 완료 상황 (성예린)", "신규 인력(이주영) 퇴사로 추가 신규 인력 투입 필요", "기존 인력(김동준) 퇴사 예정으로 7월부터 제외 필요", "7월 말 연구비 지급예정", "배준호 추가 투입"] },
  "보안유망기업": { sheet:"보안유망기업", agency:"KISA", dept:"과학기술정보통신부", biz:"2026년 AI 보안 유망기업 육성 지원사업", from:"2026-04-01", to:"2026-12-11", y2026:333334000, total:333334000, title:"멀티 LLM기반의 웹 애플리케이션 보안 취약점 자동 진단 시스템 개발", period:"2026.04.01 ~ 2026.12.11 (8.25개월)", thisYear:"2026.04.01 ~ 2026.12.11", lead:"허현", pm:"김범석", notes:["인력 변경 및 연구비 상세 내용 협약 변경 진행중 (성예린)", "성과 관리 스프레드 시트 별도 운영", "11월까지 사용 (과제는 12월 중순에 끝남)"] },
  "고성장클럽": { sheet:"고성장클럽", agency:"NIPA", dept:"과학기술정보통신부", biz:"2026년 SW고성장클럽 지원 사업", from:"2026-03-01", to:"2026-11-30", y2026:300000000, total:300000000, title:"공격표면 관리 및 침투테스트 자동화 기술", period:"2026.03.01 ~ 2026.11.30 (9개월)", thisYear:"2026.03.01 ~ 2026.11.30", lead:"최창진", pm:"김범석", notes:["성과 지표 중 인력 지표 외국어 지원 비용 성과 달성 포기 예정", "특허 출원비 사우디 상표 출원 건 추가 예정 (성예린)", "11월까지 사용"] },
  "N2SF(이노티움)": { sheet:"N2SF이노티움", agency:"KISA", dept:"과학기술정보통신부", biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:150000000, total:150000000, title:"우정사업본부 대상 N2SF 정보서비스 모델 8(클라우드 기반 통합문서체계) 도입 지원", period:"2026.06.01 ~ 2026.12.11 (6.25개월)", thisYear:"2026.06.01 ~ 2026.12.11", lead:"김연재", pm:"김연재", notes:["KT CLOUD 주관기관 안내에 따라 집행"] },
  "N2SF(휴네시온)": { sheet:"N2SF휴네시온", agency:"KISA", dept:"과학기술정보통신부", biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:100000000, total:100000000, title:"한국부동산원 대상 N2SF 정보서비스 모델 2 (업무환경에서 생성형 AI 활용) 도입 지원", period:"2026.06.01 ~ 2026.12.11 (6.25개월)", thisYear:"2026.06.01 ~ 2026.12.11", lead:"오종찬", pm:"오종찬" },
  "N2SF(프라이빗)": { sheet:"N2SF프라이빗", agency:"KISA", dept:"과학기술정보통신부", biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:150000000, total:150000000, title:"AI가상 발전소 확대를 위한 에너지 특화 클라우드 허브 구축", period:"2026.06.01 ~ 2026.12.11 (6.25개월)", thisYear:"2026.06.01 ~ 2026.12.11", lead:"김연재", pm:"김연재" },
  "N2SF(윈스테크넷)": { sheet:"N2SF윈스테크", agency:"KISA", dept:"과학기술정보통신부", biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:130000000, total:130000000, title:"안전한 연결! 똑똑한 업무! N2SF 기반 지능형 업무환경 구축", period:"2026.06.01 ~ 2026.12.11 (6.25개월)", thisYear:"2026.06.01 ~ 2026.12.11", lead:"오종찬", pm:"오종찬" },
  "사이버무력화": { sheet:"사이버무력화", agency:"국방기술진흥연구소", dept:"방위사업청", biz:"사이버 무력화 융합기술(무기체계 패키지형) 사업", from:"2026-02-11", to:"2026-12-31", y2026:323690000, total:532718397, note:"이월 209,028,397원 포함 시 532,718,397 / 올해 종료", title:"사이버 타겟 침투 및 원격 무력화 기술 개발", period:"2026.02.11 ~ 2026.12.31", thisYear:"2026.02.11 ~ 2026.12.31", mgmtNo:"2026-KRIT-31737", lead:"조정현", pm:"차현수", notes:["마지막 차년도 중도 투입 과제", "정산 시 부가세 납부 필요 (연구비 사용 금액 기준으로 부가세 변동 있음)", "재료비로 SSD 구매 건 협약 변경 필요 (주관기관에 내용 전달 완료)", "협약 변경이 상시 진행되지 않고 주관에서 안내주면 해당 기간에 변경 가능"] },
  "로봇보안": { sheet:"", agency:"산업통상자원부", dept:"산업통상자원부", biz:"제조 핵심기술 보호 AI 로봇/휴머노이드 지능보호·보안 실행 기술 개발", from:"2026-04-01", to:"2029-12-31", y2026:250000000, total:1060000000, note:"협약 예정 · RFP 수정 요청 예정", title:"제조 핵심기술 보호를 위한 AI 로봇/휴머노이드의 지능보호 및 보안 실행 기술 개발", period:"2026.04 ~ 2029.12.31", lead:"최창진", pm:"배준호", notes:["협약 예정", "RFP 수정 요청 예정"] },
  "ETRI 드론(신청중)": { sheet:"", agency:"ETRI", dept:"과기정통부", biz:"ETRI 기술 스케일업 R&BD 사업", from:"2026-10-01", to:"2027-09-30", y2026:208000000, total:208000000, note:"현금 9,600만 신청(평가) 진행중" },
};

/* 비목 체계별 기본 카탈로그 — 실적·편성이 없는 과제도 추천이 되도록 */
const RND: [string,string,string][] = [
  ["직접비","인건비","내부인건비"],["직접비","인건비","외부인건비"],
  ["직접비","활동비","출장비"],["직접비","활동비","회의비"],["직접비","활동비","전문가활용비"],
  ["직접비","활동비","학회참가비"],["직접비","활동비","세미나참가비"],["직접비","활동비","소프트웨어활용비"],
  ["직접비","활동비","클라우드이용료"],["직접비","활동비","사무용품비"],["직접비","활동비","문헌구입비"],
  ["직접비","활동비","시험인증비"],["직접비","활동비","특허출원비"],["직접비","활동비","야근식대"],
  ["직접비","활동비","인쇄비"],["직접비","재료비","재료비"],["직접비","수당","연구수당"],
  ["간접비","지원인력비","행정인력인건비"],["간접비","성과활용지원비","특허출원비"],
];
const GOV: [string,string,string][] = [
  ["인건비","인건비","내부인건비"],
  ["운영비","일반수용비","사무용품비"],["운영비","일반수용비","학회참가비"],
  ["운영비","일반수용비","전문가활용비"],["운영비","일반수용비","시험인증비"],
  ["운영비","일반수용비","인쇄비"],["운영비","일반수용비","전시회부스참가"],
  ["운영비","임차료","소프트웨어"],["운영비","공공요금및제세","클라우드이용료"],
  ["운영비","일반용역비","컨설팅비"],["운영비","일반용역비","특허비"],
  ["여비","국내여비","출장비"],["여비","국외여비","국외여비"],
  ["업무추진비","사업추진비","회의비"],
];
const GOV_PROJ = new Set(["고성장클럽","N2SF(이노티움)","N2SF(휴네시온)","N2SF(프라이빗)","N2SF(윈스테크넷)"]);

type Row = { bm: string; sm: string; budget: number; used: number; pend: number; left: number };
type Cat = { bm: string; sm: string; ssm: string; src: string; n?: number };

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ error: "method" }, 405);

  const token = req.headers.get("x-sync-token") ?? "";
  const want = Deno.env.get("SYNC_TOKEN") ?? "";
  if (!want || token !== want) {
    await new Promise((r) => setTimeout(r, 400));
    return json({ error: "unauthorized" }, 401);
  }

  let body: {
    asof?: string;
    budgets?: Record<string, Row[]>;          // 시트명 -> 2026 비목 행
    catalog?: Record<string, Cat[]>;          // 시트명 -> 편성 세세목
    history?: { proj: string; bm: string; sm: string; ssm: string; desc: string }[];
  };
  try { body = await req.json(); } catch { return json({ error: "bad json" }, 400); }

  const budgets = body.budgets ?? {};
  const catalog = body.catalog ?? {};
  const history = body.history ?? [];
  if (!Object.keys(budgets).length) return json({ error: "budgets 가 비어 있습니다" }, 400);

  // 시트명 -> 웹 과제명
  const sheet2web: Record<string, string> = {};
  for (const [web, m] of Object.entries(META)) if (m.sheet) sheet2web[String(m.sheet)] = web;

  // 집행 이력에서 (비목,세목,세세목) 빈도
  const hist: Record<string, Record<string, number>> = {};
  const samples: Record<string, string[]> = {};
  for (const h of history) {
    const web = sheet2web[h.proj] ?? h.proj;
    if (!h.ssm) continue;
    const k = `${h.bm}|${h.sm}|${h.ssm}`;
    (hist[web] ??= {})[k] = ((hist[web] ?? {})[k] ?? 0) + 1;
    if (h.desc) {
      const s = (samples[h.ssm] ??= []);
      if (s.length < 6) s.push(h.desc.slice(0, 40));
    }
  }

  const projects = Object.entries(META).map(([name, m]) => {
    const sh = String(m.sheet ?? "");
    const rows: Row[] = sh ? (budgets[sh] ?? []) : [];
    const cat: Cat[] = [];
    const seen = new Set<string>();
    const push = (bm: string, sm: string, ssm: string, src: string, n?: number) => {
      const k = `${bm}|${sm}|${ssm}`;
      if (!ssm || seen.has(k)) return;
      seen.add(k); cat.push(n ? { bm, sm, ssm, src, n } : { bm, sm, ssm, src });
    };
    for (const c of (sh ? (catalog[sh] ?? []) : [])) push(c.bm, c.sm, c.ssm, "편성");
    for (const [k, n] of Object.entries(hist[name] ?? {})) {
      const [bm, sm, ssm] = k.split("|"); push(bm, sm, ssm, "실적", n);
    }
    const pairs = rows.length ? new Set(rows.map((r) => `${r.bm}|${r.sm}`)) : null;
    for (const [bm, sm, ssm] of (GOV_PROJ.has(name) ? GOV : RND)) {
      if (pairs && !pairs.has(`${bm}|${sm}`)) continue;
      push(bm, sm, ssm, "기본");
    }
    const { sheet: _s, ...meta } = m;          // 시트명은 내보내지 않는다
    return { name, ...meta, note: m.note ?? "",
             scheme: GOV_PROJ.has(name) ? "gov" : "rnd", rows, cat };
  });

  const ds = { asof: body.asof || new Date().toISOString().slice(0, 10), projects, samples };
  const plain = JSON.stringify(ds);

  // 참여율 데이터와 같은 방식으로 암호화 (열람 비번으로 복호화 가능)
  /* 이 값은 페이지의 열람 비번과 반드시 같아야 한다(다르면 사이트가 못 읽는다).
     Secrets 로 받으면 오타 하나에 조용히 깨지므로 상수로 고정한다.
     보안상 손해는 없다 — 이 비번은 이미 배포된 페이지 안에 들어 있는 열람용 값이다. */
  const pw = "enki2026";
  const envPw = Deno.env.get("SHEET_PW") ?? "";
  const pwNote = envPw === "" ? "unset" : (envPw === pw ? "match" : "mismatch(무시함)");
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(pw), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey(
    { name: "PBKDF2", salt, iterations: 150000, hash: "SHA-256" },
    base, { name: "AES-GCM", length: 256 }, false, ["encrypt"]);
  const ct = new Uint8Array(await crypto.subtle.encrypt(
    { name: "AES-GCM", iv }, key, new TextEncoder().encode(plain)));
  const blob = JSON.stringify({ salt: toB64(salt), iv: toB64(iv), data: toB64(ct) });

  const url = Deno.env.get("SUPABASE_URL")!;
  let svc = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
  if (!svc) {
    try {
      const d = JSON.parse(Deno.env.get("SUPABASE_SECRET_KEYS") ?? "{}");
      svc = d.secret ?? d.service_role ?? (Object.values(d).find((v) => typeof v === "string") as string) ?? "";
    } catch { /* ignore */ }
  }
  if (!svc) return json({ error: "server key missing" }, 500);

  const res = await fetch(`${url}/rest/v1/budget_snapshots`, {
    method: "POST",
    headers: {
      apikey: svc, Authorization: `Bearer ${svc}`,
      "Content-Type": "application/json", Prefer: "return=representation",
    },
    body: JSON.stringify({
      blob, source: "apps-script", asof: ds.asof,
      n_projects: projects.filter((p) => p.rows.length).length, n_history: history.length,
    }),
  });
  if (!res.ok) return json({ error: "db", detail: await res.text() }, 500);
  const inserted = await res.json();

  /* 오래된 스냅샷 정리 — 페이지는 최신 1건만 읽는다.
     주기 동기화라 그냥 두면 무료 용량(500MB)을 계속 갉아먹는다. */
  const KEEP = 50;
  let pruned = 0;
  try {
    const h = { apikey: svc, Authorization: `Bearer ${svc}` };
    const q = await fetch(
      `${url}/rest/v1/budget_snapshots?select=created_at&order=created_at.desc&limit=1&offset=${KEEP}`,
      { headers: h });
    if (q.ok) {
      const old = await q.json();
      const cut = old[0]?.created_at;
      if (cut) {
        const d = await fetch(
          `${url}/rest/v1/budget_snapshots?created_at=lt.${encodeURIComponent(cut)}`,
          { method: "DELETE", headers: { ...h, Prefer: "return=representation" } });
        if (d.ok) pruned = (await d.json()).length;
      }
    }
  } catch { /* 정리 실패가 동기화를 막지는 않게 */ }
  return json({
    ok: true, id: inserted[0]?.id, asof: ds.asof, sheetPw: pwNote, pruned,
    projects: projects.length, withBudget: projects.filter((p) => p.rows.length).length,
    history: history.length, bytes: plain.length,
  });
});
