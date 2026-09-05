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

const toB64 = (b: Uint8Array) => btoa(String.fromCharCode(...b));

/* 과제 메타 — 시트에 없는 값(발주처·사업명·기간)은 여기서 붙인다.
   과제가 추가되면 이 표만 고치면 된다. */
const META: Record<string, Record<string, string | number>> = {
  "암호화트래픽":    {sheet:"암호화트래픽", agency:"IITP", dept:"과기정통부",
    biz:"암호화사이버위협대응기술연구개발", from:"2023-04-01", to:"2027-12-31", y2026:533334000, total:2139668000},
  "복원력":         {sheet:"복원력", agency:"KISA", dept:"과기정통부",
    biz:"정보보호핵심원천기술개발", from:"2024-06-01", to:"2027-12-31", y2026:430000000, total:1453667000},
  "에어갭":         {sheet:"에어갭", agency:"IITP", dept:"과기정통부",
    biz:"정보보호핵심원천기술개발", from:"2024-06-01", to:"2026-12-31", y2026:286000000, total:762667000, note:"올해 종료"},
  "시뮬레이터":      {sheet:"시뮬레이터", agency:"IITP", dept:"과기정통부",
    biz:"정보보호핵심원천기술개발사업", from:"2026-04-01", to:"2029-12-31", y2026:333334000, total:1920002000},
  "보안유망기업":    {sheet:"보안유망기업", agency:"KISA", dept:"과기정통부",
    biz:"2026년 AI 보안 유망기업 육성 지원사업", from:"2026-04-01", to:"2026-12-11", y2026:333334000, total:333334000},
  "고성장클럽":      {sheet:"고성장클럽", agency:"NIPA", dept:"과기정통부",
    biz:"2026 SW고성장클럽 지원 사업", from:"2026-03-01", to:"2026-11-30", y2026:300000000, total:300000000},
  "N2SF(이노티움)":  {sheet:"N2SF이노티움", agency:"KISA", dept:"과기정통부",
    biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:150000000, total:150000000},
  "N2SF(휴네시온)":  {sheet:"N2SF휴네시온", agency:"KISA", dept:"과기정통부",
    biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:100000000, total:100000000},
  "N2SF(프라이빗)":  {sheet:"N2SF프라이빗", agency:"KISA", dept:"과기정통부",
    biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:150000000, total:150000000},
  "N2SF(윈스테크넷)":{sheet:"N2SF윈스테크", agency:"KISA", dept:"과기정통부",
    biz:"2026년 국가 망 보안체계(N2SF) 도입 지원사업", from:"2026-06-01", to:"2026-12-11", y2026:130000000, total:130000000},
  "사이버무력화":    {sheet:"사이버무력화", agency:"국방기술진흥연구소", dept:"방위사업청",
    biz:"사이버 무력화 융합기술(무기체계 패키지형) 사업", from:"2026-02-11", to:"2026-12-31", y2026:323690000, total:532718397,
    note:"이월 209,028,397원 포함 / 올해 종료"},
  "로봇보안":        {sheet:"", agency:"산업통상자원부", dept:"산업통상자원부",
    biz:"제조 핵심기술 보호 AI 로봇/휴머노이드 지능보호·보안 실행 기술 개발", from:"2026-04-01", to:"2029-12-31",
    y2026:250000000, total:1060000000, note:"협약 예정"},
  "ETRI(제안완료)":  {sheet:"", agency:"ETRI", dept:"과기정통부",
    biz:"ETRI 기술 스케일업 R&BD 사업", from:"2026-10-01", to:"2027-09-30", y2026:208000000, total:208000000,
    note:"현금 9,600만 신청(평가) 진행중"},
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
    return {
      name, agency: m.agency, dept: m.dept, biz: m.biz, from: m.from, to: m.to,
      y2026: m.y2026, total: m.total, note: m.note ?? "",
      scheme: GOV_PROJ.has(name) ? "gov" : "rnd", rows, cat,
    };
  });

  const ds = { asof: body.asof || new Date().toISOString().slice(0, 10), projects, samples };
  const plain = JSON.stringify(ds);

  // 참여율 데이터와 같은 방식으로 암호화 (열람 비번으로 복호화 가능)
  const pw = Deno.env.get("SHEET_PW") ?? "";
  if (!pw) return json({ error: "SHEET_PW 미설정" }, 500);
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
  return json({
    ok: true, id: inserted[0]?.id, asof: ds.asof,
    projects: projects.length, withBudget: projects.filter((p) => p.rows.length).length,
    history: history.length, bytes: plain.length,
  });
});
