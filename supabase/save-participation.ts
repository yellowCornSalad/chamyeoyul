// Supabase Edge Function: save-participation
// 관리자 비밀번호를 서버에서 검증하고, 통과한 경우에만 버전을 저장한다.
// 브라우저는 이 함수를 통해서만 쓸 수 있고(anon INSERT 차단), 작성자는 서버가 판정해 기록한다.

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, apikey, x-client-info, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });

const b64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const toB64 = (b: Uint8Array) => btoa(String.fromCharCode(...b));

async function pbkdf2(pw: string, salt: Uint8Array) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(pw), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", salt, iterations: 150000, hash: "SHA-256" }, key, 256);
  return toB64(new Uint8Array(bits));
}
function timingSafeEqual(a: string, b: string) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ error: "method" }, 405);

  let body: { pw?: string; blob?: string };
  try { body = await req.json(); } catch { return json({ error: "bad json" }, 400); }
  const pw = body.pw ?? "";
  const blob = body.blob ?? "";

  if (typeof blob !== "string" || blob.length === 0) return json({ error: "blob required" }, 400);
  if (blob.length > 200000) return json({ error: "blob too large" }, 413);
  try { JSON.parse(blob); } catch { return json({ error: "blob must be JSON" }, 400); }

  // 비밀번호 → 신원 (관리자만 통과)
  const admins = JSON.parse(Deno.env.get("ADMINS") ?? "[]") as
    { id: string; name: string; salt: string; hash: string }[];
  let who: { id: string; name: string } | null = null;
  for (const a of admins) {
    const h = await pbkdf2(pw, b64(a.salt));
    if (timingSafeEqual(h, a.hash)) { who = { id: a.id, name: a.name }; break; }
  }
  if (!who) {
    await new Promise((r) => setTimeout(r, 400));           // 무차별 대입 완화
    return json({ error: "수정 권한이 없습니다." }, 403);
  }

  // service_role 로 저장 (작성자는 서버가 판정한 값으로 기록 → 위조 불가)
  const url = Deno.env.get("SUPABASE_URL")!;
  // 신·구 두 방식 모두 지원 (SERVICE_ROLE_KEY 는 deprecated 표기됨)
  let key = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
  if (!key) {
    try {
      const d = JSON.parse(Deno.env.get("SUPABASE_SECRET_KEYS") ?? "{}");
      key = d.secret ?? d.service_role ?? Object.values(d).find((v) => typeof v === "string") as string ?? "";
    } catch { /* ignore */ }
  }
  if (!key) return json({ error: "server key missing" }, 500);
  const res = await fetch(`${url}/rest/v1/participation_versions`, {
    method: "POST",
    headers: {
      apikey: key, Authorization: `Bearer ${key}`,
      "Content-Type": "application/json", Prefer: "return=representation",
    },
    body: JSON.stringify({ blob, editor: who.name, editor_id: who.id }),
  });
  if (!res.ok) return json({ error: "db", detail: await res.text() }, 500);
  const rows = await res.json();
  return json({ id: rows[0]?.id, editor: who.name });
});
