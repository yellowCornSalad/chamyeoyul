-- 예산 집행 현황 스냅샷 (구글 시트 -> Apps Script -> Edge Function -> 여기)
-- 참여율과 같은 방식: 암호문만 저장, append-only, anon 은 읽기만.

create table if not exists public.budget_snapshots (
  id          uuid primary key default gen_random_uuid(),
  created_at  timestamptz not null default now(),
  blob        text not null,                 -- AES-256-GCM 암호문 (enki2026 유도키)
  source      text,                          -- 'apps-script' 등 출처 표시
  asof        text,                          -- 시트 기준일
  n_projects  int,
  n_history   int
);

comment on table  public.budget_snapshots is '연구과제 예산·집행 현황 스냅샷 (참여율 관리 웹). 구글 시트에서 Apps Script 가 밀어넣는다. 본문은 암호문.';
comment on column public.budget_snapshots.blob   is 'AES-256-GCM 암호문. 평문은 과제별 비목 예산/집행/잔여 + 집행 이력 요약.';
comment on column public.budget_snapshots.source is '적재 경로. apps-script = 구글 시트 자동 동기화.';

create index if not exists budget_snapshots_created_idx
  on public.budget_snapshots (created_at desc);

alter table public.budget_snapshots enable row level security;

-- 읽기만 공개 (본문이 암호문이라 안전). 쓰기는 Edge Function 의 service_role 로만.
drop policy if exists "budget read" on public.budget_snapshots;
create policy "budget read" on public.budget_snapshots for select to anon using (true);

revoke insert, update, delete on public.budget_snapshots from anon;
