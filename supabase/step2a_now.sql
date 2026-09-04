-- ═══════════════════════════════════════════════════════════
-- 참여율 관리 웹앱 — 2단계 (A) : 지금 실행
-- 컬럼 추가 + DB 라벨링. 기존 동작에 영향 없음(안전).
-- ═══════════════════════════════════════════════════════════

-- 작성자 기록용 컬럼 (Edge Function 이 채움 → 클라이언트가 위조 불가)
alter table public.participation_versions
  add column if not exists editor     text,
  add column if not exists editor_id  text;

-- 이 테이블이 무엇인지 표식 (Table Editor 설명에 표시됨)
comment on table public.participation_versions is
  '[참여율 관리] 국가과제 참여연구원 참여율 데이터. 웹앱 https://yellowcornsalad.github.io/chamyeoyul/ 전용. '
  '내용은 AES-256-GCM 암호문이라 여기서 직접 읽을 수 없음. 수정할 때마다 새 행이 쌓이는 append-only 이력. '
  '쓰기는 Edge Function save-participation 을 통해서만 가능(관리자 비밀번호 서버 검증). 담당: 연구기획팀 배준호';

comment on column public.participation_versions.id         is '버전 ID';
comment on column public.participation_versions.created_at is '저장 시각';
comment on column public.participation_versions.blob       is '암호화된 참여율 데이터 (JSON: {salt,iv,data} 또는 {v:2,iv,data})';
comment on column public.participation_versions.editor     is '수정자 이름 — 서버가 비밀번호 검증 후 기록 (위조 불가)';
comment on column public.participation_versions.editor_id  is '수정자 ID (goun / hyerim)';

select '컬럼 추가 완료' as 결과,
       (select count(*) from public.participation_versions) as 저장된_버전수;
