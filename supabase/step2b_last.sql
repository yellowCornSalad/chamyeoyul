-- ═══════════════════════════════════════════════════════════
-- 참여율 관리 웹앱 — 2단계 (B) : 맨 마지막에 실행
-- 브라우저의 직접 쓰기를 차단. 웹페이지 배포가 끝난 뒤에 실행할 것!
-- (먼저 실행하면 새 웹페이지가 배포되기 전까지 저장이 막힘)
-- ═══════════════════════════════════════════════════════════

drop policy if exists "append" on public.participation_versions;
revoke insert on public.participation_versions from anon;

-- 읽기는 유지 (내용은 암호문이라 안전)
do $$
begin
  if not exists (select 1 from pg_policies
                 where schemaname='public' and tablename='participation_versions' and policyname='read') then
    create policy "read" on public.participation_versions for select to anon using (true);
  end if;
end $$;

select '쓰기 차단 완료' as 결과,
       (select count(*) from information_schema.table_privileges
         where table_name='participation_versions' and grantee='anon' and privilege_type='INSERT') as anon_INSERT권한_남은수;
