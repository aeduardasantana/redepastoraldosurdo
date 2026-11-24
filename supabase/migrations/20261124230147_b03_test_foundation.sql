create table public.b03_test_profiles (
  user_id uuid primary key references auth.users(id) on delete cascade deferrable initially immediate,
  display_name text not null check (char_length(display_name) between 1 and 100),
  enabled boolean not null default true,
  created_at timestamptz not null default now()
);
create table public.b03_test_notes (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null default auth.uid() references public.b03_test_profiles(user_id) on delete cascade,
  body text not null check (char_length(body) between 1 and 500),
  created_at timestamptz not null default now()
);
create index b03_test_notes_owner_idx on public.b03_test_notes(owner_id);
alter table public.b03_test_profiles enable row level security;
alter table public.b03_test_notes enable row level security;
revoke all on public.b03_test_profiles, public.b03_test_notes from public, anon, authenticated;
grant select on public.b03_test_profiles to authenticated;
grant select, insert, delete on public.b03_test_notes to authenticated;
grant update(body) on public.b03_test_notes to authenticated;
create policy b03_profile_read_own on public.b03_test_profiles for select to authenticated
  using ((select auth.uid()) = user_id and enabled and coalesce((select (auth.jwt()->>'is_anonymous')::boolean),false)=false);
create policy b03_note_read_own on public.b03_test_notes for select to authenticated
  using ((select auth.uid()) = owner_id and exists (select 1 from public.b03_test_profiles p where p.user_id = owner_id and p.enabled));
create policy b03_note_insert_own on public.b03_test_notes for insert to authenticated
  with check ((select auth.uid()) = owner_id and exists (select 1 from public.b03_test_profiles p where p.user_id = owner_id and p.enabled));
create policy b03_note_update_own on public.b03_test_notes for update to authenticated
  using ((select auth.uid()) = owner_id and exists (select 1 from public.b03_test_profiles p where p.user_id = owner_id and p.enabled))
  with check ((select auth.uid()) = owner_id and exists (select 1 from public.b03_test_profiles p where p.user_id = owner_id and p.enabled));
create policy b03_note_delete_own on public.b03_test_notes for delete to authenticated
  using ((select auth.uid()) = owner_id and exists (select 1 from public.b03_test_profiles p where p.user_id = owner_id and p.enabled));
