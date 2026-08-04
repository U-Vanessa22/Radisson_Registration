create extension if not exists pgcrypto;

create table if not exists public.workspaces (
  code text primary key,
  name text not null,
  owner_id uuid not null references auth.users(id) on delete cascade,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.workspace_members (
  workspace_code text not null references public.workspaces(code) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null default 'member',
  joined_at timestamptz not null default now(),
  primary key (workspace_code, user_id)
);

create table if not exists public.app_preferences (
  workspace_key text primary key,
  theme text not null default 'light',
  officer_name text not null default 'HR Officer',
  officer_role text not null default 'People & Culture',
  hotel_name text not null default 'Radisson Blu Hotel',
  daily_target integer not null default 18,
  workspace_label text not null default 'radisson-registration',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.employee_directory (
  id uuid primary key default gen_random_uuid(),
  workspace_key text not null,
  employee_id text not null,
  full_name text not null,
  department text not null,
  role text not null,
  phone text not null default '',
  email text not null default '',
  status text not null default 'Active',
  source_file_name text not null default '',
  source_file_path text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_key, employee_id)
);

create table if not exists public.registrations (
  id uuid primary key,
  workspace_key text not null,
  employee_id text,
  full_name text not null,
  department text not null,
  role text not null,
  phone text not null default '',
  shift text not null,
  date date not null,
  registered_at text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.employee_uploads (
  id uuid primary key default gen_random_uuid(),
  workspace_key text not null,
  file_name text not null,
  bucket_path text not null,
  row_count integer not null default 0,
  created_at timestamptz not null default now()
);

create index if not exists employee_directory_workspace_idx on public.employee_directory (workspace_key, full_name);
create index if not exists registrations_workspace_idx on public.registrations (workspace_key, date desc);
create index if not exists employee_uploads_workspace_idx on public.employee_uploads (workspace_key, created_at desc);

insert into storage.buckets (id, name, public)
values ('employee-imports', 'employee-imports', false)
on conflict (id) do nothing;

create or replace function public.ensure_workspace_membership(
  workspace_code text,
  workspace_name text default null
)
returns void
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  current_user_id uuid := auth.uid();
begin
  if current_user_id is null then
    raise exception 'User must be authenticated';
  end if;

  insert into public.workspaces (code, name, owner_id)
  values (
    workspace_code,
    coalesce(nullif(workspace_name, ''), workspace_code),
    current_user_id
  )
  on conflict (code) do nothing;

  insert into public.workspace_members (workspace_code, user_id, role)
  values (workspace_code, current_user_id, 'owner')
  on conflict (workspace_code, user_id) do nothing;
end;
$$;

grant execute on function public.ensure_workspace_membership(text, text) to authenticated;

alter table public.workspaces enable row level security;
alter table public.workspace_members enable row level security;
alter table public.app_preferences enable row level security;
alter table public.employee_directory enable row level security;
alter table public.registrations enable row level security;
alter table public.employee_uploads enable row level security;

drop policy if exists "workspace members can read workspaces" on public.workspaces;
create policy "workspace members can read workspaces"
on public.workspaces
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = workspaces.code
      and members.user_id = auth.uid()
  )
);

drop policy if exists "owners can create workspaces" on public.workspaces;
create policy "owners can create workspaces"
on public.workspaces
for insert
to authenticated
with check (auth.uid() = owner_id);

drop policy if exists "owners can update workspaces" on public.workspaces;
create policy "owners can update workspaces"
on public.workspaces
for update
to authenticated
using (auth.uid() = owner_id)
with check (auth.uid() = owner_id);

drop policy if exists "members can read memberships" on public.workspace_members;
create policy "members can read memberships"
on public.workspace_members
for select
to authenticated
using (user_id = auth.uid());

drop policy if exists "workspace members can read preferences" on public.app_preferences;
create policy "workspace members can read preferences"
on public.app_preferences
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = app_preferences.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can manage preferences" on public.app_preferences;
create policy "workspace members can manage preferences"
on public.app_preferences
for all
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = app_preferences.workspace_key
      and members.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = app_preferences.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can read employees" on public.employee_directory;
create policy "workspace members can read employees"
on public.employee_directory
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_directory.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can manage employees" on public.employee_directory;
create policy "workspace members can manage employees"
on public.employee_directory
for all
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_directory.workspace_key
      and members.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_directory.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can read registrations" on public.registrations;
create policy "workspace members can read registrations"
on public.registrations
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = registrations.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can manage registrations" on public.registrations;
create policy "workspace members can manage registrations"
on public.registrations
for all
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = registrations.workspace_key
      and members.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = registrations.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can read uploads" on public.employee_uploads;
create policy "workspace members can read uploads"
on public.employee_uploads
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_uploads.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can manage uploads" on public.employee_uploads;
create policy "workspace members can manage uploads"
on public.employee_uploads
for all
to authenticated
using (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_uploads.workspace_key
      and members.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.workspace_members members
    where members.workspace_code = employee_uploads.workspace_key
      and members.user_id = auth.uid()
  )
);

drop policy if exists "workspace members can manage employee imports" on storage.objects;
create policy "workspace members can manage employee imports"
on storage.objects
for all
to authenticated
using (
  bucket_id = 'employee-imports'
  and exists (
    select 1
    from public.workspace_members members
    where members.user_id = auth.uid()
      and members.workspace_code = split_part(name, '/', 1)
  )
)
with check (
  bucket_id = 'employee-imports'
  and exists (
    select 1
    from public.workspace_members members
    where members.user_id = auth.uid()
      and members.workspace_code = split_part(name, '/', 1)
  )
);
