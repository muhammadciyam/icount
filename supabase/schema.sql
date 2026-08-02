-- Run this once in your Supabase project's SQL Editor (Database > SQL Editor > New query).
-- Creates the two tables the app needs for shared, cross-device stock counting,
-- and turns on Realtime so all devices see updates instantly.

create table if not exists stock_counts (
  item_id text primary key,
  qty numeric not null,
  counted_by text,
  updated_at timestamptz not null default now()
);

create table if not exists stock_items (
  id text primary key,
  name text not null,
  loc text,
  qty numeric not null default 0,
  cost numeric not null default 0,
  price numeric not null default 0,
  -- 'base' = part of the original bundled catalog, 'custom' = added by
  -- staff (via "+ Add item" or Excel import). "Delete all items" in the
  -- app only ever touches 'custom' rows, so the base catalog can't be
  -- wiped by that bulk action.
  source text not null default 'custom' check (source in ('base', 'custom')),
  updated_at timestamptz not null default now()
);

-- Safe to re-run on a database that already has stock_items without this column.
alter table stock_items
  add column if not exists source text not null default 'custom' check (source in ('base', 'custom'));

-- Row Level Security: the app is already protected by its own staff login
-- screen (not Supabase auth), so we allow the public "anon" key full access
-- to just these two tables. Don't reuse this anon key for anything sensitive.
alter table stock_counts enable row level security;
alter table stock_items enable row level security;

create policy "anon full access" on stock_counts
  for all using (true) with check (true);

create policy "anon full access" on stock_items
  for all using (true) with check (true);

-- Realtime: lets every open tab/device get pushed updates as soon as
-- someone else counts an item or imports items.
alter publication supabase_realtime add table stock_counts;
alter publication supabase_realtime add table stock_items;

-- ============================================================
-- Staff accounts: powers the login screen and the admin-only
-- "manage staff" panel (create / delete / reset password).
--
-- Passwords are hashed with pgcrypto and never exposed to the client.
-- The table has RLS enabled with NO policies, so the public anon key
-- cannot read or write it directly — every access goes through the
-- SECURITY DEFINER functions below, which independently re-check the
-- caller's own email + password before doing anything sensitive. This
-- matters because the anon key is public (it ships in the client JS),
-- so without that re-check anyone could call these functions and
-- claim to be an admin.
-- ============================================================
create extension if not exists pgcrypto with schema extensions;

create table if not exists staff_accounts (
  id uuid primary key default gen_random_uuid(),
  email text not null unique,
  name text not null,
  password_hash text not null,
  role text not null default 'staff' check (role in ('admin', 'staff')),
  created_at timestamptz not null default now()
);

alter table staff_accounts enable row level security;
-- Intentionally no policies — see comment above.

-- Login: returns the account (without the password) on success, or
-- zero rows on failure.
create or replace function staff_login(p_email text, p_password text)
returns table (id uuid, email text, name text, role text)
language sql
security definer
set search_path = public
as $$
  select id, email, name, role
  from staff_accounts
  where email = lower(p_email)
    and password_hash = extensions.crypt(p_password, password_hash);
$$;

-- Directory listing for the "manage staff" panel. No secrets in here
-- (no password column), so it's safe to expose to anyone signed in.
create or replace function staff_list()
returns table (id uuid, email text, name text, role text, created_at timestamptz)
language sql
security definer
set search_path = public
as $$
  select id, email, name, role, created_at from staff_accounts order by created_at asc;
$$;

-- Admin-only: create a new staff or admin account.
create or replace function staff_create_user(
  p_admin_email text,
  p_admin_password text,
  p_new_email text,
  p_new_password text,
  p_new_name text,
  p_new_role text default 'staff'
)
returns table (id uuid, email text, name text, role text)
language plpgsql
security definer
set search_path = public
as $$
begin
  -- Table alias (sa) is required here: this function's OUT parameters are
  -- named id/email/name/role (from `returns table`), which PL/pgSQL turns
  -- into local variables of the same names — an unqualified "email" or
  -- "role" below would collide with those and raise "ambiguous column".
  if not exists (
    select 1 from staff_accounts sa
    where sa.email = lower(p_admin_email)
      and sa.password_hash = extensions.crypt(p_admin_password, sa.password_hash)
      and sa.role = 'admin'
  ) then
    raise exception 'Not authorized';
  end if;

  return query
  insert into staff_accounts (email, name, password_hash, role)
  values (lower(p_new_email), p_new_name, extensions.crypt(p_new_password, extensions.gen_salt('bf')), p_new_role)
  returning staff_accounts.id, staff_accounts.email, staff_accounts.name, staff_accounts.role;
end;
$$;

-- Admin-only: reset another account's password.
create or replace function staff_reset_password(
  p_admin_email text,
  p_admin_password text,
  p_target_id uuid,
  p_new_password text
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if not exists (
    select 1 from staff_accounts
    where email = lower(p_admin_email)
      and password_hash = extensions.crypt(p_admin_password, password_hash)
      and role = 'admin'
  ) then
    raise exception 'Not authorized';
  end if;

  update staff_accounts
  set password_hash = extensions.crypt(p_new_password, extensions.gen_salt('bf'))
  where id = p_target_id;
end;
$$;

-- Admin-only: delete an account. Refuses to delete the last admin so
-- the team can't accidentally lock itself out.
create or replace function staff_delete_user(
  p_admin_email text,
  p_admin_password text,
  p_target_id uuid
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if not exists (
    select 1 from staff_accounts
    where email = lower(p_admin_email)
      and password_hash = extensions.crypt(p_admin_password, password_hash)
      and role = 'admin'
  ) then
    raise exception 'Not authorized';
  end if;

  if (select role from staff_accounts where id = p_target_id) = 'admin'
     and (select count(*) from staff_accounts where role = 'admin') <= 1 then
    raise exception 'Cannot delete the last admin account';
  end if;

  delete from staff_accounts where id = p_target_id;
end;
$$;

-- Seed the admin account (safe to re-run — does nothing if it already exists).
insert into staff_accounts (email, name, password_hash, role)
values ('siyante003@gmail.com', 'Admin', extensions.crypt('229022#', extensions.gen_salt('bf')), 'admin')
on conflict (email) do nothing;

-- The table itself stays locked down; only these functions are callable.
grant execute on function staff_login(text, text) to anon;
grant execute on function staff_list() to anon;
grant execute on function staff_create_user(text, text, text, text, text, text) to anon;
grant execute on function staff_reset_password(text, text, uuid, text) to anon;
grant execute on function staff_delete_user(text, text, uuid) to anon;
