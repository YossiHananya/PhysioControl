-- PhysioControl Supabase Schema
-- IMPORTANT: Create a NEW Supabase project for PhysioControl
-- Then replace SUPA_URL and SUPA_KEY in js/db.js with the new project credentials

create table if not exists public.users (
  id text primary key, name text not null, email text not null unique,
  password text not null, role text not null default 'physio' check (role in ('physio','admin')),
  seniority numeric not null default 0, scope numeric not null default 100,
  avatar text, created_at timestamptz default now());

create table if not exists public.rules (
  id text primary key, name text not null,
  category text not null check (category in ('assessment','followup','general','non-clinical')),
  points numeric not null default 1,
  points_type text not null default 'fixed' check (points_type in ('fixed','timed')),
  icon text default '|', created_at timestamptz default now());

create table if not exists public.logs (
  id text primary key,
  user_id text not null references public.users(id) on delete cascade,
  rule_id text not null references public.rules(id) on delete restrict,
  points numeric not null, date date not null, notes text,
  status text default 'approved' check (status in ('approved','pending','flagged')),
  created_at timestamptz default now());

alter table public.users  enable row level security;
alter table public.rules  enable row level security;
alter table public.logs   enable row level security;
create policy allow_all_users on public.users  for all using (true) with check (true);
create policy allow_all_rules on public.rules  for all using (true) with check (true);
create policy allow_all_logs  on public.logs   for all using (true) with check (true);

insert into public.users (id,name,email,password,role,seniority,scope,avatar)
values ('admin_1','Admin','admin@physiocontrol.com','admin123','admin',0,100,'AD')
on conflict (id) do nothing;
