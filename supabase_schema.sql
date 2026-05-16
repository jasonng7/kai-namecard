create extension if not exists pgcrypto;

create table if not exists public.drive_connections (
  workspace_id uuid primary key,
  google_email text,
  refresh_token text not null,
  scopes text[] default '{}',
  connected_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.namecard_scans (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null,
  drive_folder_id text not null,
  drive_folder_name text,
  drive_file_id text not null,
  drive_file_name text not null,
  drive_file_md5 text,
  drive_modified_time timestamptz,
  drive_mime_type text,
  source_fingerprint text not null,
  processed_at timestamptz not null default now(),
  extraction jsonb not null,
  unique (workspace_id, drive_file_id)
);

create index if not exists namecard_scans_workspace_idx
  on public.namecard_scans (workspace_id, processed_at);

create index if not exists namecard_scans_fingerprint_idx
  on public.namecard_scans (workspace_id, source_fingerprint);

alter table public.drive_connections enable row level security;
alter table public.namecard_scans enable row level security;
