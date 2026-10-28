-- pg_cron_grn_tab_dispatch.sql
-- The GRN DATA writer (.github/workflows/grn_tab.yml -> grn_tab_writer.py) is STARTED by the database's clock
-- (pg_cron + pg_net), not by GitHub's.
--
-- WHY. The workflow's own schedule ("45 1 * * *" = 07:15 IST) started 4.5-6 h late every day from 15 to 28 Oct 2026,
-- so the tab was rewritten at 11:45-13:05 IST: the ops team read the day-before-yesterday's GRNs in the morning,
-- and GRNs that land in the afternoon (HOT and BigBasket mails ~15:30 IST) waited a full day. On 2026-10-28 someone
-- reported "no data after the 26th". A workflow_dispatch starts in seconds. Keyur, 2026-10-28: "set it up twice a day".
--
-- WHEN (IST): 08:00 (overnight feed loads in, before the ops team starts) and 18:00 (after the afternoon mail feeds
-- and the 3-hourly sheet loads). A run takes ~3-6 min. The workflow's concurrency group queues a run behind one in
-- progress, so the GitHub schedule kept as a fallback (it lands ~12:00 IST) can never overlap a dispatched run.
--
-- TOKEN. The ERP sync's vault secret 'github_bc_sync_dispatch' (o2c/pg_cron_bc_sync_dispatch.sql in
-- bc-supabase-sync): a keyur-tbd fine-grained token with Actions: Read and write, whose repository access must list
-- BOTH keyur-tbd/bc-supabase-sync and keyur-tbd/pod-tracker. 403 "Resource not accessible by personal access token"
-- in the log = pod-tracker is not in that list; 401 = expired (rotate as that file says, it covers both jobs).
--
-- Apply as postgres, whole file, idempotent. Check with:
--   select * from public.grn_tab_dispatch_log order by id desc limit 10;

create extension if not exists pg_net with schema extensions;

create table if not exists public.grn_tab_dispatch_log (
  id             bigserial primary key,
  requested_at   timestamptz not null default now(),
  skipped_reason text,
  net_request_id bigint,
  http_status    integer,
  response       text
);
alter table public.grn_tab_dispatch_log enable row level security;
revoke all on public.grn_tab_dispatch_log from anon, authenticated;
comment on table public.grn_tab_dispatch_log is
  'One row per pg_cron tick of grn-tab-dispatch (pod-tracker/pg_cron_grn_tab_dispatch.sql): the GitHub workflow_dispatch '
  'request for grn_tab.yml, or why it was skipped. http_status 204 = GitHub accepted it; 401/403/404 = the vault token '
  'is wrong, expired or not granted on keyur-tbd/pod-tracker.';

create or replace function public.grn_tab_dispatch()
  returns jsonb
  language plpgsql
  security definer
  set search_path = pg_catalog, public, pg_temp
as $function$
declare
  v_token text;
  v_id    bigint;
begin
  -- earlier requests: pg_net answers asynchronously and keeps responses for 6 hours
  update public.grn_tab_dispatch_log l
     set http_status = r.status_code,
         response = left(coalesce(nullif(r.content, ''), r.error_msg, case when r.timed_out then 'timed out' end), 500)
    from net._http_response r
   where r.id = l.net_request_id and l.http_status is null;

  select ds.decrypted_secret into v_token from vault.decrypted_secrets ds where ds.name = 'github_bc_sync_dispatch';
  if v_token is null then
    insert into public.grn_tab_dispatch_log (skipped_reason) values ('no github_bc_sync_dispatch secret in the vault');
    return jsonb_build_object('dispatched', false, 'reason', 'no token');
  end if;

  -- no inputs: the writer rewrites the current and previous month's trackers
  v_id := net.http_post(
    url := 'https://api.github.com/repos/keyur-tbd/pod-tracker/actions/workflows/grn_tab.yml/dispatches',
    body := jsonb_build_object('ref', 'main'),
    headers := jsonb_build_object('Authorization', 'Bearer ' || v_token,
                                  'Accept', 'application/vnd.github+json',
                                  'X-GitHub-Api-Version', '2022-11-28',
                                  'User-Agent', 'supabase-pg-cron-grn-tab',
                                  'Content-Type', 'application/json'),
    timeout_milliseconds := 15000);
  insert into public.grn_tab_dispatch_log (net_request_id) values (v_id);
  return jsonb_build_object('dispatched', true, 'net_request_id', v_id);
end
$function$;
revoke execute on function public.grn_tab_dispatch() from public, anon, authenticated;

select cron.unschedule(jobname) from cron.job where jobname = 'grn-tab-dispatch';
-- pg_cron's clock is GMT: 02:30 = 08:00 IST, 12:30 = 18:00 IST
select cron.schedule('grn-tab-dispatch', '30 2,12 * * *', $$select public.grn_tab_dispatch()$$);
