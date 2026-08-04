-- ===========================================================================
-- Painel ANALISE BOLETOS -- backend no projeto pyrniqluywejmgzqkari
-- (o mesmo do SE2 e do Fluxo Geraldo).
--
-- Modelo dos outros paineis: o index.html do Pages nasce vazio e a carteira
-- so chega depois do login. Aqui ficam as tabelas, o RLS e a funcao de acesso.
--
-- ATENCAO -- dois pontos precisam ser conferidos contra o que ja existe no
-- projeto antes de rodar (marcados com CONFERIR abaixo):
--   1. o nome real da tabela de bloqueios usada pelos outros paineis;
--   2. se a org ja tem o schema `private` (o Fluxo Geraldo usa).
-- Rodar em uma transacao para nao deixar meio caminho.
-- ===========================================================================

begin;

create schema if not exists private;

-- ---------------------------------------------------------------------------
-- 1. Quem pode ver ESTE painel
-- ---------------------------------------------------------------------------
create table if not exists public.analise_boletos_autorizados (
  email      text primary key,
  incluido_em timestamptz not null default now()
);

alter table public.analise_boletos_autorizados enable row level security;

-- Ninguem le essa lista pelo Data API: ela so e consultada de dentro da
-- funcao de acesso, que e SECURITY DEFINER.
revoke all on public.analise_boletos_autorizados from anon, authenticated;

insert into public.analise_boletos_autorizados (email)
values ('lucas.abnersd@gmail.com')
on conflict (email) do nothing;

-- ---------------------------------------------------------------------------
-- 2. Funcao de acesso (modelo do private.fluxo_geraldo_acesso)
--
-- Fica no schema `private` de proposito: uma funcao SECURITY DEFINER em
-- `public` seria chamavel por anon/authenticated, porque o Postgres concede
-- EXECUTE a PUBLIC por padrao.
-- ---------------------------------------------------------------------------
create or replace function private.analise_boletos_acesso()
returns boolean
language sql
security definer
set search_path = ''
stable
as $$
  select exists (
    select 1
      from auth.users u
      join auth.sessions s on s.user_id = u.id
     where u.id = (select auth.uid())
       and u.banned_until is null
       and u.deleted_at  is null
       and s.id = ((select auth.jwt()) ->> 'session_id')::uuid
       -- sessao real de no maximo 12 h, como no painel do Geraldo
       and s.created_at > now() - interval '12 hours'
       and lower(u.email) in (
             select lower(email) from public.analise_boletos_autorizados)
       -- CONFERIR: conferir o nome real desta tabela no projeto antes de rodar.
       -- Se nao existir, apagar as 2 linhas abaixo (o resto continua valendo).
       and not exists (
             select 1 from public.seguranca_bloqueios b
              where lower(b.email) = lower(u.email))
  );
$$;

revoke all on function private.analise_boletos_acesso() from public, anon, authenticated;
grant execute on function private.analise_boletos_acesso() to authenticated;

-- ---------------------------------------------------------------------------
-- 3. A carteira (so leitura pelo painel; quem grava e a Edge Function)
-- ---------------------------------------------------------------------------
create table if not exists public.analise_boletos_titulos (
  uuid          text primary key,
  dados         jsonb not null,
  atualizado_em timestamptz not null default now()
);

alter table public.analise_boletos_titulos enable row level security;

revoke all on public.analise_boletos_titulos from anon;
grant select on public.analise_boletos_titulos to authenticated;

drop policy if exists ab_titulos_ler on public.analise_boletos_titulos;
create policy ab_titulos_ler on public.analise_boletos_titulos
  for select to authenticated
  using ( private.analise_boletos_acesso() );

-- Sem policy de INSERT/UPDATE/DELETE de proposito: a carga passa pela Edge
-- Function, que usa service_role por dentro. O token que fica na maquina so
-- chama a funcao -- nao da para escrever direto na tabela.

-- ---------------------------------------------------------------------------
-- 4. Marcacoes -- UMA LINHA POR TITULO
--
-- Nao usar um JSON unico de estado: com duas pessoas mexendo, o ultimo a
-- salvar apagaria a marcacao do outro (licao do painel do Geraldo).
-- ---------------------------------------------------------------------------
create table if not exists public.analise_boletos_marcacoes (
  uuid       text primary key
             references public.analise_boletos_titulos(uuid) on delete cascade,
  feito      boolean not null default false,
  quem       text,
  atualizado_em timestamptz not null default now()
);

alter table public.analise_boletos_marcacoes enable row level security;

revoke all on public.analise_boletos_marcacoes from anon;
grant select, insert, update on public.analise_boletos_marcacoes to authenticated;

drop policy if exists ab_marc_ler on public.analise_boletos_marcacoes;
create policy ab_marc_ler on public.analise_boletos_marcacoes
  for select to authenticated
  using ( private.analise_boletos_acesso() );

drop policy if exists ab_marc_criar on public.analise_boletos_marcacoes;
create policy ab_marc_criar on public.analise_boletos_marcacoes
  for insert to authenticated
  with check ( private.analise_boletos_acesso() );

-- UPDATE precisa de USING **e** WITH CHECK: so com USING, daria para gravar
-- uma linha que a propria policy nao aceitaria depois.
drop policy if exists ab_marc_alterar on public.analise_boletos_marcacoes;
create policy ab_marc_alterar on public.analise_boletos_marcacoes
  for update to authenticated
  using      ( private.analise_boletos_acesso() )
  with check ( private.analise_boletos_acesso() );

-- Quem marcou, sem depender do cliente mandar a verdade.
create or replace function private.ab_carimbar_marcacao()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  new.quem := coalesce((select auth.jwt() ->> 'email'), 'desconhecido');
  new.atualizado_em := now();
  return new;
end;
$$;

drop trigger if exists trg_ab_carimbar on public.analise_boletos_marcacoes;
create trigger trg_ab_carimbar
  before insert or update on public.analise_boletos_marcacoes
  for each row execute function private.ab_carimbar_marcacao();

commit;

-- ===========================================================================
-- Depois de rodar, conferir:
--   select private.analise_boletos_acesso();          -- logado  -> true
--   select count(*) from public.analise_boletos_titulos;
-- E rodar os advisors (get_advisors / supabase db advisors) antes de publicar.
-- ===========================================================================
