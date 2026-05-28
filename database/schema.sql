create extension if not exists pgcrypto;

create table if not exists public.import_batches (
  id uuid primary key default gen_random_uuid(),
  source text not null,
  source_updated_at timestamptz,
  imported_at timestamptz not null default now(),
  metadata jsonb not null default '{}'::jsonb
);

create table if not exists public.obra_registros (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  data date,
  status text,
  obra text not null,
  bloco text,
  tipo text,
  planejado_m numeric,
  executado_m numeric,
  pv numeric,
  profundidade_m numeric,
  economias_previstas numeric,
  economias_recebidas numeric,
  raw jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.obra_producao_mensal (
  id bigserial primary key,
  obra_registro_id bigint not null references public.obra_registros(id) on delete cascade,
  mes date not null,
  produzido_m numeric not null default 0,
  unique (obra_registro_id, mes)
);

create table if not exists public.eap_producao_mensal (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  ano integer,
  mes integer,
  competencia date,
  eap numeric,
  produzido numeric,
  economias_eap numeric,
  economias_recebidas numeric,
  saldo_mes numeric,
  saldo_economias numeric,
  saldo_acum numeric,
  raw jsonb not null default '{}'::jsonb
);

create table if not exists public.pds_apontamentos (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  data date,
  obra text,
  equipe text,
  atividade text,
  trecho text,
  pv text,
  raw jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.funcionarios (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  mes date,
  setor text,
  equipe text,
  responsavel_direto text,
  matricula text,
  nome text,
  sexo text,
  funcao text,
  admissao date,
  rescisao date,
  status text,
  regime text,
  tipo text,
  salario numeric,
  vr numeric,
  frota numeric,
  combustivel numeric,
  valor_mensal numeric,
  custo_frota numeric,
  categoria_frota text,
  tipo_equipamento text,
  custo_total numeric,
  horario text,
  veiculo text,
  placa text,
  raw jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.medicao_series (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  tipo text not null,
  competencia date,
  descricao text,
  valor numeric,
  acumulado numeric,
  saldo numeric,
  raw jsonb not null default '{}'::jsonb
);

create table if not exists public.almoxarifado_produtos (
  id bigserial primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  codigo text,
  descricao text,
  estoque_atual numeric,
  estoque_minimo numeric,
  entradas numeric,
  saidas numeric,
  quantidade numeric,
  situacao_planilha text,
  raw jsonb not null default '{}'::jsonb
);

create table if not exists public.reclamacoes (
  id_text text primary key,
  batch_id uuid references public.import_batches(id) on delete set null,
  data date,
  obra text,
  morador text,
  endereco text,
  tipo_dano text,
  descricao text,
  valor_estimado numeric,
  valor_pago numeric,
  status text,
  responsavel text,
  observacao text,
  raw jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

create table if not exists public.dashboard_snapshots (
  id bigserial primary key,
  source text not null,
  source_updated_at timestamptz,
  payload jsonb not null,
  created_at timestamptz not null default now()
);

create index if not exists idx_obra_registros_obra on public.obra_registros (obra);
create index if not exists idx_obra_registros_status on public.obra_registros (status);
create index if not exists idx_obra_producao_mensal_mes on public.obra_producao_mensal (mes);
create index if not exists idx_pds_apontamentos_data on public.pds_apontamentos (data);
create index if not exists idx_pds_apontamentos_obra on public.pds_apontamentos (obra);
create index if not exists idx_funcionarios_nome on public.funcionarios (nome);
create index if not exists idx_funcionarios_status on public.funcionarios (status);
create index if not exists idx_funcionarios_mes on public.funcionarios (mes);

alter table public.import_batches enable row level security;
alter table public.obra_registros enable row level security;
alter table public.obra_producao_mensal enable row level security;
alter table public.eap_producao_mensal enable row level security;
alter table public.pds_apontamentos enable row level security;
alter table public.funcionarios enable row level security;
alter table public.medicao_series enable row level security;
alter table public.almoxarifado_produtos enable row level security;
alter table public.reclamacoes enable row level security;
alter table public.dashboard_snapshots enable row level security;

create policy "authenticated read import_batches" on public.import_batches for select to authenticated using (true);
create policy "authenticated read obra_registros" on public.obra_registros for select to authenticated using (true);
create policy "authenticated read obra_producao_mensal" on public.obra_producao_mensal for select to authenticated using (true);
create policy "authenticated read eap_producao_mensal" on public.eap_producao_mensal for select to authenticated using (true);
create policy "authenticated read pds_apontamentos" on public.pds_apontamentos for select to authenticated using (true);
create policy "authenticated read funcionarios" on public.funcionarios for select to authenticated using (true);
create policy "authenticated read medicao_series" on public.medicao_series for select to authenticated using (true);
create policy "authenticated read almoxarifado_produtos" on public.almoxarifado_produtos for select to authenticated using (true);
create policy "authenticated read reclamacoes" on public.reclamacoes for select to authenticated using (true);
create policy "authenticated read dashboard_snapshots" on public.dashboard_snapshots for select to authenticated using (true);
