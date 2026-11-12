-- =====================================================================
-- Prorrateo de seguros automotor por centro de costo
-- Tablas con prefijo seg_ en el esquema public.
-- =====================================================================

-- ---------- Usuarios habilitados ----------
create table if not exists public.seg_usuarios (
  email      text primary key,
  nombre     text,
  creado_en  timestamptz not null default now()
);
comment on table public.seg_usuarios is 'Emails autorizados a usar la aplicación (en minúsculas).';

create or replace function public.seg_es_usuario()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1 from public.seg_usuarios u
    where u.email = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

-- ---------- Bases maestras ----------
create table if not exists public.seg_centros_costo (
  codigo       text primary key,
  descripcion  text not null default '',
  habilitado   boolean not null default true
);

create table if not exists public.seg_zonas (
  patente  text primary key,
  zona     text not null
);

create table if not exists public.seg_asignaciones (
  patente         text primary key,
  cc              text not null,
  actualizado_en  timestamptz not null default now(),
  actualizado_por text default (auth.jwt() ->> 'email')
);
comment on table public.seg_asignaciones is 'Último CC conocido por patente; se usa para precargar los períodos nuevos.';

create table if not exists public.seg_bases (
  tipo        text primary key check (tipo in ('centros_costo', 'zonas')),
  archivo     text,
  filas       integer,
  detalle     jsonb not null default '{}'::jsonb,
  cargado_en  timestamptz not null default now(),
  cargado_por text default (auth.jwt() ->> 'email')
);

-- ---------- Períodos (una factura de renovación) ----------
create table if not exists public.seg_periodos (
  id                 text primary key,             -- N° de factura, ej. 0027-00233598
  poliza             text,
  tomador            text,
  vigencia_desde     date,
  vigencia_hasta     date,
  total_certificados integer,
  premio_final       numeric(16,2),
  factura_fecha      date,
  factura_detalle    text,
  cae                text,
  neto               numeric(16,2) not null,
  iva                numeric(16,2) not null default 0,
  percepcion         numeric(16,2) not null default 0,
  otros_impuestos    numeric(16,2) not null default 0,
  total              numeric(16,2) not null,
  metodo             text not null default 'valor' check (metodo in ('valor', 'premio')),
  archivo_poliza     text,
  archivo_factura    text,
  creado_en          timestamptz not null default now(),
  creado_por         text default (auth.jwt() ->> 'email')
);

create table if not exists public.seg_periodo_vehiculos (
  periodo_id  text not null references public.seg_periodos(id) on delete cascade,
  patente     text not null,
  certificado integer,
  marca       text not null default '',
  anio        text,
  valor       numeric(16,2) not null default 0,
  premio      numeric(16,2) not null default 0,
  cc          text,
  nueva       boolean not null default false,
  primary key (periodo_id, patente)
);
create index if not exists seg_periodo_vehiculos_cc_idx on public.seg_periodo_vehiculos (periodo_id, cc);

-- ---------- Prorrateo (mayor resto, al centavo) ----------
create or replace view public.seg_v_prorrateo
with (security_invoker = true) as
with b as (
  select v.*, z.zona, p.metodo,
         case when p.metodo = 'premio' then v.premio else v.valor end as peso,
         sum(case when p.metodo = 'premio' then v.premio else v.valor end) over (partition by v.periodo_id) as peso_total,
         round(p.neto * 100) as c_neto, round(p.iva * 100) as c_iva,
         round(p.percepcion * 100) as c_perc, round(p.otros_impuestos * 100) as c_otros
  from public.seg_periodo_vehiculos v
  join public.seg_periodos p on p.id = v.periodo_id
  left join public.seg_zonas z on z.patente = v.patente
), r as (
  select b.*,
         case when peso_total > 0 then peso / peso_total else 0 end as porcentaje,
         case when peso_total > 0 then c_neto  * peso / peso_total else 0 end as r_neto,
         case when peso_total > 0 then c_iva   * peso / peso_total else 0 end as r_iva,
         case when peso_total > 0 then c_perc  * peso / peso_total else 0 end as r_perc,
         case when peso_total > 0 then c_otros * peso / peso_total else 0 end as r_otros
  from b
), f as (
  select r.*,
    c_neto  - sum(floor(r_neto))  over w as s_neto,
    c_iva   - sum(floor(r_iva))   over w as s_iva,
    c_perc  - sum(floor(r_perc))  over w as s_perc,
    c_otros - sum(floor(r_otros)) over w as s_otros,
    row_number() over (partition by periodo_id order by r_neto  - floor(r_neto)  desc, certificado, patente) as k_neto,
    row_number() over (partition by periodo_id order by r_iva   - floor(r_iva)   desc, certificado, patente) as k_iva,
    row_number() over (partition by periodo_id order by r_perc  - floor(r_perc)  desc, certificado, patente) as k_perc,
    row_number() over (partition by periodo_id order by r_otros - floor(r_otros) desc, certificado, patente) as k_otros
  from r
  window w as (partition by periodo_id)
), m as (
  select f.*,
    ((floor(r_neto)  + case when peso_total > 0 and k_neto  <= s_neto  then 1 else 0 end) / 100)::numeric(16,2) as neto,
    ((floor(r_iva)   + case when peso_total > 0 and k_iva   <= s_iva   then 1 else 0 end) / 100)::numeric(16,2) as iva,
    ((floor(r_perc)  + case when peso_total > 0 and k_perc  <= s_perc  then 1 else 0 end) / 100)::numeric(16,2) as percepcion,
    ((floor(r_otros) + case when peso_total > 0 and k_otros <= s_otros then 1 else 0 end) / 100)::numeric(16,2) as otros_impuestos
  from f
)
select periodo_id, certificado, patente, marca, anio, zona, valor, premio, metodo,
       porcentaje::numeric(12,10) as porcentaje,
       neto, iva, percepcion, otros_impuestos,
       (neto + iva + percepcion + otros_impuestos)::numeric(16,2) as total,
       cc, nueva
from m;

create or replace view public.seg_v_totales_cc
with (security_invoker = true) as
select v.periodo_id,
       coalesce(v.cc, '') as cc,
       c.descripcion,
       c.habilitado,
       count(*)                 as patentes,
       sum(v.neto)              as neto,
       sum(v.iva)               as iva,
       sum(v.percepcion)        as percepcion,
       sum(v.otros_impuestos)   as otros_impuestos,
       sum(v.total)             as total,
       string_agg(distinct v.zona, ', ' order by v.zona) as zonas
from public.seg_v_prorrateo v
left join public.seg_centros_costo c on c.codigo = v.cc
group by v.periodo_id, coalesce(v.cc, ''), c.descripcion, c.habilitado;

-- ---------- Funciones de carga (transaccionales) ----------
create or replace function public.seg_reemplazar_centros_costo(p_items jsonb, p_archivo text)
returns integer
language plpgsql
set search_path = ''
as $$
declare n integer;
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  delete from public.seg_centros_costo where true;
  insert into public.seg_centros_costo (codigo, descripcion, habilitado)
  select distinct on (trim(x->>'codigo')) trim(x->>'codigo'), coalesce(x->>'descripcion', ''), coalesce((x->>'habilitado')::boolean, true)
  from jsonb_array_elements(p_items) x
  where coalesce(trim(x->>'codigo'), '') <> '';
  get diagnostics n = row_count;
  insert into public.seg_bases (tipo, archivo, filas, cargado_en, cargado_por)
  values ('centros_costo', p_archivo, n, now(), auth.jwt() ->> 'email')
  on conflict (tipo) do update set archivo = excluded.archivo, filas = excluded.filas, cargado_en = now(), cargado_por = excluded.cargado_por;
  return n;
end $$;

create or replace function public.seg_reemplazar_zonas(p_items jsonb, p_archivo text, p_conflictos jsonb default '[]'::jsonb)
returns integer
language plpgsql
set search_path = ''
as $$
declare n integer;
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  delete from public.seg_zonas where true;
  insert into public.seg_zonas (patente, zona)
  select distinct on (x->>'patente') x->>'patente', x->>'zona'
  from jsonb_array_elements(p_items) with ordinality as t(x, i)
  where coalesce(x->>'patente', '') <> '' and coalesce(x->>'zona', '') <> ''
  order by x->>'patente', i desc;
  get diagnostics n = row_count;
  insert into public.seg_bases (tipo, archivo, filas, detalle, cargado_en, cargado_por)
  values ('zonas', p_archivo, n, jsonb_build_object('conflictos', p_conflictos), now(), auth.jwt() ->> 'email')
  on conflict (tipo) do update set archivo = excluded.archivo, filas = excluded.filas, detalle = excluded.detalle, cargado_en = now(), cargado_por = excluded.cargado_por;
  return n;
end $$;

-- Guarda (o reemplaza) un período. Conserva los CC ya asignados en ese período
-- y precarga el resto con el último CC conocido de cada patente.
create or replace function public.seg_guardar_periodo(p_periodo jsonb, p_vehiculos jsonb)
returns jsonb
language plpgsql
set search_path = ''
as $$
declare
  v_id text := p_periodo->>'id';
  v_hay_previo boolean;
  v_total integer; v_con_cc integer;
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  if coalesce(v_id, '') = '' then raise exception 'Falta el número de factura'; end if;

  select exists (select 1 from public.seg_asignaciones) into v_hay_previo;

  insert into public.seg_periodos as p (id, poliza, tomador, vigencia_desde, vigencia_hasta, total_certificados, premio_final,
    factura_fecha, factura_detalle, cae, neto, iva, percepcion, otros_impuestos, total, archivo_poliza, archivo_factura)
  values (v_id, p_periodo->>'poliza', p_periodo->>'tomador',
    to_date(nullif(p_periodo->>'vigencia_desde', ''), 'DD/MM/YYYY'), to_date(nullif(p_periodo->>'vigencia_hasta', ''), 'DD/MM/YYYY'),
    (p_periodo->>'total_certificados')::integer, (p_periodo->>'premio_final')::numeric,
    to_date(nullif(p_periodo->>'factura_fecha', ''), 'DD/MM/YYYY'), p_periodo->>'factura_detalle', p_periodo->>'cae',
    (p_periodo->>'neto')::numeric, coalesce((p_periodo->>'iva')::numeric, 0), coalesce((p_periodo->>'percepcion')::numeric, 0),
    coalesce((p_periodo->>'otros_impuestos')::numeric, 0), (p_periodo->>'total')::numeric,
    p_periodo->>'archivo_poliza', p_periodo->>'archivo_factura')
  on conflict (id) do update set
    poliza = excluded.poliza, tomador = excluded.tomador, vigencia_desde = excluded.vigencia_desde, vigencia_hasta = excluded.vigencia_hasta,
    total_certificados = excluded.total_certificados, premio_final = excluded.premio_final, factura_fecha = excluded.factura_fecha,
    factura_detalle = excluded.factura_detalle, cae = excluded.cae, neto = excluded.neto, iva = excluded.iva,
    percepcion = excluded.percepcion, otros_impuestos = excluded.otros_impuestos, total = excluded.total,
    archivo_poliza = excluded.archivo_poliza, archivo_factura = excluded.archivo_factura;

  delete from public.seg_periodo_vehiculos v
  where v.periodo_id = v_id
    and v.patente not in (
      select upper(regexp_replace(x->>'patente', '[^A-Za-z0-9]', '', 'g')) from jsonb_array_elements(p_vehiculos) x);

  insert into public.seg_periodo_vehiculos as v (periodo_id, patente, certificado, marca, anio, valor, premio, cc, nueva)
  select distinct on (n.patente) v_id, n.patente, n.certificado, n.marca, n.anio, n.valor, n.premio, a.cc,
         (v_hay_previo and a.cc is null)
  from (
    select upper(regexp_replace(x->>'patente', '[^A-Za-z0-9]', '', 'g')) as patente,
           (x->>'certificado')::integer as certificado, coalesce(x->>'marca', '') as marca, x->>'anio' as anio,
           coalesce((x->>'valor')::numeric, 0) as valor, coalesce((x->>'premio')::numeric, 0) as premio
    from jsonb_array_elements(p_vehiculos) x
  ) n
  left join public.seg_asignaciones a on a.patente = n.patente
  where n.patente <> ''
  order by n.patente, n.certificado
  on conflict (periodo_id, patente) do update set
    certificado = excluded.certificado, marca = excluded.marca, anio = excluded.anio,
    valor = excluded.valor, premio = excluded.premio,
    cc = coalesce(v.cc, excluded.cc);

  select count(*), count(cc) into v_total, v_con_cc from public.seg_periodo_vehiculos where periodo_id = v_id;
  return jsonb_build_object('id', v_id, 'vehiculos', v_total, 'con_cc', v_con_cc);
end $$;

-- Asigna (o quita, con p_cc nulo o vacío) el CC de varias patentes de un período
-- y lo recuerda para los períodos siguientes.
create or replace function public.seg_asignar_cc(p_periodo text, p_patentes text[], p_cc text)
returns integer
language plpgsql
set search_path = ''
as $$
declare n integer; v_cc text := nullif(trim(p_cc), '');
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  update public.seg_periodo_vehiculos set cc = v_cc, nueva = false
  where periodo_id = p_periodo and patente = any(p_patentes);
  get diagnostics n = row_count;
  if v_cc is null then
    delete from public.seg_asignaciones where patente = any(p_patentes);
  else
    insert into public.seg_asignaciones (patente, cc, actualizado_en, actualizado_por)
    select unnest(p_patentes), v_cc, now(), auth.jwt() ->> 'email'
    on conflict (patente) do update set cc = excluded.cc, actualizado_en = now(), actualizado_por = excluded.actualizado_por;
  end if;
  return n;
end $$;

create or replace function public.seg_cambiar_metodo(p_periodo text, p_metodo text)
returns void
language plpgsql
set search_path = ''
as $$
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  update public.seg_periodos set metodo = p_metodo where id = p_periodo;
end $$;

-- ---------- Seguridad ----------
alter table public.seg_usuarios          enable row level security;
alter table public.seg_centros_costo     enable row level security;
alter table public.seg_zonas             enable row level security;
alter table public.seg_asignaciones      enable row level security;
alter table public.seg_bases             enable row level security;
alter table public.seg_periodos          enable row level security;
alter table public.seg_periodo_vehiculos enable row level security;

drop policy if exists seg_usuarios_propio on public.seg_usuarios;
create policy seg_usuarios_propio on public.seg_usuarios
  for select to authenticated using (email = lower(coalesce(auth.jwt() ->> 'email', '')));

do $$
declare t text;
begin
  foreach t in array array['seg_centros_costo','seg_zonas','seg_asignaciones','seg_bases','seg_periodos','seg_periodo_vehiculos'] loop
    execute format('drop policy if exists %I on public.%I', t || '_usuarios', t);
    execute format('create policy %I on public.%I for all to authenticated using ((select public.seg_es_usuario())) with check ((select public.seg_es_usuario()))', t || '_usuarios', t);
  end loop;
end $$;

revoke all on public.seg_usuarios, public.seg_centros_costo, public.seg_zonas, public.seg_asignaciones,
  public.seg_bases, public.seg_periodos, public.seg_periodo_vehiculos, public.seg_v_prorrateo, public.seg_v_totales_cc from anon;
grant select on public.seg_usuarios to authenticated;
grant select, insert, update, delete on public.seg_centros_costo, public.seg_zonas, public.seg_asignaciones,
  public.seg_bases, public.seg_periodos, public.seg_periodo_vehiculos to authenticated;
grant select on public.seg_v_prorrateo, public.seg_v_totales_cc to authenticated;

revoke execute on function public.seg_es_usuario(), public.seg_reemplazar_centros_costo(jsonb, text),
  public.seg_reemplazar_zonas(jsonb, text, jsonb), public.seg_guardar_periodo(jsonb, jsonb),
  public.seg_asignar_cc(text, text[], text), public.seg_cambiar_metodo(text, text) from public, anon;
grant execute on function public.seg_es_usuario(), public.seg_reemplazar_centros_costo(jsonb, text),
  public.seg_reemplazar_zonas(jsonb, text, jsonb), public.seg_guardar_periodo(jsonb, jsonb),
  public.seg_asignar_cc(text, text[], text), public.seg_cambiar_metodo(text, text) to authenticated;
