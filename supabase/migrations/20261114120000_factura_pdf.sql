-- Guarda el PDF de la factura de cada período para poder verlo desde la aplicación.
-- La tabla no se lee directamente: solo a través de estas funciones, que controlan
-- que quien consulta sea un usuario habilitado.
-- Se puede ejecutar más de una vez: si ya se ejecutó una versión anterior, la actualiza.

create table if not exists public.seg_facturas_pdf (
  periodo_id  text primary key references public.seg_periodos(id) on delete cascade,
  archivo     text not null default '',
  contenido   bytea not null,
  bytes       integer not null,
  cargado_en  timestamptz not null default now(),
  cargado_por text
);
alter table public.seg_facturas_pdf enable row level security;
revoke all on public.seg_facturas_pdf from anon, authenticated;

-- Correo de quien hace el cambio: el encabezado x-usuario que envía la app o, si no está, el del token
create or replace function public.seg_quien()
returns text
language sql
stable
set search_path = ''
as $$
  select lower(coalesce(
    nullif(current_setting('request.headers', true), '')::json ->> 'x-usuario',
    auth.jwt() ->> 'email'));
$$;

-- ¿Quien consulta está habilitado? Se acepta el correo del encabezado x-usuario (como
-- ingresa la app) si figura en seg_usuarios y no está dado de baja. Si la base ya tiene
-- public.seg_es_usuario() y responde que sí, también se acepta.
create or replace function public.seg_pdf_habilitado()
returns boolean
language plpgsql
stable
security definer
set search_path = ''
as $$
declare v_email text := public.seg_quien(); v_ok boolean := false;
begin
  if coalesce(v_email, '') <> '' then
    select coalesce((to_jsonb(u) ->> 'activo')::boolean, true) into v_ok
    from public.seg_usuarios u where lower(u.email) = v_email;
    if coalesce(v_ok, false) then return true; end if;
  end if;
  begin
    execute 'select public.seg_es_usuario()' into v_ok;
  exception when others then v_ok := false;
  end;
  return coalesce(v_ok, false);
end $$;

-- Guarda (o reemplaza) el PDF de la factura de un período. p_pdf va en base64.
create or replace function public.seg_guardar_factura_pdf(p_periodo text, p_archivo text, p_pdf text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare v bytea;
begin
  if not public.seg_pdf_habilitado() then raise exception 'Sin permiso para ver las facturas: el correo % no está habilitado', coalesce(public.seg_quien(), '(sin correo)'); end if;
  if not exists (select 1 from public.seg_periodos where id = p_periodo) then
    raise exception 'El período % no existe', p_periodo;
  end if;
  v := decode(p_pdf, 'base64');
  if length(v) > 10 * 1024 * 1024 then raise exception 'El PDF supera los 10 MB'; end if;
  if substring(v from 1 for 5) <> '\x255044462d'::bytea then raise exception 'El archivo no es un PDF'; end if;
  insert into public.seg_facturas_pdf (periodo_id, archivo, contenido, bytes, cargado_en, cargado_por)
  values (p_periodo, coalesce(p_archivo, ''), v, length(v), now(), public.seg_quien())
  on conflict (periodo_id) do update set archivo = excluded.archivo, contenido = excluded.contenido,
    bytes = excluded.bytes, cargado_en = now(), cargado_por = excluded.cargado_por;
  return length(v);
end $$;

-- Devuelve el PDF de la factura de un período en base64, o null si no se guardó.
create or replace function public.seg_factura_pdf(p_periodo text)
returns jsonb
language plpgsql
stable
security definer
set search_path = ''
as $$
declare r jsonb;
begin
  if not public.seg_pdf_habilitado() then raise exception 'Sin permiso para ver las facturas: el correo % no está habilitado', coalesce(public.seg_quien(), '(sin correo)'); end if;
  select jsonb_build_object('archivo', f.archivo, 'bytes', f.bytes, 'cargado_en', f.cargado_en,
                            'contenido', encode(f.contenido, 'base64'))
  into r from public.seg_facturas_pdf f where f.periodo_id = p_periodo;
  return r;
end $$;

-- Qué períodos tienen el PDF guardado (sin descargar el contenido).
create or replace function public.seg_facturas_pdf_info()
returns table (periodo_id text, archivo text, bytes integer, cargado_en timestamptz)
language plpgsql
stable
security definer
set search_path = ''
as $$
begin
  if not public.seg_pdf_habilitado() then raise exception 'Sin permiso para ver las facturas: el correo % no está habilitado', coalesce(public.seg_quien(), '(sin correo)'); end if;
  return query select f.periodo_id, f.archivo, f.bytes, f.cargado_en from public.seg_facturas_pdf f;
end $$;

revoke execute on function public.seg_quien(), public.seg_pdf_habilitado(), public.seg_guardar_factura_pdf(text, text, text),
  public.seg_factura_pdf(text), public.seg_facturas_pdf_info() from public;
grant execute on function public.seg_quien(), public.seg_pdf_habilitado(), public.seg_guardar_factura_pdf(text, text, text),
  public.seg_factura_pdf(text), public.seg_facturas_pdf_info() to anon, authenticated;

-- Avisa a la API de Supabase que hay funciones nuevas
notify pgrst, 'reload schema';

-- Control: tienen que aparecer las cinco funciones
select p.proname as funcion, pg_get_function_identity_arguments(p.oid) as argumentos
from pg_proc p join pg_namespace n on n.oid = p.pronamespace
where n.nspname = 'public' and p.proname in ('seg_quien', 'seg_pdf_habilitado', 'seg_guardar_factura_pdf', 'seg_factura_pdf', 'seg_facturas_pdf_info')
order by 1;
