-- =====================================================================
-- Roles de usuario, registro de actividad y hoja «Usuarios»
--
-- Roles:
--   admin     Administrador: todo, incluida la hoja Usuarios. Es el único que sube archivos.
--   operador  Operador: asigna CC fila por fila o en bloque.
--   lectura   Lectura: ve todas las hojas, exporta a Excel y abre el PDF de la factura.
--
-- Además:
--   * Deja en el repositorio las funciones que hoy están en producción
--     (ingreso con el encabezado x-usuario y seg_usuario_actual).
--   * Aplica el control de CC habilitados de 20261113120000_cc_solo_habilitados.sql,
--     que no estaba aplicado en producción.
--
-- Se puede ejecutar más de una vez: no cambia los roles ya asignados.
-- =====================================================================

-- ---------- 1. Rol de cada usuario ----------
alter table public.seg_usuarios add column if not exists activo boolean not null default true;
alter table public.seg_usuarios add column if not exists rol text;
alter table public.seg_usuarios add column if not exists modificado_en timestamptz;
alter table public.seg_usuarios add column if not exists modificado_por text;

-- Roles iniciales: solo se completan los que todavía no tienen rol
update public.seg_usuarios
set rol = case when email = 'cg@fabricasrl.com.ar' then 'admin' else 'operador' end
where rol is null;

alter table public.seg_usuarios alter column rol set default 'lectura';
alter table public.seg_usuarios alter column rol set not null;
alter table public.seg_usuarios drop constraint if exists seg_usuarios_rol_valido;
alter table public.seg_usuarios add constraint seg_usuarios_rol_valido check (rol in ('admin', 'operador', 'lectura'));

-- ---------- 2. Registro de actividad ----------
create table if not exists public.seg_actividad (
  id          bigint generated always as identity primary key,
  email       text not null,
  accion      text not null,
  detalle     text not null default '',
  periodo_id  text,
  en          timestamptz not null default now()
);
comment on table public.seg_actividad is 'Ingresos y cambios hechos desde la aplicación. Solo se lee con seg_actividad_listar.';
create index if not exists seg_actividad_en_idx on public.seg_actividad (en desc);
create index if not exists seg_actividad_email_idx on public.seg_actividad (email, accion, en desc);
alter table public.seg_actividad enable row level security;
revoke all on public.seg_actividad from anon, authenticated;

-- La primera vez, recupera lo que ya quedó registrado en las tablas
do $$
begin
  if not exists (select 1 from public.seg_actividad) then
    insert into public.seg_actividad (email, accion, detalle, periodo_id, en)
    select lower(creado_por), 'periodo',
           format('Factura %s · %s vehículos', p.id, (select count(*) from public.seg_periodo_vehiculos v where v.periodo_id = p.id)),
           p.id, p.creado_en
    from public.seg_periodos p where coalesce(creado_por, '') <> ''
    union all
    select lower(cargado_por), case tipo when 'zonas' then 'zonas' else 'base_cc' end,
           format('%s · %s %s', coalesce(archivo, ''), filas, case tipo when 'zonas' then 'patentes' else 'CC' end),
           null, cargado_en
    from public.seg_bases where coalesce(cargado_por, '') <> ''
    union all
    select lower(cargado_por), 'factura_pdf', format('Factura %s · %s', periodo_id, archivo), periodo_id, cargado_en
    from public.seg_facturas_pdf where coalesce(cargado_por, '') <> ''
    union all
    select lower(actualizado_por), 'asignar_cc',
           format('CC %s a %s %s (registro previo)', cc, count(*), case when count(*) = 1 then 'patente' else 'patentes' end),
           null, max(actualizado_en)
    from public.seg_asignaciones where coalesce(actualizado_por, '') <> ''
    group by lower(actualizado_por), cc, date_trunc('minute', actualizado_en);
  end if;
end $$;

-- ---------- 3. Quién consulta y qué puede hacer ----------
-- Correo del encabezado x-usuario, si está en seg_usuarios y activo (ya existía en producción)
create or replace function public.seg_usuario_actual()
returns text
language sql
stable
security definer
set search_path = ''
as $$
  select u.email from public.seg_usuarios u
  where u.activo
    and u.email = lower(trim(coalesce(nullif(current_setting('request.headers', true), '')::json ->> 'x-usuario', '')));
$$;

create or replace function public.seg_rol()
returns text
language sql
stable
security definer
set search_path = ''
as $$
  select u.rol from public.seg_usuarios u where u.email = public.seg_usuario_actual();
$$;

-- Permisos por acción. Si cambia el reparto, se ajusta acá y en PERMISOS de src/main.js.
create or replace function public.seg_puede(p_accion text)
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select coalesce(case p_accion
    when 'ver'      then public.seg_rol() in ('admin', 'operador', 'lectura')
    when 'asignar'  then public.seg_rol() in ('admin', 'operador')
    when 'zonas'    then public.seg_rol() = 'admin'
    when 'cargar'   then public.seg_rol() = 'admin'
    when 'usuarios' then public.seg_rol() = 'admin'
  end, false);
$$;

-- Corta la operación si no tiene permiso; si lo tiene, devuelve el correo
create or replace function public.seg_exigir(p_accion text)
returns text
language plpgsql
stable
security definer
set search_path = ''
as $$
declare v_email text := public.seg_usuario_actual();
begin
  if v_email is null then raise exception 'Usuario no habilitado' using errcode = '42501'; end if;
  if not public.seg_puede(p_accion) then
    raise exception 'Tu rol no permite %', case p_accion
      when 'asignar'  then 'asignar centros de costo'
      when 'zonas'    then 'actualizar las zonas'
      when 'cargar'   then 'cargar períodos, la base de CC ni el PDF de la factura'
      when 'usuarios' then 'administrar usuarios'
      else p_accion end
    using errcode = '42501';
  end if;
  return v_email;
end $$;

create or replace function public.seg_registrar(p_email text, p_accion text, p_detalle text, p_periodo text default null)
returns void
language sql
security definer
set search_path = ''
as $$
  insert into public.seg_actividad (email, accion, detalle, periodo_id) values (p_email, p_accion, coalesce(p_detalle, ''), p_periodo);
$$;

-- Lista corta de patentes para el detalle de la actividad
create or replace function public.seg_lista_patentes(p_patentes text[])
returns text
language sql
immutable
set search_path = ''
as $$
  select array_to_string(p_patentes[1:8], ', ') || case when cardinality(p_patentes) > 8 then format(' y %s más', cardinality(p_patentes) - 8) else '' end;
$$;

-- ---------- 4. Ingreso ----------
-- Valida el correo, devuelve nombre y rol, y registra el ingreso
-- (abrir la app varias veces en 30 minutos cuenta como un solo ingreso).
create or replace function public.seg_ingresar(p_email text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare u public.seg_usuarios;
begin
  select * into u from public.seg_usuarios where activo and email = lower(trim(coalesce(p_email, '')));
  if not found then return null; end if;
  if not exists (select 1 from public.seg_actividad a
                 where a.email = u.email and a.accion = 'ingreso' and a.en > now() - interval '30 minutes') then
    perform public.seg_registrar(u.email, 'ingreso', '');
  end if;
  return jsonb_build_object('email', u.email, 'nombre', u.nombre, 'rol', u.rol);
end $$;

-- Datos del usuario actual sin registrar un ingreso (la app lo consulta al recargar)
create or replace function public.seg_mi_usuario()
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select jsonb_build_object('email', u.email, 'nombre', u.nombre, 'rol', u.rol)
  from public.seg_usuarios u where u.email = public.seg_usuario_actual();
$$;

-- ---------- 5. Funciones de la aplicación con su permiso ----------
create or replace function public.seg_asignar_cc(p_periodo text, p_patentes text[], p_cc text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare n integer; v_cc text := nullif(trim(p_cc), ''); v_email text := public.seg_exigir('asignar');
begin
  update public.seg_periodo_vehiculos set cc = v_cc, nueva = false
  where periodo_id = p_periodo and patente = any(p_patentes);
  get diagnostics n = row_count;
  if v_cc is null then
    delete from public.seg_asignaciones where patente = any(p_patentes);
  else
    insert into public.seg_asignaciones (patente, cc, actualizado_en, actualizado_por)
    select unnest(p_patentes), v_cc, now(), v_email
    on conflict (patente) do update set cc = excluded.cc, actualizado_en = now(), actualizado_por = excluded.actualizado_por;
  end if;
  if n > 0 then
    perform public.seg_registrar(v_email, case when v_cc is null then 'quitar_cc' else 'asignar_cc' end,
      case when v_cc is null then '' else format('CC %s a ', v_cc) end
        || format('%s %s: %s', n, case when n = 1 then 'patente' else 'patentes' end, public.seg_lista_patentes(p_patentes)),
      p_periodo);
  end if;
  return n;
end $$;

create or replace function public.seg_cambiar_metodo(p_periodo text, p_metodo text)
returns void
language plpgsql
security definer
set search_path = ''
as $$
begin
  perform public.seg_exigir('cargar');
  -- Se mantiene por compatibilidad: el prorrateo es siempre por premio
  update public.seg_periodos set metodo = 'premio' where id = p_periodo;
end $$;

create or replace function public.seg_reemplazar_centros_costo(p_items jsonb, p_archivo text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare n integer; v_email text := public.seg_exigir('cargar');
begin
  delete from public.seg_centros_costo where true;
  insert into public.seg_centros_costo (codigo, descripcion, habilitado)
  select distinct on (trim(x->>'codigo')) trim(x->>'codigo'), coalesce(x->>'descripcion', ''), coalesce((x->>'habilitado')::boolean, true)
  from jsonb_array_elements(p_items) x
  where coalesce(trim(x->>'codigo'), '') <> '';
  get diagnostics n = row_count;
  insert into public.seg_bases (tipo, archivo, filas, cargado_en, cargado_por)
  values ('centros_costo', p_archivo, n, now(), v_email)
  on conflict (tipo) do update set archivo = excluded.archivo, filas = excluded.filas, cargado_en = now(), cargado_por = excluded.cargado_por;
  perform public.seg_registrar(v_email, 'base_cc', format('%s · %s CC', coalesce(p_archivo, ''), n));
  return n;
end $$;

create or replace function public.seg_reemplazar_zonas(p_items jsonb, p_archivo text, p_conflictos jsonb default '[]'::jsonb)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare n integer; v_email text := public.seg_exigir('zonas');
begin
  delete from public.seg_zonas where true;
  insert into public.seg_zonas (patente, zona)
  select distinct on (x->>'patente') x->>'patente', x->>'zona'
  from jsonb_array_elements(p_items) with ordinality as t(x, i)
  where coalesce(x->>'patente', '') <> '' and coalesce(x->>'zona', '') <> ''
  order by x->>'patente', i desc;
  get diagnostics n = row_count;
  insert into public.seg_bases (tipo, archivo, filas, detalle, cargado_en, cargado_por)
  values ('zonas', p_archivo, n, jsonb_build_object('conflictos', p_conflictos), now(), v_email)
  on conflict (tipo) do update set archivo = excluded.archivo, filas = excluded.filas, detalle = excluded.detalle, cargado_en = now(), cargado_por = excluded.cargado_por;
  perform public.seg_registrar(v_email, 'zonas', format('%s · %s patentes', coalesce(p_archivo, ''), n));
  return n;
end $$;

create or replace function public.seg_guardar_periodo(p_periodo jsonb, p_vehiculos jsonb)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_id text := p_periodo->>'id';
  v_email text := public.seg_exigir('cargar');
  v_hay_previo boolean;
  v_total integer; v_con_cc integer;
begin
  if coalesce(v_id, '') = '' then raise exception 'Falta el número de factura'; end if;

  select exists (select 1 from public.seg_asignaciones) into v_hay_previo;

  insert into public.seg_periodos as p (id, poliza, tomador, vigencia_desde, vigencia_hasta, total_certificados, premio_final,
    factura_fecha, factura_detalle, cae, neto, iva, percepcion, otros_impuestos, total, archivo_poliza, archivo_factura, creado_por)
  values (v_id, p_periodo->>'poliza', p_periodo->>'tomador',
    to_date(nullif(p_periodo->>'vigencia_desde', ''), 'DD/MM/YYYY'), to_date(nullif(p_periodo->>'vigencia_hasta', ''), 'DD/MM/YYYY'),
    (p_periodo->>'total_certificados')::integer, (p_periodo->>'premio_final')::numeric,
    to_date(nullif(p_periodo->>'factura_fecha', ''), 'DD/MM/YYYY'), p_periodo->>'factura_detalle', p_periodo->>'cae',
    (p_periodo->>'neto')::numeric, coalesce((p_periodo->>'iva')::numeric, 0), coalesce((p_periodo->>'percepcion')::numeric, 0),
    coalesce((p_periodo->>'otros_impuestos')::numeric, 0), (p_periodo->>'total')::numeric,
    p_periodo->>'archivo_poliza', p_periodo->>'archivo_factura', v_email)
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
  perform public.seg_registrar(v_email, 'periodo', format('Factura %s · %s vehículos, %s con CC', v_id, v_total, v_con_cc), v_id);
  return jsonb_build_object('id', v_id, 'vehiculos', v_total, 'con_cc', v_con_cc);
end $$;

-- PDF de la factura: verlo, cualquier rol; guardarlo, solo Administrador
create or replace function public.seg_pdf_habilitado()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select public.seg_puede('ver');
$$;

create or replace function public.seg_guardar_factura_pdf(p_periodo text, p_archivo text, p_pdf text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare v bytea; v_email text := public.seg_exigir('cargar');
begin
  if not exists (select 1 from public.seg_periodos where id = p_periodo) then
    raise exception 'El período % no existe', p_periodo;
  end if;
  v := decode(p_pdf, 'base64');
  if length(v) > 10 * 1024 * 1024 then raise exception 'El PDF supera los 10 MB'; end if;
  if substring(v from 1 for 5) <> '\x255044462d'::bytea then raise exception 'El archivo no es un PDF'; end if;
  insert into public.seg_facturas_pdf (periodo_id, archivo, contenido, bytes, cargado_en, cargado_por)
  values (p_periodo, coalesce(p_archivo, ''), v, length(v), now(), v_email)
  on conflict (periodo_id) do update set archivo = excluded.archivo, contenido = excluded.contenido,
    bytes = excluded.bytes, cargado_en = now(), cargado_por = excluded.cargado_por;
  perform public.seg_registrar(v_email, 'factura_pdf', format('Factura %s · %s', p_periodo, coalesce(p_archivo, '')), p_periodo);
  return length(v);
end $$;

-- ---------- 6. Hoja Usuarios (solo Administrador) ----------
create or replace function public.seg_usuarios_listar()
returns table (email text, nombre text, rol text, activo boolean, creado_en timestamptz, ingresos bigint, ultimo_ingreso timestamptz)
language plpgsql
stable
security definer
set search_path = ''
as $$
begin
  perform public.seg_exigir('usuarios');
  return query
  select u.email, u.nombre, u.rol, u.activo, u.creado_en, count(a.id), max(a.en)
  from public.seg_usuarios u
  left join public.seg_actividad a on a.email = u.email and a.accion = 'ingreso'
  group by u.email, u.nombre, u.rol, u.activo, u.creado_en
  order by u.activo desc, lower(coalesce(nullif(u.nombre, ''), u.email));
end $$;

-- p_tipo: 'cambios' (todo menos ingresos), 'ingresos' o 'todo'
create or replace function public.seg_actividad_listar(p_email text default null, p_tipo text default 'cambios', p_limite integer default 200)
returns table (id bigint, email text, nombre text, accion text, detalle text, periodo_id text, en timestamptz)
language plpgsql
stable
security definer
set search_path = ''
as $$
begin
  perform public.seg_exigir('usuarios');
  return query
  select a.id, a.email, u.nombre, a.accion, a.detalle, a.periodo_id, a.en
  from public.seg_actividad a
  left join public.seg_usuarios u on u.email = a.email
  where (nullif(p_email, '') is null or a.email = lower(p_email))
    and case coalesce(p_tipo, 'cambios')
          when 'ingresos' then a.accion = 'ingreso'
          when 'todo' then true
          else a.accion <> 'ingreso' end
  order by a.en desc, a.id desc
  limit least(greatest(coalesce(p_limite, 200), 1), 1000);
end $$;

create or replace function public.seg_rol_nombre(p_rol text)
returns text
language sql
immutable
set search_path = ''
as $$
  select case p_rol when 'admin' then 'Administrador' when 'operador' then 'Operador' when 'lectura' then 'Lectura' else p_rol end;
$$;

create or replace function public.seg_controlar_admins()
returns void
language plpgsql
security definer
set search_path = ''
as $$
begin
  if not exists (select 1 from public.seg_usuarios where activo and rol = 'admin') then
    raise exception 'Tiene que quedar al menos un administrador activo' using errcode = 'check_violation';
  end if;
end $$;

create or replace function public.seg_usuario_alta(p_email text, p_nombre text, p_rol text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_admin text := public.seg_exigir('usuarios');
  v_email text := lower(trim(coalesce(p_email, '')));
  v_nombre text := nullif(trim(regexp_replace(coalesce(p_nombre, ''), '\s+', ' ', 'g')), '');
begin
  if v_email !~ '^[^\s@]+@[^\s@]+\.[^\s@]+$' then raise exception 'El correo % no es válido', v_email; end if;
  if v_nombre is null then raise exception 'Falta el nombre'; end if;
  if p_rol not in ('admin', 'operador', 'lectura') then raise exception 'Rol desconocido: %', p_rol; end if;
  if exists (select 1 from public.seg_usuarios where email = v_email) then
    raise exception 'Ya existe un usuario con el correo %', v_email using errcode = 'unique_violation';
  end if;
  insert into public.seg_usuarios (email, nombre, rol, activo, modificado_en, modificado_por)
  values (v_email, v_nombre, p_rol, true, now(), v_admin);
  perform public.seg_registrar(v_admin, 'usuario_alta', format('%s (%s) · %s', v_nombre, v_email, public.seg_rol_nombre(p_rol)));
  return jsonb_build_object('email', v_email, 'nombre', v_nombre, 'rol', p_rol, 'activo', true);
end $$;

-- Cambia nombre, rol o estado. Los parámetros nulos no se tocan.
create or replace function public.seg_usuario_cambiar(p_email text, p_nombre text default null, p_rol text default null, p_activo boolean default null)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_admin text := public.seg_exigir('usuarios');
  v_email text := lower(trim(coalesce(p_email, '')));
  v_nombre text := nullif(trim(regexp_replace(coalesce(p_nombre, ''), '\s+', ' ', 'g')), '');
  u public.seg_usuarios;
  cambios text[] := '{}';
begin
  select * into u from public.seg_usuarios where email = v_email for update;
  if not found then raise exception 'No existe el usuario %', v_email; end if;
  if p_rol is not null and p_rol not in ('admin', 'operador', 'lectura') then raise exception 'Rol desconocido: %', p_rol; end if;

  if v_nombre is not null and v_nombre is distinct from u.nombre then cambios := cambios || format('nombre: %s → %s', coalesce(u.nombre, '—'), v_nombre); end if;
  if p_rol is not null and p_rol <> u.rol then cambios := cambios || format('rol: %s → %s', public.seg_rol_nombre(u.rol), public.seg_rol_nombre(p_rol)); end if;
  if p_activo is not null and p_activo <> u.activo then cambios := cambios || case when p_activo then 'reactivado' else 'desactivado' end; end if;
  if cardinality(cambios) = 0 then
    return jsonb_build_object('email', u.email, 'nombre', u.nombre, 'rol', u.rol, 'activo', u.activo);
  end if;

  update public.seg_usuarios set
    nombre = coalesce(v_nombre, nombre), rol = coalesce(p_rol, rol), activo = coalesce(p_activo, activo),
    modificado_en = now(), modificado_por = v_admin
  where email = v_email
  returning * into u;

  perform public.seg_controlar_admins();
  perform public.seg_registrar(v_admin, 'usuario_cambio', format('%s (%s) · %s', coalesce(u.nombre, ''), u.email, array_to_string(cambios, ', ')));
  return jsonb_build_object('email', u.email, 'nombre', u.nombre, 'rol', u.rol, 'activo', u.activo);
end $$;

-- ---------- 7. Solo CC habilitados (de 20261113120000_cc_solo_habilitados.sql) ----------
create or replace function public.seg_validar_cc_habilitado()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  if new.cc is not null and new.cc is distinct from old.cc
     and not exists (select 1 from public.seg_centros_costo c where c.codigo = new.cc and c.habilitado) then
    raise exception 'El centro de costo % no está habilitado', new.cc using errcode = 'check_violation';
  end if;
  return new;
end $$;

drop trigger if exists seg_periodo_vehiculos_cc_habilitado on public.seg_periodo_vehiculos;
create trigger seg_periodo_vehiculos_cc_habilitado
  before update of cc on public.seg_periodo_vehiculos
  for each row execute function public.seg_validar_cc_habilitado();

-- ---------- 8. Permisos de ejecución ----------
-- Internas: solo las usan las otras funciones
revoke execute on function public.seg_rol(), public.seg_puede(text), public.seg_exigir(text),
  public.seg_registrar(text, text, text, text), public.seg_lista_patentes(text[]), public.seg_rol_nombre(text),
  public.seg_controlar_admins(), public.seg_validar_cc_habilitado()
  from public, anon, authenticated;

-- Las que llama la aplicación
revoke execute on function public.seg_ingresar(text), public.seg_mi_usuario(), public.seg_usuario_actual(),
  public.seg_asignar_cc(text, text[], text), public.seg_cambiar_metodo(text, text),
  public.seg_reemplazar_centros_costo(jsonb, text), public.seg_reemplazar_zonas(jsonb, text, jsonb),
  public.seg_guardar_periodo(jsonb, jsonb), public.seg_pdf_habilitado(), public.seg_guardar_factura_pdf(text, text, text),
  public.seg_usuarios_listar(), public.seg_actividad_listar(text, text, integer),
  public.seg_usuario_alta(text, text, text), public.seg_usuario_cambiar(text, text, text, boolean)
  from public;
grant execute on function public.seg_ingresar(text), public.seg_mi_usuario(), public.seg_usuario_actual(),
  public.seg_asignar_cc(text, text[], text), public.seg_cambiar_metodo(text, text),
  public.seg_reemplazar_centros_costo(jsonb, text), public.seg_reemplazar_zonas(jsonb, text, jsonb),
  public.seg_guardar_periodo(jsonb, jsonb), public.seg_pdf_habilitado(), public.seg_guardar_factura_pdf(text, text, text),
  public.seg_usuarios_listar(), public.seg_actividad_listar(text, text, integer),
  public.seg_usuario_alta(text, text, text), public.seg_usuario_cambiar(text, text, text, boolean)
  to anon, authenticated;

-- Las tablas no se modifican directamente desde la API: solo con estas funciones
revoke insert, update, delete, truncate on public.seg_usuarios, public.seg_centros_costo, public.seg_zonas,
  public.seg_asignaciones, public.seg_bases, public.seg_periodos, public.seg_periodo_vehiculos, public.seg_facturas_pdf
  from anon, authenticated;

notify pgrst, 'reload schema';

-- Control: usuarios con su rol
select email, nombre, public.seg_rol_nombre(rol) as rol, activo from public.seg_usuarios order by rol, nombre;
