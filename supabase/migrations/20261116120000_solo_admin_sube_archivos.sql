-- Solo el Administrador sube archivos: póliza, factura, base de CC, zonas y PDF de la factura.
-- El Operador asigna CC a mano (fila por fila o en bloque).
-- La importación de asignaciones desde Excel usa la misma función que la asignación a mano
-- (seg_asignar_cc), así que la limita la app; para la base es una asignación en bloque más.
-- Requiere 20261115120000_roles_usuarios.sql. Se puede ejecutar más de una vez.

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

revoke execute on function public.seg_puede(text) from public, anon, authenticated;
notify pgrst, 'reload schema';

-- Control: tiene que decir «admin» en zonas
select 'zonas' as accion, 'admin' as esperado,
       case when pg_get_functiondef('public.seg_puede(text)'::regprocedure) like '%''zonas''    then public.seg_rol() = ''admin''%' then 'admin' else 'REVISAR' end as aplicado;
