-- Solo se pueden asignar centros de costo habilitados.
-- La aplicación ya ofrece únicamente los CC habilitados; este control impide
-- que una asignación con un CC inhabilitado o inexistente se guarde por otra vía.
-- Se aplica al cambiar el CC de un vehículo (seg_asignar_cc). No afecta a los
-- períodos nuevos que heredan el CC del período anterior: si ese CC quedó
-- inhabilitado, la patente aparece marcada para reasignar.
-- Se puede ejecutar más de una vez.

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

revoke execute on function public.seg_validar_cc_habilitado() from public, anon;

drop trigger if exists seg_periodo_vehiculos_cc_habilitado on public.seg_periodo_vehiculos;
create trigger seg_periodo_vehiculos_cc_habilitado
  before update of cc on public.seg_periodo_vehiculos
  for each row execute function public.seg_validar_cc_habilitado();
