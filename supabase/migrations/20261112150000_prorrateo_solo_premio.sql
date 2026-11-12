-- El prorrateo se hace siempre por el premio de cada certificado.
-- Se puede ejecutar más de una vez.

-- 1. Los períodos ya cargados pasan a prorratearse por premio
update public.seg_periodos set metodo = 'premio' where metodo is distinct from 'premio';

-- 2. Los períodos nuevos se crean con premio
alter table public.seg_periodos alter column metodo set default 'premio';

-- 3. Ya no se puede volver a "valor"
alter table public.seg_periodos drop constraint if exists seg_periodos_solo_premio;
alter table public.seg_periodos add constraint seg_periodos_solo_premio check (metodo = 'premio');

create or replace function public.seg_cambiar_metodo(p_periodo text, p_metodo text)
returns void
language plpgsql
set search_path = ''
as $$
begin
  if not public.seg_es_usuario() then raise exception 'Usuario no autorizado'; end if;
  -- Se mantiene por compatibilidad: el prorrateo es siempre por premio
  update public.seg_periodos set metodo = 'premio' where id = p_periodo;
end $$;
