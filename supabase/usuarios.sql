-- Usuarios habilitados para ingresar a la aplicación (solo con el correo).
-- Se puede ejecutar más de una vez: actualiza el nombre y reactiva el correo.
-- Para dar de baja a alguien:  update public.seg_usuarios set activo = false where email = 'correo@fabricasrl.com.ar';

insert into public.seg_usuarios (email, nombre, activo) values
  ('nguaglino@fabricasrl.com.ar', 'Norma Guaglino', true),
  ('cg@fabricasrl.com.ar', 'Gabriel Cardozo', true),
  ('mguidini@fabricasrl.com.ar', 'Marco Guidini', true),
  ('arealegal3@fabricasrl.com.ar', 'Jose Luis Romero', true)
on conflict (email) do update set nombre = excluded.nombre, activo = true;

select email, nombre, activo from public.seg_usuarios order by nombre;
