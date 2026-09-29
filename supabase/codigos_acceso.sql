-- ============================================================
-- Códigos de acceso para la liga pública de PALM
-- ============================================================
-- La página manda el código en el header "x-codigo-acceso".
-- Las políticas RLS de lectura solo entregan filas si ese código
-- existe en privado.codigos_acceso, está activo y no ha expirado.
--
-- Los códigos NO se guardan en este repo (es público). Se administran
-- desde Supabase (Table Editor → esquema "privado" → codigos_acceso, o SQL):
--
--   -- Dar acceso a alguien
--   insert into privado.codigos_acceso (codigo, nombre) values ('ABCD2345', 'Nombre de la persona');
--   -- Quitar acceso
--   update privado.codigos_acceso set activo = false where codigo = 'ABCD2345';
--   -- Acceso temporal
--   insert into privado.codigos_acceso (codigo, nombre, expira) values ('XYZ98765', 'Visita', now() + interval '7 days');
--   -- Ver quién ha entrado
--   select nombre, codigo, activo, usos, ultimo_uso from privado.codigos_acceso order by ultimo_uso desc nulls last;

-- ---------- Paso 1: tabla y funciones (no rompe nada) ----------
create schema if not exists privado;
revoke all on schema privado from public, anon, authenticated;

create table if not exists privado.codigos_acceso (
  codigo     text primary key check (codigo = upper(btrim(codigo)) and length(codigo) >= 6),
  nombre     text not null,
  activo     boolean not null default true,
  expira     timestamptz,
  creado     timestamptz not null default now(),
  ultimo_uso timestamptz,
  usos       integer not null default 0
);
alter table privado.codigos_acceso enable row level security;
revoke all on privado.codigos_acceso from public, anon, authenticated;

-- ¿La petición actual trae un código válido? (se usa dentro de las políticas RLS)
create or replace function public.codigo_acceso_valido()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1
    from privado.codigos_acceso c
    where c.codigo = upper(btrim(coalesce(
            nullif(current_setting('request.headers', true), '')::json ->> 'x-codigo-acceso', '')))
      and c.activo
      and (c.expira is null or c.expira > now())
  );
$$;
revoke all on function public.codigo_acceso_valido() from public;
grant execute on function public.codigo_acceso_valido() to anon, authenticated;

-- Login: valida el código, registra el uso y regresa el nombre (null si no es válido).
create or replace function public.entrar_con_codigo(p_codigo text)
returns text
language plpgsql
volatile
security definer
set search_path = ''
as $$
declare
  v_nombre text;
begin
  update privado.codigos_acceso
     set ultimo_uso = now(), usos = usos + 1
   where codigo = upper(btrim(coalesce(p_codigo, '')))
     and activo
     and (expira is null or expira > now())
  returning nombre into v_nombre;
  if v_nombre is null then
    perform pg_sleep(1); -- frena intentos de adivinar códigos
  end if;
  return v_nombre;
end;
$$;
revoke all on function public.entrar_con_codigo(text) from public;
grant execute on function public.entrar_con_codigo(text) to anon, authenticated;

-- ---------- Paso 2: cerrar la lectura pública ----------
-- Sin la política "using (true)", Postgres sí evalúa la política de admin para
-- anónimos, así que necesitan poder ejecutar app_role() (regresa null para ellos).
grant execute on function public.app_role() to anon, authenticated;

-- Cambia todas las políticas "public read" (using true) para exigir código.
do $$
declare r record;
begin
  for r in
    select tablename from pg_policies
    where schemaname = 'public' and policyname = 'public read'
  loop
    execute format('drop policy "public read" on public.%I', r.tablename);
    execute format(
      'create policy "lectura con codigo" on public.%I for select using ((select public.codigo_acceso_valido()))',
      r.tablename);
  end loop;
end;
$$;
