-- ============================================================================
-- S&S IA — Borrado de datos de prueba antes de la prueba de validación final
-- ============================================================================
-- Corre esto en Supabase → tu proyecto → SQL Editor → "New query".
--
-- Módulos que SÍ se vacían (tal como los pediste):
--   Registro (venta, compra, gasto, cobro, pago, nómina)  -> tabla operaciones
--   Comercial (cotizaciones, terceros)                    -> tablas cotizaciones, clientes
--   Proyectos                                             -> tabla proyectos
--   Cuentas x cobrar                                      -> tabla operaciones (mismas ventas/cobros de Registro)
--   Cuentas x pagar                                       -> tabla operaciones (mismas compras/pagos) + proveedores
--   Presupuesto                                           -> tabla costos_presupuesto
--   Nómina y prestaciones (nómina, cesantías, prima,
--     vacaciones, liquidación final)                      -> tablas operaciones (nómina) + liquidaciones
--
--   Dos tablas que agrego por completitud, aunque no las nombraste una por
--   una (avísame si NO quieres que se toquen):
--   - tareas_comerciales: los recordatorios/seguimientos de las cotizaciones
--     (parte del módulo Comercial) — si no se borran, quedan tareas de
--     seguimiento apuntando a cotizaciones que ya no van a existir.
--   - pasivos_financieros: pasivos/deudas financieras — es el mismo tipo de
--     registro que maneja el módulo Cuentas x pagar, además de las compras.
--
-- Qué NO se toca (se conserva tal cual):
--   usuarios (logins admin/comercial/gerencia — si se borra, pierdes el
--             acceso a la app)
--   Calendario — no tiene tabla propia, todas sus fechas se calculan al
--             vuelo (IVA, retención, ICA, seguridad social), así que no hay
--             nada que borrar ahí aunque quisieras.
--   Producción (materias_primas, productos_terminados, lista_materiales,
--             ordenes_produccion) — no lo mencionaste en la lista, así que
--             lo dejo intacto. Si también quieres limpiar Producción,
--             dímelo y agrego esas 4 tablas.
--   colaboradores — la nómina de empleados en sí (no es lo mismo que los
--             registros de nómina/liquidaciones que sí se borran) —
--             normalmente son personas reales, no datos de prueba.
--   catalogo_items — el catálogo de ítems/precios para cotizar (lista de
--             referencia, no un registro transaccional de prueba).
--
-- Esto es IRREVERSIBLE una vez que corras la parte de DELETE — no hay
-- "deshacer". Si tienes clientes o proveedores reales ya cargados (no solo
-- de prueba), este script también los borra, porque no hay forma de
-- distinguir "prueba" de "real" en la base de datos.
-- ============================================================================


-- PASO 1 (opcional, recomendado) — revisa cuántos registros hay en cada
-- tabla ANTES de borrar, para confirmar que tiene sentido lo que vas a
-- perder. Corre solo este bloque primero si quieres revisar antes:

select 'operaciones' as tabla, count(*) from operaciones
union all select 'cotizaciones', count(*) from cotizaciones
union all select 'clientes', count(*) from clientes
union all select 'proveedores', count(*) from proveedores
union all select 'proyectos', count(*) from proyectos
union all select 'costos_presupuesto', count(*) from costos_presupuesto
union all select 'liquidaciones', count(*) from liquidaciones
union all select 'tareas_comerciales', count(*) from tareas_comerciales
union all select 'pasivos_financieros', count(*) from pasivos_financieros;


-- PASO 2 — el borrado en sí. Selecciona TODO este bloque (desde BEGIN
-- hasta COMMIT) y dale "Run". Si algo sale mal a mitad de camino, Supabase
-- revierte todo el bloque solo (no queda a medias).
--
-- Orden: primero las tablas que "apuntan a" clientes/proveedores
-- (cotizaciones, operaciones, proyectos, etc.), y clientes/proveedores de
-- últimas — así nunca queda nada huérfano a mitad del borrado.

begin;

delete from operaciones;
delete from costos_presupuesto;
delete from proyectos;
delete from tareas_comerciales;
delete from cotizaciones;
delete from liquidaciones;
delete from pasivos_financieros;
delete from clientes;
delete from proveedores;

commit;
