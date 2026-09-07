# Prompts de implementación — Skybridge ERP/CRM

Uno por sección de [BACKLOG_TECNICO.md](BACKLOG_TECNICO.md), pensados para
pegar tal cual en la sesión que va a programarlos. Orden recomendado: 1 → 2
→ 3 → 4 en secuencia (2 antes que 3: la columna "Cliente" del prompt 3
necesita que el prompt 2 ya esté vinculando bien los clientes). El 5 es
limpieza, sin apuro, se puede hacer en cualquier momento o en paralelo.

**Estado (2026-09-06): Prompts 1, 2, 3 y 5 (ítems 1 y 2) ya están
implementados y verificados** — ver la nota al final de cada uno. Queda
pendiente el Prompt 4 (Honorario Skybridge) y el ítem 3 del Prompt 5
(documentar columnas reutilizadas del schema).

Contexto compartido para quien programe: app Streamlit (`app.py`) + SQLite
(`db.py`) + motor de embudo (`crm.py`) del ERP/CRM de Skybridge, una
consultora de comercio exterior (no el importador — ver el comentario al
principio de `crm.py` si hace falta el contexto de negocio completo).
Convenciones ya establecidas a respetar: nombres en español, comentarios
que explican el POR QUÉ de decisiones no obvias (no qué hace el código),
columnas nuevas se agregan vía el dict `COLUMNAS_NUEVAS` en `db.py`
(ALTER TABLE en cada arranque), y el criterio de "nunca retroceder de
etapa" (`crm.es_avance`) se respeta en cualquier cambio de etapa nuevo.

---

## Prompt 1 — Fix: filtro de inactividad marca mal la actividad de hoy

En `app.py`, la función `_dias_desde_actividad(c)` devuelve un `int` (0 si
la última actividad fue hoy mismo) o `None` si no hay ninguna fecha. Más
abajo, el filtro "Mostrar solo sin actividad hace 14+ días" usa:

```python
visibles = [c for c in visibles if (_dias_desde_actividad(c) or 10 ** 9) >= 14]
```

Bug: en Python `0 or 10**9` da `10**9`, porque `0` es falsy. Un contacto
con actividad HOY (`_dias_desde_actividad` = 0) queda tratado como si
tuviera "10 mil millones de días sin actividad" y aparece incorrectamente
en ese filtro.

Arreglar para que solo se considere "inactivo 14+ días" cuando el valor
realmente sea `None` (sin ninguna actividad registrada nunca) o un número
≥ 14 — nunca por el efecto colateral de que 0 sea falsy. Mismo archivo,
buscar el checkbox `crm_filtro_inactivo` para ubicar el bloque exacto.

**Cómo probarlo:** en CRM → Contactos, tildar "Mostrar solo sin actividad
hace 14+ días". Un contacto con una nota/actividad cargada hoy NO debe
aparecer en la lista filtrada.

**✅ Resuelto (2026-09-06).** El bug puntual se corrigió primero, pero el
checkbox que lo tenía ya **no existe**: a pedido del usuario se eliminó ese
filtro (y el de "seguimiento vencido") y se reemplazó por un orden
automático en `_render_lista_contactos` — cada columna del Kanban ordena
sola por días desde la última actividad, sin ningún filtro que activar a
mano. "Cómo probarlo" de arriba ya no aplica tal cual (el checkbox no
está), pero el resultado que buscaba (que un contacto activo hoy no
aparezca como "el más urgente") se cumple igual, ahora por diseño.

---

## Prompt 2 — Sincronía comercial ↔ administrativo

Objetivo: un cliente puede seguir naciendo por dos puertas (directo en
Clientes, o desde un contacto del CRM) — las dos coexisten — pero las dos
tienen que terminar siempre con un contacto de CRM vinculado y en la etapa
correcta, sin que nadie tenga que acordarse de hacerlo a mano.

Ya existe el patrón a reutilizar en `app.py`: `_asegurar_contacto_para_cliente(cliente_id)`
(busca el contacto vinculado a ese cliente, o lo crea si no existe) y
`_avanzar_etapa_si_corresponde(contact_id, etapa_objetivo, autor)` (sube de
etapa solo si es un avance real, nunca retrocede). `_cotizar_contacto` ya
usa exactamente este patrón en el sentido CRM → Clientes.

1. **Alta directa de cliente:** donde se procesa el formulario "＋ Nuevo
   cliente" en `vista_clientes()` (llamada a `db.create_cliente(...)`),
   después de crear el cliente llamar a `_asegurar_contacto_para_cliente(new_id)`
   para que nazca su contacto de CRM en "Nuevo" en el mismo momento.

2. **Cotizar directo desde Clientes:** en `_cotizar_cliente(cliente_id)`,
   después de `db.create_cotizacion_borrador(...)`, agregar el mismo
   patrón que ya usa `_cotizar_contacto` en sentido inverso: asegurar el
   contacto vinculado y avanzarlo a "Cotizado" con `_avanzar_etapa_si_corresponde`.

3. **Retroactivo:** un script/paso de migración (puede vivir en `db.py` o
   correrse una sola vez) que recorra todos los `clientes` sin contacto de
   CRM vinculado (`db.get_contact_by_cliente_id` devuelve `None`) y les
   cree uno — **pero no todos en "Nuevo"**: el contacto retroactivo tiene
   que nacer en la etapa que corresponda a la evidencia real que ya tiene
   ese cliente (si ya tiene alguna importación → "Cliente"/Ganado; si tiene
   una cotización Aprobada → "Cliente"/Ganado; si tiene cualquier
   cotización → "Cotizado"; si no tiene nada → "Nuevo"). Guiarse por la
   misma prioridad de señales que ya usa `crm.estado_comercial()`. Ejemplo
   real para probar: el cliente "BARWATEX SRL" ya tiene varias
   importaciones cargadas — después de la migración tiene que aparecer
   directo en la columna "Cliente" del CRM, no en "Nuevo".

**Nota, no hace falta tocar nada acá:** ya existe un tercer lugar que
cambia el estado de una cotización aparte del editor completo —
`_cambiar_estado_cotizacion` (el control "Cambiar" en cada fila del
historial) — y ya dispara `_marcar_ganado_por_cliente` correctamente
cuando el estado pasa a "Aprobada". Mantener ese mismo criterio, no
duplicarlo ni tocarlo.

**Cómo probarlo:** crear un cliente nuevo directo desde Clientes → debe
aparecer inmediato en CRM → Contactos → Nuevos. Cotizarlo desde su ficha →
debe pasar a Cotizados. Correr la migración retroactiva → BARWATEX SRL
(y cualquier otro cliente viejo con importaciones) debe aparecer en la
columna Cliente, no en Nuevos.

**✅ Resuelto y verificado (2026-09-04).** Los 3 puntos implementados tal
cual el prompt. La migración retroactiva se corrió contra la base real:
BARWATEX SRL y el resto de los clientes viejos con importaciones
aparecieron en "Cliente", confirmado por prueba manual. Se sumó además un
botón en Panel de Control → Configuración para poder volver a correrla
desde la UI (es idempotente, no duplica nada si se corre de nuevo).

---

## Prompt 3 — Vocabulario y vista del embudo (CRM)

### Renombrar "Ganado" → "Cliente", "Perdido" → "Descartado"

Recomendación de implementación (para no migrar datos ni arriesgar romper
una comparación de texto que quede sin actualizar en algún lado): **no
cambiar el valor interno guardado** (`contacts.etapa` sigue guardando
literalmente "Ganado"/"Perdido", igual que hoy, y todas las comparaciones
en `crm.py`/`app.py` — `STAGES`, `COLOR_ETAPA`, `ETAPA_SLUG`, los `if etapa
== "Ganado"`, etc. — se mantienen sin tocar). En su lugar, sumar una capa
de **etiqueta de visualización** (ej. un dict `ETAPA_LABEL` en `crm.py`
que mapee `"Ganado" → "Cliente"` y `"Perdido" → "Descartado"`, identidad
para el resto) y usarla en todos los lugares donde una etapa se muestra al
usuario: las badges (`_badge_etapa`), el nombre de columna/chip del
kanban, el selectbox de etapa en la ficha del contacto, y el texto de
"cambio_etapa" en el timeline de actividad. Mismo criterio ya usado en
`cotizaciones.contenedor`/`etd_eta`/`carrier`/`freetime` (columnas que
guardan algo distinto de su nombre literal) — es un patrón ya conocido en
este código, no uno nuevo.

Si se prefiere en cambio un rename real (cambiar el valor guardado), hace
falta además una migración `UPDATE contacts SET etapa='Cliente' WHERE
etapa='Ganado'` (ídem 'Descartado') y revisar el texto ya guardado en
`activity_log.texto` de eventos `cambio_etapa` viejos (quedarían con el
texto viejo salvo que también se reescriban). Más trabajo y más riesgo de
dejar algo sin migrar — no es la opción recomendada acá.

### "Negociación" con su propio chip

Hoy `crm.kanban_grupo()` agrupa "Cotizado" y "Negociación" bajo el mismo
grupo `"cotizados"`. Separarlos: "Negociación" pasa a su propio grupo
(ej. `"negociacion"`), sumado a `KANBAN_COLUMNAS` en `app.py` con su
propio ícono/label. El CSS por etapa en `_inyectar_estilos` (borde
izquierdo coloreado por `crm.COLOR_ETAPA`) ya es genérico — no hace falta
tocar nada ahí, se acomoda solo.

### Botones manuales para "Negociación" y "Descartado"

En la tarjeta de un contacto (`_render_fila_contacto_crm`):
- Cuando la etapa es "Cotizado", sumar un botón "Marcar en negociación"
  (mismo estilo que "Avanzar etapa") que llame a
  `db.change_etapa_contacto(id, "Negociación", autor)`.
- En cualquier etapa antes de "Cliente"/Ganado, sumar un botón para
  "Descartado" — hoy solo se puede marcar desde el selector libre en la
  ficha completa o desde la sugerencia en la pestaña Seguimiento; que
  quede igual de accesible que "Avanzar etapa" directo en la tarjeta.

### Columna "Cliente" visible con resumen operativo en vivo

Hoy `crm.kanban_grupo()` devuelve `None` para etapa "Ganado" (se excluye
de la vista). Darle su propio grupo/columna, pero con un tratamiento
distinto al resto: en vez de los botones normales (Avanzar etapa/Ver
ficha), cada tarjeta ahí muestra un resumen de solo lectura tipo *"N
importación(es) en curso · última actividad hace X día(s)"*, calculado del
lado administrativo (reutilizar/extender `db.contar_importaciones_por_cliente`
o una consulta similar a la que ya arma `en_proceso_por_cliente` en
`vista_clientes()`). Esto reemplaza la idea de un reporte aparte para
comercial — mirando su propio CRM ya ve el estado operativo de cada
cliente ganado.

**Cómo probarlo:** en CRM → Contactos deben verse 6 chips (Nuevos,
Contactados, Cotizados, Negociación, Cliente, Descartado) con las
etiquetas nuevas donde corresponda. Un contacto en "Cotizado" debe poder
pasar a "Negociación" con un clic, y a "Descartado" desde cualquier etapa
anterior a "Cliente". La columna "Cliente" debe mostrar el resumen
operativo, no los botones de avance normales.

**✅ Resuelto y verificado (2026-09-04/06).** Los 6 chips existen con las
etiquetas nuevas, `crm.ETAPA_LABEL` cubre badges/selectbox/timeline/
Analíticas, y la columna "Cliente" muestra el resumen operativo en vez de
botones normales. Cambio de nombre respecto al prompt original: el botón
quedó "Negociación" (no "Marcar en negociación") — ese texto más largo
rompía la palabra a la mitad en la mitad angosta de la card en pantallas
chicas.

---

## Prompt 4 — Honorario Skybridge (ingreso propio)

Contexto de negocio importante: Skybridge es una consultora, no el
importador — `ganancia_bruta_usd`/`rentabilidad_pct` que ya calcula
`calculo.py` es el margen del CLIENTE, no ingreso de Skybridge. Esto es
información nueva y separada.

1. **Nueva tabla** `importacion_honorarios` en `db.py` (agregar al
   `SCHEMA`, mismo patrón que `importacion_documentos`): `id`,
   `importacion_id` (FK a `importaciones`, `ON DELETE CASCADE`),
   `concepto` (TEXT), `monto` (REAL), `moneda` (TEXT, default 'USD'),
   `estado` (TEXT, default 'Pendiente' — valores: Pendiente/Facturado/Cobrado),
   `creado_en` (default datetime now). Sumar las funciones CRUD
   correspondientes (`add_honorario`, `list_honorarios`,
   `update_honorario`, `delete_honorario`), mismo estilo que las de
   `importacion_documentos`.

2. **UI**: nueva sección dentro del tab **"🧾 Facturas y pagos"** que ya
   existe en la ficha de importación (`_render_contenido_importacion`,
   junto a `DOC_CATEGORIA_FACTURAS`/`DOC_CATEGORIA_PAGOS`) — no crear un
   lugar nuevo. Lista simple de honorarios (concepto, monto, estado) con
   alta/edición/baja, mismo lenguaje visual que `_render_gastos()` del
   cotizador (una fila por honorario, expandible). Al marcar un honorario
   como "Cobrado", si ya hay algún documento cargado en la categoría
   "Pagos" de esa misma importación, permitir enlazarlo como comprobante
   en vez de tener que subir el archivo de nuevo.

3. **No** sumar por defecto un honorario en cero al crear la importación
   (a diferencia de `DEFAULT_GASTOS` en las cotizaciones) — evaluar en
   cambio si tiene sentido auto-sugerir (no forzar) el concepto
   "Importación realizada" cuando el estado de la importación pasa a
   "Entregada", ya que hoy es el único hecho que efectivamente genera
   cobro. Confirmar este detalle si genera duda antes de implementarlo.

**Cómo probarlo:** abrir una importación, ir a "Facturas y pagos", cargar
un honorario con concepto/monto, marcarlo Cobrado, recargar la página y
confirmar que persiste.

---

## Prompt 5 — Limpieza y optimización (sin apuro, no bloquea nada)

1. **PDFs regenerados en cada rerun:** `_generar_pdfs_cotizacion` (llamada
   desde `_render_fila_cotizacion` en el historial) recalcula la
   cotización y arma 2 PDFs con reportlab en cada rerun del script,
   incluida cada letra tipeada en el buscador. Cachear con `st.cache_data`,
   clave `(cotizacion_id, actualizado_en)` — así una cotización sin
   cambios no se recalcula ni regenera en cada interacción.
   **✅ Resuelto (2026-09-06)**, tal cual: `@st.cache_data` con
   `(cot_id, actualizado_en)` como parámetros de la función.

2. **Nombre de archivo PDF duplicado:** la misma lógica de armar
   `numero_pdf`/`nombre_pdf` está una vez en `_generar_pdfs_cotizacion` y
   calcada de nuevo dentro de `_render_cotizador_editor`. Extraer a una
   función compartida `(cab, resultado) -> nombre_pdf` y usarla en los dos
   lugares.
   **✅ Resuelto (2026-09-06)**: `_nombre_pdf_cotizacion(cab, resultado)`.
   Bonus: se encontró de paso un tercer botón de "Eliminar cotización"
   (en el editor completo) que borraba sin diálogo de confirmación ni
   backup — se corrigió para que reutilice el mismo diálogo del historial.

3. **Columnas reutilizadas sin documentar:** en `cotizaciones`,
   `contenedor` guarda "condición de venta", `etd_eta` guarda "tipo de
   envío", `carrier`/`freetime` guardan las fechas ETD/ETA — funciona bien
   pero solo se entiende leyendo el formulario. Agregar un comentario
   claro arriba de esas columnas en el `SCHEMA` de `db.py` explicando la
   reutilización (mismo criterio que se va a usar para "Ganado"/"Perdido"
   en el prompt 3) — no hace falta migrar nada, alcanza con dejarlo
   documentado donde se define el schema.

**Cómo probarlo:** no cambia comportamiento visible — alcanza con
descargar un PDF y guardar una cotización una vez para confirmar que
sigue andando igual.
