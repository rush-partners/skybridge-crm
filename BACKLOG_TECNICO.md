# Backlog técnico — Skybridge ERP/CRM

Cambios agendados para implementar **antes** de sumar el login (ver
[Metodología Skybridge](https://claude.ai/code/artifact/3afe1e36-fe1a-4f98-8776-008f0d0ae7a2)
para el porqué de cada uno). Nada de esto está implementado todavía — es
el registro de lo que hay que pedir, en orden razonable.

## 1. Errores a corregir

- [x] **Filtro "sin actividad hace 14+ días" en Contactos marca mal a los
      contactos con actividad HOY.** ~~`_dias_desde_actividad()` devuelve
      `0` para actividad de hoy, y `0 or 10**9` en Python da `10**9` (0 es
      falsy)~~ — **resuelto de dos formas, no solo una**: primero se
      corrigió el bug puntual (`_es_inactivo_14_dias` explícito, sin el
      `or`), pero después el checkbox "Mostrar solo con seguimiento
      vencido"/"Mostrar solo sin actividad hace 14+ días" se **eliminó por
      completo** a pedido del usuario (no convencía la UX) y se reemplazó
      por un orden automático en `_render_lista_contactos`: cada columna
      del Kanban ordena sola por días desde la última actividad (o desde el
      alta), sin ningún filtro manual que activar. El bug ya no puede
      volver a pasar porque el código que lo tenía no existe más.

## 2. Sincronía comercial ↔ administrativo — ✅ resuelto (2026-09-04)

Objetivo: las dos puertas de entrada (CRM primero, o Clientes primero)
siguen coexistiendo, pero todo cliente termina siempre con un contacto de
CRM vinculado, sin carga doble. Reemplaza y resuelve de raíz el bug
original "clientes cotizados directo desde Clientes quedan invisibles"
(`app.py:1443`/`app.py:1622`).

- [x] Al crear un cliente directo desde "＋ Nuevo cliente" (módulo
      Clientes), crear también su contacto de CRM en la etapa "Nuevo" en el
      mismo momento — no esperar a que pase otra cosa.
      Implementado: `vista_clientes()` llama a
      `_asegurar_contacto_para_cliente(new_id)` justo después de
      `db.create_cliente(...)`.
- [x] Al generar una cotización desde la ficha de Clientes (`_cotizar_cliente`),
      hacer avanzar el contacto vinculado a "Cotizado" — mismo criterio que
      ya usa `_cotizar_contacto` en el sentido inverso (CRM → Clientes).
      Implementado tal cual: asegura el contacto y llama a
      `_avanzar_etapa_si_corresponde(..., "Cotizado", ...)`.
- [x] Nota de revisión (2026-09-04): ya existe un tercer lugar que cambia el
      estado de una cotización aparte del editor —
      `_cambiar_estado_cotizacion` (control rápido "Cambiar" en cada fila
      del historial) — y ya dispara `_marcar_ganado_por_cliente`
      correctamente cuando el estado pasa a "Aprobada". Se mantuvo ese
      criterio sin tocarlo, tal como pedía la nota.
- [x] Aplicar esto **retroactivamente**: `_migrar_contactos_retroactivos()`
      (app.py) ya existe, es idempotente (salta clientes que ya tienen
      contacto vinculado) y clasifica por la misma prioridad de señales que
      `crm.estado_comercial()`. **Ya se corrió contra la base real**:
      BARWATEX SRL y el resto de los clientes viejos con importaciones
      aparecieron directo en "Cliente", no en "Nuevo" — confirmado. Además
      se le sumó un botón en Panel de Control → Configuración
      ("▶️ Ejecutar migración") para poder correrla desde la UI cuando haga
      falta (ej. después de una carga masiva de clientes viejos), en vez de
      depender de que alguien la dispare a mano por consola.

## 3. Vocabulario y vista del embudo (CRM) — ✅ resuelto (2026-09-04)

- [x] Renombrar en toda la app (base, UI, PDFs si corresponde) "Ganado" →
      **"Cliente"** y "Perdido" → **"Descartado"**. Mismo color y misma
      lógica de cada etapa, solo cambia el texto.
      Implementado exactamente como recomendaba este ítem: capa de
      visualización (`crm.ETAPA_LABEL`), sin tocar el valor guardado en
      `contacts.etapa` ni ninguna comparación existente. Aplicado en
      badges, chips del kanban, selectbox de etapa, timeline de
      `cambio_etapa` y Analíticas (funnel, tasas, tiempos).
- [x] "Negociación" pasa a tener su propio chip en la vista de Contactos,
      separado de "Cotizado". `crm.kanban_grupo()` la separa en su propio
      grupo `"negociacion"`. El botón manual quedó con el label
      "Negociación" (no "Marcar en negociación" — el texto largo rompía la
      palabra a la mitad en la mitad angosta de la card en pantallas
      chicas; ver fix de `overflow-wrap` en el CSS de `filacrm_`).
- [x] La columna/chip "Cliente" deja de estar oculta del CRM y muestra un
      resumen operativo en vivo por cliente ("N importación(es) en curso ·
      última actividad hace X día(s)", vía
      `db.contar_importaciones_en_curso_por_cliente`) en vez de los
      botones normales de avance.

## 4. Honorario Skybridge (ingreso propio, no el del cliente)

- [ ] Nueva lista de honorarios por importación: **concepto + monto +
      estado de cobro** (Pendiente/Facturado/Cobrado), en vez de un campo
      único — pensada para poder sumar conceptos nuevos (gestión, búsqueda
      de proveedores, cotización como servicio) más adelante sin rediseñar.
      Vive en `importaciones`, no en la cotización — ver la nota de
      [[business-model-skybridge]] sobre por qué `ganancia_bruta_usd`/
      `rentabilidad_pct` del cotizador NO son ingreso de Skybridge.
- [ ] Concepto inicial a cargar: "Importación realizada" — es hoy el único
      hecho que genera cobro real.
- [ ] Ubicación en la UI: la ficha de importación ya tiene un tab dedicado
      **"🧾 Facturas y pagos"** (`app.py:1147-1153`, con las categorías de
      documentos `DOC_CATEGORIA_FACTURAS`/`DOC_CATEGORIA_PAGOS` — revisado
      2026-09-04) — el honorario de Skybridge va ahí, como una sección
      más, no en un lugar nuevo. De paso, un honorario "Cobrado" puede
      referenciar/adjuntar un archivo ya subido a la categoría "Pagos" de
      ese mismo tab como comprobante, en vez de duplicar la carga de
      archivos.

## 5. Optimizaciones (no bloquean nada, hacer cuando haya lugar)

- [x] **Regeneración de PDFs en cada rerun del historial de cotizaciones.**
      ~~Cada fila visible genera 2 PDFs con reportlab en cada rerun del
      script~~ — resuelto (2026-09-06): `_generar_pdfs_cotizacion` ahora
      tiene `@st.cache_data`, clave `(cotizacion_id, actualizado_en)` (ya
      presente en la fila de `db.list_cotizaciones`, se bumpea en cada
      guardado o cambio de estado).
- [x] **Lógica de nombre de archivo PDF duplicada** entre
      `_generar_pdfs_cotizacion` y el bloque calcado dentro de
      `_render_cotizador_editor` — resuelto (2026-09-06): extraída a
      `_nombre_pdf_cotizacion(cab, resultado)`, usada en los dos lugares.
      De paso se encontró y arregló un tercer punto de baja de cotización
      (botón "🗑️ Eliminar cotización" del editor completo) que borraba sin
      ningún diálogo de confirmación ni backup — ahora reutiliza el mismo
      `_dialog_eliminar_cotizacion` que ya usaba el historial.
- [ ] **Columnas de `cotizaciones` reutilizadas sin documentar en el
      schema:** `contenedor` guarda "condición de venta", `etd_eta` guarda
      "tipo de envío", `carrier`/`freetime` guardan las fechas ETD/ETA. Es
      consistente en todos lados pero solo se entiende leyendo el
      formulario (`app.py:1887-1903`) y el PDF (`pdf_export.py:102-107`).
      Ver `db.py:22-49`. Evaluar `ALTER TABLE ... RENAME COLUMN` o, como
      mínimo, comentar el `SCHEMA`.

## No hace falta hacer aparte

- El campo "autor" de texto libre en las notas del CRM (`crm_autor`) se
  resuelve solo con el login — una vez que cada quien entra identificado,
  no hace falta un selector manual de nombre aparte. No sumar como tarea
  propia.

## Después de esto (fuera de este backlog)

- Login de los 3 usuarios + migración SQLite → MySQL (Hostinger) +
  despliegue en Render — el usuario lo va a subir él mismo, se planifica
  aparte cuando este backlog esté cerrado.
- Sin autenticación en toda la app todavía — cualquiera en la red local
  accede a clientes, CUIT y cotizaciones. Se resuelve con el login de arriba.
