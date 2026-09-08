# Prompt de implementación — Cotizador: seguro configurable, TC Venta para simulación, e Incoterms (EXW suma a CIF)

Contexto para quien lo programe: motor de cálculo en `calculo.py` (replica
"Simulador Skybridge.xlsx", hojas `Presupuesto` y `Costeo`), consumido desde
`_render_cotizador_editor()` en `app.py`. Viene de una auditoría formula por
fórmula contra el Excel original — lo que sigue son 4 cambios que el dueño
del negocio confirmó explícitamente después de esa auditoría, no hipótesis.
**No tocar ninguna fórmula fuera de lo que se describe acá** — el resto de
`calculo.py` ya calza exacto con el Excel y así debe quedar.

Convenciones a respetar (ya establecidas en el proyecto): nombres en
español, comentarios que explican el POR QUÉ, columnas nuevas en SQLite vía
el dict `COLUMNAS_NUEVAS` en `db.py` (ALTER TABLE en cada arranque),
persistir todo campo nuevo en `save_cotizacion()`.

Antes de tocar código: correr `python3 -m py_compile app.py calculo.py db.py`
después de cada cambio, y verificar con una cotización de prueba (no la base
real) que el "CHECK de consistencia" (`check_diferencia`, sección 9️⃣ del
editor, "Ver costeo unitario por producto") sigue dando ~$0 — es el mismo
check que ya trae el Excel (`Costeo!B19`) y detecta cualquier fórmula que
quede inconsistente.

---

## Cambio 1 — Seguro: modo Automático / No cobrar / Manual

**Confirmado por el dueño:** el 0,3% se calcula sobre el FOB DECLARADO (no
el FOB real — así quedó la app, y así debe seguir), con un piso de USD 75,
más IVA. Esa fórmula **ya está bien** en `calculo.py` (líneas ~109-115,
`seguro_declarado = max(fob_declarado * seguro_pct, SEGURO_MINIMO_USD) if
fob_declarado > 0 else 0.0`) — **no tocarla**. Lo que falta es la opción de
no cobrarlo o cargarlo a mano.

1. **`db.py`**: agregar a `COLUMNAS_NUEVAS` dos columnas en `cotizaciones`:
   `seguro_modo TEXT DEFAULT 'auto'` (valores: `'auto' | 'ninguno' |
   'manual'`) y `seguro_manual_usd REAL DEFAULT 0`.

2. **`calculo.py`**, en `calcular()`, junto a donde se leen `seguro_pct` /
   `arancel_sim` (línea ~36-39): agregar
   ```python
   seguro_modo = cab.get("seguro_modo") or "auto"
   seguro_manual_usd = _f(cab.get("seguro_manual_usd"))
   ```
   Y reemplazar el cálculo de `seguro_declarado` (línea ~115) por:
   ```python
   if seguro_modo == "ninguno":
       seguro_declarado = 0.0
   elif seguro_modo == "manual":
       seguro_declarado = seguro_manual_usd
   else:  # "auto" — misma fórmula de siempre, sin tocar
       seguro_declarado = max(fob_declarado * seguro_pct, SEGURO_MINIMO_USD) if fob_declarado > 0 else 0.0
   ```
   Todo lo que ya consume `seguro_declarado` (CIF, gasto operativo "Seguro"
   sincronizado) sigue funcionando igual, sin más cambios — el valor
   simplemente llega calculado distinto según el modo.

3. **`app.py`**: en `_render_cotizador_editor()`, sección "3️⃣ Tarifas
   flete" (donde hoy está el campo disabled "Seguro (USD)", buscar
   `segurohdr_key` — línea ~2569-2573 actual): agregar un
   `st.selectbox("Seguro", ["Automático", "No cobrar", "Manual"], ...)` y,
   solo cuando esté en "Manual", un `st.number_input` editable para el
   monto. El campo "Seguro (USD)" que ya existe (disabled, mostrando
   `resultado["seguro_declarado"]`) queda igual para los modos Automático y
   No cobrar (sigue siendo de solo lectura, ahí se ve el resultado); en modo
   Manual, el input editable reemplaza a ese campo disabled (o el disabled
   pasa a reflejar lo que el usuario tipeó, a gusto de quien lo programe).
   Pasar `seguro_modo` y `seguro_manual_usd` a `cab_calc` (línea ~2462) y
   guardarlos en `_guardar_cotizacion()` junto con el resto de la cabecera.

**Cómo probarlo:** cargar una cotización con productos, dejar Seguro en
Automático y confirmar que no cambia nada vs. hoy. Cambiar a "No cobrar" y
confirmar que el CIF baja exactamente el monto que tenía el seguro antes, y
que el gasto operativo "Seguro" también queda en 0. Cambiar a "Manual",
poner un monto arbitrario, y confirmar que ese mismo monto aparece tanto en
el CIF como en el gasto operativo "Seguro".

---

## Cambio 2 — TC Venta (para la Simulación de venta) + fix de Ganancia bruta (ARS)

Son dos cosas relacionadas pero **con fórmulas distintas** — no mezclarlas.

### 2a. TC Venta — nueva tasa, usada SOLO adentro de "🔟 Simulación de venta"

El Excel original tenía 3 tipos de cambio independientes (Tributos,
Operativos, Venta) porque mercadería, tributos y venta se pagan/cobran en
fechas distintas. La app había eliminado el tercero reusando TC Operativos
(línea 34 de `calculo.py`: `tc_venta = tc_oper`). El dueño confirmó: la
Simulación de venta es informativa (se manda al cliente para que pruebe
escenarios), así que el TC Venta se reincorpora ahí — pero **el costo final
(c/IVA, s/IVA, incidencia s/FOB) no depende de esto y no se toca**.

1. **`db.py`**: agregar a `COLUMNAS_NUEVAS`: `tc_venta REAL DEFAULT 0`
   (mismo patrón que `tc_tributos`/`tc_operativos`).

2. **`calculo.py`**, línea 34: reemplazar
   `tc_venta = tc_oper  # TC Venta ARS se eliminó del formulario...`
   por
   `tc_venta = _f(cab.get("tc_venta"))` (default 0, igual que el resto).

3. **`calculo.py`**, en el loop por producto (donde se arman
   `pv_sugerido_ars`, `pv_final_ars`, `ganancia_unit_ars`,
   `ganancia_total_ars` — líneas ~238-241):
   - `pv_sugerido_ars` queda **igual** (`pv_sugerido * tc_oper if tc_oper
     else 0.0`) — en el Excel, `Costeo!D42 = B42*(1+C42)` usa el costo ya
     convertido a ARS por TC Operativos, no TC Venta. Es el único de los
     cuatro que NO cambia.
   - `pv_final_ars`: cambiar de `pv_final * tc_venta` (que hoy equivale a
     `pv_final * tc_oper` porque `tc_venta==tc_oper`) a que dependa del
     nuevo `tc_venta` real: `pv_final * tc_venta if tc_venta else 0.0`
     (con el `tc_venta` ya corregido en el punto anterior, esta línea de
     código no cambia en sí — cambia lo que `tc_venta` vale).
   - `ganancia_unit_ars`: **esta fórmula sí hay que cambiarla.** Hoy es
     `ganancia_unit * tc_oper if tc_oper else 0.0` (multiplica la ganancia
     en USD por TC Operativos). En el Excel, `Costeo!H42 =
     IF(E42>0,E42-B42,0)` — es la RESTA de dos valores en ARS que usan
     tasas distintas (`E42`=PV Final ARS a TC Venta, `B42`=Costo ARS a TC
     Operativos), no la ganancia en USD multiplicada por una sola tasa.
     Reemplazar por (usando `pv_final_ars` y `costo_civa_unit_ars`, que ya
     están calculados en ese punto del loop):
     ```python
     ganancia_unit_ars = (pv_final_ars - costo_civa_unit_ars) if (pv_final > 0 and tc_venta > 0) else 0.0
     ```
   - `ganancia_total_ars = ganancia_unit_ars * cantidad` — la fórmula en sí
     no cambia, solo hereda el nuevo `ganancia_unit_ars`.

4. **`app.py`**: agregar el input "TC Venta (ARS, para la simulación)" en
   la sección "🔟 Simulación de venta" (`_render_cotizador_editor`, justo
   antes de llamar a `_render_simulacion_venta(resultado)` — línea ~2681),
   con su propio `key=f"tcventa_{cot_id}"`, agregarlo a `cab_calc` y a
   `_guardar_cotizacion()`. Default 0 (igual que el Excel: con TC Venta en
   0, todas las cifras ARS de esta sección dan 0, exactamente el gate `IF
   (B$40>0,...,0)` que ya tiene el Excel).

### 2b. Ganancia bruta (ARS) — fix, independiente de TC Venta

Esto es un bug encontrado en la auditoría, no pedido explícitamente pero el
dueño lo asoció al punto anterior — en realidad es una fórmula aparte y
**sigue usando TC Operativos, no TC Venta** (así lo hace el Excel:
`Costeo!C70 = B70*Presupuesto!B$86`, con `B$86`=TC Operativos).

Hoy (`calculo.py` línea 248): `ganancia_bruta_ars = ganancia_bruta_usd *
tc_oper if tc_oper else 0.0` — esto da mal apenas TC Tributos ≠ TC
Operativos, porque `total_c_iva_ars` (con el que se calculó `total_c_iva_usd`
que arma `ganancia_bruta_usd`) mezcla TC Tributos para los tributos y TC
Operativos para el resto — multiplicar el delta en USD por una sola tasa
pierde esa mezcla. El Excel (`Costeo!C71 = IF(C70>0,C70-C69,0)`) resta
directamente dos totales en ARS ya bien calculados.

Reemplazar (línea ~246-248):
```python
venta_total_ars = venta_total_usd * tc_oper if tc_oper else 0.0  # sin cambios, ya es igual al Excel
ganancia_bruta_ars = (venta_total_ars - total_c_iva_ars) if venta_total_ars > 0 else 0.0
```
(`total_c_iva_ars` ya está calculado correctamente más arriba, línea 187 —
solo hay que restarlo en vez de reconstruirlo desde el lado USD).

**Cómo probarlo:** cargar una cotización con TC Tributos y TC Operativos
DISTINTOS (ej. 1000 y 1300) y un PV Final > 0 en al menos un producto.
"Ganancia bruta (ARS)" en la sección 🔟 debe coincidir con Venta total
(ARS) menos TOTAL desembolsado (ARS) de la sección 9️⃣ — hoy no coincide,
después del fix sí. Para el TC Venta: dejarlo en 0 y confirmar que PV
Final/Ganancia unitaria (ARS) por producto dan $0 (gate igual al Excel);
ponerle un valor y confirmar que esos dos números cambian, pero que el
costo c/IVA, costo s/IVA e incidencia s/FOB de las secciones 9️⃣ no se
mueven un centavo.

---

## Cambio 3 — PV Mercado: activarlo en la Simulación de venta

Hoy `calculo.py` lee `pv_mercado_usd` del producto pero no lo usa para
nada, y `app.py` no lo muestra en ningún lado (campo huérfano). El Excel sí
lo tiene y lo convierte a ARS con el mismo TC Venta del Cambio 2
(`Costeo!F42 = IF(B$40>0,F23*B$40,0)`).

1. **`calculo.py`**, junto al resto de campos ARS por producto (línea
   ~238-241): agregar
   ```python
   pv_mercado_ars=pv_mercado * tc_venta if tc_venta else 0.0,
   ```
   (usando la misma variable local `pv_mercado` que ya se lee de
   `p["pv_mercado_usd"]`, si no existe como variable local agregarla igual
   que `pv_final`).

2. **`app.py`**, en `_render_simulacion_venta()` (línea ~2304): agregar un
   `st.number_input("PV Mercado (USD)", ...)` para `p["pv_mercado_usd"]`
   (mismo patrón que el `pv_final_usd` ya existente, línea ~2317) y mostrar
   su equivalente en ARS junto a los demás `st.metric` de esa fila. Es
   informativo (referencia de precio de mercado/competencia), no entra en
   ningún cálculo de costo ni de margen — mismo rol que tiene en el Excel.

**Cómo probarlo:** cargar un PV Mercado en un producto, confirmar que se
guarda y se muestra en USD y en ARS (usando el TC Venta del Cambio 2), y
que no mueve ningún otro número de la cotización.

---

## Cambio 4 — Incoterms: EXW suma Gastos en origen al CIF; FOB/FCA no cobran gastos en origen

**Confirmado por el dueño:** si la Condición de venta es FOB o FCA, no se
cobran gastos en origen. Si es EXW, sí se cobran, y esos gastos en origen
entran a la base del CIF Declarado (junto con FOB declarado + flete
declarado + seguro declarado) — es decir, Derechos, Tasa Estadística,
Antidumping, IVA, Ganancias e IIBB por producto pasan a calcularse sobre un
CIF más alto. El Costo financiero **NO se toca**: sigue siendo únicamente
`FOB Total (mercadería) × Costo financiero %` — el dueño fue explícito en
que el costo financiero es aparte (el spread de la transferencia
internacional) y no participa ni del CIF ni de los tributos.

1. **`app.py`**, `cab_calc` (línea ~2462-2466): agregar
   `"gastos_origen": gastos_origen, "condicion_venta": contenedor,` (la
   variable de Condición de venta ya existe en el editor con el nombre
   `contenedor`, viene del selectbox `COT_CONDICIONES_VENTA` — línea
   ~2388-2393 — es nombre heredado, no hace falta renombrarla).

2. **`calculo.py`**, junto a donde se leen `tarifa_flete` / `pct_certificacion`
   (línea ~38-39): agregar
   ```python
   gastos_origen = _f(cab.get("gastos_origen"))
   condicion_venta = cab.get("condicion_venta") or ""
   gastos_origen_taxable = gastos_origen if condicion_venta == "EXW" else 0.0
   ```

3. **`calculo.py`**, CIF declarado agregado (línea ~116):
   ```python
   cif_declarado = fob_declarado + flete_declarado + seguro_declarado + gastos_origen_taxable
   ```

4. **`calculo.py`**, CIF por producto (línea ~121, dentro del loop de
   derechos/tasas/antidumping):
   ```python
   cif_prod = p["fob_decl_total"] + (flete_declarado + seguro_declarado + gastos_origen_taxable) * p["pct_partic"]
   ```
   (mismo patrón de prorrateo por `pct_partic` que ya usan flete y seguro —
   así el CIF agregado sigue siendo exactamente la suma de los CIF por
   producto, sin inconsistencia).

5. **`app.py`**, sección "3️⃣ Tarifas flete" (campo "Gastos en origen
   (USD)", `tf4` — línea ~2557-2562): cuando `contenedor` (Condición de
   venta) sea "FOB" o "FCA", deshabilitar ese input y forzar
   `gastos_origen = 0` (mismo criterio que ya se usa en otros campos
   disabled del formulario, ej. "Tarifa flete certificado"). Agregar un
   `st.caption` corto explicando por qué está deshabilitado (ej. "No aplica
   con FOB/FCA — el exportador ya cubre los gastos hasta el puerto de
   origen"). Con EXW, el campo queda editable como hoy.

**Cómo probarlo:** con Condición de venta = FOB, confirmar que "Gastos en
origen" queda en 0 y no se puede tipear. Cambiar a EXW, cargar un monto de
gastos en origen, y confirmar que (a) el CIF Declarado (sección 4️⃣) subió
exactamente ese monto, (b) Derechos/Tasa/Antidumping/IVA/Ganancias/IIBB de
cada producto subieron en proporción a su % de participación, y (c) el
Costo financiero (sección 9️⃣, "① Mercadería + fin.") **no se movió**.

---

## Orden recomendado

1 → 4 → 2a → 2b → 3. El 4 es independiente y de bajo riesgo (solo suma un
término más al CIF, mismo patrón que flete/seguro). El 2b conviene hacerlo
junto con el 2a porque toca las mismas líneas del loop de productos. El 3
es el más chico y puede ir en cualquier momento, incluso en paralelo.

Después de los 4, correr de nuevo el CHECK de consistencia (`Costeo!B19`
equivalente, sección 9️⃣ del editor) con al menos dos cotizaciones de
prueba: una con Condición de venta EXW y TC Tributos ≠ TC Operativos (para
ejercitar los 4 cambios a la vez), y confirmar que sigue dando ~$0.
