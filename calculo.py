"""Motor de cálculo del cotizador de importación.

Replica fielmente las fórmulas de 'Simulador Skybridge.xlsx' (hojas
'Presupuesto' y 'Costeo'):
  - Detalle de mercadería -> FOB total, FOB declarado, % participación
  - Costo financiero (recargo % s/FOB)
  - Valor declarado en aduana (FOB + flete + seguro declarados => CIF)
  - Derechos, tasa estadística y antidumping por producto
  - IVA, IVA adicional, Ganancias e IIBB (percepciones) por producto
  - Costos operativos (prorrateo por FOB o por peso/volumen) e IVA de terceros
  - Totales de desembolso (con y sin IVA/percepciones), en USD y ARS
  - Costeo unitario por producto y simulación de venta (margen sugerido/real)
  - Resultado de la operación (ganancia bruta, rentabilidad %)
"""
from copy import deepcopy


def _f(v, default=0.0):
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def calcular(cabecera: dict, productos_in: list, gastos_in: list) -> dict:
    cab = cabecera or {}
    productos = [dict(p) for p in productos_in] if productos_in else []
    gastos = [dict(g) for g in gastos_in] if gastos_in else []

    tc_trib = _f(cab.get("tc_tributos"))
    tc_oper = _f(cab.get("tc_operativos"))
    tc_venta = _f(cab.get("tc_venta"))  # tasa propia, usada solo en Simulación de venta (ver Costeo!B40)
    cf_pct = _f(cab.get("costo_financiero_pct"), 0.025)
    seguro_pct = _f(cab.get("seguro_pct"), 0.003)
    seguro_modo = cab.get("seguro_modo") or "auto"
    seguro_manual_usd = _f(cab.get("seguro_manual_usd"))
    arancel_sim = _f(cab.get("arancel_sim"), 10)
    tarifa_flete = _f(cab.get("tarifa_flete"))
    pct_certificacion = _f(cab.get("pct_certificacion"), 0.5)
    # Con FOB/FCA el exportador ya cubre los gastos hasta el puerto de
    # origen — no se cobran, y por lo tanto no entran a la base del CIF
    # Declarado. Con EXW sí corren por cuenta propia, y entran a esa base
    # (junto con FOB + flete + seguro declarados), igual que el Excel.
    gastos_origen = _f(cab.get("gastos_origen"))
    condicion_venta = cab.get("condicion_venta") or ""
    gastos_origen_taxable = gastos_origen if condicion_venta == "EXW" else 0.0

    for p in productos:
        p["fob_unit"] = _f(p.get("fob_unit"))
        p["cantidad"] = _f(p.get("cantidad"))
        p["fob_decl_unit"] = _f(p.get("fob_decl_unit"))
        p["peso_kg"] = _f(p.get("peso_kg"))
        p["volumen_m3"] = _f(p.get("volumen_m3"))
        # Peso cobrable (regla W/M de flete marítimo LCL): se compara el peso
        # en toneladas contra el volumen en m3 (equivalencia 1 t = 1 m3) y se
        # toma el mayor, que es lo que efectivamente factura el forwarder.
        p["peso_vol"] = max(p["peso_kg"] / 1000.0, p["volumen_m3"])
        p["pct_derechos"] = _f(p.get("pct_derechos"))
        p["pct_tasa_estadistica"] = _f(p.get("pct_tasa_estadistica"), 0.03)
        p["pct_antidumping"] = _f(p.get("pct_antidumping"))
        p["pct_iva"] = _f(p.get("pct_iva"), 0.21)
        p["pct_iva_adicional"] = _f(p.get("pct_iva_adicional"))
        p["pct_ganancias"] = _f(p.get("pct_ganancias"))
        p["pct_iibb"] = _f(p.get("pct_iibb"), 0.05)
        p["margen_pct"] = _f(p.get("margen_pct"), 0.3)
        p["pv_final_usd"] = _f(p.get("pv_final_usd"))
        p["pv_mercado_usd"] = _f(p.get("pv_mercado_usd"))
        p["fob_total"] = p["fob_unit"] * p["cantidad"]
        p["fob_decl_total"] = p["fob_decl_unit"] * p["cantidad"]

    for g in gastos:
        g["monto"] = _f(g.get("monto"))
        g["pct_iva"] = _f(g.get("pct_iva"), 0.21)
        g["moneda"] = g.get("moneda") or "USD"
        g["prorrateo"] = g.get("prorrateo") or "FOB"
        g["iva_incluido"] = g.get("iva_incluido") or "NO"

    fob_total_sum = sum(p["fob_total"] for p in productos)       # E33
    fob_decl_total_sum = sum(p["fob_decl_total"] for p in productos)  # G33
    peso_total_sum = sum(p["peso_vol"] for p in productos)       # I33

    for p in productos:
        p["pct_partic"] = (p["fob_decl_total"] / fob_decl_total_sum) if fob_decl_total_sum > 0 else 0.0

    costo_financiero = fob_total_sum * cf_pct               # E34
    fob_mas_financiero = fob_total_sum + costo_financiero    # G34

    # ---- Costos operativos (USD/ARS por línea) ----
    def gasto_usd(g):
        # Antes dividía siempre por 1.21 (21% fijo) sin importar el % de IVA
        # propio del gasto — con un gasto "IVA incluido" a una tasa distinta
        # de 21% (ej. 10.5%), el neto quedaba mal calculado.
        base_incl = g["monto"] / (1 + g["pct_iva"]) if g["iva_incluido"] == "SI" else g["monto"]
        if g["moneda"] == "USD":
            return base_incl
        return (base_incl / tc_oper) if tc_oper > 0 else 0.0

    def gasto_ars(g):
        base_incl = g["monto"] / (1 + g["pct_iva"]) if g["iva_incluido"] == "SI" else g["monto"]
        if g["moneda"] == "ARS":
            return base_incl
        return base_incl * tc_oper

    for g in gastos:
        g["usd"] = gasto_usd(g)
        g["ars"] = gasto_ars(g)

    # ---- Valor declarado en aduana ----
    # Tarifa flete certificada = Tarifa flete (Costos operativos, gasto real
    # pagado al forwarder) x % Certificación. Antes era un 50% fijo (flete/2);
    # ahora el % es configurable, pero con 50% por defecto reproduce el mismo
    # cálculo que tenía el Excel original.
    tarifa_flete_certificado = tarifa_flete * pct_certificacion
    flete_declarado = tarifa_flete_certificado  # B39
    fob_declarado = fob_decl_total_sum       # B38
    # Seguro sobre el FOB declarado (no el FOB real): así el CIF Declarado
    # queda armado 100% sobre base declarada, sin mezclar bases distintas
    # dentro del mismo valor que se presenta en aduana. Prima mínima de
    # USD 75 (piso habitual de póliza), salvo que no haya mercadería
    # cargada todavía (fob_declarado=0), donde no corresponde cobrar nada.
    SEGURO_MINIMO_USD = 75.0
    if seguro_modo == "ninguno":
        seguro_declarado = 0.0
    elif seguro_modo == "manual":
        seguro_declarado = seguro_manual_usd
    else:  # "auto" — misma fórmula de siempre, sin tocar
        seguro_declarado = max(fob_declarado * seguro_pct, SEGURO_MINIMO_USD) if fob_declarado > 0 else 0.0  # B40
    cif_declarado = fob_declarado + flete_declarado + seguro_declarado + gastos_origen_taxable  # B41

    # ---- Derechos, tasas y antidumping por producto ----
    subtotal_der = subtotal_tasa = subtotal_antid = subtotal_dertasas = 0.0
    for p in productos:
        cif_prod = p["fob_decl_total"] + (flete_declarado + seguro_declarado + gastos_origen_taxable) * p["pct_partic"]  # B45
        derechos = cif_prod * p["pct_derechos"]
        tasa = cif_prod * p["pct_tasa_estadistica"]
        antid = cif_prod * p["pct_antidumping"]
        subtotal_i = derechos + tasa + antid
        p.update(cif_producto=cif_prod, derechos=derechos, tasa_estadistica=tasa,
                  antidumping=antid, subtotal_der_tasas=subtotal_i)
        subtotal_der += derechos
        subtotal_tasa += tasa
        subtotal_antid += antid
        subtotal_dertasas += subtotal_i

    base_iva_total = cif_declarado + subtotal_dertasas  # B61

    # ---- Impuestos por producto (IVA, IVA Adicional, Ganancias, IIBB) ----
    subtotal_iva = subtotal_iva_ad = subtotal_gcias = subtotal_iibb = subtotal_impk = 0.0
    for p in productos:
        base_iva_prod = p["cif_producto"] + p["subtotal_der_tasas"]  # B65
        iva = base_iva_prod * p["pct_iva"]
        iva_ad = base_iva_prod * p["pct_iva_adicional"]
        gcias = base_iva_prod * p["pct_ganancias"]
        iibb = base_iva_prod * p["pct_iibb"]
        subk = iva + iva_ad + gcias + iibb
        p.update(base_iva_prod=base_iva_prod, iva=iva, iva_adicional=iva_ad,
                  ganancias_percep=gcias, iibb_percep=iibb, subtotal_impuestos=subk)
        subtotal_iva += iva
        subtotal_iva_ad += iva_ad
        subtotal_gcias += gcias
        subtotal_iibb += iibb
        subtotal_impk += subk

    total_tributos_vep_usd = subtotal_dertasas + subtotal_impk + arancel_sim  # B82
    total_tributos_vep_ars = total_tributos_vep_usd * tc_trib                # C82

    # ---- Subtotal operativos y prorrateo (por FOB o por peso) ----
    subtotal_oper_usd = sum(g["usd"] for g in gastos)
    subtotal_oper_ars = sum(g["ars"] for g in gastos)
    prorr_fob_usd = sum(g["usd"] for g in gastos if g["prorrateo"] == "FOB")
    prorr_peso_usd = sum(g["usd"] for g in gastos if g["prorrateo"] == "PESO")

    # ---- IVA de terceros (crédito fiscal sobre gastos operativos) ----
    subtotal_ivater_usd = subtotal_ivater_ars = 0.0
    for g in gastos:
        if g["pct_iva"] > 0:
            if g["iva_incluido"] == "SI":
                if g["moneda"] == "USD":
                    iva_usd = g["monto"] - g["usd"]
                else:
                    iva_usd = (g["monto"] - g["ars"]) / tc_oper if tc_oper > 0 else 0.0
                if g["moneda"] == "ARS":
                    iva_ars = g["monto"] - g["ars"]
                else:
                    iva_ars = (g["monto"] - g["usd"]) * tc_oper
            else:
                iva_usd = g["usd"] * g["pct_iva"]
                iva_ars = g["ars"] * g["pct_iva"]
        else:
            iva_usd = 0.0
            iva_ars = 0.0
        g["iva_terceros_usd"] = iva_usd
        g["iva_terceros_ars"] = iva_ars
        subtotal_ivater_usd += iva_usd
        subtotal_ivater_ars += iva_ars

    # ---- Totales de desembolso ----
    total_c_iva_usd = fob_total_sum + costo_financiero + total_tributos_vep_usd + subtotal_oper_usd + subtotal_ivater_usd  # B126
    total_c_iva_ars = fob_mas_financiero * tc_oper + total_tributos_vep_ars + subtotal_oper_ars + subtotal_ivater_ars      # C126
    total_s_iva_usd = total_c_iva_usd - subtotal_iva - subtotal_iva_ad - subtotal_gcias - subtotal_iibb - subtotal_ivater_usd  # B127
    total_s_iva_ars = (
        total_c_iva_ars - subtotal_iva * tc_trib - subtotal_iva_ad * tc_trib
        - subtotal_gcias * tc_trib - subtotal_iibb * tc_trib - subtotal_ivater_ars
    )  # C127

    incidencia_fob = ((total_tributos_vep_usd + subtotal_oper_usd + subtotal_ivater_usd) / fob_total_sum) if fob_total_sum > 0 else 0.0

    # ---- Costeo unitario por producto + simulación de venta ----
    total_costeo_check = 0.0
    venta_total_usd = 0.0
    for p in productos:
        share_fob = (p["fob_total"] / fob_total_sum) if fob_total_sum > 0 else 0.0
        share_peso = (p["peso_vol"] / peso_total_sum) if peso_total_sum > 0 else share_fob
        cantidad = p["cantidad"]
        if cantidad > 0:
            costo_fob_unit = p["fob_unit"] + costo_financiero * share_fob / cantidad
            imp_gas_prorr_unit = (
                p["subtotal_der_tasas"] + p["subtotal_impuestos"]
                + arancel_sim * share_fob + prorr_fob_usd * share_fob
                + prorr_peso_usd * share_peso + subtotal_ivater_usd * share_fob
            ) / cantidad
            costo_civa_unit = costo_fob_unit + imp_gas_prorr_unit
            recuperable_unit = (
                p["iva"] + p["iva_adicional"] + p["ganancias_percep"] + p["iibb_percep"]
                + subtotal_ivater_usd * share_fob
            ) / cantidad
            costo_sviva_unit = costo_civa_unit - recuperable_unit
        else:
            costo_fob_unit = imp_gas_prorr_unit = costo_civa_unit = costo_sviva_unit = 0.0

        p.update(
            costo_fob_unit=costo_fob_unit,
            imp_gas_prorr_unit=imp_gas_prorr_unit,
            costo_civa_unit=costo_civa_unit,
            costo_sviva_unit=costo_sviva_unit,
            costo_civa_unit_ars=costo_civa_unit * tc_oper,
            costo_sviva_unit_ars=costo_sviva_unit * tc_oper,
        )
        total_costeo_check += costo_civa_unit * cantidad

        pv_sugerido = costo_civa_unit * (1 + p["margen_pct"])
        pv_final = p["pv_final_usd"]
        pv_mercado = p["pv_mercado_usd"]
        margen_real = ((pv_final - costo_civa_unit) / costo_civa_unit) if (costo_civa_unit > 0 and pv_final > 0) else 0.0
        ganancia_unit = (pv_final - costo_civa_unit) if pv_final > 0 else 0.0
        pv_final_ars = pv_final * tc_venta if tc_venta else 0.0
        # Costeo!H42 = IF(E42>0,E42-B42,0) — RESTA de dos valores en ARS que
        # usan tasas distintas (E42=PV Final a TC Venta, B42=Costo a TC
        # Operativos), no la ganancia en USD multiplicada por una sola tasa
        # (eso mezclaría mal las dos tasas apenas TC Venta ≠ TC Operativos).
        ganancia_unit_ars = (pv_final_ars - costo_civa_unit * tc_oper) if (pv_final > 0 and tc_venta > 0) else 0.0
        p.update(
            pv_sugerido_usd=pv_sugerido,
            margen_real_pct=margen_real,
            ganancia_unit_usd=ganancia_unit,
            ganancia_total_usd=ganancia_unit * cantidad,
            pv_sugerido_ars=pv_sugerido * tc_oper if tc_oper else 0.0,
            pv_final_ars=pv_final_ars,
            ganancia_unit_ars=ganancia_unit_ars,
            ganancia_total_ars=ganancia_unit_ars * cantidad,
            # Informativo (referencia de precio de mercado/competencia): no
            # entra en ningún cálculo de costo ni de margen, mismo rol que
            # tiene en el Excel (Costeo!F42 = IF(B$40>0,F23*B$40,0)).
            pv_mercado_ars=pv_mercado * tc_venta if tc_venta else 0.0,
        )
        venta_total_usd += pv_final * cantidad

    check_diferencia = total_c_iva_usd - total_costeo_check
    venta_total_ars = venta_total_usd * tc_oper if tc_oper else 0.0  # sin cambios, ya es igual al Excel
    ganancia_bruta_usd = (venta_total_usd - total_c_iva_usd) if venta_total_usd > 0 else 0.0
    # Costeo!C71 = IF(C70>0,C70-C69,0) — resta directo de dos totales ARS ya
    # bien calculados. Antes multiplicaba la ganancia en USD por TC
    # Operativos, pero total_c_iva_ars mezcla TC Tributos (tributos) y TC
    # Operativos (el resto) — esa mezcla se pierde al reconstruir desde el
    # lado USD, y da mal apenas TC Tributos ≠ TC Operativos.
    ganancia_bruta_ars = (venta_total_ars - total_c_iva_ars) if venta_total_ars > 0 else 0.0
    iva_a_recuperar = subtotal_iva + subtotal_iva_ad + subtotal_ivater_usd
    percep_a_recuperar = subtotal_gcias + subtotal_iibb
    rentabilidad_pct = (ganancia_bruta_usd / total_c_iva_usd) if (total_c_iva_usd > 0 and venta_total_usd > 0) else 0.0

    return {
        "cab": {
            "tc_trib": tc_trib, "tc_oper": tc_oper, "tc_venta": tc_venta,
            "cf_pct": cf_pct, "seguro_pct": seguro_pct, "arancel_sim": arancel_sim,
        },
        "productos": productos,
        "gastos": gastos,
        "fob_total_sum": fob_total_sum,
        "fob_decl_total_sum": fob_decl_total_sum,
        "peso_total_sum": peso_total_sum,
        "costo_financiero": costo_financiero,
        "fob_mas_financiero": fob_mas_financiero,
        "tarifa_flete": tarifa_flete,
        "pct_certificacion": pct_certificacion,
        "tarifa_flete_certificado": tarifa_flete_certificado,
        "flete_declarado": flete_declarado,
        "seguro_declarado": seguro_declarado,
        "fob_declarado": fob_declarado,
        "cif_declarado": cif_declarado,
        "subtotal_der": subtotal_der,
        "subtotal_tasa": subtotal_tasa,
        "subtotal_antid": subtotal_antid,
        "subtotal_dertasas": subtotal_dertasas,
        "base_iva_total": base_iva_total,
        "subtotal_iva": subtotal_iva,
        "subtotal_iva_ad": subtotal_iva_ad,
        "subtotal_gcias": subtotal_gcias,
        "subtotal_iibb": subtotal_iibb,
        "subtotal_impk": subtotal_impk,
        "total_tributos_vep_usd": total_tributos_vep_usd,
        "total_tributos_vep_ars": total_tributos_vep_ars,
        "subtotal_oper_usd": subtotal_oper_usd,
        "subtotal_oper_ars": subtotal_oper_ars,
        "prorr_fob_usd": prorr_fob_usd,
        "prorr_peso_usd": prorr_peso_usd,
        "subtotal_ivater_usd": subtotal_ivater_usd,
        "subtotal_ivater_ars": subtotal_ivater_ars,
        "total_c_iva_usd": total_c_iva_usd,
        "total_c_iva_ars": total_c_iva_ars,
        "total_s_iva_usd": total_s_iva_usd,
        "total_s_iva_ars": total_s_iva_ars,
        "incidencia_fob": incidencia_fob,
        "check_diferencia": check_diferencia,
        "venta_total_usd": venta_total_usd,
        "venta_total_ars": venta_total_ars,
        "ganancia_bruta_usd": ganancia_bruta_usd,
        "ganancia_bruta_ars": ganancia_bruta_ars,
        "iva_a_recuperar": iva_a_recuperar,
        "percep_a_recuperar": percep_a_recuperar,
        "rentabilidad_pct": rentabilidad_pct,
    }
