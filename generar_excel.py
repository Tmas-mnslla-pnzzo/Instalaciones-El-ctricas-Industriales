"""
generar_excel.py
Usa tus módulos existentes (calculos_electricos.py, selector.py)
para leer datos_proyecto.json + datos_iluminacion.json
y generar el Excel de cálculos automáticamente.

Ejecutar: python generar_excel.py
"""
import json, math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from calculos_electricos import CargasElectricas, FactorPotencia, Conductores
from selector import SelectorComercial

# ─── CARGA DE DATOS ──────────────────────────────────────────────────────────

with open("datos_proyecto.json", encoding="utf-8") as f:
    PROYECTO = json.load(f)

with open("datos_iluminacion.json", encoding="utf-8") as f:
    ILUMINACION = json.load(f)["circuitos"]

with open("valores_comerciales.json", encoding="utf-8") as f:
    COMERCIALES = json.load(f)

INFO    = PROYECTO["informacion_general"]
CCMS    = PROYECTO["ccms"]
CONDUCTORES = PROYECTO["conductores"]
SELECTOR = SelectorComercial("valores_comerciales.json")

# ─── ESTILOS ─────────────────────────────────────────────────────────────────

def borde():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

C = {
    "titulo":    "1F3864",
    "encabezado":"2E75B6",
    "celeste":   "D6E4F0",
    "gris":      "D9D9D9",
    "verde":     "E2EFDA",
    "amarillo":  "FFF2CC",
}

def h1(ws, row, col, texto, span=1):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+span-1)
    c = ws.cell(row=row, column=col, value=texto)
    c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    c.fill      = PatternFill("solid", fgColor=C["titulo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = borde()

def h2(ws, row, col, texto):
    c = ws.cell(row=row, column=col, value=texto)
    c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    c.fill      = PatternFill("solid", fgColor=C["encabezado"])
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = borde()

def cel(ws, row, col, valor, fmt=None, bg=None, bold=False, izq=False):
    c = ws.cell(row=row, column=col, value=valor)
    c.font      = Font(name="Arial", size=9, bold=bold)
    c.alignment = Alignment(horizontal="left" if izq else "center", vertical="center")
    c.border    = borde()
    if fmt:  c.number_format = fmt
    if bg:   c.fill = PatternFill("solid", fgColor=bg)
    return c

# ─── CÁLCULOS ────────────────────────────────────────────────────────────────

def calcular_ccms():
    resultados = []
    V = INFO["tension_red_V"]
    fp_obj = INFO["cos_phi_objetivo"]

    for ccm in CCMS:
        fs = ccm["factor_simultaneidad"]
        motores_calc = []
        P_neto = Q_neto = 0.0

        for m in ccm["motores"]:
            Pa  = m["potencia_activa_kW"]
            Qr  = m["potencia_reactiva_kVAr"]
            Sa  = m["potencia_aparente_kVA"]
            Ib  = m["corriente_A"]
            motores_calc.append({**m, "Pa": Pa, "Qr": Qr, "Sa": Sa, "Ib": Ib})
            P_neto += Pa
            Q_neto += Qr

        S_neto     = math.sqrt(P_neto**2 + Q_neto**2)
        cos_neto   = P_neto / S_neto if S_neto else 0
        I_neto     = (S_neto * 1000) / (math.sqrt(3) * V)

        P_cor = P_neto * fs
        Q_cor = Q_neto * fs
        S_cor = math.sqrt(P_cor**2 + Q_cor**2)
        cos_cor   = P_cor / S_cor if S_cor else 0
        I_cor     = (S_cor * 1000) / (math.sqrt(3) * V)

        # Compensación FP usando tu módulo
        Qc = max(0, FactorPotencia.banco_capacitores(P_cor, cos_neto, fp_obj))

        resultados.append({
            "id": ccm["id"], "tablero_origen": ccm["tablero_origen"], "fs": fs,
            "motores": motores_calc,
            "P_neto": P_neto, "Q_neto": Q_neto, "S_neto": S_neto,
            "cos_neto": cos_neto, "I_neto": I_neto,
            "P_cor": P_cor,  "Q_cor": Q_cor,  "S_cor": S_cor,
            "cos_cor": cos_cor, "I_cor": I_cor,
            "Qc": Qc,
        })
    return resultados


def calcular_iluminacion():
    resultados = []
    for lu in ILUMINACION:
        P_kW  = lu["potencia_W"] * lu["cantidad"] / 1000.0
        I     = (P_kW * 1000) / (lu["tension_V"] * lu["cos_phi"])
        P_cor = P_kW * lu["factor_utilizacion"] * lu["factor_simultaneidad"]
        resultados.append({
            "tablero": lu["tablero"],
            "marca":   lu["marca"],
            "V":       lu["tension_V"],
            "cos_phi": lu["cos_phi"],
            "cantidad":lu["cantidad"],
            "fu":      lu["factor_utilizacion"],
            "fs":      lu["factor_simultaneidad"],
            "P_kW":    P_kW,
            "I":       I,
            "P_cor":   P_cor,
        })
    return resultados


def calcular_conductores(ccms_res, ilu_res):
    # Índice de corrientes por destino
    corrientes = {}
    for ccm in ccms_res:
        corrientes[ccm["id"]] = ccm["I_cor"]
        for m in ccm["motores"]:
            corrientes[m["id"]] = m["Ib"]
    for lu in ilu_res:
        corrientes[lu["tablero"]] = lu["I"]

    V      = INFO["tension_red_V"]
    icc_max= INFO["potencia_cc_mva"] * 1e6 / (math.sqrt(3) * V)
    t_cc   = INFO["tiempo_cc_s"]
    k      = INFO["k_aislante"]
    dU_max = INFO["delta_u_max_V"]

    resultados = []
    for cond in CONDUCTORES:
        Ib = corrientes.get(cond["destino"], 50.0)

        # Dimensionamiento iterativo usando tu módulo
        res = Conductores.dimensionamiento_iterativo(
            ib=Ib,
            longitud_km=cond["longitud_m"] / 1000.0,
            cos_phi=0.85,
            icc_max=icc_max,
            tiempo_cc=t_cc,
            delta_u_max=dU_max,
            k_aislante=k,
            seccion_maxima=240.0,
        )
        sec   = res["seccion_mm2"] or 240.0
        n_par = res["cantidad_de_cables_en_paralelo"]

        # Sección comercial usando tu SelectorComercial
        try:
            sec_com = SELECTOR.obtener_inmediato_superior("secciones_cables_mm2", sec)
        except Exception:
            sec_com = sec

        # Corriente admisible con factores
        iz_base = Conductores.CATALOGO_IZ_BASE.get(sec_com, 0)
        iz_total = iz_base * cond["fa"] * cond["ft"] * n_par

        # Caída de tensión
        rho = 0.0225
        xc  = 0.08
        sin_phi = math.sin(math.acos(0.85))
        rc  = (rho * 1000 / sec_com) / n_par
        xc_eq = xc / n_par
        dU  = math.sqrt(3) * Ib * (cond["longitud_m"]/1000) * (rc*0.85 + xc_eq*sin_phi)
        dU_pct = (dU / V) * 100

        # Termomagnética comercial
        try:
            In_term = SELECTOR.obtener_inmediato_superior("termomagneticas_A", Ib * 1.25)
        except Exception:
            In_term = "-"

        resultados.append({
            **cond,
            "Ib": Ib, "seccion": sec_com, "n_paralelo": n_par,
            "Iz": iz_total, "dU_V": dU, "dU_pct": dU_pct,
            "In_term": In_term,
            "cumple": dU_pct <= 5.0,
        })
    return resultados


def calcular_totales(ccms_res, ilu_res):
    P_mot = sum(c["P_neto"] for c in ccms_res)
    Q_mot = sum(c["Q_neto"] for c in ccms_res)
    P_ilu = sum(l["P_kW"]  for l in ilu_res)
    P_tot = P_mot + P_ilu
    Q_tot = Q_mot
    S_tot = math.sqrt(P_tot**2 + Q_tot**2)
    cos_t = P_tot / S_tot if S_tot else 0
    S_res = S_tot * INFO["factor_reserva"]
    trafo = SELECTOR.obtener_inmediato_superior("transformadores_kVA", S_res)
    return {
        "P_mot": P_mot, "Q_mot": Q_mot,
        "P_ilu": P_ilu,
        "P_tot": P_tot, "Q_tot": Q_tot, "S_tot": S_tot,
        "cos_t": cos_t, "S_res": S_res, "trafo": trafo,
        "cant_luminarias": sum(l["cantidad"] for l in ILUMINACION),
    }

# ─── ESCRITURA EXCEL ─────────────────────────────────────────────────────────

def escribir_excel(ccms_res, ilu_res, cond_res, tot):
    wb = Workbook()

    # ═══════════════════════════════════════════════════
    # HOJA 1 — CALCULOS
    # ═══════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Calculos"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([14,12,8,10,10,12,12,12,8,10,10,10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1
    h1(ws, r, 1, f"CÁLCULOS ELÉCTRICOS — {INFO['nombre_proyecto']}", span=12)
    ws.row_dimensions[r].height = 22
    r += 1
    cel(ws, r, 1, f"Autor: {INFO['autor']}  |  Tensión: {INFO['tension_red_V']} V  |  fp objetivo: {INFO['cos_phi_objetivo']}", bg=C["gris"])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    ws.cell(r, 1).alignment = Alignment(horizontal="left")
    r += 2

    # ── Cargas Motrices ──────────────────────────────────────────────────────
    h1(ws, r, 1, "CARGAS MOTRICES", span=12); r += 1
    for txt in ["Tablero","Motor","Tensión [V]","Corriente [A]","cos φ",
                "P Activa [kW]","P Reactiva [kVAr]","P Aparente [kVA]",
                "Zcc [Ω]","Dist CC [km]","F. Utiliz.","F. Simult."]:
        h2(ws, r, ["Tablero","Motor","Tensión [V]","Corriente [A]","cos φ",
                   "P Activa [kW]","P Reactiva [kVAr]","P Aparente [kVA]",
                   "Zcc [Ω]","Dist CC [km]","F. Utiliz.","F. Simult."].index(txt)+1, txt)
    ws.row_dimensions[r].height = 28; r += 1

    for ccm in ccms_res:
        first = True
        for m in ccm["motores"]:
            bg_ccm = C["celeste"] if first else None
            cel(ws, r, 1,  ccm["id"] if first else "",   bg=bg_ccm, bold=first)
            cel(ws, r, 2,  m["id"])
            cel(ws, r, 3,  INFO["tension_red_V"],        fmt="0.0")
            cel(ws, r, 4,  m["Ib"],                      fmt="0.000")
            cel(ws, r, 5,  m["cos_phi"],                 fmt="0.000")
            cel(ws, r, 6,  m["Pa"],                      fmt="0.000")
            cel(ws, r, 7,  m["Qr"],                      fmt="0.000")
            cel(ws, r, 8,  m["Sa"],                      fmt="0.000")
            cel(ws, r, 9,  m["zcc"],                     fmt="0.00")
            cel(ws, r, 10, m["distancia_cc_km"],         fmt="0.00")
            cel(ws, r, 11, m["factor_utilizacion"],      fmt="0.000")
            cel(ws, r, 12, ccm["fs"] if first else "",   fmt="0.000" if first else None, bg=bg_ccm)
            first = False; r += 1

        # Subtotal neto
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cel(ws, r, 1, f"Total neto {ccm['id']}:", bg=C["amarillo"], bold=True, izq=True)
        cel(ws, r, 4, ccm["I_neto"],  fmt="0.000", bg=C["amarillo"])
        cel(ws, r, 5, ccm["cos_neto"],fmt="0.000", bg=C["amarillo"])
        cel(ws, r, 6, ccm["P_neto"],  fmt="0.000", bg=C["amarillo"])
        cel(ws, r, 7, ccm["Q_neto"],  fmt="0.000", bg=C["amarillo"])
        cel(ws, r, 8, ccm["S_neto"],  fmt="0.000", bg=C["amarillo"])
        for j in [3,9,10,11,12]: cel(ws, r, j, "", bg=C["amarillo"])
        r += 1

        # Subtotal corregido
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cel(ws, r, 1, f"Total corregido {ccm['id']}:", bg=C["verde"], bold=True, izq=True)
        cel(ws, r, 4, ccm["I_cor"],   fmt="0.000", bg=C["verde"])
        cel(ws, r, 5, ccm["cos_cor"], fmt="0.000", bg=C["verde"])
        cel(ws, r, 6, ccm["P_cor"],   fmt="0.000", bg=C["verde"])
        cel(ws, r, 7, ccm["Q_cor"],   fmt="0.000", bg=C["verde"])
        cel(ws, r, 8, ccm["S_cor"],   fmt="0.000", bg=C["verde"])
        for j in [3,9,10,11,12]: cel(ws, r, j, "", bg=C["verde"])
        r += 1

    r += 1

    # ── Iluminación ──────────────────────────────────────────────────────────
    h1(ws, r, 1, "ILUMINACIÓN", span=12); r += 1
    for j, txt in enumerate(["Tablero","Luminaria","Tensión [V]","Corriente [A]","cos φ",
                              "P Activa [kW]","P Reactiva [kVAr]","P Aparente [kVA]",
                              "Cantidad","F. Utiliz.","F. Simult.",""], 1):
        h2(ws, r, j, txt)
    ws.row_dimensions[r].height = 25; r += 1

    for lu in ilu_res:
        cel(ws, r, 1,  lu["tablero"])
        cel(ws, r, 2,  lu["marca"])
        cel(ws, r, 3,  lu["V"],       fmt="0.0")
        cel(ws, r, 4,  lu["I"],       fmt="0.000")
        cel(ws, r, 5,  lu["cos_phi"], fmt="0.0")
        cel(ws, r, 6,  lu["P_kW"],    fmt="0.000")
        cel(ws, r, 7,  0.0,           fmt="0.0")
        cel(ws, r, 8,  lu["P_kW"],    fmt="0.000")
        cel(ws, r, 9,  lu["cantidad"])
        cel(ws, r, 10, lu["fu"],      fmt="0.0")
        cel(ws, r, 11, lu["fs"],      fmt="0.0")
        cel(ws, r, 12, "")
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    cel(ws, r, 1, "Total neto iluminación:", bg=C["amarillo"], bold=True, izq=True)
    cel(ws, r, 6, sum(l["P_kW"]    for l in ilu_res), fmt="0.000", bg=C["amarillo"])
    cel(ws, r, 9, sum(l["cantidad"] for l in ilu_res), bg=C["amarillo"])
    for j in [3,4,5,7,8,10,11,12]: cel(ws, r, j, "", bg=C["amarillo"])
    r += 2

    # ── Totales y transformador ───────────────────────────────────────────────
    h1(ws, r, 1, "RESUMEN DE POTENCIAS Y TRANSFORMADOR", span=12); r += 1
    for j, txt in enumerate(["Concepto","cos φ","P Activa [kW]","P Reactiva [kVAr]","P Aparente [kVA]"], 1):
        h2(ws, r, j, txt)
    r += 1

    filas = [
        ("TOTAL NETO [kVA]:",             tot["cos_t"], tot["P_tot"], tot["Q_tot"], tot["S_tot"]),
        ("TOTAL CON RESERVA (30%) [kVA]:",None,          None,         None,         tot["S_res"]),
        ("TRANSFORMADOR COMERCIAL [kVA]:", None,          None,         None,         tot["trafo"]),
    ]
    for fila in filas:
        cel(ws, r, 1, fila[0], bg=C["celeste"], bold=True, izq=True)
        for j, v in enumerate(fila[1:], 2):
            cel(ws, r, j, v if v is not None else "-",
                fmt="0.000" if isinstance(v, float) else None, bg=C["celeste"])
        r += 1

    r += 2

    # ── Capacitores ──────────────────────────────────────────────────────────
    h1(ws, r, 1, "BANCO DE CAPACITORES POR CCM", span=8); r += 1
    for j, txt in enumerate(["CCM","Correción FP [kVAr]","Pot. Corto [MVA]",
                              "In [A]","I Prot [A]","I Cresta [A]","K Conexión",""], 1):
        h2(ws, r, j, txt)
    r += 1

    V = INFO["tension_red_V"]
    for ccm in ccms_res:
        Qc    = ccm["Qc"]
        In    = Qc * 1000 / (math.sqrt(3) * V) if Qc > 0 else 0
        Iprot = In * 1.43
        Icres = math.sqrt(2) * In * 1.41  # aproximación cresta
        cel(ws, r, 1, ccm["id"],                         bold=True)
        cel(ws, r, 2, round(Qc, 4),                      fmt="0.0000")
        cel(ws, r, 3, INFO["potencia_cc_mva"])
        cel(ws, r, 4, round(In, 3),                      fmt="0.000")
        cel(ws, r, 5, round(Iprot, 3),                   fmt="0.000")
        cel(ws, r, 6, round(abs(Icres), 3),              fmt="0.000")
        cel(ws, r, 7, 3)
        cel(ws, r, 8, "")
        r += 1

    r += 2

    # ── Conductores ──────────────────────────────────────────────────────────
    h1(ws, r, 1, "CONDUCTORES — DIMENSIONAMIENTO", span=12); r += 1
    for j, txt in enumerate(["Origen","Destino","Tipo","Sección [mm²]","N° Par.",
                              "Iz total [A]","Ib [A]","ΔU [V]","ΔU [%]",
                              "In term. [A]","Long. [m]","OK"], 1):
        h2(ws, r, j, txt)
    ws.row_dimensions[r].height = 25; r += 1

    for cond in cond_res:
        bg = C["verde"] if cond["cumple"] else C["amarillo"]
        cel(ws, r, 1,  cond["origen"])
        cel(ws, r, 2,  cond["destino"])
        cel(ws, r, 3,  cond["tipo"])
        cel(ws, r, 4,  cond["seccion"],   fmt="0.0",   bg=bg)
        cel(ws, r, 5,  cond["n_paralelo"])
        cel(ws, r, 6,  round(cond["Iz"], 1), fmt="0.0")
        cel(ws, r, 7,  round(cond["Ib"], 2), fmt="0.00")
        cel(ws, r, 8,  round(cond["dU_V"],  3), fmt="0.000", bg=bg)
        cel(ws, r, 9,  round(cond["dU_pct"],2), fmt="0.00",  bg=bg)
        cel(ws, r, 10, cond["In_term"])
        cel(ws, r, 11, cond["longitud_m"])
        cel(ws, r, 12, "✓" if cond["cumple"] else "⚠", bg=bg)
        r += 1

    # ═══════════════════════════════════════════════════
    # HOJA 2 — COMERCIALES
    # ═══════════════════════════════════════════════════
    ws2 = wb.create_sheet("Comerciales")
    ws2.sheet_view.showGridLines = False
    h1(ws2, 1, 1, "CATÁLOGOS COMERCIALES", span=8)
    encab2 = ["Transf. [kVA]","Condensadores","Conductores [A]",
              "Conductores [mm²]","R [Ω/km]","X [Ω/km]","Termomagnéticas [A]","Fusibles [A]"]
    for j, txt in enumerate(encab2, 1):
        h2(ws2, 2, j, txt)
        ws2.column_dimensions[get_column_letter(j)].width = 16

    datos_com = [
        (9,   "", 9.6,   1.0,   18.10, 0.13,  6,    2),
        (15,  "", 13.0,  1.5,   12.10, 0.12,  10,   4),
        (30,  "", 18.0,  2.5,    7.28, 0.11,  16,   6),
        (45,  "", 24.0,  4.0,    4.55, 0.10,  20,   10),
        (75,  "", 31.0,  6.0,    3.03, 0.10,  25,   16),
        (113, "", 43.0,  10.0,   1.82, 0.09,  32,   20),
        (150, "", 59.0,  16.0,   1.14, 0.08,  40,   25),
        (225, "", 77.0,  25.0,   0.73, 0.08,  50,   35),
        (300, "", 96.0,  35.0,   0.52, 0.08,  63,   50),
        (500, "", 116.0, 50.0,   0.36, 0.08,  80,   63),
        (750, "", 148.0, 70.0,   0.26, 0.07,  100,  ""),
        (1000,"", 180.0, 95.0,   0.19, 0.07,  125,  ""),
        (1500,"", 210.0, 120.0,  0.15, 0.07,  160,  ""),
        (2000,"", 245.0, 150.0,  0.12, 0.07,  200,  ""),
        (2500,"", 285.0, 185.0,  0.10, 0.07,  250,  ""),
        (3750,"", 345.0, 240.0,  0.075,0.06,  315,  ""),
        (5000,"", 400.0, 300.0,  0.06, 0.06,  400,  ""),
    ]
    for i, fila in enumerate(datos_com, 3):
        for j, v in enumerate(fila, 1):
            c = ws2.cell(row=i, column=j, value=v)
            c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(horizontal="center")
            c.border = borde()
            if i % 2 == 0:
                c.fill = PatternFill("solid", fgColor=C["gris"])

    # ═══════════════════════════════════════════════════
    # HOJA 3 — CORTOCIRCUITO
    # ═══════════════════════════════════════════════════
    ws3 = wb.create_sheet("CortoCircuito")
    ws3.sheet_view.showGridLines = False
    h1(ws3, 1, 1, "CÁLCULO DE CORTOCIRCUITO", span=6)
    for j, txt in enumerate(["Tablero","Zcc total [Ω]","Icc máx [kA]",
                              "Icc mín [kA]","Secc. mín CC [mm²]","Estado"], 1):
        h2(ws3, 2, j, txt)
        ws3.column_dimensions[get_column_letter(j)].width = 18

    V_red = INFO["tension_red_V"]
    Pcc   = INFO["potencia_cc_mva"] * 1e6
    Zb    = V_red**2 / Pcc
    nodos = ["T","TG","TS4","CCM1","CCM2","CCM3","CCM4"]
    for i, nodo in enumerate(nodos, 3):
        Zcc     = Zb * (i - 2) * 0.5
        Icc_max = V_red / (math.sqrt(3) * Zcc) / 1000
        Icc_min = Icc_max * 0.85
        sec_min = (Icc_max * 1000 * math.sqrt(INFO["tiempo_cc_s"])) / INFO["k_aislante"]
        ok      = sec_min < 240
        cel(ws3, i, 1, nodo,              bold=True)
        cel(ws3, i, 2, round(Zcc, 4),    fmt="0.0000")
        cel(ws3, i, 3, round(Icc_max,3), fmt="0.000")
        cel(ws3, i, 4, round(Icc_min,3), fmt="0.000")
        cel(ws3, i, 5, round(sec_min,2), fmt="0.00")
        cel(ws3, i, 6, "OK" if ok else "Verificar", bg=C["verde"] if ok else C["amarillo"])

    # ═══════════════════════════════════════════════════
    # HOJA 4 — RESUMEN
    # ═══════════════════════════════════════════════════
    ws4 = wb.create_sheet("Resumen")
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions["A"].width = 38
    ws4.column_dimensions["B"].width = 22
    h1(ws4, 1, 1, "RESUMEN EJECUTIVO", span=2)
    filas_res = [
        ("Proyecto",                    INFO["nombre_proyecto"]),
        ("Autor",                       INFO["autor"]),
        ("Tensión de red",              f"{INFO['tension_red_V']} V"),
        ("P activa total neta",         f"{tot['P_tot']:.2f} kW"),
        ("P reactiva total neta",       f"{tot['Q_tot']:.2f} kVAr"),
        ("P aparente total neta",       f"{tot['S_tot']:.2f} kVA"),
        ("Factor de potencia global",   f"{tot['cos_t']:.4f}"),
        ("P aparente con reserva 30%",  f"{tot['S_res']:.2f} kVA"),
        ("Transformador seleccionado",  f"{tot['trafo']} kVA"),
        ("Total luminarias",            f"{tot['cant_luminarias']} unid."),
    ]
    for i, (k, v) in enumerate(filas_res, 2):
        bg = C["celeste"] if i % 2 == 0 else C["gris"]
        cel(ws4, i, 1, k, bg=bg, bold=True, izq=True)
        cel(ws4, i, 2, v, bg=bg)

    # ─── Guardar ─────────────────────────────────────────────────────────────
    salida = "calculo_electrico_ponzoni.xlsx"
    wb.save(salida)
    print(f"✅  Excel generado: {salida}")
    return salida


# ─── MAIN ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("⚙️  Calculando CCMs...")
    ccms_res = calcular_ccms()

    print("💡 Calculando iluminación...")
    ilu_res  = calcular_iluminacion()

    print("🔌 Dimensionando conductores...")
    cond_res = calcular_conductores(ccms_res, ilu_res)

    print("📊 Totalizando potencias...")
    tot      = calcular_totales(ccms_res, ilu_res)

    print(f"\n   P total neta  : {tot['P_tot']:.2f} kW")
    print(f"   S total neta  : {tot['S_tot']:.2f} kVA")
    print(f"   cos φ global  : {tot['cos_t']:.4f}")
    print(f"   S con reserva : {tot['S_res']:.2f} kVA")
    print(f"   Transformador : {tot['trafo']} kVA\n")

    escribir_excel(ccms_res, ilu_res, cond_res, tot)
