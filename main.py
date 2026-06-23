from calculos_electricos import Conductores, CargasElectricas, MemoriaTecnica
from selector import SelectorComercial
from lector_proyecto import LectorProyecto

print("🔌 Iniciando cálculos automáticos de la Planta...")

# 1. Cargamos las bases de datos
catalogo = SelectorComercial("valores_comerciales.json")
planta = LectorProyecto("datos_proyecto.json")
reporte = MemoriaTecnica(planta.info["nombre_proyecto"], "Tu Nombre")

# 2. Bucle Mágico: Procesar todas las cargas automáticamente
for motor in planta.cargas:
    print(f"\n⚙️ Procesando: {motor['nombre']} ({motor['id']})")
    
    # A. Demanda Eléctrica
    potencia_dem = CargasElectricas.demanda_motor(
        potencia_nominal=motor["potencia_kW"],
        rendimiento=motor["rendimiento"],
        factor_utilizacion=motor["factor_utilizacion"]
    )
    
    # B. Cálculo de Corriente (I = P / (sqrt(3) * V * cos_phi))
    # Aproximación básica trifásica en Amperes
    voltaje = planta.info["tension_red_V"]
    corriente_ib = (potencia_dem * 1000) / (1.732 * voltaje * motor["cos_phi"])
    
    # C. Cálculo de Distancia (automático según coordenadas)
    longitud_km = planta.calcular_longitud_cable(motor)
    
    # D. Dimensionamiento Iterativo del Cable (Motor 4 de calculos_electricos.py)
    resultado_cable = Conductores.dimensionamiento_iterativo(
        ib=corriente_ib,
        longitud_km=longitud_km,
        cos_phi=motor["cos_phi"],
        icc_max=5000.0, # Cortocircuito estimado (se puede parametrizar luego)
        tiempo_cc=0.02,
        delta_u_max=20.0, # 5% de 400V
        catalogo_iz=Conductores.CATALOGO_IZ_BASE
    )
    
    # E. Ajuste Comercial del Cable
    seccion_comercial = catalogo.obtener_inmediato_superior("secciones_cables_mm2", resultado_cable['seccion_mm2'])
    
    # F. Registrar en la Memoria Técnica
    seccion_nombre = f"Circuito {motor['id']} - {motor['nombre']}"
    reporte.registrar_parametro(seccion_nombre, "Alimentación desde", motor['tablero_origen'])
    reporte.registrar_parametro(seccion_nombre, "Distancia calculada", f"{round(longitud_km * 1000, 1)} metros")
    reporte.registrar_parametro(seccion_nombre, "Corriente de Proyecto (Ib)", f"{round(corriente_ib, 2)} A")
    reporte.registrar_parametro(seccion_nombre, "Cable Comercial Asignado", f"{resultado_cable['cantidad_de_cables_en_paralelo']} x {seccion_comercial} mm2")
    
    print(f"✅ Cable asignado: {resultado_cable['cantidad_de_cables_en_paralelo']} x {seccion_comercial} mm2 a {round(longitud_km * 1000, 1)}m de distancia.")

# 3. Exportamos el reporte final
print("\n📝 Generando documentación...")
reporte.exportar_md()