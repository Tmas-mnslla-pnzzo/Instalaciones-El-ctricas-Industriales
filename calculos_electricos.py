"""
Librería: calculos_electricos.py
Motor de cálculo para Instalaciones Eléctricas Industriales (Normativa AEA).
"""
import math

class CargasElectricas:
    """Módulo 1: Estimación de Cargas y Potencia Instalada"""
    @staticmethod
    def demanda_motor(potencia_nominal: float, rendimiento: float, factor_utilizacion: float = 1.0) -> float:
        return (potencia_nominal * factor_utilizacion) / rendimiento

    @staticmethod
    def dpms(potencias_demandadas: list, factor_simultaneidad: float = 1.0) -> float:
        return sum(potencias_demandadas) * factor_simultaneidad


class FactorPotencia:
    """Módulo 2: Compensación del Factor de Potencia"""
    @staticmethod
    def banco_capacitores(potencia_activa: float, cos_phi_actual: float, cos_phi_deseado: float = 0.95) -> float:
        phi_actual = math.acos(cos_phi_actual)
        phi_deseado = math.acos(cos_phi_deseado)
        return potencia_activa * (math.tan(phi_actual) - math.tan(phi_deseado))


class Iluminacion:
    """Módulo 3: Cálculo de Iluminación (Cavidades Zonales)"""
    @staticmethod
    def numero_luminarias(em: float, superficie: float, factor_util: float, factor_mant: float, flujo_lampara: float) -> float:
        return (em * superficie) / (factor_util * factor_mant * flujo_lampara)


class Conductores:
    """Módulo 4: Dimensionamiento de Conductores (Iterativo y Paralelo)"""
    CATALOGO_IZ_BASE = {
        1.5: 13.5, 2.5: 18.0, 4.0: 24.0, 6.0: 31.0, 10.0: 42.0,
        16.0: 56.0, 25.0: 73.0, 35.0: 89.0, 50.0: 108.0, 70.0: 136.0,
        95.0: 164.0, 120.0: 188.0, 150.0: 216.0, 185.0: 245.0, 240.0: 286.0
    }

    @staticmethod
    def _factor_agrupamiento_paralelo(n_cables: int) -> float:
        if n_cables == 1: return 1.0
        elif n_cables == 2: return 0.80
        elif n_cables == 3: return 0.70
        elif n_cables == 4: return 0.65
        else: return 0.60 

    @classmethod
    def dimensionamiento_iterativo(
        cls, ib: float, longitud_km: float, cos_phi: float, 
        icc_max: float, tiempo_cc: float, delta_u_max: float, 
        seccion_maxima: float = 35.0, k_aislante: float = 115.0, 
        trifasica: bool = True, catalogo_iz: dict = None
    ) -> dict:
        catalogo = catalogo_iz if catalogo_iz else cls.CATALOGO_IZ_BASE
        secciones_disponibles = sorted([s for s in catalogo.keys() if s <= seccion_maxima])
        
        if not secciones_disponibles:
            raise ValueError("La sección máxima exigida es menor a la mínima del catálogo.")

        n_paralelo = 1
        limite_paralelos = 10 
        k_sis = math.sqrt(3) if trifasica else 2.0
        sin_phi = math.sin(math.acos(cos_phi))
        rho_cu = 0.0225 
        xc_std = 0.08   

        while n_paralelo <= limite_paralelos:
            factor_agrupamiento = cls._factor_agrupamiento_paralelo(n_paralelo)
            
            for seccion in secciones_disponibles:
                iz_un_cable = catalogo[seccion]
                iz_total = iz_un_cable * n_paralelo * factor_agrupamiento
                
                if iz_total < ib:
                    continue 
                
                rc_equiv = (rho_cu * 1000 / seccion) / n_paralelo 
                xc_equiv = xc_std / n_paralelo
                delta_u_calculada = k_sis * ib * longitud_km * ((rc_equiv * cos_phi) + (xc_equiv * sin_phi))
                
                if delta_u_calculada > delta_u_max:
                    continue 

                seccion_total = seccion * n_paralelo
                seccion_min_cc = (icc_max * math.sqrt(tiempo_cc)) / k_aislante
                
                if seccion_total < seccion_min_cc:
                    continue 

                detalle = (f"Iz total: {round(iz_total, 1)}A | "
                           f"Caída: {round(delta_u_calculada, 2)}V | "
                           f"Scc_min: {round(seccion_min_cc, 1)}mm2")
                
                return {
                    'seccion_mm2': seccion,
                    'cantidad_de_cables_en_paralelo': n_paralelo,
                    'cumple_todas_las_condiciones': True,
                    'detalle': detalle
                }
            n_paralelo += 1

        return {
            'seccion_mm2': None,
            'cantidad_de_cables_en_paralelo': n_paralelo,
            'cumple_todas_las_condiciones': False,
            'detalle': "Imposible dimensionar. Requiere más de 10 conductores en paralelo."
        }


class Cortocircuito:
    """Módulo 5: Cálculo de Cortocircuito"""
    @staticmethod
    def impedancia_total(r_total: float, x_total: float) -> float:
        return math.sqrt(r_total**2 + x_total**2)

    @staticmethod
    def corriente_cc_max_trifasica(tension_linea: float, zcc: float) -> float:
        return tension_linea / (math.sqrt(3) * zcc)


class Protecciones:
    """Módulo 6: Dimensionamiento de Protecciones"""
    @staticmethod
    def validar_sobrecarga(ib: float, iz: float, in_interruptor: float, i2_interruptor: float = None) -> bool:
        condicion_1 = ib <= in_interruptor <= iz
        if i2_interruptor is None:
            i2_interruptor = 1.45 * in_interruptor
        condicion_2 = i2_interruptor <= (1.45 * iz)
        return condicion_1 and condicion_2

    @staticmethod
    def validar_cortocircuito(icu: float, icc_max: float, imag: float, icc_min: float) -> bool:
        return (icu >= icc_max) and (imag <= icc_min)


class Tableros:
    """Módulo 7: Diseño de Barras en Tableros"""
    @staticmethod
    def fuerza_electrodinamica_fh(is_pico: float, longitud_soportes: float, distancia_fases: float) -> float:
        return 0.2 * (is_pico**2) * (longitud_soportes / distancia_fases)


class PuestaATierra:
    """Módulo 8: Puesta a Tierra (PAT)"""
    @staticmethod
    def resistencia_jabalina_vertical(rho: float, l: float, d: float) -> float:
        return (rho / (2 * math.pi * l)) * math.log((2 * l) / d)


class GradoElectrificacion:
    """Módulo 9: Grados de Electrificación (AEA)"""
    @staticmethod
    def clasificar_local(superficie: float) -> dict:
        if superficie <= 60.0: return {"Grado": "Mínimo", "Circuitos_Minimos": 2}
        elif superficie <= 130.0: return {"Grado": "Medio", "Circuitos_Minimos": 3}
        elif superficie <= 200.0: return {"Grado": "Elevado", "Circuitos_Minimos": 5}
        else: return {"Grado": "Superior", "Circuitos_Minimos": 6}


class Canalizaciones:
    """Módulo 10: Dimensionamiento de Canalizaciones"""
    @staticmethod
    def dimensionar_cano(seccion_cable: float, cantidad_cables: float, ocupacion_maxima: float = 0.35) -> float:
        area_cano_necesaria = (seccion_cable * cantidad_cables) / ocupacion_maxima
        return math.sqrt((4 * area_cano_necesaria) / math.pi)


class AcometidaEPRE:
    """Módulo 11: Selección de Acometida (EPRE)"""
    @staticmethod
    def definir_acometida(potencia_kw: float) -> dict:
        if potencia_kw <= 10.0: return {"Tarifa": "T1", "Bajada": "Caño H°G° 1 1/4\""}
        elif potencia_kw <= 50.0: return {"Tarifa": "T2", "Bajada": "Caño H°G° 2\" o superior"}
        else: return {"Tarifa": "T3", "Bajada": "Acometida subterránea"}


class MemoriaTecnica:
    """Módulo 12: Generador de Memoria Técnica"""
    def __init__(self, nombre_proyecto: str, autor: str):
        self.nombre_proyecto = nombre_proyecto
        self.autor = autor
        self.datos_recolectados = {}
        self.materiales = []

    def registrar_parametro(self, seccion: str, clave: str, valor: str):
        if seccion not in self.datos_recolectados:
            self.datos_recolectados[seccion] = {}
        self.datos_recolectados[seccion][clave] = valor

    def exportar_md(self, ruta_archivo: str = "Memoria_Tecnica.md"):
        with open(ruta_archivo, 'w', encoding='utf-8') as f:
            f.write(f"# Memoria Técnica: {self.nombre_proyecto}\n")
            f.write(f"**Autor:** {self.autor}\n\n---\n\n")
            for seccion, datos in self.datos_recolectados.items():
                f.write(f"### {seccion}\n")
                for clave, valor in datos.items():
                    f.write(f"* **{clave}:** {valor}\n")
                f.write("\n")
        print(f"✅ Memoria Técnica exportada con éxito en: {ruta_archivo}")