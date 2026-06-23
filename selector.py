import json
from typing import Dict, List

class LimiteCatalogoError(Exception):
    """
    Excepción personalizada lanzada cuando un valor calculado 
    supera el máximo disponible en el catálogo comercial.
    """
    pass

class SelectorComercial:
    """
    Clase encargada de vincular los cálculos teóricos con los 
    valores comerciales reales disponibles en el mercado.
    """

    def __init__(self, ruta_json: str = "valores_comerciales.json"):
        """
        Inicializa el selector cargando el archivo JSON en memoria.
        """
        self.ruta_json = ruta_json
        self.catalogos: Dict[str, List[float]] = self._cargar_json()

    def _cargar_json(self) -> Dict[str, List[float]]:
        """Lee y parsea el archivo JSON con manejo de errores."""
        try:
            with open(self.ruta_json, "r", encoding="utf-8") as archivo:
                return json.load(archivo)
        except FileNotFoundError:
            raise FileNotFoundError(f"🚨 Error: No se encontró el catálogo comercial en la ruta: {self.ruta_json}")
        except json.JSONDecodeError:
            raise ValueError(f"🚨 Error: El archivo {self.ruta_json} está corrupto o no tiene formato JSON válido.")

    def obtener_inmediato_superior(self, categoria: str, valor_calculado: float) -> float:
        """
        Busca el primer valor comercial que sea mayor o igual al valor calculado.

        Args:
            categoria (str): Llave del catálogo (ej: 'secciones_cables_mm2').
            valor_calculado (float): El valor teórico obtenido por fórmula.

        Returns:
            float: El valor comercial estandarizado.
            
        Raises:
            KeyError: Si la categoría no existe en el JSON.
            LimiteCatalogoError: Si el valor supera el máximo del catálogo.
        """
        if categoria not in self.catalogos:
            opciones = ", ".join(self.catalogos.keys())
            raise KeyError(f"La categoría '{categoria}' no existe. Opciones válidas: {opciones}")

        valores_disponibles = self.catalogos[categoria]

        for valor_comercial in valores_disponibles:
            if float(valor_comercial) >= valor_calculado:
                return float(valor_comercial)

        # Si el bucle termina sin retornar, el valor es más grande que el catálogo
        maximo_disponible = valores_disponibles[-1]
        raise LimiteCatalogoError(
            f"⚠️ El valor solicitado ({valor_calculado}) para '{categoria}' supera "
            f"el máximo comercial disponible ({maximo_disponible}). "
            "Debe aplicar división de circuitos, celdas adicionales o cables en paralelo."
        )

# --- EJEMPLO DE USO ---
if __name__ == "__main__":
    # Instanciamos la clase (Carga el JSON automáticamente)
    selector = SelectorComercial("valores_comerciales.json")

    try:
        # Caso 1: Sección teórica de cable de 21.4 mm2
        seccion = selector.obtener_inmediato_superior("secciones_cables_mm2", 21.4)
        print(f"✅ Sección comercial asignada: {seccion} mm2")

        # Caso 2: Transformador para una potencia teórica de 430 kVA
        trafo = selector.obtener_inmediato_superior("transformadores_kVA", 430.0)
        print(f"✅ Transformador asignado: {trafo} kVA")

        # Caso 3: Provocar un error intencional superando el catálogo
        cable_gigante = selector.obtener_inmediato_superior("secciones_cables_mm2", 800.0)

    except LimiteCatalogoError as e:
        print(f"\n{e}")