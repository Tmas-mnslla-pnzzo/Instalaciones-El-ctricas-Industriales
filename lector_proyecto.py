import json
import math

class LectorProyecto:
    def __init__(self, ruta_json: str):
        with open(ruta_json, "r", encoding="utf-8") as f:
            self.datos = json.load(f)
            
        # Creamos un diccionario rápido para encontrar tableros por su ID
        self.tableros = {t["id"]: t for t in self.datos["tableros"]}
        self.cargas = self.datos["cargas_motrices"]
        self.info = self.datos["informacion_general"]

    def calcular_longitud_cable(self, carga: dict, margen_seguridad: float = 1.10) -> float:
        """
        Calcula la longitud del cable desde el tablero hasta la carga.
        Usa geometría ortogonal (Manhattan) multiplicada por un margen de seguridad 
        para considerar caídas y subidas de bandeja.
        """
        id_tablero = carga["tablero_origen"]
        if id_tablero not in self.tableros:
            raise ValueError(f"El tablero {id_tablero} no existe en la base de datos.")
            
        tablero = self.tableros[id_tablero]
        
        # Distancia en "L" (Manhattan)
        distancia_x = abs(carga["ubicacion_x"] - tablero["ubicacion_x"])
        distancia_y = abs(carga["ubicacion_y"] - tablero["ubicacion_y"])
        
        longitud_base = distancia_x + distancia_y
        
        # Retornamos en kilómetros (necesario para la fórmula de caída de tensión)
        longitud_total_km = (longitud_base * margen_seguridad) / 1000.0
        return longitud_total_km