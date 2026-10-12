"""Generador del dataset sintético de Empaquétalo.

Simula envíos de una paquetería con valores realistas de peso, dimensiones,
distancia y clima, y los etiqueta con el medio de envío (dron, moto o
carro/camión) según una política de costo, tiempo y seguridad definida por el
equipo. Produce un archivo "crudo" que incluye errores de captura típicos
(valores faltantes, texto inconsistente, unidades equivocadas y duplicados)
para que el preprocesamiento tenga que resolverlos.

Uso: python src/generar_dataset.py
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEMILLA = 42
N_ENVIOS = 5200
RAIZ = Path(__file__).resolve().parents[1]

# Límites físicos de cada medio (supuestos del proyecto, Avance 1).
LIMITES = {
    "dron": {"peso": 2.5, "dims": (30, 30, 20), "distancia": 15, "viento": 30,
             "climas": {"despejado", "nublado"}},
    "moto": {"peso": 20, "dims": (50, 40, 40), "distancia": 40, "viento": None,
             "climas": {"despejado", "nublado", "lluvia"}},
    "carro/camion": {"peso": 800, "dims": None, "distancia": None, "viento": None,
                     "climas": {"despejado", "nublado", "lluvia", "tormenta"}},
}
MEDIOS = list(LIMITES)


def es_factible(medio, peso, dims, distancia, clima, viento):
    """True si el medio puede hacer el envío sin violar ningún límite."""
    lim = LIMITES[medio]
    if peso > lim["peso"] or clima not in lim["climas"]:
        return False
    if lim["dims"] is not None and any(d > m for d, m in zip(sorted(dims, reverse=True), lim["dims"])):
        return False
    if lim["distancia"] is not None and distancia > lim["distancia"]:
        return False
    if lim["viento"] is not None and viento >= lim["viento"]:
        return False
    return True


def generar_paquetes(rng, n):
    tipos = rng.choice(["sobre", "pequeno", "mediano", "grande"], size=n, p=[0.22, 0.38, 0.25, 0.15])
    peso, largo, ancho, alto = (np.empty(n) for _ in range(4))
    for i, t in enumerate(tipos):
        if t == "sobre":
            peso[i] = rng.uniform(0.05, 0.6)
            largo[i], ancho[i], alto[i] = rng.uniform(18, 32), rng.uniform(12, 25), rng.uniform(0.5, 4)
        elif t == "pequeno":
            peso[i] = np.clip(rng.lognormal(np.log(1.4), 0.45), 0.2, 5)
            largo[i], ancho[i], alto[i] = rng.uniform(15, 36), rng.uniform(10, 30), rng.uniform(5, 22)
        elif t == "mediano":
            peso[i] = np.clip(rng.lognormal(np.log(8), 0.5), 2, 25)
            largo[i], ancho[i], alto[i] = rng.uniform(30, 60), rng.uniform(25, 45), rng.uniform(15, 42)
        else:
            peso[i] = np.clip(rng.lognormal(np.log(45), 0.9), 12, 600)
            largo[i], ancho[i], alto[i] = rng.uniform(50, 180), rng.uniform(40, 110), rng.uniform(30, 100)
    distancia = np.clip(rng.lognormal(np.log(9), 0.85, size=n), 0.3, 120)
    clima = rng.choice(["despejado", "nublado", "lluvia", "tormenta"], size=n, p=[0.45, 0.25, 0.20, 0.10])
    media_viento = {"despejado": 10, "nublado": 16, "lluvia": 24, "tormenta": 48}
    desv_viento = {"despejado": 6, "nublado": 8, "lluvia": 10, "tormenta": 14}
    viento = np.array([max(0.0, rng.normal(media_viento[c], desv_viento[c])) for c in clima])
    return pd.DataFrame({
        "peso_kg": peso.round(2), "largo_cm": largo.round(1), "ancho_cm": ancho.round(1),
        "alto_cm": alto.round(1), "distancia_km": distancia.round(2), "clima": clima,
        "viento_kmh": viento.round(1),
    })


def costo(medio, peso, dims, distancia, clima, viento):
    """Puntaje de costo-tiempo-riesgo de la política del equipo (menor es mejor).

    Las interacciones no lineales son intencionales: el viento castiga al dron
    de forma creciente, la lluvia castiga más a la moto cuanto más pesado y
    voluminoso es el paquete, y el carro sólo compensa a partir de cierta
    distancia o tamaño.
    """
    volumen_l = dims[0] * dims[1] * dims[2] / 1000
    if medio == "dron":
        c = 16 + 4.0 * distancia
        c += 0.08 * max(0.0, viento - 12) ** 2           # inestabilidad por viento
        c += 14 * (peso / 2.5) ** 3                       # batería y riesgo cerca del límite
        c += 6 if clima == "nublado" else 0
        return c
    if medio == "moto":
        c = 30 + 3.6 * distancia
        c += 0.45 * volumen_l + 0.6 * peso                # maniobrabilidad
        if clima == "lluvia":
            c += 12 + 1.2 * peso + 0.4 * volumen_l + 0.25 * distancia
        c += 0.25 * max(0.0, viento - 35) ** 1.5
        return c
    c = 78 + 2.2 * distancia + 0.04 * peso               # carro/camión
    c += 10 if distancia < 3 else 0                       # tráfico y estacionamiento en tramos cortos
    return c


def etiquetar(df, rng, ruido=7.0):
    etiquetas = []
    for r in df.itertuples(index=False):
        dims = (r.largo_cm, r.ancho_cm, r.alto_cm)
        mejor, mejor_costo = None, np.inf
        for m in MEDIOS:
            if not es_factible(m, r.peso_kg, dims, r.distancia_km, r.clima, r.viento_kmh):
                continue
            # El ruido simula el criterio variable de los despachadores.
            c = costo(m, r.peso_kg, dims, r.distancia_km, r.clima, r.viento_kmh) + rng.normal(0, ruido)
            if c < mejor_costo:
                mejor, mejor_costo = m, c
        etiquetas.append(mejor)
    return etiquetas


def ensuciar(df, rng):
    """Agrega errores de captura realistas para el archivo crudo."""
    df = df.copy()
    n = len(df)
    # Texto de clima inconsistente (mayúsculas, espacios, acentos, sinónimos).
    variantes = {"despejado": ["Despejado", "DESPEJADO", " despejado", "soleado"],
                 "nublado": ["Nublado", "nublado ", "NUBLADO"],
                 "lluvia": ["Lluvia", "lluvioso", " LLUVIA"],
                 "tormenta": ["Tormenta", "tormenta eléctrica"]}
    idx = rng.choice(n, size=int(0.06 * n), replace=False)
    df.loc[idx, "clima"] = [rng.choice(variantes[c]) for c in df.loc[idx, "clima"]]
    # Valores faltantes.
    df.loc[rng.choice(n, size=int(0.025 * n), replace=False), "viento_kmh"] = np.nan
    df.loc[rng.choice(n, size=int(0.008 * n), replace=False), "clima"] = np.nan
    df.loc[rng.choice(n, size=int(0.006 * n), replace=False), "peso_kg"] = np.nan
    # Dimensiones capturadas en milímetros por error.
    idx = rng.choice(n, size=int(0.006 * n), replace=False)
    df.loc[idx, "largo_cm"] = df.loc[idx, "largo_cm"] * 10
    # Pesos imposibles (cero o negativos).
    idx = rng.choice(n, size=int(0.004 * n), replace=False)
    df.loc[idx, "peso_kg"] = -df.loc[idx, "peso_kg"].abs().round(2) * rng.choice([0, 1], size=len(idx))
    # Registros duplicados.
    dup = df.sample(frac=0.012, random_state=SEMILLA)
    df = pd.concat([df, dup]).sample(frac=1, random_state=SEMILLA).reset_index(drop=True)
    return df


def main():
    rng = np.random.default_rng(SEMILLA)
    df = generar_paquetes(rng, N_ENVIOS)
    df["medio"] = etiquetar(df, rng)
    df.insert(0, "id_envio", [f"ENV-{i:05d}" for i in range(1, len(df) + 1)])
    crudo = ensuciar(df, rng)
    (RAIZ / "data").mkdir(exist_ok=True)
    crudo.to_csv(RAIZ / "data" / "envios_crudo.csv", index=False)
    print(f"Crudo: {len(crudo)} filas -> data/envios_crudo.csv")
    print(df["medio"].value_counts(normalize=True).round(3))


if __name__ == "__main__":
    main()
