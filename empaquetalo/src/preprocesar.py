"""Preprocesamiento inicial del dataset de Empaquétalo.

Lee data/envios_crudo.csv, limpia los errores de captura, crea las variables
derivadas, separa entrenamiento/prueba y escala las variables numéricas.
Imprime una bitácora de cada paso para documentarla en el avance.

Uso: python src/preprocesar.py
"""
import json
import unicodedata
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

RAIZ = Path(__file__).resolve().parents[1]
SEMILLA = 42
CLIMAS = ["despejado", "nublado", "lluvia", "tormenta"]
SINONIMOS = {"soleado": "despejado", "lluvioso": "lluvia", "tormenta electrica": "tormenta"}
NUMERICAS = ["peso_kg", "dim_mayor_cm", "dim_media_cm", "dim_menor_cm", "volumen_l",
             "distancia_km", "viento_kmh"]


def normalizar_clima(valor):
    if pd.isna(valor):
        return np.nan
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode()
    texto = " ".join(texto.lower().split())
    return SINONIMOS.get(texto, texto)


def limpiar(df, bitacora):
    n0 = len(df)
    bitacora["filas_crudas"] = n0

    # 1. Duplicados: mismo id_envio registrado dos veces.
    df = df.drop_duplicates(subset="id_envio", keep="first")
    bitacora["duplicados_eliminados"] = n0 - len(df)

    # 2. Texto de clima inconsistente.
    antes = df["clima"].nunique(dropna=True)
    df = df.assign(clima=df["clima"].map(normalizar_clima))
    bitacora["variantes_clima_antes"] = int(antes)
    bitacora["variantes_clima_despues"] = int(df["clima"].nunique(dropna=True))

    # 3. Registros sin peso o clima: no se pueden imputar sin inventar el caso.
    sin_dato = df["peso_kg"].isna() | df["clima"].isna()
    bitacora["eliminados_sin_peso_o_clima"] = int(sin_dato.sum())
    df = df[~sin_dato]

    # 4. Pesos imposibles (<= 0).
    invalido = df["peso_kg"] <= 0
    bitacora["eliminados_peso_invalido"] = int(invalido.sum())
    df = df[~invalido]

    # 5. Largo capturado en mm: ninguna caja del negocio supera 300 cm ni es
    #    más de 6 veces más larga que su otra dimensión mayor.
    en_mm = (df["largo_cm"] > 300) | (df["largo_cm"] > 6 * df[["ancho_cm", "alto_cm"]].max(axis=1))
    bitacora["largo_corregido_mm_a_cm"] = int(en_mm.sum())
    df.loc[en_mm, "largo_cm"] = df.loc[en_mm, "largo_cm"] / 10

    # 6. Viento faltante: se imputa con la mediana del mismo clima.
    faltan = df["viento_kmh"].isna()
    bitacora["viento_imputado"] = int(faltan.sum())
    df["viento_kmh"] = df["viento_kmh"].fillna(df.groupby("clima")["viento_kmh"].transform("median"))

    bitacora["filas_limpias"] = len(df)
    return df.reset_index(drop=True)


def derivar(df):
    dims = np.sort(df[["largo_cm", "ancho_cm", "alto_cm"]].to_numpy(), axis=1)[:, ::-1]
    df = df.assign(dim_mayor_cm=dims[:, 0], dim_media_cm=dims[:, 1], dim_menor_cm=dims[:, 2])
    df["volumen_l"] = (df["largo_cm"] * df["ancho_cm"] * df["alto_cm"] / 1000).round(2)
    for c in CLIMAS:
        df[f"clima_{c}"] = (df["clima"] == c).astype(int)
    return df


def main():
    bitacora = {}
    crudo = pd.read_csv(RAIZ / "data" / "envios_crudo.csv")
    df = derivar(limpiar(crudo, bitacora))
    df.to_csv(RAIZ / "data" / "envios_limpio.csv", index=False)

    columnas = NUMERICAS + [f"clima_{c}" for c in CLIMAS]
    X, y = df[columnas], df["medio"]
    X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.2, stratify=y, random_state=SEMILLA)

    # El escalador se ajusta sólo con entrenamiento para no filtrar información de prueba.
    escalador = StandardScaler().fit(X_tr[NUMERICAS])
    X_tr_esc, X_te_esc = X_tr.copy(), X_te.copy()
    X_tr_esc[NUMERICAS] = escalador.transform(X_tr[NUMERICAS])
    X_te_esc[NUMERICAS] = escalador.transform(X_te[NUMERICAS])

    salida = RAIZ / "data" / "procesado"
    salida.mkdir(exist_ok=True)
    X_tr.assign(medio=y_tr).to_csv(salida / "entrenamiento.csv", index=False)
    X_te.assign(medio=y_te).to_csv(salida / "prueba.csv", index=False)
    X_tr_esc.assign(medio=y_tr).to_csv(salida / "entrenamiento_escalado.csv", index=False)
    X_te_esc.assign(medio=y_te).to_csv(salida / "prueba_escalado.csv", index=False)
    joblib.dump(escalador, salida / "escalador.joblib")

    bitacora["entrenamiento"] = len(X_tr)
    bitacora["prueba"] = len(X_te)
    bitacora["clases"] = y.value_counts().to_dict()
    with open(salida / "bitacora_preprocesamiento.json", "w", encoding="utf-8") as f:
        json.dump(bitacora, f, ensure_ascii=False, indent=2)
    print(json.dumps(bitacora, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
