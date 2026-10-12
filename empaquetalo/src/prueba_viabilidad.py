"""Prueba de viabilidad del Avance 2.

Responde la pregunta de la revisión con evidencia: ¿los datos permiten
resolver el problema? Compara el sistema base de reglas simples contra los
cuatro modelos propuestos con hiperparámetros por defecto (el ajuste fino se
hace en las semanas 9 a 11). También genera las figuras del reporte.

Uso: python src/prueba_viabilidad.py
"""
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix, f1_score, classification_report
from sklearn.model_selection import StratifiedKFold, cross_val_score
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.tree import DecisionTreeClassifier

RAIZ = Path(__file__).resolve().parents[1]
PROC = RAIZ / "data" / "procesado"
FIG = RAIZ / "figuras"
SEMILLA = 42
ORDEN = ["dron", "moto", "carro/camion"]
COLOR = {"dron": "#3567B5", "moto": "#C9702E", "carro/camion": "#7A5195"}
TINTA, TINTA_2, REJILLA = "#1F2A44", "#5B6475", "#E3E6EC"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": REJILLA,
    "axes.labelcolor": TINTA_2, "xtick.color": TINTA_2, "ytick.color": TINTA_2,
    "axes.spines.top": False, "axes.spines.right": False, "axes.titleweight": "bold",
    "axes.titlecolor": TINTA, "axes.titlesize": 11, "axes.titlelocation": "left",
})


def sistema_base(df):
    """Reglas simples como las que usa hoy la paquetería: sólo peso, distancia y clima."""
    def regla(r):
        if r.peso_kg <= 2.5 and r.distancia_km <= 15 and r.clima_despejado + r.clima_nublado == 1:
            return "dron"
        if r.peso_kg <= 20 and r.distancia_km <= 40 and r.clima_tormenta == 0:
            return "moto"
        return "carro/camion"
    return np.array([regla(r) for r in df.itertuples(index=False)])


def figuras_datos(limpio):
    FIG.mkdir(exist_ok=True)
    # Distribución de clases.
    conteo = limpio["medio"].value_counts().reindex(ORDEN)
    fig, ax = plt.subplots(figsize=(6, 2.6))
    ax.barh(range(3), conteo.values, color=[COLOR[m] for m in ORDEN], height=0.55)
    for i, v in enumerate(conteo.values):
        ax.text(v + 30, i, f"{v:,}  ({v / conteo.sum():.0%})", va="center", color=TINTA, fontsize=9)
    ax.set_yticks(range(3), ["Dron", "Moto", "Carro/camión"])
    ax.invert_yaxis()
    ax.set_xlim(0, conteo.max() * 1.3)
    ax.xaxis.grid(True, color=REJILLA); ax.set_axisbelow(True)
    ax.set_title("Envíos por medio de transporte (n = {:,})".format(conteo.sum()))
    fig.tight_layout(); fig.savefig(FIG / "distribucion_clases.png", dpi=200); plt.close(fig)

    # Peso vs distancia, coloreado por medio.
    fig, ax = plt.subplots(figsize=(6.4, 4.2))
    for m in reversed(ORDEN):
        d = limpio[limpio["medio"] == m]
        ax.scatter(d["distancia_km"], d["peso_kg"], s=9, alpha=0.55, color=COLOR[m],
                   edgecolors="white", linewidths=0.3, label={"carro/camion": "Carro/camión"}.get(m, m.capitalize()))
    ax.set_xscale("log"); ax.set_yscale("log")
    ax.axhline(2.5, color=TINTA_2, lw=0.8, ls="--"); ax.text(130, 2.75, "límite dron 2.5 kg", fontsize=8, color=TINTA_2, ha="right")
    ax.axhline(20, color=TINTA_2, lw=0.8, ls="--"); ax.text(130, 22, "límite moto 20 kg", fontsize=8, color=TINTA_2, ha="right")
    ax.set_xlabel("Distancia (km, escala log)"); ax.set_ylabel("Peso (kg, escala log)")
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles[::-1], labels[::-1], frameon=False, loc="upper center", fontsize=9, markerscale=2.2,
              ncol=3, bbox_to_anchor=(0.5, -0.17))
    for h in ax.get_legend().legend_handles:
        h.set_alpha(1)
    ax.set_title("Peso contra distancia: las fronteras no son cortes simples")
    fig.tight_layout(); fig.savefig(FIG / "peso_vs_distancia.png", dpi=200); plt.close(fig)

    # Clima contra medio (proporción).
    tabla = pd.crosstab(limpio["clima"], limpio["medio"], normalize="index").reindex(
        ["despejado", "nublado", "lluvia", "tormenta"])[ORDEN]
    fig, ax = plt.subplots(figsize=(6.4, 2.8))
    izquierda = np.zeros(len(tabla))
    for m in ORDEN:
        ax.barh(range(len(tabla)), tabla[m], left=izquierda, color=COLOR[m], height=0.6,
                edgecolor="white", linewidth=1.5, label={"carro/camion": "Carro/camión"}.get(m, m.capitalize()))
        for i, v in enumerate(tabla[m]):
            if v >= 0.07:
                ax.text(izquierda[i] + v / 2, i, f"{v:.0%}", ha="center", va="center", color="white", fontsize=8.5)
        izquierda += tabla[m].values
    ax.set_yticks(range(len(tabla)), [c.capitalize() for c in tabla.index]); ax.invert_yaxis()
    ax.set_xlim(0, 1); ax.set_xticks([0, .25, .5, .75, 1], ["0%", "25%", "50%", "75%", "100%"])
    ax.legend(frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.15), fontsize=9)
    ax.set_title("Medio elegido según el clima")
    fig.tight_layout(); fig.savefig(FIG / "clima_vs_medio.png", dpi=200); plt.close(fig)


def main():
    limpio = pd.read_csv(RAIZ / "data" / "envios_limpio.csv")
    tr, te = pd.read_csv(PROC / "entrenamiento.csv"), pd.read_csv(PROC / "prueba.csv")
    tr_e, te_e = pd.read_csv(PROC / "entrenamiento_escalado.csv"), pd.read_csv(PROC / "prueba_escalado.csv")
    X_tr, y_tr, X_te, y_te = tr.drop(columns="medio"), tr["medio"], te.drop(columns="medio"), te["medio"]
    Xe_tr, Xe_te = tr_e.drop(columns="medio"), te_e.drop(columns="medio")

    figuras_datos(limpio)

    modelos = {
        "Árbol de decisión": (DecisionTreeClassifier(random_state=SEMILLA), False),
        "KNN (k=5)": (KNeighborsClassifier(), True),
        "Random Forest": (RandomForestClassifier(random_state=SEMILLA, n_jobs=-1), False),
        "Red neuronal (MLP)": (MLPClassifier(hidden_layer_sizes=(32, 16), max_iter=800, random_state=SEMILLA), True),
    }
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEMILLA)
    resultados = {}
    pred_base = sistema_base(X_te)
    resultados["Sistema base (reglas)"] = {
        "f1_cv": None, "f1_prueba": f1_score(y_te, pred_base, average="macro"),
        "f1_dron": f1_score(y_te, pred_base, labels=["dron"], average="macro")}
    predicciones = {"Sistema base (reglas)": pred_base}
    for nombre, (modelo, escalado) in modelos.items():
        A_tr, A_te = (Xe_tr, Xe_te) if escalado else (X_tr, X_te)
        f1_cv = cross_val_score(modelo, A_tr, y_tr, cv=cv, scoring="f1_macro")
        modelo.fit(A_tr, y_tr)
        pred = modelo.predict(A_te)
        predicciones[nombre] = pred
        resultados[nombre] = {"f1_cv": f1_cv.mean(), "f1_cv_std": f1_cv.std(),
                              "f1_prueba": f1_score(y_te, pred, average="macro"),
                              "f1_dron": f1_score(y_te, pred, labels=["dron"], average="macro")}
    for k, v in resultados.items():
        print(f"{k:24s} " + "  ".join(f"{a}={b:.3f}" for a, b in v.items() if b is not None))

    rf = modelos["Random Forest"][0]
    importancias = pd.Series(rf.feature_importances_, index=X_tr.columns).sort_values(ascending=False)
    print(importancias.round(3))
    print(classification_report(y_te, predicciones["Random Forest"], digits=3))
    print(classification_report(y_te, pred_base, digits=3))

    # Gráfica comparativa de F1-macro en prueba.
    nombres = list(resultados)
    vals = [resultados[n]["f1_prueba"] for n in nombres]
    fig, ax = plt.subplots(figsize=(6.4, 3.0))
    colores = ["#9AA3B2"] + ["#2E4A7D"] * 4
    ax.barh(range(len(nombres)), vals, color=colores, height=0.55)
    for i, v in enumerate(vals):
        ax.text(v + 0.01, i, f"{v:.2f}", va="center", color=TINTA, fontsize=9)
    ax.axvline(0.85, color="#C9702E", lw=1.2, ls="--"); ax.text(0.853, len(nombres) - 0.55, "meta 0.85", color="#C9702E", fontsize=8.5)
    ax.set_yticks(range(len(nombres)), nombres); ax.invert_yaxis(); ax.set_xlim(0, 1.05)
    ax.xaxis.grid(True, color=REJILLA); ax.set_axisbelow(True)
    ax.set_title("F1-macro en prueba (modelos sin ajustar)")
    fig.tight_layout(); fig.savefig(FIG / "comparacion_f1.png", dpi=200); plt.close(fig)

    # Matrices de confusión: sistema base vs Random Forest.
    fig, axes = plt.subplots(1, 2, figsize=(7.4, 3.3))
    etiquetas = ["Dron", "Moto", "Carro"]
    for ax, nombre in zip(axes, ["Sistema base (reglas)", "Random Forest"]):
        cm = confusion_matrix(y_te, predicciones[nombre], labels=ORDEN)
        ax.imshow(cm, cmap="Blues", vmin=0, vmax=cm.max())
        for i in range(3):
            for j in range(3):
                ax.text(j, i, cm[i, j], ha="center", va="center", fontsize=10,
                        color="white" if cm[i, j] > cm.max() * 0.55 else TINTA)
        ax.set_xticks(range(3), etiquetas); ax.set_yticks(range(3), etiquetas)
        ax.set_xlabel("Predicho"); ax.set_ylabel("Real")
        ax.set_title(nombre, fontsize=10)
        for s in ax.spines.values(): s.set_visible(False)
    fig.tight_layout(); fig.savefig(FIG / "matrices_confusion.png", dpi=200); plt.close(fig)

    salida = {"resultados": resultados, "importancias_rf": importancias.round(4).to_dict(),
              "confusion_base": confusion_matrix(y_te, pred_base, labels=ORDEN).tolist(),
              "confusion_rf": confusion_matrix(y_te, predicciones["Random Forest"], labels=ORDEN).tolist()}
    with open(RAIZ / "data" / "procesado" / "resultados_viabilidad.json", "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=2, default=float)


if __name__ == "__main__":
    main()
