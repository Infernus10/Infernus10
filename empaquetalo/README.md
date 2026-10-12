# Empaquétalo · Avance 2

Recopilación de datos y elección del paradigma para el clasificador del medio de envío (dron, moto o carro/camión).

## Entregables

- `entregables/Avance2_Empaquetalo.docx`: documento del avance.
- `entregables/Avance2_Empaquetalo.pptx`: presentación.

## Datos

| Archivo | Contenido |
|---|---|
| `data/envios_crudo.csv` | 5,262 registros con errores de captura simulados |
| `data/envios_limpio.csv` | 5,109 envíos limpios con características derivadas |
| `data/procesado/` | División 80/20 (normal y escalada), escalador, bitácora y resultados de viabilidad |

## Reproducir

```bash
pip install -r requirements.txt
python src/generar_dataset.py      # dataset sintético (semilla 42)
python src/preprocesar.py          # limpieza, características y división
python src/prueba_viabilidad.py    # sistema base vs. 4 modelos + figuras
```

Para regenerar el Word y la presentación (`npm install` primero):

```bash
node src/documentos/crear_word.js
PPTX_SKILL_DIR=<ruta con scripts/apply_theme.js> node src/documentos/crear_presentacion.js
```
