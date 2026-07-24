# Resultados de osmolalidad — gráficos Excel

Archivos en `osmolality_results/`:

| Archivo | Descripción |
|---------|-------------|
| `osmolality_results.xlsx` | Libro Excel serie DO (datos + gráfico nativo + imagen) |
| `osmolality_chart.png` | Gráfico estilo Excel — productos DO1–DO8 |
| `generate_osmolality_chart.py` | Script para regenerar la serie DO |
| `la_osmolality_results.xlsx` | Libro Excel serie LA (datos + gráfico nativo + imagen) |
| `la_osmolality_chart.png` | Gráfico estilo Excel — productos LA1–LA12 |
| `generate_la_osmolality_chart.py` | Script para regenerar la serie LA |

## Cómo leer el gráfico

- **Eje X:** Producto
- **Eje Y:** Osmolalidad (mOsm/kg), divisiones cada 25 unidades
- **Banda azul:** Rango de aceptación
- **Círculos naranjas / verdes:** Lote 1 / Lote 2
- Productos sin especificación: solo se muestran las mediciones
- Serie DO: eje Y 200–620; DO8 sin especificación
- Serie LA: eje Y 150–625; LA1/LA5/LA12 sin especificación; LA1 sin Lote 2
- Serie LA: debajo de cada producto se indica el Laboratorio (A–J)

## Regenerar

```bash
pip install openpyxl matplotlib numpy
python osmolality_results/generate_osmolality_chart.py
python osmolality_results/generate_la_osmolality_chart.py
```
