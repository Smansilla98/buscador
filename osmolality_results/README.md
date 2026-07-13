# Resultados de osmolalidad — gráfico Excel

Archivos en `osmolality_results/`:

| Archivo | Descripción |
|---------|-------------|
| `osmolality_results.xlsx` | Libro Excel con datos, estado Pass/Fail, gráfico nativo y hoja con imagen |
| `osmolality_chart.png` | Gráfico estilo Excel (rangos sombreados + marcadores circulares) |
| `generate_osmolality_chart.py` | Script para regenerar el Excel y el PNG |

## Cómo leer el gráfico

- **Eje X:** productos DO1–DO8
- **Eje Y:** osmolalidad (mOsm/kg)
- **Banda azul:** Rango de aceptación de cada producto
- **Círculos naranjas:** Lote 1
- **Círculos verdes:** Lote 2
- **DO8:** sin especificación (solo se muestran las mediciones)
- **Eje Y:** rango 200–620 mOsm/kg, divisiones cada 25 unidades

## Regenerar

```bash
pip install openpyxl matplotlib numpy
python osmolality_results/generate_osmolality_chart.py
```
