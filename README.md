# App Streamlit - Validación de Órdenes de Trabajo

Aplicación para revisar los campos esenciales definidos en las Órdenes de Trabajo (OT) y generar un reporte ejecutivo de faltantes y valores inválidos.

## Campos validados

### Anverso

- Horómetro.
- Motivo de detención del equipo.
- En cada fila utilizada del bloque **Información del trabajo**:
  - Descripción del síntoma.
  - Código trabajo.
  - Código síntoma.
  - Código causa.
- Descripción de actividades.

### Reverso

- Firma del jefe de turno: nombre y RUT.
- Firma del técnico responsable: nombre y RUT.

No se genera ninguna observación por los demás campos del anverso o reverso.

## Reglas especiales

- Los códigos causa **6.6** y **7.1** se consideran inválidos porque corresponden a la categoría **Otros**.
- La descripción de actividades se considera completa cuando existe información en al menos una de las líneas del bloque.
- En el resumen por OT se incluye la columna **Campos con observación**, inmediatamente después de **Estado**, con el detalle de los campos faltantes o inválidos.

## Campos críticos

- Código trabajo.
- Código síntoma.
- Código causa.
- Firma jefe turno (nombre + RUT).
- Firma técnico responsable (nombre + RUT).

El horómetro, el motivo de detención, la descripción del síntoma y la descripción de actividades se clasifican como campos estándar.

La aplicación muestra KPIs, un gráfico de cumplimiento por campo, estado de las OT, resumen por campo y detalle de hallazgos. También permite descargar reportes en Excel y PDF.

## Ejecución local recomendada

Usar Python 3.12:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Despliegue en Streamlit Community Cloud

1. Crear o actualizar el repositorio en GitHub.
2. Subir `app.py`, `requirements.txt` y este `README.md`.
3. En Streamlit Community Cloud seleccionar el repositorio.
4. Definir `app.py` como archivo principal.
5. Seleccionar Python 3.12 en las opciones avanzadas y desplegar.

## Coordenadas del formato

Las coordenadas se encuentran en `validate_work_order()`:

- Horómetro: `G13`.
- Motivo de detención: `AB25`.
- Filas de información del trabajo: desde la fila 42, cada cuatro filas.
- Código trabajo: columna `W`.
- Descripción del síntoma: columna `Z`.
- Código síntoma: columna `AO`.
- Código causa: columna `BH`.
- Descripción de actividades: bloque `B98:BZ124`.
- Firma jefe de turno: `C238` y `C244`.
- Firma técnico responsable: `BD239` y `BD243`.
