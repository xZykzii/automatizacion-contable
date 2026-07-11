# Automatización Contable — Diario → Mayor → Balance de 8 Columnas

Trabajo de Gestión Contable: automatización del ciclo contable de la empresa "CIM".

**Página en vivo:** https://xzykzii.github.io/automatizacion-contable/

## ¿Qué hace?

1. **Libro Diario**: se sube un archivo Excel con el libro diario (sirve el formato del
   ejercicio original o la plantilla tabular descargable). Los asientos quedan editables
   en pantalla: al cambiar cualquier monto, cuenta o asiento, todo se recalcula al instante.
2. **Libro Mayor**: se genera automáticamente — centraliza los movimientos por cuenta
   (cuentas T), suma Debe/Haber y determina el saldo deudor (SD) o acreedor (SA).
3. **Balance de 8 Columnas**: se construye desde los saldos del Mayor, clasifica cada
   cuenta en Activos / Pasivos / Pérdidas / Ganancias y calcula la utilidad o pérdida
   del ejercicio con Totales Iguales.
4. **Extra**: Estado de Resultados automático.

## Exportaciones a Excel

- Libro Diario
- Libro Mayor (solo) o Diario + Mayor en un mismo archivo
- Balance de 8 Columnas (solo) o **todo junto**: Diario + Mayor + Balance + Estado de Resultados

## Validaciones y alertas

- Partida doble por asiento (Debe = Haber) con alerta de asientos descuadrados.
- Cuadratura global del Diario y del Balance ("BALANCE CUADRADO / DESCUADRADO").
- Detección automática de la clasificación de cuentas, con alerta para confirmar
  las cuentas no reconocidas.
- Tolerancia de $1 para diferencias por redondeo (ej. IVA).

## Tecnología

HTML + CSS + JavaScript puro con [SheetJS (xlsx-js-style)](https://github.com/gitbrent/xlsx-js-style)
para leer y generar los Excel. Todo corre en el navegador: **no se envían datos a ningún servidor**.

## Uso local

Basta abrir `index.html` en el navegador, o servir la carpeta:

```
python -m http.server 8000
```
