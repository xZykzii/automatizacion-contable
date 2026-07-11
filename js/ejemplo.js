// Ejercicio de ejemplo: Empresa "CIM" (desarrollo del certamen de contabilidad).
// Es el mismo libro diario del archivo "Ejercicio para trabajo (1).xlsx".
const EJEMPLO_CIM = [
  {
    num: 1, fecha: "01.02.2012", glosa: "Asiento de apertura",
    lineas: [
      { cuenta: "Caja", debe: 5000000, haber: 0 },
      { cuenta: "Mercaderías", debe: 10000000, haber: 0 },
      { cuenta: "Vehículo", debe: 6500000, haber: 0 },
      { cuenta: "Letras por Pagar", debe: 0, haber: 650000 },
      { cuenta: "Capital", debe: 0, haber: 20850000 },
    ],
  },
  {
    num: 2, fecha: "02.02.2012", glosa: "Apertura de cuenta corriente (90% del efectivo)",
    lineas: [
      { cuenta: "Banco", debe: 4500000, haber: 0 },
      { cuenta: "Caja", debe: 0, haber: 4500000 },
    ],
  },
  {
    num: 3, fecha: "03.02.2012", glosa: "Contabiliza venta de mercaderías según factura N° 7",
    lineas: [
      { cuenta: "Caja", debe: 1950000, haber: 0 },
      { cuenta: "Clientes", debe: 4550000, haber: 0 },
      { cuenta: "Ingresos por Ventas", debe: 0, haber: 5462185 },
      { cuenta: "IVA Débito Fiscal", debe: 0, haber: 1037815 },
    ],
  },
  {
    num: 4, fecha: "03.02.2012", glosa: "Contabiliza el costo de la mercadería vendida (40% de los materiales)",
    lineas: [
      { cuenta: "Costo de Ventas", debe: 4000000, haber: 0 },
      { cuenta: "Mercaderías", debe: 0, haber: 4000000 },
    ],
  },
  {
    num: 5, fecha: "04.02.2012", glosa: "Depósito en cuenta corriente de la venta del día 03.02",
    lineas: [
      { cuenta: "Banco", debe: 1950000, haber: 0 },
      { cuenta: "Caja", debe: 0, haber: 1950000 },
    ],
  },
  {
    num: 6, fecha: "10.02.2012", glosa: "Se cancela luz ($27.000) y agua ($15.500) con cheques N° 1 y 2",
    lineas: [
      { cuenta: "Gastos Generales", debe: 42500, haber: 0 },
      { cuenta: "Banco", debe: 0, haber: 42500 },
    ],
  },
  {
    num: 7, fecha: "16.02.2012", glosa: "Cancela arriendo del mes y 2 meses anticipados",
    lineas: [
      { cuenta: "Arriendo", debe: 90000, haber: 0 },
      { cuenta: "Arriendo Anticipado", debe: 180000, haber: 0 },
      { cuenta: "Caja", debe: 0, haber: 90000 },
      { cuenta: "Banco", debe: 0, haber: 180000 },
    ],
  },
  {
    num: 8, fecha: "28.02.2012", glosa: "Cancela remuneraciones del mes con cheque",
    lineas: [
      { cuenta: "Remuneraciones", debe: 1000000, haber: 0 },
      { cuenta: "Banco", debe: 0, haber: 1000000 },
    ],
  },
  {
    num: 9, fecha: "30.04.2012", glosa: "Compra repuesto y mantención del vehículo, factura N° 3047",
    lineas: [
      { cuenta: "Gastos Generales", debe: 254454, haber: 0 },
      { cuenta: "IVA Crédito Fiscal", debe: 48346, haber: 0 },
      { cuenta: "Banco", debe: 0, haber: 302800 },
    ],
  },
  {
    num: 10, fecha: "30.06.2012", glosa: "Contabiliza cargo bancario por intereses de cuenta corriente",
    lineas: [
      { cuenta: "Gastos Financieros", debe: 32000, haber: 0 },
      { cuenta: "Banco", debe: 0, haber: 32000 },
    ],
  },
  {
    num: 11, fecha: "28.07.2012", glosa: "Compra herramientas: 50% con cheque y 50% al crédito simple",
    lineas: [
      { cuenta: "Máquinas y Herramientas", debe: 2100840, haber: 0 },
      { cuenta: "IVA Crédito Fiscal", debe: 399159.6, haber: 0 },
      { cuenta: "Banco", debe: 0, haber: 1250000 },
      { cuenta: "Acreedores", debe: 0, haber: 1250000 },
    ],
  },
  {
    num: 12, fecha: "31.12.2012", glosa: "Recibe pago del cliente de la venta del 03.02",
    lineas: [
      { cuenta: "Banco", debe: 4550000, haber: 0 },
      { cuenta: "Clientes", debe: 0, haber: 4550000 },
    ],
  },
  {
    num: 13, fecha: "31.12.2012", glosa: "Cierra cuentas de IVA y determina impuesto a pagar",
    lineas: [
      { cuenta: "IVA Débito Fiscal", debe: 1037815, haber: 0 },
      { cuenta: "IVA Crédito Fiscal", debe: 0, haber: 447506 },
      { cuenta: "Impuesto a Pagar", debe: 0, haber: 590309 },
    ],
  },
];
