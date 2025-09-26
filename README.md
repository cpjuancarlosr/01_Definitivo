# Sistema contable en Google Sheets · Scripts mejorados

Este repositorio contiene una versión mejorada del código Apps Script para operar el
sistema contable multiempresa descrito por la Jefatura Contable (CFDI 3.3 / 4.0).

## Contenido

- `apps_script/Code.gs` – Lógica principal de carga y parseo de CFDI.
- `apps_script/Picker.html` – Interfaz para subir CFDI desde el equipo mediante Google Picker.

## Novedades principales

- Soporte simultáneo para CFDI 3.3 y 4.0, detectando automáticamente el namespace del
  comprobante y de los complementos.
- Inserción por lotes y validaciones para evitar duplicados por UUID antes de escribir
en las hojas.
- Parseo completo de conceptos y complementos de pago 1.0 / 2.0, alimentando las hojas
  **Conceptos** y **Pagos** con el detalle requerido para CxC/CxP.
- Registro básico en la hoja **Bitácora** para cada importación o error.
- Toasts informativos con resumen de resultados y tiempos de procesamiento.

## Requisitos previos

1. Crear las hojas de cálculo con encabezados idénticos a los definidos en la
   especificación (CFDI, Conceptos, Pagos, Bitácora, etc.).
2. Definir las carpetas de Drive en la hoja **Setup** (celdas B13–B19) y otorgar
   permisos a la cuenta del script.
3. Activar la API de Google Picker y, si se desea, completar `developerKey` y
   `client_id` en `Picker.html` para omitir los diálogos de autorización.

## Uso

1. Abrir la hoja y ejecutar `onOpen` (o recargar) para que aparezca el menú
   **Jefatura Contable**.
2. Subir XML con **Cargar XML desde equipo…** o procesar los existentes con
   **Procesar carpeta del periodo**.
3. Revisar los resultados en las hojas **CFDI**, **Conceptos**, **Pagos** y la
   **Bitácora**.

> Las funciones de pólizas, conciliación y exportación se mantienen como _placeholders_
> para integrarse con los módulos existentes.
