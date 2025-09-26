# Sistema Contable en Google Sheets (Versión Ligera)

Este repositorio contiene el código Apps Script para una plantilla contable simplificada en Google Sheets, diseñada para agilidad y eficiencia. El sistema se enfoca en las operaciones esenciales: carga de CFDI, generación de pólizas, conciliación bancaria y reportes clave.

## Contenido

-   `apps_script/Code.gs` – Lógica principal para el menú, carga y parseo de CFDI XML.
-   `apps_script/Picker.html` – Interfaz para subir archivos CFDI desde el equipo.

## Características Principales

-   **Estructura Simplificada:** Opera sobre 5 hojas principales: `Setup`, `CFDI`, `Pólizas`, `Bancos` y `Reporte`.
-   **Carga Rápida de CFDI:** Permite subir archivos XML desde el equipo, los cuales son parseados y registrados en la hoja `CFDI`.
-   **Flujo de Pólizas Integrado:** Genera asientos contables en estado "Borrador" directamente en la hoja `Pólizas`, listos para ser emitidos.
-   **Enfoque en lo Esencial:** El parseo de XML extrae los datos clave del encabezado y los totales de impuestos, sin desglosar cada concepto para mantener la agilidad.
-   **Menú Intuitivo:** Ofrece un menú simple en Google Sheets para las acciones más comunes:
    -   Cargar XML
    -   Generar Previa de Pólizas
    -   Emitir Pólizas
    -   Importar Estado de Cuenta
    -   Realizar Conciliación
    -   Refrescar Reportes

## Uso

1.  **Configuración:** Abre la hoja de `Setup` y define el RFC activo y otros parámetros básicos.
2.  **Cargar XML:** Usa el menú `Jefatura Contable > Cargar XML…` para seleccionar y subir los archivos CFDI del periodo.
3.  **Generar Pólizas:** Selecciona los CFDI y usa `Generar Previa` para crear los asientos contables en la hoja `Pólizas`.
4.  **Emitir:** Cambia el estado de las pólizas de "Borrador" a "Emitida".
5.  **Conciliar y Reportar:** Importa tus estados de cuenta y utiliza las funciones de conciliación y el `Reporte` para analizar la información.