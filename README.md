# Módulo de Control de Bloqueos - Zona Franca Santander

Este módulo es una aplicación web desarrollada con Streamlit diseñada para automatizar la estructuración, filtrado y visualización de los datos extraídos del archivo `ReportePW` de PICIZ. 

Su función principal es transformar la data cruda de los reportes de ingresos de vehículos en un panel de control interactivo con alertas visuales basadas en semaforización (rojo, amarillo, verde), permitiendo un monitoreo eficiente del vencimiento de tiempos (Días restantes) de las operaciones.

## Funcionalidades Principales
- **Carga de Archivos:** Permite cargar archivos Excel (`ReportePW_...xlsx`) descargados desde PICIZ.
- **Limpieza de Datos:** Automáticamente detecta la fila de encabezados y filtra las columnas relevantes (Nombre Compañía, Placa, Fecha Registro, Tipo Ingreso, Número Documento, Tránsito, Fecha Báscula).
- **Cálculo de Tiempos:** Calcula los días restantes basados en un límite de 5 días hábiles.
- **Semaforización (Alertas Visuales):** 
  - 🔴 **Rojo (<= 0 días):** Vencido o vence hoy.
  - 🟡 **Amarillo (1 a 2 días):** Próximo a vencer.
  - 🟢 **Verde (>= 3 días):** A tiempo.
- **Exportación:** Opción para descargar la tabla consolidada y estructurada a un nuevo archivo Excel (`.xlsx`).