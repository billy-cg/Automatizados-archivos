Descargar

[Descargar última versión (Windows)](https://github.com/billy-cg/Automatizados-archivos/releases)



Automatizador de Archivos – Conversión y Exportación

Aplicación de escritorio desarrollada en **Python + Tkinter** que permite procesar, convertir y exportar distintos tipos de archivos de forma simple y directa, sin depender de plataformas web.



Origen del proyecto

Este proyecto surge a partir del trabajo cotidiano con **planillas de Excel y otros tipos de archivos**, donde se repetían tareas manuales como correcciones de datos, movimiento de archivos entre carpetas y generación de informes.

Con el objetivo de **simplificar estos procesos** y evitar el uso constante de plataformas de internet, se desarrolló esta **aplicación de escritorio**, que permite trabajar directamente sobre los archivos locales, agilizando tareas comunes y reduciendo errores.

La aplicación facilita operaciones básicas sobre archivos de Excel y la **generación de informes en PDF**, además de permitir la conversión y exportación entre distintos formatos de una forma práctica, rápida y centralizada dentro de la misma carpeta de trabajo.


Aclaración importante

La aplicación fue desarrollada con **asistencia de una inteligencia artificial**, ya que no cuento con experiencia previa en Python ni en el uso de sus librerías.

El enfoque principal del proyecto fue **resolver una necesidad real**, comprender la lógica de los procesos involucrados y aprender durante el desarrollo, priorizando la funcionalidad, la claridad del código y la utilidad práctica de la herramienta.



Funcionalidades principales

Selección de archivos
- Selección mediante explorador
- Arrastrar y soltar archivos
- Visualización del archivo cargado

Procesamiento de datos
- Procesa archivos **.xlsx / .xls / .csv**
- Elimina duplicados
- Elimina filas vacías
- Genera un archivo Excel procesado

Conversión directa (sin procesar)
- **Word (.docx) → PDF**
- **PDF → Word (.docx)**  
(No requiere procesamiento previo)

Exportación
Desde el último archivo procesado:
- PDF
- Word
- CSV
- TXT

Interfaz
- Tema claro / oscuro
- Interfaz simple e intuitiva
- No requiere conexión a internet



Requisitos del sistema

- Windows 10 / 11
- Para usuarios finales: ejecutar el archivo **.exe**
- Para desarrollo: Python 3.11 o superior



Uso de la aplicación

1. Ejecutar la aplicación
2. Seleccionar o arrastrar un archivo
3. Opcional: procesar archivo Excel / CSV
4. Elegir conversión directa o exportación
5. El archivo generado se guarda en la misma carpeta



Consideraciones importantes

- Windows puede mostrar advertencias de seguridad si el `.exe` no está firmado  
  → “Más información” → “Ejecutar de todas formas”
- La conversión PDF → Word extrae **texto plano**
- Los archivos originales no se modifican



Instalación (modo desarrollador)

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py


Licencia

Este proyecto se distribuye bajo la licencia **MIT**.
