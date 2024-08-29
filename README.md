# Ejecución del Script de Extracción de Datos del Sistema

# Objetivo

El objetivo es guiar al usuario en la ejecución de un script de PowerShell que extrae información clave del sistema, como el modelo del PC, la cantidad de RAM, los discos instalados, y otros datos relevantes. Los datos se almacenan en un archivo Excel especificado.

# Requisitos Previos

Acceso a PowerShell: Debes tener acceso a PowerShell en el equipo donde se ejecutará el script.
Permisos de Administrador: Es necesario que tengas permisos de administrador en el equipo, ya que se requiere cambiar temporalmente la política de ejecución de scripts.
PowerShell 5.1 o superior: Asegúrate de que el equipo tenga PowerShell 5.1 o una versión superior instalada.
Módulo ImportExcel: El script instalará automáticamente el módulo ImportExcel si no está presente.

# Pasos para Ejecutar el Script
# Paso 1: Abrir PowerShell como Administrador

Buscar PowerShell:

En el menú de inicio, busca "PowerShell".
Haz clic derecho sobre "Windows PowerShell" y selecciona "Ejecutar como administrador".

# Paso 2: Preparar y Ejecutar el Script

Guardar el Script en un Archivo:

Copia el código proporcionado y guárdalo en un archivo con extensión .ps1, por ejemplo, ReporteSistema.ps1.
Modificar el Script si es Necesario:

Asegúrate de que la ruta del archivo Excel ($excelFilePath) apunte al lugar correcto en tu sistema.
Si el archivo Excel no existe, créalo antes de ejecutar el script o modifica el script para que lo cree automáticamente.

