# Reporte de Sistema en Excel
Este repositorio contiene un script de PowerShell que extrae información clave del sistema, como el modelo del PC, la cantidad de RAM, los discos instalados, y otros datos relevantes. Los datos se almacenan en un archivo Excel especificado.

# Requisitos
PowerShell 5.1 o superior.
Permisos de administrador para ejecutar el script.
Módulo ImportExcel: El script instalará automáticamente el módulo ImportExcel si no está presente en el sistema.
# Instrucciones de Uso
# Paso 1: Abrir PowerShell como Administrador
Abrir PowerShell:
En Windows, busca "PowerShell" en el menú de inicio.
Haz clic derecho en "Windows PowerShell" y selecciona "Ejecutar como administrador".
Verificar la Política de Ejecución:

Ejecuta el siguiente comando en PowerShell para verificar la política de ejecución actual:
powershell
Copiar código
Get-ExecutionPolicy
Anota la política mostrada (por ejemplo, Restricted, RemoteSigned, etc.)
# Paso 2: Ejecutar el Script
Modificar la Ruta del Archivo Excel:
Abre el script ReporteSistema.ps1 en un editor de texto.
Asegúrate de que la ruta del archivo Excel ($excelFilePath) apunte al lugar correcto en tu sistema.
Ejecutar el Script:

Navega hasta la ubicación del script en PowerShell:
powershell
Copiar código
cd "C:\ruta\al\directorio\del\script"
Ejecuta el script:
powershell
Copiar código
.\ReporteSistema.ps1
Ingreso del RUT:

Cuando se ejecute el script, te pedirá que ingreses el RUT sin puntos y con guion. Ejemplo: 12345678-9.
Ingresa el RUT en el formato correcto y presiona Enter.
# Paso 3: Verificación de los Resultados
Verificar el Archivo Excel:
Una vez que el script termine de ejecutarse, abre el archivo Excel en la ruta especificada (D:\ReportePCs.xlsx por defecto).
Verifica que la información del sistema se haya registrado correctamente.
# Restaurar la Política de Ejecución:
El script restaura automáticamente la política de ejecución a su valor original. Puedes verificar esto ejecutando:
powershell
Copiar código
Get-ExecutionPolicy
# Solución de Problemas
# Error de Ejecución de Scripts:
Si el script no se ejecuta, asegúrate de que PowerShell se esté ejecutando como administrador y de que la política de ejecución se haya configurado correctamente.
# Falta de Permisos:
Asegúrate de tener permisos de administrador en el sistema. Si no tienes los permisos necesarios, consulta con tu administrador de sistemas.
# Contribuciones
Las contribuciones son bienvenidas. Si tienes alguna mejora o encuentras un error, no dudes en crear un pull request o abrir un issue.
# Licencia
Este proyecto está licenciado bajo la Licencia MIT - consulta el archivo LICENSE para más detalles.
