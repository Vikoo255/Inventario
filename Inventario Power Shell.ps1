# Guardar la política de ejecución actual
$originalExecutionPolicy = Get-ExecutionPolicy

# Cambiar la política de ejecución a RemoteSigned (o Unrestricted)
Set-ExecutionPolicy RemoteSigned -Scope Process -Force

try {
    # Verificar si el módulo ImportExcel está instalado, si no, instalarlo
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "El módulo ImportExcel no está instalado. Procediendo con la instalación..."
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }

    Import-Module ImportExcel

    # Ruta al archivo Excel
    $excelFilePath = "D:\ReportePCs.xlsx"

    function Reporte(){
        # Solicitar el RUT al usuario
        $rut = Read-Host "Ingrese el RUT (sin puntos, con guion, por ejemplo: 12345678-9)"
        
        # Validar que el RUT tenga el formato correcto
        if ($rut -match '^\d{1,8}-[\dkK]{1}$') {

            # Obtener el modelo del PC utilizando Get-WmiObject y seleccionando el primer resultado
            $modeloPC = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model)

            # Obtener los DeviceID y Tamaños de los discos en bytes, luego convertirlos a GB
            $tamaniosDiscos = Get-WmiObject Win32_DiskDrive | Select-Object DeviceID, Size | ForEach-Object {
                $tamanioGB = [math]::Round($_.Size / 1GB, 2)
                "$($_.DeviceID): $tamanioGB GB"
            }

            # Unir todos los tamaños en una cadena separada por comas
            $tamaniosDiscos = $tamaniosDiscos -join ", "

            # Obtener la cantidad total de RAM en GB
            $totalRAM = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)

            # Crear un objeto con los datos extraídos del sistema
            $reporte = [PSCustomObject]@{
                RUT                     = $rut  # RUT en la columna A
                Fecha                   = Get-Date
                UsuarioActual           = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).name
                SistemaOperativo        = (Get-CimInstance Win32_OperatingSystem).Caption
                Marca                   = (Get-WmiObject Win32_ComputerSystem).Manufacturer
                ModeloPC                = $modeloPC  # Usando Get-WmiObject para obtener el modelo del PC
                NumeroSeriePC           = (Get-WmiObject win32_bios).SerialNumber
                DireccionMAC            = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled -eq "True"}).MACAddress
                Procesador              = (Get-CimInstance Win32_Processor).Name
                TarjetaVideo            = (Get-WmiObject win32_VideoController).Name -join ", "
                Discos                  = $tamaniosDiscos  # Tamaños de discos en GB con sus DeviceID, separados por comas
                MemoriaRAM              = "$totalRAM GB"  # Total de RAM en GB
                Impresoras              = (get-printer).name -join ", "
            }

            # Escribir los datos en la siguiente fila disponible del archivo Excel
            $reporte | Export-Excel -Path $excelFilePath -WorksheetName 'Reporte' -Append -AutoSize
        } else {
            Write-Host "El RUT ingresado no tiene el formato correcto. Inténtelo de nuevo."
        }
    }

    # Ejecutar el reporte
    Reporte
}
finally {
    # Restaurar la política de ejecución original
    Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
}



