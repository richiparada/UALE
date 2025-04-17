# Rutas
$carpetaOrigen = "C:\Ruta\De\Tus\Archivos"
$carpetaDestino = "C:\Ruta\Donde\Guardar\CSV"

# Crear carpeta destino si no existe
if (!(Test-Path -Path $carpetaDestino)) {
    New-Item -ItemType Directory -Path $carpetaDestino | Out-Null
}

# Fecha y hora para el nombre del CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$archivoCSV = Join-Path $carpetaDestino "resultado_$timestamp.csv"

# Cargar módulo de AD si no está
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}

# Lista de resultados
$resultados = @()

# Procesar archivos
Get-ChildItem -Path $carpetaOrigen -Filter *.txt | ForEach-Object {
    $nombreArchivo = $_.Name

    # Asegurarse que sea un arreglo de líneas
    $lineas = @(Get-Content $_.FullName)

    # Buscar la última línea válida (desde el final)
    $ultimaValida = $null
    for ($i = $lineas.Count - 1; $i -ge 0; $i--) {
        $linea = $lineas[$i].Trim()
        if ($linea -match ".+,.+") {
            $ultimaValida = $linea
            break
        }
    }

    if ($ultimaValida) {
        $partes = $ultimaValida -split ","

        if ($partes.Count -eq 2) {
            $usuario = $partes[0].Trim()
            $equipo = $partes[1].Trim()
            
            # Obtener nombre completo y OUs desde AD
            try {
                $adUser = Get-ADUser -Identity $usuario -Properties Name, distinguishedName -ErrorAction Stop
                $nombreCompleto = $adUser.Name
                $distinguishedName = $adUser.distinguishedName
                
                $ouParts = $distinguishedName -split ","
                $ous = $ouParts | Where-Object { $_ -like "OU=*" } | ForEach-Object { $_.Substring(3) }
                $organizationalUnits = $ous -join "\"
            } catch {
                $nombreCompleto = "[Usuario no encontrado]"
                $organizationalUnits = "[No encontrado en AD]"
            }

            # Obtener IP desde AD (del equipo)
            try {
                $adEquipo = Get-ADComputer -Identity $equipo -Properties IPv4Address -ErrorAction Stop
                $ipEquipo = $adEquipo.IPv4Address
                if (-not $ipEquipo) { $ipEquipo = "[Sin IP registrada]" }
            } catch {
                $ipEquipo = "[Equipo no encontrado]"
            }

            # Agregar los datos al resultado
            $resultados += [PSCustomObject]@{
                Archivo            = $nombreArchivo
                Usuario            = $usuario
                Equipo             = $equipo
                NombreCompleto     = $nombreCompleto
                IP_Equipo          = $ipEquipo
                OUs                = $organizationalUnits
            }
        } else {
            $resultados += [PSCustomObject]@{
                Archivo            = $nombreArchivo
                Usuario            = "[Formato inválido]"
                Equipo             = "[Formato inválido]"
                NombreCompleto     = "[No procesado]"
                IP_Equipo          = "[No procesado]"
                OUs                = "[No procesado]"
            }
        }
    } else {
        $resultados += [PSCustomObject]@{
            Archivo            = $nombreArchivo
            Usuario            = "[Sin línea válida]"
            Equipo             = "[Sin línea válida]"
            NombreCompleto     = "[No procesado]"
            IP_Equipo          = "[No procesado]"
            OUs                = "[No procesado]"
        }
    }
}

# Exportar resultados
$resultados | Export-Csv -Path $archivoCSV -NoTypeInformation -Encoding UTF8

Write-Host "✅ CSV generado exitosamente: $archivoCSV"
