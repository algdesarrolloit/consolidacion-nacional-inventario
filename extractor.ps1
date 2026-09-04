# Configuración de conexión a SQL Server
$server = "SRVCAG09"
$database = "AGRCA"

$usuarioSql = "reportes"
$passwordSql = "123/01"

$connectionString = "Server=$server;Database=$database;User Id=$usuarioSql;Password=$passwordSql;"

$queryDetalle = "SELECT * FROM KRINVDIS"

# Función auxiliar para convertir valores numéricos con formato a Double de forma segura
function To-Double ($val) {
    if ($null -eq $val -or [string]::IsNullOrWhiteSpace("$val")) { return 0.0 }
    if ($val -is [double] -or $val -is [float] -or $val -is [decimal] -or $val -is [int] -or $val -is [long]) {
        return [double]$val
    }
    $valStr = "$val".Trim().Replace('.', '').Replace(',', '.')
    $num = 0.0
    if ([double]::TryParse($valStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        return $num
    }
    return 0.0
}

try {
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = $queryDetalle
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $connection.Close()

    $dt = $dataset.Tables[0]

    # Nombres de columnas según la vista
    $colCodigo = $dt.Columns[0].ColumnName
    $colDesc   = $dt.Columns[1].ColumnName
    $colUm     = $dt.Columns[2].ColumnName

    # Lista de almacenes a procesar
    $almacenesKeys = @("BARQUISIMETO", "CAGUA", "CAPITAL", "ANDES", "MARGARITA", "ZULIA", "BARCELONA", "BOLIVAR")

    # Identificar dinámicamente las filas de resumen (sin importar en qué posición estén)
    $rowTotalesCajas = $null
    $rowTotalesKg = $null
    $rowsProductos = @()

    foreach ($r in $dt.Rows) {
        $codigo = [string]$r[$colCodigo]
        $desc = [string]$r[$colDesc]

        if ($codigo -like "*TOTAL ALMACEN CAJAS*" -or $desc -like "*TOTAL ALMACEN CAJAS*") {
            $rowTotalesCajas = $r
        }
        elseif ($codigo -like "*TOTAL ALMACEN KG*" -or $desc -like "*TOTAL ALMACEN KG*") {
            $rowTotalesKg = $r
        }
        else {
            $rowsProductos += $r
        }
    }

    # Mapeo de ítems regulares y cálculo de Línea Verde
    $itemsList = @()
    $lineaVerdeList = @()

    # Factores de conversión para Línea Verde (KG por Bidón)
    $factoresBidon = @{
        "1VEG001" = 60.0   # ACEITUNAS ENTERAS
        "1VEG002" = 55.0   # ACEITUNAS RELLENAS
        "1VEG003" = 180.0  # ALCAPARRAS
    }

    foreach ($r in $rowsProductos) {
        $codigo = [string]$r[$colCodigo]
        $desc   = [string]$r[$colDesc]
        $um     = [string]$r[$colUm]

        $itemDict = @{
            codigo      = $codigo
            descripcion = $desc
            um          = $um
            barquisimeto = [math]::Round((To-Double $r["BARQUISIMETO"]), 2)
            cagua        = [math]::Round((To-Double $r["CAGUA"]), 2)
            capital      = [math]::Round((To-Double $r["CAPITAL"]), 2)
            andes        = [math]::Round((To-Double $r["ANDES"]), 2)
            margarita    = [math]::Round((To-Double $r["MARGARITA"]), 2)
            zulia        = [math]::Round((To-Double $r["ZULIA"]), 2)
            barcelona    = [math]::Round((To-Double $r["BARCELONA"]), 2)
            bolivar      = [math]::Round((To-Double $r["BOLIVAR"]), 2)
        }

        $itemsList += $itemDict

        # Si el código pertenece a Línea Verde, calculamos sus equivalencias en Bidones
        if ($factoresBidon.ContainsKey($codigo.Trim())) {
            $factor = $factoresBidon[$codigo.Trim()]
            
            $itemLV = @{
                codigo      = $codigo
                descripcion = $desc
                um          = $um
                factor_bidon = $factor
                existencias = @{}
            }

            foreach ($alm in $almacenesKeys) {
                $kg = To-Double $r[$alm]
                $bidones = if ($factor -gt 0) { [math]::Round(($kg / $factor), 2) } else { 0.0 }
                $itemLV.existencias[$alm.ToLower()] = @{
                    kg      = [math]::Round($kg, 2)
                    bidones = $bidones
                }
            }
            $lineaVerdeList += $itemLV
        }
    }

    # Procesamiento dinámico de totales por almacén (Cajas y KG)
    $almacenesData = @()
    $totalGeneralCajas = 0.0
    $totalGeneralKg = 0.0

    $codigosAlm = @{
        "BARQUISIMETO" = "BTO"; "CAGUA" = "CAG"; "CAPITAL" = "CAP"; "ANDES" = "AND";
        "MARGARITA" = "MAR"; "ZULIA" = "ZUL"; "BARCELONA" = "BAR"; "BOLIVAR" = "BOL"
    }

    foreach ($alm in $almacenesKeys) {
        $cajasAlm = if ($rowTotalesCajas) { [math]::Round((To-Double $rowTotalesCajas[$alm]), 2) } else { 0.0 }
        $kgAlm    = if ($rowTotalesKg)    { [math]::Round((To-Double $rowTotalesKg[$alm]), 2) } else { 0.0 }

        $totalGeneralCajas += $cajasAlm
        $totalGeneralKg    += $kgAlm

        $almacenesData += @{
            almacen = $alm
            campo   = $alm.ToLower()
            codigo  = $codigosAlm[$alm]
            cajas   = $cajasAlm
            kg      = $kgAlm
        }
    }

    $data = @{
        ultima_actualizacion = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
        total_general_cajas  = $totalGeneralCajas
        total_general_kg     = $totalGeneralKg
        almacenes            = $almacenesData
        linea_verde          = $lineaVerdeList
        productos            = $itemsList
    }

    $json = $data | ConvertTo-Json -Depth 5
    $jsContent = "var DATA_INVENTARIO = $json;"
    $jsPath = Join-Path -Path $PSScriptRoot -ChildPath "datos.js"
    [System.IO.File]::WriteAllText($jsPath, $jsContent, [System.Text.Encoding]::UTF8)

    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] datos.js actualizado con éxito." -ForegroundColor Green
}
catch {
    Write-Host "Error al extraer datos: $_" -ForegroundColor Red
}