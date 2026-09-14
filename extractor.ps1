# Configuración de conexión a SQL Server
$server = "SRVCAG09"
$database = "AGRCA"

$usuarioSql = "reportes"
$passwordSql = "123/01"

$connectionString = "Server=$server;Database=$database;User Id=$usuarioSql;Password=$passwordSql;"

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

    # 1. EXTRACCIÓN PRINCIPAL (KRINVDIS) PARA EXISTENCIAS Y TARJETAS
    $commandInv = $connection.CreateCommand()
    $commandInv.CommandText = "SELECT * FROM KRINVDIS"
    $adapterInv = New-Object System.Data.SqlClient.SqlDataAdapter($commandInv)
    $datasetInv = New-Object System.Data.DataSet
    $adapterInv.Fill($datasetInv) | Out-Null

    # 2. EXTRACCIÓN SECUNDARIA (KRLOTDIS2) PARA LOTES Y VENCIMIENTOS
    $commandLotes = $connection.CreateCommand()
    $commandLotes.CommandText = "SELECT * FROM AGRCA.DBO.KRLOTDIS2"
    $adapterLotes = New-Object System.Data.SqlClient.SqlDataAdapter($commandLotes)
    $datasetLotes = New-Object System.Data.DataSet
    $adapterLotes.Fill($datasetLotes) | Out-Null

    $connection.Close()

    $dt = $datasetInv.Tables[0]
    $dtLotes = $datasetLotes.Tables[0]

    # Nombres de columnas según la vista principal
    $colCodigo = $dt.Columns[0].ColumnName
    $colDesc   = $dt.Columns[1].ColumnName
    $colUm     = $dt.Columns[2].ColumnName

    # Lista de almacenes a procesar
    $almacenesKeys = @("BARQUISIMETO", "CAGUA", "CAPITAL", "ANDES", "MARGARITA", "ZULIA", "BARCELONA", "BOLIVAR")

    # Identificar dinámicamente las filas de resumen
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

    $factoresBidon = @{
        "1VEG001" = 60.0
        "1VEG002" = 55.0
        "1VEG003" = 180.0
    }

    foreach ($r in $rowsProductos) {
        $codigo = [string]$r[$colCodigo]
        $desc   = [string]$r[$colDesc]
        $um     = [string]$r[$colUm]

        $itemDict = @{
            codigo       = $codigo
            descripcion  = $desc
            um           = $um
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

        if ($factoresBidon.ContainsKey($codigo.Trim())) {
            $factor = $factoresBidon[$codigo.Trim()]
            
            $itemLV = @{
                codigo       = $codigo
                descripcion  = $desc
                um           = $um
                factor_bidon = $factor
                existencias = @{}
            }

            foreach ($alm in $almacenesKeys) {
                $kg = To-Double $r[$alm]
                if ($factor -gt 0) {
                    $bidones = [math]::Round(($kg / $factor), 2)
                } else {
                    $bidones = 0.0
                }
                $itemLV.existencias[$alm.ToLower()] = @{
                    kg      = [math]::Round($kg, 2)
                    bidones = $bidones
                }
            }
            $lineaVerdeList += $itemLV
        }
    }

    # Procesamiento de lotes desde KRLOTDIS2 de forma segura
    $lotesList = @()
    foreach ($rLote in $dtLotes.Rows) {
        $codLote = [string]$rLote[0]
        if ([string]::IsNullOrWhiteSpace($codLote) -or $codLote -like "*TOTAL*") { continue }

        $valUbicacion = ""
        if ($dtLotes.Columns.Contains("UBICACION")) { $valUbicacion = [string]$rLote["UBICACION"] }

        $valLoteProd = "N/A"
        if ($dtLotes.Columns.Contains("LOTE DE PRODUCTO")) { $valLoteProd = [string]$rLote["LOTE DE PRODUCTO"] }

        $valFecFab = "N/A"
        if ($dtLotes.Columns.Contains("FECHA REC/FAB")) { $valFecFab = [string]$rLote["FECHA REC/FAB"] }

        $valFecVenc = "N/A"
        if ($dtLotes.Columns.Contains("FECHA DE VENCIMIENTO")) { $valFecVenc = [string]$rLote["FECHA DE VENCIMIENTO"] }

        $valCantDisp = 0.0
        if ($dtLotes.Columns.Contains("CANTIDAD DISPONIBLE")) { $valCantDisp = To-Double $rLote["CANTIDAD DISPONIBLE"] }

        $valUmBase = "CJ"
        if ($dtLotes.Columns.Contains("U D M BASE")) { $valUmBase = [string]$rLote["U D M BASE"] }

        $lotesList += @{
            ubicacion         = $valUbicacion
            lote              = $valLoteProd
            codigo            = [string]$rLote[0]
            descripcion       = [string]$rLote[1]
            fecha_fabricacion = $valFecFab
            fecha_vencimiento = $valFecVenc
            cantidad          = $valCantDisp
            um                = $valUmBase
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
        if ($rowTotalesCajas) {
            $cajasAlm = [math]::Round((To-Double $rowTotalesCajas[$alm]), 2)
        } else {
            $cajasAlm = 0.0
        }

        if ($rowTotalesKg) {
            $kgAlm = [math]::Round((To-Double $rowTotalesKg[$alm]), 2)
        } else {
            $kgAlm = 0.0
        }

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
        lotes                = $lotesList
    }

    $json = $data | ConvertTo-Json -Depth 5
    $jsContent = "var DATA_INVENTARIO = $json;"
    $jsPath = Join-Path -Path $PSScriptRoot -ChildPath "datos.js"
    [System.IO.File]::WriteAllText($jsPath, $jsContent, [System.Text.Encoding]::UTF8)

    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] datos.js actualizado localmente con éxito." -ForegroundColor Green

    # -------------------------------------------------------------
    # SECUENCIA DE PUBLICACIÓN AUTOMÁTICA EN GITHUB (ACTIVADA)
    # -------------------------------------------------------------
    Set-Location -Path $PSScriptRoot
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Preparando cambios en Git..." -ForegroundColor Yellow
    git add datos.js
    $status = git status --porcelain
    if ($status) {
        $fechaActual = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        git commit -m "Actualizacion automatica de inventario y lotes - $fechaActual"
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Enviando datos a GitHub..." -ForegroundColor Yellow
        git push origin main --force
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Publicación en GitHub completada con éxito." -ForegroundColor Green
    } else {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] No hay cambios nuevos en datos.js para enviar." -ForegroundColor Cyan
    }

}
catch {
    Write-Host "Error al extraer datos: $_" -ForegroundColor Red
}