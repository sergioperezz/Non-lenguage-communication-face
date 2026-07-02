' ============================================================================
'  Macro del PANEL GENÉRICO (Fase 1) + copia a PowerPoint como EMF
'  Modelo: Tipo entidad -> Entidad(1..3) -> Grupo -> Métrica -> Dimensión -> ...
'
'  INSTALACIÓN (una vez):
'   1) PARTE A -> clic derecho en la pestaña "Panel" -> "Ver código" y pega ahí.
'   2) PARTE B -> Insertar -> Módulo, y pega ahí.
'   3) PARTE C -> clic derecho en la pestaña "Tablas" -> "Ver código" y pega ahí.
'   4) PARTE D (Fase 2, opcional) -> Insertar -> Módulo NUEVO y pega ahí.
'   5) Guarda como .xlsm. (Opcional: botón con la macro "CopiarAPowerPoint".)
'
'  Controles: B3 Tipo entidad · B4 Entidad 1 · B5/B6 Entidad 2/3 (opc.)
'   · B7 Grupo · B8 Métrica · B9 Dimensión · B10 Filtro tipo activo · B11 Periodo
'   · B12 Benchmark (Con/Sin) · B13 Estilo benchmark · B14 Tipo de gráfico
' ============================================================================


' =====================  PARTE A: en la hoja "Panel"  ========================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Salir
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Al cambiar el TIPO de entidad: ajustar Entidad 1 y limpiar comparadores inválidos.
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        AjustarSeleccion "B4", "Ent_" & Me.Range("B3").Value
        Me.Range("B5").Value = "(ninguna)"   ' resetear comparadores al cambiar de tipo
        Me.Range("B6").Value = "(ninguna)"
    End If
    ' Al cambiar el GRUPO de métrica: ajustar la métrica.
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        AjustarSeleccion "B8", "Grupo_" & Me.Range("B7").Value
    End If

    If Not Intersect(Target, Me.Range("B3:B14")) Is Nothing Then AplicarGrafico

    ' --- Fase 2: refrescar la vista previa de la SQL (si está pegada la PARTE D).
    '     No conecta a BigQuery; solo construye el texto de la consulta.
    If Not Intersect(Target, Me.Range("B3:B11")) Is Nothing Then
        On Error Resume Next
        Application.Run "ActualizarSQL"
        On Error GoTo Salir
    End If
    ' Para traer datos reales de BigQuery: botón con la macro "RefrescarDatos" (PARTE D).

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    AplicarGrafico
    Application.ScreenUpdating = True
End Sub

' Si "celda" no está en el rango con nombre, la pone al primer elemento válido.
Private Sub AjustarSeleccion(ByVal celda As String, ByVal nombreLista As String)
    Dim rng As Range, c As Range, valido As Boolean
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nombreLista).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If c.Value = Me.Range(celda).Value Then valido = True
    Next c
    If Not valido Then Me.Range(celda).Value = rng.Cells(1, 1).Value
End Sub

' Añade una serie de entidad (columna colLetter) si el slot tiene una entidad
' real (ni vacío ni "(ninguna)").
Private Sub AddEnt(ByVal ch As Chart, ByVal slotCell As String, ByVal colLetter As String, ByVal lastRow As Long)
    Dim s As Series, v As String
    v = Trim(Me.Range(slotCell).Value)
    If v = "" Or v = "(ninguna)" Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "=Panel!$" & colLetter & "$2"          ' cabecera = nombre de la entidad
    s.Values = "=Panel!$" & colLetter & "$3:$" & colLetter & "$" & lastRow
    s.XValues = "=Panel!$D$3:$D$" & lastRow
End Sub

' Reconstruye series (multi-entidad + benchmark), tipo, estilo, título y ejes.
Private Sub AplicarGrafico()
    Dim ch As Chart, s As Series, conBench As Boolean, tipo As String
    Dim vis As Long, lastRow As Long, benchIdx As Long
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    conBench = (Me.Range("B12").Value = "Con benchmark")
    tipo = LCase(Trim(Me.Range("B14").Value))
    vis = Me.Range("B16").Value
    If vis < 1 Then vis = 1
    lastRow = 2 + vis

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ' Una serie por entidad seleccionada (E=Ent1, F=Ent2, G=Ent3).
    AddEnt ch, "B4", "E", lastRow
    AddEnt ch, "B5", "F", lastRow
    AddEnt ch, "B6", "G", lastRow

    ' Benchmark (H) de la Entidad 1, si procede.
    benchIdx = 0
    If conBench And Trim(Me.Range("B4").Value) <> "" Then
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "=Panel!$H$2"
        s.Values = "=Panel!$H$3:$H$" & lastRow
        s.XValues = "=Panel!$D$3:$D$" & lastRow
        benchIdx = ch.SeriesCollection.Count
    End If

    ' Tipo de gráfico base (B14).
    Select Case tipo
        Case "barras":            ch.ChartType = xlBarClustered
        Case "líneas", "lineas":  ch.ChartType = xlLineMarkers
        Case "área", "area":      ch.ChartType = xlArea
        Case "circular":          ch.ChartType = xlPie
        Case "anillo":            ch.ChartType = xlDoughnut
        Case "radar":             ch.ChartType = xlRadarMarkers
        Case "apiladas":          ch.ChartType = xlColumnStacked
        Case "100% apiladas":     ch.ChartType = xlColumnStacked100
        Case Else:                ch.ChartType = xlColumnClustered  ' "Columnas"
    End Select

    ' Estilo del benchmark (B13): su serie se dibuja como Barras, Líneas o Puntos.
    If benchIdx > 0 Then
        On Error Resume Next
        Select Case LCase(Trim(Me.Range("B13").Value))
            Case "líneas", "lineas"
                ch.FullSeriesCollection(benchIdx).ChartType = xlLineMarkers
            Case "puntos"
                ch.FullSeriesCollection(benchIdx).ChartType = xlLineMarkers
                ch.FullSeriesCollection(benchIdx).Format.Line.Visible = msoFalse  ' solo marcadores
            Case Else  ' Barras
                ch.FullSeriesCollection(benchIdx).ChartType = xlColumnClustered
        End Select
        On Error GoTo 0
    End If

    ' Título (A19) y títulos de eje (Y = métrica B8, X = dimensión B9).
    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = Me.Range("A19").Value
    If tipo <> "circular" And tipo <> "anillo" And tipo <> "radar" Then
        ch.Axes(xlValue).HasTitle = True
        ch.Axes(xlValue).AxisTitle.Text = Me.Range("B8").Value
        ch.Axes(xlCategory).HasTitle = True
        ch.Axes(xlCategory).AxisTitle.Text = Me.Range("B9").Value
    End If
    On Error GoTo 0
End Sub


' =================  PARTE B: en un Módulo estándar  =========================
' Copia el gráfico a PowerPoint como EMF (vectorial, sin vínculos -> no se rompe).

Public Sub CopiarAPowerPoint()
    Dim ch As ChartObject
    On Error Resume Next
    Set ch = ThisWorkbook.Sheets("Panel").ChartObjects(1)
    On Error GoTo 0
    If ch Is Nothing Then
        MsgBox "No encuentro el gráfico en la hoja 'Panel'.", vbExclamation
        Exit Sub
    End If

    ch.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture   ' EMF en Windows

    Dim ppt As Object, pres As Object, sld As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True

    If ppt.Presentations.Count = 0 Then
        Set pres = ppt.Presentations.Add
    Else
        Set pres = ppt.ActivePresentation
    End If

    On Error Resume Next
    Set sld = ppt.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        If pres.Slides.Count = 0 Then
            Set sld = pres.Slides.Add(1, 12)   ' 12 = ppLayoutBlank
        Else
            Set sld = pres.Slides(pres.Slides.Count)
        End If
    End If

    Dim shp As Object
    Set shp = sld.Shapes.PasteSpecial(DataType:=2)   ' 2 = ppPasteEnhancedMetafile

    On Error Resume Next
    shp.Left = (pres.PageSetup.SlideWidth - shp.Width) / 2
    shp.Top = (pres.PageSetup.SlideHeight - shp.Height) / 2
    On Error GoTo 0
End Sub


' =====================  PARTE C: en la hoja "Tablas"  =======================
' Generador de tablas configurable: al cambiar un desplegable (o al abrir la
' pestaña) ajusta columnas visibles, decimales y el estilo de formato condicional.
' Pegar clic derecho en la pestaña "Tablas" -> "Ver código".

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B3:B10")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Salir
    FormatearTablas
Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    On Error Resume Next
    FormatearTablas
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' Selectores: B3 tipo entidad · B4 métrica · B5 dimensión (columnas) · B6 serie
'  · B7 filtro tipo activo · B8 periodo · B9 estilo · B10 decimales.
Private Sub FormatearTablas()
    Const HDR As Long = 17, ROW0 As Long = 18, LASTROW As Long = 21
    Const FIRSTC As Long = 2, LASTC As Long = 13      ' columnas B..M
    Dim body As Range, c As Long, fmt As String

    Set body = Me.Range(Me.Cells(ROW0, FIRSTC), Me.Cells(LASTROW, LASTC))

    ' 1) Decimales (B10) -> formato de número.
    Select Case CLng(Me.Range("B10").Value)
        Case 0: fmt = "0"
        Case 1: fmt = "0.0"
        Case Else: fmt = "0.00"
    End Select
    body.NumberFormat = fmt

    ' 2) Ocultar columnas de categoría vacías (según dimensión + periodo).
    '    La columna B nunca se oculta (siempre hay >=1 bucket y ahí no hay selectores).
    For c = FIRSTC + 1 To LASTC
        Me.Columns(c).Hidden = (Trim(CStr(Me.Cells(HDR, c).Value)) = "")
    Next c

    ' 3) Estilo (B9): limpiar y aplicar el elegido.
    body.FormatConditions.Delete
    Select Case LCase(Trim(Me.Range("B9").Value))
        Case "mapa de calor"
            With body.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
                .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)   ' rojo
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)   ' amarillo
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)    ' verde
            End With
        Case "barras de datos"
            With body.FormatConditions.AddDatabar
                .BarColor.Color = RGB(0, 114, 206)
                On Error Resume Next
                .BarFillType = xlDataBarFillGradient
                On Error GoTo 0
            End With
        Case "signos +/-"
            With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Font.Color = RGB(16, 124, 65)       ' verde
            End With
            With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Font.Color = RGB(192, 0, 0)         ' rojo
            End With
        Case Else
            ' "Sin formato": ya se ha limpiado, no se añade nada.
    End Select
End Sub


' =============  PARTE D: en un MÓDULO ESTÁNDAR NUEVO (Fase 2)  ==============
' Con los parámetros del Panel construye la SQL contra el MODELO HOMOLOGADO real
' (proyecto go-cam-beg-camd9-camcd9p01-pro), como tu Consulta de Power Query.
' Puedes: (a) copiar la SQL/M a tu Power Query (Odbc.Query), o (b) usar
' RefrescarDatos para traerla por ODBC (ADODB) a la hoja "BQ_Resultado".
'
' INSTALACIÓN: Insertar -> Módulo (uno NUEVO, distinto al de la PARTE B) y pega
' todo esto. Revisa el bloque CONFIG (DSN, proyecto, datasets, filtros).
'
' Botones sugeridos (Insertar -> Forma -> Asignar macro):
'   · "Ver SQL"           -> VerSQL          (muestra la consulta SQL)
'   · "Ver M"             -> VerM            (muestra el M de Power Query)
'   · "Traer de BigQuery" -> RefrescarDatos  (conecta por ODBC y trae los datos)
'
' MAPEO (diccionario + query de ejemplo):
'   · Rentabilidad / Rentab. acum. -> CAM_TX_PERFORMANCE_FIGURES_PD: TWR_<per>,
'     TWR_<per>_BMK y DIFERENCIAL_<per>. Filtros fijos PK_NAV_GNAV='GNAV' y
'     BENCHMARK='Benchmark 1'. Se filtra por PK_PORTFOLIO_ID y se toma la última fecha.
'   · Volatilidad -> VOL_1Y_260 ; Beta -> BETA (misma tabla).
'   · Duración -> CAM_TX_RISK_FIG_AGG_PD: VALOR por PK_ETIQUETA_AGREGACION
'     (PK_VARIABLE_TARGET='Duración Modificada', PK_CRITERIO_AGREGACION segun dimensión).
'   · TIR -> CAM_TM_PORTFOLIOS_PD.TIR_VALORACION (otra tabla)          [pendiente].
'   · Peso/Composición -> CAM_TX_PORTFOLIOS_COMP_PD / CAM_TX_BENCHMARK_COMP_PD  [pendiente].
' NOTA: fondos y carteras están en tablas distintas; aquí se cubre performance
' de portfolios. El nombre visible se traduce a PK_PORTFOLIO_ID vía MapaEntidades.
' ---------------------------------------------------------------------------

' ==== CONFIG (valores reales del entorno; ajústalos si cambian) ====
Private Const BQ_DSN As String = "Conexión_BQ"                          ' DSN ODBC ya configurado
Private Const BQ_CONN As String = "DSN=Conexión_BQ;"                    ' cadena de conexión ADODB
Private Const BQ_PROJECT As String = "go-cam-beg-camd9-camcd9p01-pro"   ' proyecto de BigQuery
Private Const DS_PROD As String = "productosdatosdecontratos_ds01"      ' rentabilidad, riesgo, patrimonios
Private Const DS_OPER As String = "operativafinanciera_ds01"            ' benchmark, indices, lookthrough
Private Const DS_MERC As String = "informaciondemercado_ds01"           ' maestros de valores/precios
Private Const LANDING As String = "BQ_Resultado"                        ' hoja donde se vuelcan los datos
' Filtros obligatorios de la tabla de performance (si no, filas duplicadas):
Private Const F_NAV_GNAV As String = "GNAV"
Private Const F_BENCHMARK As String = "Benchmark 1"
' Tabla de rentabilidades:
Private Const T_PERF As String = "CAM_TX_PERFORMANCE_FIGURES_PD"
' Tabla de riesgo (VALOR = número; PK_VARIABLE_TARGET = métrica; criterio/etiqueta = desglose):
Private Const T_RISK As String = "CAM_TX_RISK_FIG_AGG_PD"
Private Const RISK_COL_FONDOBMK As String = "PK_TIPOGAMAN1"         ' columna FONDO/BENCHMARK
' Composición por sector (posiciones x maestro de valores):
Private Const T_POS As String = "CAM_TM_PORTFOLIOS_PD"             ' posiciones (dataset DS_PROD)
Private Const T_VALORES As String = "CAM_TM_MSTR_VALORES_PD"       ' maestro de valores (dataset DS_MERC)
Private Const POS_VALOR As String = "VALUATION_PC"                 ' valor de la posición (peso); PC=divisa base cartera
Private Const SECTOR_COL As String = "CLASSIFICATION_GICS"         ' estándar sectorial (GICS/BICS/ICB)
Private Const RATING_COL As String = "COMPOSITERATINGSPCOMPOSITE"  ' rating (S&P; o MOODYS/FITCH/WORST...)
' =====================================================

' Duplica comillas simples para evitar romper la cadena SQL.
Private Function Esc(ByVal s As String) As String
    Esc = Replace(CStr(s), "'", "''")
End Function

' Nombre de tabla cualificado: `proyecto.dataset.TABLA`
Private Function Tbl(ByVal ds As String, ByVal t As String) As String
    Tbl = "`" & BQ_PROJECT & "." & ds & "." & t & "`"
End Function

' Traduce el nombre de entidad a su PK_PORTFOLIO_ID vía el rango MapaEntidades
' (Listas!AJ:AK). Devuelve "" si no está mapeado.
Private Function IdEntidad(ByVal nombre As String) As String
    Dim rng As Range, c As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names("MapaEntidades").RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    For Each c In rng.Columns(1).Cells
        If Trim(CStr(c.Value)) = nombre Then
            IdEntidad = Trim(CStr(c.Offset(0, 1).Value))
            Exit Function
        End If
    Next c
End Function

' Lista de PK_PORTFOLIO_ID de las entidades seleccionadas (B4/B5/B6), saltando
' vacías y "(ninguna)". Si una no está en MapaEntidades, usa el nombre (placeholder).
Private Function ListaEntidades(ws As Worksheet) As String
    Dim celda As Variant, v As String, idp As String, out As String
    For Each celda In Array("B4", "B5", "B6")
        v = Trim(CStr(ws.Range(celda).Value))
        If v <> "" And v <> "(ninguna)" Then
            idp = IdEntidad(v)
            If Len(idp) = 0 Then idp = v
            If Len(out) > 0 Then out = out & ", "
            out = out & "'" & Esc(idp) & "'"
        End If
    Next celda
    ListaEntidades = out
End Function

' Periodo del Panel (B11) -> sufijo de columna TWR del DWH ("" si no hay columna).
Private Function SufijoPeriodo(ByVal p As String) As String
    Select Case UCase(Trim(p))
        Case "MTD": SufijoPeriodo = "MTD"
        Case "YTD": SufijoPeriodo = "YTD"
        Case "1M":  SufijoPeriodo = "1M"
        Case "1A":  SufijoPeriodo = "1Y"
        Case "3A":  SufijoPeriodo = "3Y"
        Case "5A":  SufijoPeriodo = "5Y"
        Case Else:  SufijoPeriodo = ""     ' 2M/3M/4M/5M/6M, 2A/4A/6A: sin columna directa
    End Select
End Function

' Métrica del Panel -> columnas del DWH (fondo, benchmark, diferencial).
' Devuelve True si la métrica está mapeada.
Private Function MapMetrica(ByVal met As String, ByVal per As String, _
        ByRef colVal As String, ByRef colBmk As String, ByRef colDif As String) As Boolean
    Dim suf As String: suf = SufijoPeriodo(per)
    colVal = "": colBmk = "": colDif = ""
    Select Case met
        Case "Rentabilidad", "Rentab. acum."
            If Len(suf) = 0 Then Exit Function            ' periodo sin columna en el DWH
            colVal = "TWR_" & suf
            colBmk = "TWR_" & suf & "_BMK"
            colDif = "DIFERENCIAL_" & suf
            MapMetrica = True
        Case "Volatilidad"
            colVal = "VOL_1Y_260": MapMetrica = True
        Case "Beta"
            colVal = "BETA": MapMetrica = True
    End Select
End Function

' Dimensión del Panel (B9) -> valor de PK_CRITERIO_AGREGACION en la tabla de riesgo.
' Criterios que EXISTEN en CAM_TX_RISK_FIG_AGG_PD (SELECT DISTINCT completo):
'   AssetType, Duracion, FX, Geo, TIR. No hay Sector/Rating/Industria.
Private Function DimACriterio(ByVal dimen As String) As String
    Select Case dimen
        Case "Activo":                                      DimACriterio = "AssetType"
        Case "Geografia":                                   DimACriterio = "Geo"
        Case "Divisa":                                      DimACriterio = "FX"
        Case "Mensual", "Trimestral", "Semestral", "Anual": DimACriterio = "Duracion"  ' total
        Case Else:                                          DimACriterio = ""           ' Industria/Sector/Rating: no existen
    End Select
End Function

' SQL de riesgo desde CAM_TX_RISK_FIG_AGG_PD: VALOR por etiqueta de agregación,
' última fecha por portfolio.
'  - Duración: filtra PK_VARIABLE_TARGET=variable y criterio segun la dimensión.
'  - Métricas-total (TIR): pasa critFijo (p. ej. "TIR") y no filtra por variable.
Private Function SQLRiesgo(ws As Worksheet, ByVal variable As String, ByVal ents As String, _
        Optional ByVal critFijo As String = "") As String
    Dim dimen As String, crit As String, sql As String, wVar As String
    If Len(critFijo) > 0 Then
        crit = critFijo                 ' criterio propio de la métrica-total
        wVar = ""
    Else
        dimen = Trim(CStr(ws.Range("B9").Value))
        crit = DimACriterio(dimen)
        If Len(crit) = 0 Then
            SQLRiesgo = "-- Dimensión '" & dimen & "' no existe como desglose en CAM_TX_RISK_FIG_AGG_PD." & vbLf & _
                        "-- Criterios disponibles: AssetType (Activo), Geo (Geografia), FX (Divisa), Duracion (total)." & vbLf & _
                        "-- Sector/Rating/Industria no están en la tabla de riesgo."
            Exit Function
        End If
        wVar = "  AND PK_VARIABLE_TARGET = '" & Esc(variable) & "'" & vbLf
    End If
    sql = "SELECT PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION, VALOR" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_RISK) & vbLf & _
          "WHERE PK_CRITERIO_AGREGACION = '" & Esc(crit) & "'" & vbLf & _
          wVar & _
          "  AND " & RISK_COL_FONDOBMK & " = 'FONDO'"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & _
          "QUALIFY ROW_NUMBER() OVER (PARTITION BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION" & _
          " ORDER BY PK_FECHA_DATOS DESC) = 1" & vbLf & _
          "ORDER BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION"
    SQLRiesgo = sql
End Function

' Dimensión del Panel -> columna de clasificación del maestro de valores.
Private Function DimAClasificacion(ByVal dimen As String) As String
    Select Case dimen
        Case "Sector": DimAClasificacion = "v." & SECTOR_COL
        Case "Rating": DimAClasificacion = "v." & RATING_COL
        Case Else:     DimAClasificacion = ""
    End Select
End Function

' JOIN de posiciones con el maestro de valores (por PK_SECURITY_IK y fecha).
Private Function JoinValores() As String
    JoinValores = "JOIN " & Tbl(DS_MERC, T_VALORES) & " v" & vbLf & _
                  "  ON v.PK_SECURITY_IK = p.PK_SECURITY_IK AND v.PK_FECHA_DATOS = p.PK_FECHA_DATOS" & vbLf
End Function

' SQL de composición (Peso): posiciones (CAM_TM_PORTFOLIOS_PD) x maestro de valores,
' SUMANDO el valor de mercado por clasificación (Sector/Rating). Última fecha.
Private Function SQLComposicion(ws As Worksheet, ByVal ents As String) As String
    Dim dimen As String, grp As String, sql As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    grp = DimAClasificacion(dimen)
    If Len(grp) = 0 Then
        SQLComposicion = "-- Composición por '" & dimen & "': disponible por Sector (" & SECTOR_COL & _
                         ") y Rating (" & RATING_COL & ")." & vbLf & _
                         "-- Otras dimensiones necesitan su columna de clasificación en el maestro de valores."
        Exit Function
    End If
    sql = "SELECT p.PK_PORTFOLIO_ID, " & grp & " AS categoria, SUM(p." & POS_VALOR & ") AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & JoinValores & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID, " & grp & vbLf & _
          "ORDER BY p.PK_PORTFOLIO_ID, valor DESC"
    SQLComposicion = sql
End Function

' SQL de Spread: media PONDERADA por valor de mercado del SPREAD de las posiciones.
' Por Sector/Rating si esa es la dimensión; si no, un total por portfolio.
Private Function SQLSpread(ws As Worksheet, ByVal ents As String) As String
    Dim dimen As String, grp As String, selCat As String, grpBy As String, joinV As String, sql As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    grp = DimAClasificacion(dimen)
    If Len(grp) > 0 Then
        selCat = ", " & grp & " AS categoria"
        grpBy = ", " & grp
        joinV = JoinValores
    End If
    sql = "SELECT p.PK_PORTFOLIO_ID" & selCat & "," & vbLf & _
          "       SUM(p.SPREAD * p." & POS_VALOR & ") / NULLIF(SUM(p." & POS_VALOR & "), 0) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & joinV & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID" & grpBy & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLSpread = sql
End Function

' Construye la SQL a partir de los parámetros del Panel.
' Rentabilidad/Volatilidad/Beta -> CAM_TX_PERFORMANCE_FIGURES_PD (última fecha por
' portfolio). Duración -> CAM_TX_RISK_FIG_AGG_PD (VALOR por etiqueta de agregación).
Public Function ConstruirSQL() As String
    Dim ws As Worksheet, met As String, per As String, ents As String
    Dim colVal As String, colBmk As String, colDif As String
    Dim cols As String, sql As String, conBmk As Boolean
    Set ws = ThisWorkbook.Sheets("Panel")
    met = Trim(CStr(ws.Range("B8").Value))
    per = Trim(CStr(ws.Range("B11").Value))
    ents = ListaEntidades(ws)
    conBmk = (ws.Range("B12").Value = "Con benchmark")

    ' Riesgo: Duración Modificada / Efectiva (VALOR en CAM_TX_RISK_FIG_AGG_PD).
    ' El nombre de la métrica del Panel coincide con PK_VARIABLE_TARGET del DWH.
    If InStr(met, "Duración") = 1 Then
        ConstruirSQL = SQLRiesgo(ws, met, ents)
        Exit Function
    End If
    ' TIR: métrica-total con criterio propio (PK_CRITERIO_AGREGACION='TIR').
    If met = "TIR" Then
        ConstruirSQL = SQLRiesgo(ws, "", ents, "TIR")
        Exit Function
    End If
    ' Peso/Composición: posiciones x maestro de valores (Sector/Rating).
    If met = "Peso" Then
        ConstruirSQL = SQLComposicion(ws, ents)
        Exit Function
    End If
    ' Spread: media ponderada del SPREAD de las posiciones (por Sector/Rating o total).
    If met = "Spread" Then
        ConstruirSQL = SQLSpread(ws, ents)
        Exit Function
    End If

    If Not MapMetrica(met, per, colVal, colBmk, colDif) Then
        ConstruirSQL = _
            "-- Métrica '" & met & "' / periodo '" & per & "': aún no mapeada a BigQuery." & vbLf & _
            "-- TER -> CAM_TM_MSTR_VALORES_PD.KEYFIGURESTER (decidir: TER del fondo o look-through)." & vbLf & _
            "-- PER / DividendYield: no aparecen en el diccionario. Liquidez: figura como 'Pte'."
        Exit Function
    End If

    cols = "PK_FECHA_DATOS, PK_PORTFOLIO_ID, " & colVal
    If conBmk And Len(colBmk) > 0 Then cols = cols & ", " & colBmk
    If conBmk And Len(colDif) > 0 Then cols = cols & ", " & colDif

    sql = "SELECT " & cols & vbLf & _
          "FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
          "WHERE PK_NAV_GNAV = '" & F_NAV_GNAV & "'" & vbLf & _
          "  AND BENCHMARK = '" & F_BENCHMARK & "'"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & _
          "QUALIFY ROW_NUMBER() OVER (PARTITION BY PK_PORTFOLIO_ID ORDER BY PK_FECHA_DATOS DESC) = 1" & vbLf & _
          "ORDER BY PK_PORTFOLIO_ID"
    ConstruirSQL = sql
End Function

' Envuelve la SQL en Power Query M (Odbc.Query), como en tu Consulta actual.
Public Function ConstruirM() As String
    Dim sql As String
    sql = ConstruirSQL()
    If Left(sql, 2) = "--" Then ConstruirM = sql: Exit Function
    ConstruirM = "let" & vbLf & _
        "    Origen = Odbc.Query(""dsn=" & BQ_DSN & """, """ & Replace(sql, vbLf, " ") & """)" & vbLf & _
        "in" & vbLf & "    Origen"
End Function

' Escribe la SQL en la vista previa del Panel (A41). La llama Worksheet_Change.
Public Sub ActualizarSQL()
    On Error Resume Next
    ThisWorkbook.Sheets("Panel").Range("A41").Value = ConstruirSQL()
End Sub

' Muestra la SQL de los parámetros actuales (sin conectar).
Public Sub VerSQL()
    ActualizarSQL
    MsgBox ConstruirSQL(), vbInformation, "SQL para los parámetros actuales"
End Sub

' Muestra el código Power Query M (Odbc.Query) listo para pegar en una consulta.
Public Sub VerM()
    MsgBox ConstruirM(), vbInformation, "Power Query (M) para los parámetros actuales"
End Sub

' Conecta a BigQuery (ODBC), ejecuta la SQL y vuelca el resultado en "BQ_Resultado".
Public Sub RefrescarDatos()
    Dim cn As Object, rs As Object, ws As Object, sql As String, j As Long
    sql = ConstruirSQL()
    If Left(sql, 2) = "--" Then
        MsgBox "Esta métrica/periodo aún no está mapeada a BigQuery:" & vbLf & vbLf & sql, _
               vbExclamation, "Fase 2"
        Exit Sub
    End If
    On Error GoTo fallo

    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120
    cn.Open BQ_CONN
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, 1, 1                       ' adOpenKeyset, adLockReadOnly

    ' Hoja de aterrizaje (se crea si no existe).
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(LANDING)
    On Error GoTo fallo
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = LANDING
    End If
    Application.EnableEvents = False
    ws.Cells.ClearContents
    For j = 0 To rs.Fields.Count - 1            ' cabeceras
        ws.Cells(1, j + 1).Value = rs.Fields(j).Name
    Next j
    If Not rs.EOF Then ws.Range("A2").CopyFromRecordset rs
    rs.Close: cn.Close
    Application.EnableEvents = True

    MsgBox "Datos traídos de BigQuery a la hoja '" & LANDING & "'.", vbInformation, "Fase 2"
    Exit Sub

fallo:
    Application.EnableEvents = True
    MsgBox "No se pudo conectar/consultar BigQuery:" & vbLf & Err.Description & _
           vbLf & vbLf & "SQL:" & vbLf & sql, vbExclamation, "Fase 2"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub
