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
' (esquema en estrella), la ejecuta por ODBC (ADODB) y vuelca los datos en una
' hoja de aterrizaje "BQ_Resultado".
'
' INSTALACIÓN: Insertar -> Módulo (uno NUEVO, distinto al de la PARTE B) y pega
' todo esto. Rellena BQ_CONN y BQ_DATASET con tu entorno.
'
' Botones sugeridos (Insertar -> Forma -> Asignar macro):
'   · "Ver SQL"          -> VerSQL         (solo muestra la consulta)
'   · "Traer de BigQuery"-> RefrescarDatos (conecta y trae los datos)
'
' MAPEO (según el diccionario de datos homologado):
'   · Rentabilidad / Rentab. acum. -> CAM_TX_PERFORMANCE_FIGURES_PD, columna
'     TWR_<periodo> (MTD/QTD/YTD/1M/1Y/3Y/5Y); benchmark = TWR_<periodo>_BMK.
'   · Volatilidad -> VOL_1Y_260 ; Beta -> BETA (misma tabla, valor por portfolio).
'   · Identidad y tipo: CAM_TM_PORTFOLIOS_PD (PORTFOLIO_NAME, PTF_PORTFOLIO_TYPE).
'   · Duración/TIR/Spread -> CAM_TX_RISK_FIG_AGG_PD (PK_VARIABLE_TARGET +
'     PK_CRITERIO_AGREGACION/PK_ETIQUETA_AGREGACION)  [pendiente de mapear].
'   · Peso/Composición -> composición de CAM_TM_PORTFOLIOS_PD y benchmark en
'     CAM_TX_BENCHMARK_COMP_PD (COMPONENT, WEIGHT)          [pendiente de mapear].
' NOTA: la consulta es de CORTE TRANSVERSAL (un valor por portfolio en la última
' fecha), que es como el DWH guarda las rentabilidades rolling.
' ---------------------------------------------------------------------------

' ==== CONFIG (rellena con tu entorno de BigQuery) ====
Private Const BQ_CONN As String = "DSN=BigQuery;"                 ' DSN ODBC (o cadena Driver={...};...)
Private Const BQ_DATASET As String = "proyecto.dataset"          ' proyecto.dataset de BigQuery (sin backticks)
Private Const LANDING As String = "BQ_Resultado"                 ' hoja donde se vuelcan los datos
' Tablas del modelo homologado:
Private Const T_PERF As String = "CAM_TX_PERFORMANCE_FIGURES_PD"
Private Const T_PORT As String = "CAM_TM_PORTFOLIOS_PD"
' Columnas de identidad de portfolio:
Private Const P_ID As String = "PK_PORTFOLIO_ID"
Private Const P_NAME As String = "PORTFOLIO_NAME"
Private Const P_TYPE As String = "PTF_PORTFOLIO_TYPE"
Private Const P_FECHA As String = "PK_FECHA_DATOS"
' =====================================================

' Duplica comillas simples para evitar romper la cadena SQL.
Private Function Esc(ByVal s As String) As String
    Esc = Replace(CStr(s), "'", "''")
End Function

' Nombre de tabla cualificado: `proyecto.dataset.TABLA`
Private Function Tbl(ByVal t As String) As String
    Tbl = "`" & BQ_DATASET & "." & t & "`"
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

' Métrica del Panel -> tabla + columna del fondo + columna del benchmark del DWH.
' Devuelve True si la métrica está mapeada.
Private Function MapMetrica(ByVal met As String, ByVal per As String, _
        ByRef colVal As String, ByRef colBmk As String) As Boolean
    Dim suf As String: suf = SufijoPeriodo(per)
    colVal = "": colBmk = ""
    Select Case met
        Case "Rentabilidad", "Rentab. acum."
            If Len(suf) = 0 Then Exit Function            ' periodo sin columna en el DWH
            colVal = "TWR_" & suf
            colBmk = "TWR_" & suf & "_BMK"
            MapMetrica = True
        Case "Volatilidad"
            colVal = "VOL_1Y_260": MapMetrica = True
        Case "Beta"
            colVal = "BETA": MapMetrica = True
    End Select
End Function

' Construye la SQL a partir de los parámetros del Panel (corte transversal).
Public Function ConstruirSQL() As String
    Dim ws As Worksheet, met As String, per As String, ents As String
    Dim colVal As String, colBmk As String, whereEnt As String, sql As String
    Dim conBmk As Boolean
    Set ws = ThisWorkbook.Sheets("Panel")
    met = Trim(CStr(ws.Range("B8").Value))
    per = Trim(CStr(ws.Range("B11").Value))
    ents = ListaEntidades(ws)
    conBmk = (ws.Range("B12").Value = "Con benchmark")

    If Not MapMetrica(met, per, colVal, colBmk) Then
        ConstruirSQL = _
            "-- Métrica '" & met & "' / periodo '" & per & "': aún no mapeada a BigQuery." & vbLf & _
            "-- Riesgo (Duración/TIR/Spread) -> " & T_PERF & " y CAM_TX_RISK_FIG_AGG_PD" & vbLf & _
            "--   (PK_VARIABLE_TARGET, PK_CRITERIO_AGREGACION, PK_ETIQUETA_AGREGACION)." & vbLf & _
            "-- Peso/Composición -> CAM_TM_PORTFOLIOS_PD (composición) / CAM_TX_BENCHMARK_COMP_PD" & vbLf & _
            "--   (COMPONENT, WEIGHT). Rellena el mapeo cuando definamos criterio/etiqueta."
        Exit Function
    End If

    whereEnt = ""
    If Len(ents) > 0 Then whereEnt = " AND p." & P_ID & " IN (" & ents & ")"

    sql = "SELECT p." & P_NAME & " AS entidad, p." & P_TYPE & " AS tipo_activo," & vbLf & _
          "       '" & Esc(met) & "' AS metrica, CAST(f." & P_FECHA & " AS STRING) AS eje_valor," & vbLf & _
          "       'Cartera' AS serie, f." & colVal & " AS valor" & vbLf & _
          "FROM " & Tbl(T_PERF) & " f" & vbLf & _
          "JOIN " & Tbl(T_PORT) & " p ON p." & P_ID & " = f." & P_ID & vbLf & _
          "WHERE f." & P_FECHA & " = (SELECT MAX(" & P_FECHA & ") FROM " & Tbl(T_PERF) & ")" & whereEnt
    If conBmk And Len(colBmk) > 0 Then
        sql = sql & vbLf & "UNION ALL" & vbLf & _
          "SELECT p." & P_NAME & ", p." & P_TYPE & ", '" & Esc(met) & "', CAST(f." & P_FECHA & " AS STRING)," & vbLf & _
          "       'Benchmark', f." & colBmk & vbLf & _
          "FROM " & Tbl(T_PERF) & " f" & vbLf & _
          "JOIN " & Tbl(T_PORT) & " p ON p." & P_ID & " = f." & P_ID & vbLf & _
          "WHERE f." & P_FECHA & " = (SELECT MAX(" & P_FECHA & ") FROM " & Tbl(T_PERF) & ")" & whereEnt
    End If
    sql = sql & vbLf & "ORDER BY serie, entidad"
    ConstruirSQL = sql
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
