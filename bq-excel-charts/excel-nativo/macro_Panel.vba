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
' Con los parámetros del Panel construye la SQL de BigQuery, la ejecuta por ODBC
' (ADODB) y vuelca los datos en la hoja "Datos"; luego redibuja el gráfico.
'
' INSTALACIÓN: Insertar -> Módulo (uno NUEVO, distinto al de la PARTE B) y pega
' todo esto. Rellena el bloque CONFIG con tu entorno.
'
' Botones sugeridos (Insertar -> Forma -> Asignar macro):
'   · "Ver SQL"          -> VerSQL         (solo muestra la consulta)
'   · "Traer de BigQuery"-> RefrescarDatos (conecta y actualiza los datos)
'
' NOTA: se asume una tabla en formato largo con las MISMAS columnas que la hoja
' "Datos" (entidad, tipo_activo, metrica, eje_tipo, eje_valor, serie, valor) y
' una columna de fecha para el periodo. Si tu modelo es distinto, ajusta el
' bloque CONFIG y/o la función ConstruirSQL (por ejemplo, añadiendo un GROUP BY).
' ---------------------------------------------------------------------------

' ==== CONFIG (rellena con tu entorno de BigQuery) ====
Private Const BQ_CONN As String = "DSN=BigQuery;"          ' DSN ODBC ya configurado (o cadena Driver={...};...)
Private Const BQ_TABLA As String = "`proyecto.dataset.hechos_metricas`"
Private Const COL_ENTIDAD As String = "entidad"
Private Const COL_TIPOACTIVO As String = "tipo_activo"
Private Const COL_METRICA As String = "metrica"
Private Const COL_DIMENSION As String = "eje_tipo"        ' distingue Mensual/Trimestral/.../Geografia...
Private Const COL_EJEVALOR As String = "eje_valor"        ' etiqueta del eje X (debe casar con las listas Cat_*)
Private Const COL_SERIE As String = "serie"               ' 'Cartera' / 'Benchmark'
Private Const COL_VALOR As String = "valor"
Private Const COL_FECHA As String = "fecha"               ' para acotar el periodo
' =====================================================

' Duplica comillas simples para evitar romper la cadena SQL.
Private Function Esc(ByVal s As String) As String
    Esc = Replace(CStr(s), "'", "''")
End Function

' Lista de entidades seleccionadas (B4/B5/B6), saltando vacías y "(ninguna)".
Private Function ListaEntidades(ws As Worksheet) As String
    Dim celda As Variant, v As String, out As String
    For Each celda In Array("B4", "B5", "B6")
        v = Trim(CStr(ws.Range(celda).Value))
        If v <> "" And v <> "(ninguna)" Then
            If Len(out) > 0 Then out = out & ", "
            out = out & "'" & Esc(v) & "'"
        End If
    Next celda
    ListaEntidades = out
End Function

' Traduce el periodo (B11) a un predicado de fecha para BigQuery.
Private Function FiltroPeriodo(ByVal p As String) As String
    Dim n As Long
    p = Trim(UCase(p))
    Select Case p
        Case "MTD": FiltroPeriodo = COL_FECHA & " >= DATE_TRUNC(CURRENT_DATE(), MONTH)"
        Case "YTD": FiltroPeriodo = COL_FECHA & " >= DATE_TRUNC(CURRENT_DATE(), YEAR)"
        Case Else
            If Len(p) >= 2 And IsNumeric(Left(p, Len(p) - 1)) Then
                n = CLng(Left(p, Len(p) - 1))
                If Right(p, 1) = "M" Then
                    FiltroPeriodo = COL_FECHA & " >= DATE_SUB(CURRENT_DATE(), INTERVAL " & n & " MONTH)"
                ElseIf Right(p, 1) = "A" Then
                    FiltroPeriodo = COL_FECHA & " >= DATE_SUB(CURRENT_DATE(), INTERVAL " & n & " YEAR)"
                End If
            End If
    End Select
End Function

' Construye la SQL a partir de los parámetros del Panel.
Public Function ConstruirSQL() As String
    Dim ws As Worksheet, sql As String, ents As String, wF As String
    Dim met As String, dimen As String, filtro As String, periodo As String
    Set ws = ThisWorkbook.Sheets("Panel")
    met = Esc(ws.Range("B8").Value)
    dimen = Esc(ws.Range("B9").Value)
    filtro = Trim(CStr(ws.Range("B10").Value))
    periodo = Trim(CStr(ws.Range("B11").Value))
    ents = ListaEntidades(ws)

    sql = "SELECT " & COL_ENTIDAD & ", " & COL_TIPOACTIVO & ", " & COL_METRICA & ", " & _
          COL_DIMENSION & ", " & COL_EJEVALOR & ", " & COL_SERIE & ", " & COL_VALOR & vbLf & _
          "FROM " & BQ_TABLA & vbLf & _
          "WHERE " & COL_METRICA & " = '" & met & "'" & vbLf & _
          "  AND " & COL_DIMENSION & " = '" & dimen & "'"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND " & COL_ENTIDAD & " IN (" & ents & ")"
    If filtro <> "" And filtro <> "Todos" Then
        sql = sql & vbLf & "  AND " & COL_TIPOACTIVO & " = '" & Esc(filtro) & "'"
    End If
    wF = FiltroPeriodo(periodo)
    If Len(wF) > 0 Then sql = sql & vbLf & "  AND " & wF
    sql = sql & vbLf & "ORDER BY " & COL_SERIE & ", " & COL_EJEVALOR
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

' Conecta a BigQuery (ODBC), ejecuta la SQL y vuelca los datos en "Datos".
Public Sub RefrescarDatos()
    Dim cn As Object, rs As Object, wsD As Worksheet, sql As String
    sql = ConstruirSQL()
    On Error GoTo fallo

    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120
    cn.Open BQ_CONN
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, 1, 1                       ' adOpenKeyset, adLockReadOnly

    Set wsD = ThisWorkbook.Sheets("Datos")
    Application.EnableEvents = False
    wsD.Range("A2:G" & wsD.Rows.Count).ClearContents   ' limpia datos previos (deja cabecera)
    If Not rs.EOF Then wsD.Range("A2").CopyFromRecordset rs
    rs.Close: cn.Close
    Application.EnableEvents = True

    Application.Calculate
    ' Redibuja el gráfico re-disparando el evento del Panel.
    With ThisWorkbook.Sheets("Panel")
        .Activate
        .Range("B14").Value = .Range("B14").Value
    End With
    MsgBox "Datos actualizados desde BigQuery.", vbInformation, "Fase 2"
    Exit Sub

fallo:
    Application.EnableEvents = True
    MsgBox "No se pudo conectar/consultar BigQuery:" & vbLf & Err.Description & _
           vbLf & vbLf & "SQL:" & vbLf & sql, vbExclamation, "Fase 2"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub
