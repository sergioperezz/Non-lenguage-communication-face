' ============================================================================
'  Macro del PANEL GENÉRICO (Fase 1) + copia a PowerPoint como EMF
'  Modelo: Tipo entidad -> Entidad(1..3) -> Grupo -> Métrica -> Dimensión -> ...
'
'  INSTALACIÓN (una vez):
'   1) PARTE A -> clic derecho en la pestaña "Panel" -> "Ver código" y pega ahí.
'   2) PARTE B -> Insertar -> Módulo, y pega ahí.
'   3) Guarda como .xlsm. (Opcional: botón con la macro "CopiarAPowerPoint".)
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
        LimpiarSiInvalida "B5", "Ent_" & Me.Range("B3").Value
        LimpiarSiInvalida "B6", "Ent_" & Me.Range("B3").Value
    End If
    ' Al cambiar el GRUPO de métrica: ajustar la métrica.
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        AjustarSeleccion "B8", "Grupo_" & Me.Range("B7").Value
    End If

    If Not Intersect(Target, Me.Range("B3:B14")) Is Nothing Then AplicarGrafico

    ' --- Fase 2 (BigQuery por Power Query/ODBC): descomenta para refrescar datos ---
    ' If Not Intersect(Target, Me.Range("B3:B11")) Is Nothing Then ThisWorkbook.RefreshAll

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

' Como AjustarSeleccion pero para celdas opcionales: si no es válida, la vacía.
Private Sub LimpiarSiInvalida(ByVal celda As String, ByVal nombreLista As String)
    Dim rng As Range, c As Range, valido As Boolean
    If Me.Range(celda).Value = "" Then Exit Sub
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nombreLista).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If c.Value = Me.Range(celda).Value Then valido = True
    Next c
    If Not valido Then Me.Range(celda).Value = ""
End Sub

' Añade una serie de entidad (columna colLetter) si su slot no está vacío.
Private Sub AddEnt(ByVal ch As Chart, ByVal slotCell As String, ByVal colLetter As String, ByVal lastRow As Long)
    Dim s As Series
    If Trim(Me.Range(slotCell).Value) = "" Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "=Panel!$" & colLetter & "$2"          ' cabecera = nombre de la entidad
    s.Values = "=Panel!$" & colLetter & "$3:$" & colLetter & "$" & lastRow
    s.XValues = "=Panel!$D$3:$D$" & lastRow
End Sub

' Reconstruye series (multi-entidad + benchmark), tipo, estilo, título y ejes.
Private Sub AplicarGrafico()
    Dim ch As Chart, s As Series, conBench As Boolean, tipo As String
    Dim benchLinea As Boolean, vis As Long, lastRow As Long, benchIdx As Long
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    conBench = (Me.Range("B12").Value = "Con benchmark")
    benchLinea = (LCase(Trim(Me.Range("B13").Value)) Like "l*")
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
        Case Else:                ch.ChartType = xlColumnClustered  ' "Columnas"
    End Select

    ' Estilo del benchmark (B13): su serie se dibuja como Barras o Líneas.
    If benchIdx > 0 Then
        On Error Resume Next
        If benchLinea Then
            ch.FullSeriesCollection(benchIdx).ChartType = xlLineMarkers
        Else
            ch.FullSeriesCollection(benchIdx).ChartType = xlColumnClustered
        End If
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
