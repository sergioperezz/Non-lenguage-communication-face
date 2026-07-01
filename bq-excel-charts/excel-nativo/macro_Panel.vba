' ============================================================================
'  Macro del PANEL GENÉRICO (Fase 1) + copia a PowerPoint como EMF
'  Modelo: Tipo entidad -> Entidad -> Grupo -> Métrica -> Dimensión -> ...
'
'  INSTALACIÓN (una vez):
'   1) PARTE A -> clic derecho en la pestaña "Panel" -> "Ver código" y pega ahí.
'   2) PARTE B -> Insertar -> Módulo, y pega ahí.
'   3) Guarda como .xlsm. (Opcional: botón con la macro "CopiarAPowerPoint".)
'
'  Controles: B3 Tipo entidad · B4 Entidad · B5 Grupo · B6 Métrica · B7 Dimensión
'   · B8 Filtro tipo activo · B9 Periodo · B10 Benchmark (Con/Sin)
'   · B11 Estilo benchmark (Barras/Líneas) · B12 Tipo de gráfico
' ============================================================================


' =====================  PARTE A: en la hoja "Panel"  ========================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Salir
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Al cambiar el TIPO de entidad, asegurar que la entidad sea válida.
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        AjustarSeleccion "B4", "Ent_" & Me.Range("B3").Value
    End If
    ' Al cambiar el GRUPO de métrica, asegurar que la métrica sea válida.
    If Not Intersect(Target, Me.Range("B5")) Is Nothing Then
        AjustarSeleccion "B6", "Grupo_" & Me.Range("B5").Value
    End If

    ' Cualquier cambio en B3:B12 -> reconstruye el gráfico.
    If Not Intersect(Target, Me.Range("B3:B12")) Is Nothing Then AplicarGrafico

    ' --- Fase 2 (BigQuery por Power Query/ODBC): descomenta para refrescar datos ---
    ' If Not Intersect(Target, Me.Range("B3:B9")) Is Nothing Then ThisWorkbook.RefreshAll

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Al mostrar la hoja, renderiza el gráfico según los desplegables actuales.
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    AplicarGrafico
    Application.ScreenUpdating = True
End Sub

' Si el valor de "celda" no está en el rango con nombre "nombreLista",
' lo sustituye por el primer elemento válido.
Private Sub AjustarSeleccion(ByVal celda As String, ByVal nombreLista As String)
    Dim rng As Range, c As Range, actual As String, valido As Boolean
    actual = Me.Range(celda).Value
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nombreLista).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If c.Value = actual Then valido = True
    Next c
    If Not valido Then Me.Range(celda).Value = rng.Cells(1, 1).Value
End Sub

' Reconstruye las series (rango dinámico), aplica tipo, estilo benchmark, título y ejes.
Private Sub AplicarGrafico()
    Dim ch As Chart, s As Series, conBench As Boolean, tipo As String
    Dim benchLinea As Boolean, vis As Long, lastRow As Long
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    conBench = (Me.Range("B10").Value = "Con benchmark")
    benchLinea = (LCase(Trim(Me.Range("B11").Value)) Like "l*")   ' "Líneas"
    tipo = LCase(Trim(Me.Range("B12").Value))

    ' Nº de categorías visibles (B14) -> última fila de datos del gráfico.
    vis = Me.Range("B14").Value
    If vis < 1 Then vis = 1
    lastRow = 2 + vis

    ' 1) Reconstruir series con rango dinámico (solo categorías activas).
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    Set s = ch.SeriesCollection.NewSeries          ' Cartera
    s.Name = "=Panel!$E$2"
    s.Values = "=Panel!$E$3:$E$" & lastRow
    s.XValues = "=Panel!$D$3:$D$" & lastRow
    If conBench Then
        Set s = ch.SeriesCollection.NewSeries      ' Benchmark
        s.Name = "=Panel!$F$2"
        s.Values = "=Panel!$F$3:$F$" & lastRow
        s.XValues = "=Panel!$D$3:$D$" & lastRow
    End If

    ' 2) Tipo de gráfico base (B12).
    Select Case tipo
        Case "barras":            ch.ChartType = xlBarClustered
        Case "líneas", "lineas":  ch.ChartType = xlLineMarkers
        Case "área", "area":      ch.ChartType = xlArea
        Case "circular":          ch.ChartType = xlPie
        Case "anillo":            ch.ChartType = xlDoughnut
        Case "radar":             ch.ChartType = xlRadarMarkers
        Case Else:                ch.ChartType = xlColumnClustered  ' "Columnas"
    End Select

    ' 3) Estilo del benchmark (B11): la serie 2 se dibuja como Barras o Líneas.
    If conBench And ch.SeriesCollection.Count >= 2 Then
        On Error Resume Next
        If benchLinea Then
            ch.FullSeriesCollection(2).ChartType = xlLineMarkers
        Else
            ch.FullSeriesCollection(2).ChartType = xlColumnClustered
        End If
        On Error GoTo 0
    End If

    ' 4) Título (A17) y títulos de eje (Y = métrica B6, X = dimensión B7).
    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = Me.Range("A17").Value
    If tipo <> "circular" And tipo <> "anillo" And tipo <> "radar" Then
        ch.Axes(xlValue).HasTitle = True
        ch.Axes(xlValue).AxisTitle.Text = Me.Range("B6").Value
        ch.Axes(xlCategory).HasTitle = True
        ch.Axes(xlCategory).AxisTitle.Text = Me.Range("B7").Value
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
