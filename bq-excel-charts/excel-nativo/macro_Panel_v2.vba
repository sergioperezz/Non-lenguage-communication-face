' ============================================================================
'  Macro del PANEL GENÉRICO (Fase 1, v2) + copia a PowerPoint como EMF
'
'  INSTALACIÓN (una vez):
'   1) PARTE A -> clic derecho en la pestaña "Panel" -> "Ver código" y pega ahí.
'   2) PARTE B -> Insertar -> Módulo, y pega ahí.
'   3) Guarda como .xlsm. (Opcional: botón con la macro "CopiarAPowerPoint".)
'
'  Qué hace al cambiar cualquiera de los 5 desplegables (B3:B7):
'   - Mantiene coherente la cascada (métrica y eje X válidos para el activo).
'   - Aplica el tipo de gráfico elegido.
'   - Si es "Columnas" con "Cartera + Benchmark": convierte el benchmark en LÍNEA
'     (gráfico combinado barras + línea), que es tu caso típico.
'   Los datos ya se recalculan solos por fórmulas, así que el gráfico se actualiza.
' ============================================================================


' =====================  PARTE A: en la hoja "Panel"  ========================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Salir
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Al cambiar el tipo de activo, validar métrica y dimensión (eje X).
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        AjustarSeleccion "B4", Me.Range("B3").Value          ' métrica
        AjustarSeleccion "B5", "Eje_" & Me.Range("B3").Value ' dimensión
    End If

    ' Cualquier cambio en B3:B8 -> reconstruye la representación del gráfico.
    If Not Intersect(Target, Me.Range("B3:B8")) Is Nothing Then AplicarGrafico

    ' --- Fase 2 (BigQuery por Power Query/ODBC): descomenta para refrescar datos ---
    ' If Not Intersect(Target, Me.Range("B3:B6")) Is Nothing Then ThisWorkbook.RefreshAll

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Si el valor de "celda" no está en el rango con nombre "nombreLista",
' lo sustituye por el primer elemento válido. (Reutilizable: métrica y eje X.)
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

' Aplica tipo de gráfico y, si procede, el combo barras + línea (benchmark).
Private Sub AplicarGrafico()
    Dim ch As Chart
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    ' 1) Tipo base para todas las series (lo elige B8).
    Select Case LCase(Trim(Me.Range("B8").Value))
        Case "barras":            ch.ChartType = xlBarClustered
        Case "líneas", "lineas":  ch.ChartType = xlLineMarkers
        Case "área", "area":      ch.ChartType = xlArea
        Case "circular":          ch.ChartType = xlPie
        Case "anillo":            ch.ChartType = xlDoughnut
        Case "radar":             ch.ChartType = xlRadarMarkers
        Case Else:                ch.ChartType = xlColumnClustered  ' "Columnas"
    End Select

    ' 2) Combo: en columnas con ambas series (B7), el benchmark se dibuja como
    '    LÍNEA. (La serie 2 es "Benchmark" por el orden de las columnas E/F.)
    If LCase(Trim(Me.Range("B8").Value)) = "columnas" _
       And Me.Range("B7").Value = "Cartera + Benchmark" Then
        On Error Resume Next
        ch.FullSeriesCollection(2).ChartType = xlLineMarkers
        On Error GoTo 0
    End If
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
