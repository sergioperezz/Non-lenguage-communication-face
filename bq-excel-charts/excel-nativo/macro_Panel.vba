' ============================================================================
'  Macro del panel interactivo (Fase 1 + copia a PowerPoint como EMF)
'
'  INSTALACIÓN (una vez):
'  1) PARTE A  -> clic derecho en la pestaña "Panel" -> "Ver código" y pega ahí.
'  2) PARTE B  -> menú Insertar -> Módulo, y pega ahí (módulo estándar).
'  3) Guarda como .xlsm (libro habilitado para macros).
'  4) (Opcional) Inserta una forma/botón y asígnale la macro "CopiarAPowerPoint".
' ============================================================================


' =====================  PARTE A: en la hoja "Panel"  ========================
' (Worksheet_Change: reacciona al cambiar los desplegables.)

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Salir
    Application.EnableEvents = False
    Application.ScreenUpdating = False          ' sin parpadeo -> cambio "instantáneo"

    ' Al cambiar el tipo de activo, asegurar que la métrica siga siendo válida.
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then AjustarMetrica

    ' Al cambiar cualquier parámetro, aplicar el tipo de gráfico elegido.
    If Not Intersect(Target, Me.Range("B3:B5")) Is Nothing Then AplicarTipoGrafico

    ' --- Fase 2 (BigQuery por Power Query/ODBC): descomenta para refrescar datos ---
    ' If Not Intersect(Target, Me.Range("B3:B4")) Is Nothing Then ThisWorkbook.RefreshAll

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Si la métrica no pertenece al tipo de activo, pone la primera válida.
Private Sub AjustarMetrica()
    Dim tipo As String, met As String, rng As Range, c As Range, valida As Boolean
    tipo = Me.Range("B3").Value
    met = Me.Range("B4").Value
    On Error Resume Next
    Set rng = ThisWorkbook.Names(tipo).RefersToRange   ' rango con nombre RF / RV
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If c.Value = met Then valida = True
    Next c
    If Not valida Then Me.Range("B4").Value = rng.Cells(1, 1).Value
End Sub

' Cambia el tipo de gráfico según B5. Cambiar .ChartType de UN solo gráfico es
' lo más rápido (no se reconstruye nada).
Private Sub AplicarTipoGrafico()
    Dim ch As Chart
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    Select Case LCase(Trim(Me.Range("B5").Value))
        Case "barras":            ch.ChartType = xlBarClustered      ' horizontales
        Case "líneas", "lineas":  ch.ChartType = xlLineMarkers
        Case "área", "area":      ch.ChartType = xlArea
        Case "circular":          ch.ChartType = xlPie
        Case "anillo":            ch.ChartType = xlDoughnut
        Case "radar":             ch.ChartType = xlRadarMarkers
        Case Else:                ch.ChartType = xlColumnClustered   ' "Columnas" (vertical)
    End Select
End Sub


' =================  PARTE B: en un Módulo estándar  =========================
' Copia el gráfico del panel a PowerPoint como EMF (vectorial -> no se rompe,
' nítido a cualquier tamaño, sin vínculo al Excel). Asígnala a un botón.

Public Sub CopiarAPowerPoint()
    Dim ch As ChartObject
    On Error Resume Next
    Set ch = ThisWorkbook.Sheets("Panel").ChartObjects(1)
    On Error GoTo 0
    If ch Is Nothing Then
        MsgBox "No encuentro el gráfico en la hoja 'Panel'.", vbExclamation
        Exit Sub
    End If

    ' 1) Copiar el gráfico como imagen VECTORIAL (en Windows, xlPicture = metafile/EMF).
    ch.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' 2) Conectar con PowerPoint (lo abre si no está abierto). Late binding: sin
    '    necesidad de activar referencias.
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

    ' Usa la diapositiva activa; si no hay, crea una en blanco (layout 12 = ppLayoutBlank).
    On Error Resume Next
    Set sld = ppt.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        If pres.Slides.Count = 0 Then
            Set sld = pres.Slides.Add(1, 12)
        Else
            Set sld = pres.Slides(pres.Slides.Count)
        End If
    End If

    ' 3) Pegar como EMF (DataType 2 = ppPasteEnhancedMetafile) -> imagen vectorial.
    Dim shp As Object
    Set shp = sld.Shapes.PasteSpecial(DataType:=2)

    ' 4) Centrar en la diapositiva (opcional).
    On Error Resume Next
    shp.Left = (pres.PageSetup.SlideWidth - shp.Width) / 2
    shp.Top = (pres.PageSetup.SlideHeight - shp.Height) / 2
    On Error GoTo 0
End Sub
