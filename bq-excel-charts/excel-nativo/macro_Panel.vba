' ============================================================================
'  Macro del panel interactivo (Fase 1)
'  Dónde pegarla: en Excel, clic derecho en la pestaña "Panel" -> "Ver código"
'  y pega TODO esto en esa ventana. Luego guarda como .xlsm (libro con macros).
'
'  Qué hace, automáticamente al cambiar un desplegable:
'   - B3 Tipo de activo  -> si la métrica deja de ser válida, pone la primera válida
'   - B5 Tipo de gráfico -> cambia el gráfico entre verticales y horizontales
'  (Los datos y la métrica ya se actualizan solos por fórmulas; esto añade el
'   cambio de tipo de gráfico y la coherencia de la cascada.)
' ============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Salir
    Application.EnableEvents = False

    ' Al cambiar el tipo de activo, asegurar que la métrica sea válida
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        AjustarMetrica
        AplicarTipoGrafico
    End If

    ' Al cambiar la métrica o el tipo de gráfico, refrescar el gráfico
    If Not Intersect(Target, Me.Range("B4:B5")) Is Nothing Then
        AplicarTipoGrafico
    End If

    ' --- Fase 2 (cuando conectes BigQuery por Power Query/ODBC): descomenta ---
    ' If Not Intersect(Target, Me.Range("B3:B4")) Is Nothing Then
    '     ThisWorkbook.RefreshAll
    ' End If

Salir:
    Application.EnableEvents = True
End Sub

' Si la métrica seleccionada no pertenece al tipo de activo, pone la 1ª válida.
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

' Cambia el tipo de gráfico según B5 (Verticales / Horizontales).
Private Sub AplicarTipoGrafico()
    Dim ch As Chart
    On Error Resume Next
    Set ch = Me.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub
    Select Case LCase(Trim(Me.Range("B5").Value))
        Case "horizontales": ch.ChartType = xlBarClustered     ' barras horizontales
        Case Else:           ch.ChartType = xlColumnClustered  ' barras verticales
    End Select
End Sub
