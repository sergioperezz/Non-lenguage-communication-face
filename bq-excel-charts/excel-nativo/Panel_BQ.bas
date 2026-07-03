Attribute VB_Name = "PanelBQ"
' ============================================================================
'  Panel de graficos + Fase 2 (BigQuery) - MODULO UNICO IMPORTABLE (.bas)
'
'  INSTALACION (una vez):
'   1) Abre el .xlsx y guarda como .xlsm (Libro habilitado para macros).
'   2) Alt+F11 -> File -> Import File... -> elige "Panel_BQ.bas".
'   3) Revisa el bloque CONFIG (DSN, proyecto, datasets).
'   4) En la hoja Panel: Insertar -> Formas -> un rectangulo -> clic derecho ->
'      Asignar macro -> "Actualizar"  (el boton que lo hace TODO en Fase 2).
'      Otro boton opcional -> "DibujarGrafico" (redibuja con los datos actuales).
'      En la hoja Tablas: un boton -> "FormatearTablas".
'
'  Todo va en este unico modulo: no hay que pegar nada en las hojas.
'  El comportamiento es por BOTON (no automatico al cambiar un desplegable).
' ============================================================================

Option Explicit

' ====================  CONFIG (entorno BigQuery)  ===========================
Private Const BQ_DSN As String = "Conexion_BQ"
Private Const BQ_CONN As String = "DSN=Conexion_BQ;"
Private Const BQ_PROJECT As String = "go-cam-beg-camd9-camcd9p01-pro"
Private Const DS_PROD As String = "productosdatosdecontratos_ds01"
Private Const DS_OPER As String = "operativafinanciera_ds01"
Private Const DS_MERC As String = "informaciondemercado_ds01"
Private Const F_NAV_GNAV As String = "GNAV"
Private Const F_BENCHMARK As String = "Benchmark 1"
Private Const T_PERF As String = "CAM_TX_PERFORMANCE_FIGURES_PD"
Private Const T_RISK As String = "CAM_TX_RISK_FIG_AGG_PD"
Private Const RISK_COL_FONDOBMK As String = "PK_TIPOGAMAN1"
Private Const T_POS As String = "CAM_TM_PORTFOLIOS_PD"
Private Const T_VALORES As String = "CAM_TM_MSTR_VALORES_PD"
Private Const POS_VALOR As String = "VALUATION_PC"
Private Const SECTOR_COL As String = "CLASSIFICATION_GICS"
Private Const RATING_COL As String = "COMPOSITERATINGSPCOMPOSITE"
Private Const TER_COL As String = "KEYFIGURESTER"
Private Const T_FONDOS As String = "CAM_TM_MSTR_FONDOS_PD"
Private Const TER_FONDO_EXPR As String = "f.COMISION_DE_GESTION_DIRECTA + f.COMISION_DEPOSITARIA_DIRECTA"
' Hoja de carteras (la que copiaras): nombre -> id. Se busca por varios nombres.
Private Const ACTIVOS_SHEET As String = "cartera"          ' nombre principal a buscar
Private Const ACTIVOS_COL_NOMBRE As String = "nombre_elemento"
Private Const ACTIVOS_COL_ID As String = "id_elemento"    ' columna que va a PK_PORTFOLIO_ID
' ===========================================================================

Private Function Panel() As Worksheet
    Set Panel = ThisWorkbook.Sheets("Panel")
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(CStr(s), "'", "''")
End Function

' Normaliza a minusculas SIN acentos (usa codigos de caracter, asi el propio
' codigo no lleva tildes y no depende de la codificacion al importar el .bas).
Private Function Fold(ByVal s As String) As String
    s = LCase(Trim(CStr(s)))
    s = Replace(s, ChrW(225), "a")   ' a con tilde
    s = Replace(s, ChrW(233), "e")   ' e con tilde
    s = Replace(s, ChrW(237), "i")   ' i con tilde
    s = Replace(s, ChrW(243), "o")   ' o con tilde
    s = Replace(s, ChrW(250), "u")   ' u con tilde
    s = Replace(s, ChrW(241), "n")   ' ene
    Fold = s
End Function

Private Function Tbl(ByVal ds As String, ByVal t As String) As String
    Tbl = "`" & BQ_PROJECT & "." & ds & "." & t & "`"
End Function

' Normaliza un nombre para comparar (evita fallos tipicos del copy-paste):
' minusculas, sin acentos, sin espacios duros (nbsp) ni espacios dobles.
Private Function NormNom(ByVal s As String) As String
    s = CStr(s)
    s = Replace(s, ChrW(160), " ")          ' espacio duro (nbsp) -> espacio normal
    s = Replace(s, Chr(9), " ")             ' tabulador -> espacio
    Do While InStr(s, "  ") > 0             ' colapsa espacios dobles
        s = Replace(s, "  ", " ")
    Loop
    NormNom = Fold(s)                        ' Fold ya hace LCase + Trim + quita tildes
End Function

' Nombres de entidad que no se pudieron traducir a id (para avisar al usuario).
Private mAvisoEnt As String

' Devuelve la hoja de carteras exista con el nombre que exista.
Private Function HojaMaestro() As Worksheet
    Dim nombres As Variant, nm As Variant, ws As Worksheet
    nombres = Array(ACTIVOS_SHEET, "cartera", "carteras", "activos")
    For Each nm In nombres
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then Set HojaMaestro = ws: Exit Function
    Next nm
End Function

' Localiza una columna por el texto de su cabecera (fila 1).
Private Function ColPorCabecera(ws As Worksheet, ByVal cab As String) As Long
    Dim c As Long
    For c = 1 To 50
        If LCase(Trim(CStr(ws.Cells(1, c).Value))) = LCase(cab) Then ColPorCabecera = c: Exit Function
    Next c
End Function

' Hoja auxiliar oculta (se crea si no existe).
Private Function HojaAux(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set HojaAux = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If HojaAux Is Nothing Then
        Set HojaAux = ThisWorkbook.Sheets.Add
        HojaAux.Name = nm
    End If
    HojaAux.Visible = xlSheetHidden
End Function

' Pone (o reemplaza) la validacion de lista de una celda apuntando a un rango.
Private Sub PonerDV(celda As Range, ByVal f As String)
    With celda.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=f
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' ---- Traduccion nombre de entidad -> id (id_elemento = pk_portfolio_id) ----
' 1) Hoja "cartera": busca en nombre_elemento y devuelve id_elemento.
'    La comparacion es robusta (NormNom): ignora mayus/acentos/espacios/nbsp.
' 2) Si le pasas ya un id_elemento, lo acepta tal cual.
' 3) Si no esta, cae al rango MapaEntidades (mock de ejemplo).
' Devuelve "" si no lo encuentra (NUNCA el nombre: el nombre no vale como id).
Private Function IdEntidad(ByVal nombre As String) As String
    Dim wa As Worksheet, colN As Long, colI As Long, lastR As Long, r As Long
    Dim clave As String, rng As Range, c As Range
    IdEntidad = ""
    clave = NormNom(nombre)
    If clave = "" Or clave = NormNom("(ninguna)") Then Exit Function

    Set wa = HojaMaestro()
    If Not wa Is Nothing Then
        colN = ColPorCabecera(wa, ACTIVOS_COL_NOMBRE)
        colI = ColPorCabecera(wa, ACTIVOS_COL_ID)
        If colI > 0 Then
            lastR = wa.Cells(wa.Rows.Count, colI).End(xlUp).Row
            ' 1) por nombre_elemento -> devuelve id_elemento
            If colN > 0 Then
                For r = 2 To lastR
                    If NormNom(wa.Cells(r, colN).Value) = clave Then
                        IdEntidad = Trim(CStr(wa.Cells(r, colI).Value)): Exit Function
                    End If
                Next r
            End If
            ' 2) por si ya te pasan directamente el id_elemento
            For r = 2 To lastR
                If NormNom(wa.Cells(r, colI).Value) = clave Then
                    IdEntidad = Trim(CStr(wa.Cells(r, colI).Value)): Exit Function
                End If
            Next r
        End If
    End If

    ' 3) fallback: mapa de ejemplo (solo para los datos ficticios)
    On Error Resume Next
    Set rng = ThisWorkbook.Names("MapaEntidades").RefersToRange
    On Error GoTo 0
    If Not rng Is Nothing Then
        For Each c In rng.Columns(1).Cells
            If NormNom(c.Value) = clave Then IdEntidad = Trim(CStr(c.Offset(0, 1).Value)): Exit Function
        Next c
    End If
End Function

' Construye la lista de ids para el IN(...) SOLO con ids reales (id_elemento).
' Si un nombre no se encuentra en 'cartera', se ignora y se apunta en mAvisoEnt
' (NUNCA se mete el nombre en PK_PORTFOLIO_ID: eso devolveria 0 filas).
Private Function ListaEntidades(ws As Worksheet) As String
    Dim celda As Variant, v As String, idp As String, out As String
    mAvisoEnt = ""
    For Each celda In Array("B4", "B5", "B6")
        v = Trim(CStr(ws.Range(celda).Value))
        If v <> "" And v <> "(ninguna)" Then
            idp = IdEntidad(v)
            If Len(idp) = 0 Then
                mAvisoEnt = mAvisoEnt & vbLf & "   - " & v
            Else
                If Len(out) > 0 Then out = out & ", "
                out = out & "'" & Esc(idp) & "'"
            End If
        End If
    Next celda
    ListaEntidades = out
End Function

Private Function SufijoPeriodo(ByVal p As String) As String
    Select Case UCase(Trim(p))
        Case "MTD": SufijoPeriodo = "MTD"
        Case "YTD": SufijoPeriodo = "YTD"
        Case "1M":  SufijoPeriodo = "1M"
        Case "1A":  SufijoPeriodo = "1Y"
        Case "3A":  SufijoPeriodo = "3Y"
        Case "5A":  SufijoPeriodo = "5Y"
        Case Else:  SufijoPeriodo = ""
    End Select
End Function

Private Function MapMetrica(ByVal met As String, ByVal per As String, _
        ByRef colVal As String, ByRef colBmk As String, ByRef colDif As String) As Boolean
    Dim suf As String: suf = SufijoPeriodo(per)
    colVal = "": colBmk = "": colDif = ""
    Select Case met
        Case "Rentabilidad", "Rentab. acum."
            If Len(suf) = 0 Then Exit Function
            colVal = "TWR_" & suf: colBmk = "TWR_" & suf & "_BMK": colDif = "DIFERENCIAL_" & suf
            MapMetrica = True
        Case "Volatilidad": colVal = "VOL_1Y_260": MapMetrica = True
        Case "Beta":        colVal = "BETA": MapMetrica = True
    End Select
End Function

Private Function DimACriterio(ByVal dimen As String) As String
    Select Case dimen
        Case "Activo":                                      DimACriterio = "AssetType"
        Case "Geografia":                                   DimACriterio = "Geo"
        Case "Divisa":                                      DimACriterio = "FX"
        Case "Mensual", "Trimestral", "Semestral", "Anual": DimACriterio = "Duracion"
        Case Else:                                          DimACriterio = ""
    End Select
End Function

Private Function SQLRiesgo(ws As Worksheet, ByVal variable As String, ByVal ents As String, _
        Optional ByVal critFijo As String = "") As String
    Dim dimen As String, crit As String, sql As String, wVar As String
    If Len(critFijo) > 0 Then
        crit = critFijo: wVar = ""
    Else
        dimen = Trim(CStr(ws.Range("B9").Value))
        crit = DimACriterio(dimen)
        If Len(crit) = 0 Then
            SQLRiesgo = "-- Dimension '" & dimen & "' no existe como desglose en CAM_TX_RISK_FIG_AGG_PD." & vbLf & _
                        "-- Criterios: AssetType (Activo), Geo (Geografia), FX (Divisa), Duracion (total)."
            Exit Function
        End If
        ' 'Duracion' es un criterio-total autocontenido: NO se combina con
        ' PK_VARIABLE_TARGET (esa combinacion devuelve 0 filas).
        If crit = "Duracion" Then
            wVar = ""
        Else
            wVar = "  AND PK_VARIABLE_TARGET = '" & Esc(variable) & "'" & vbLf
        End If
    End If
    sql = "SELECT PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION AS categoria, CAST(VALOR AS FLOAT64) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_RISK) & vbLf & _
          "WHERE PK_CRITERIO_AGREGACION = '" & Esc(crit) & "'" & vbLf & _
          wVar & "  AND " & RISK_COL_FONDOBMK & " = 'FONDO'"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "QUALIFY ROW_NUMBER() OVER (PARTITION BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION" & _
          " ORDER BY PK_FECHA_DATOS DESC) = 1" & vbLf & "ORDER BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION"
    SQLRiesgo = sql
End Function

Private Function DimAClasificacion(ByVal dimen As String) As String
    Select Case dimen
        Case "Sector": DimAClasificacion = "v." & SECTOR_COL
        Case "Rating": DimAClasificacion = "v." & RATING_COL
        Case Else:     DimAClasificacion = ""
    End Select
End Function

Private Function JoinValores() As String
    JoinValores = "JOIN " & Tbl(DS_MERC, T_VALORES) & " v" & vbLf & _
                  "  ON v.PK_SECURITY_IK = p.PK_SECURITY_IK AND v.PK_FECHA_DATOS = p.PK_FECHA_DATOS" & vbLf
End Function

Private Function SQLComposicion(ws As Worksheet, ByVal ents As String) As String
    Dim dimen As String, grp As String, sql As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    grp = DimAClasificacion(dimen)
    If Len(grp) = 0 Then
        SQLComposicion = "-- Composicion por '" & dimen & "': disponible por Sector y Rating."
        Exit Function
    End If
    sql = "SELECT p.PK_PORTFOLIO_ID, " & grp & " AS categoria, CAST(SUM(p." & POS_VALOR & ") AS FLOAT64) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & JoinValores & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID, " & grp & vbLf & "ORDER BY p.PK_PORTFOLIO_ID, valor DESC"
    SQLComposicion = sql
End Function

Private Function SQLSpread(ws As Worksheet, ByVal ents As String) As String
    Dim dimen As String, grp As String, selCat As String, grpBy As String, joinV As String, sql As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    grp = DimAClasificacion(dimen)
    If Len(grp) > 0 Then
        selCat = ", " & grp & " AS categoria": grpBy = ", " & grp: joinV = JoinValores
    End If
    sql = "SELECT p.PK_PORTFOLIO_ID" & selCat & "," & vbLf & _
          "       CAST(SUM(p.SPREAD * p." & POS_VALOR & ") / NULLIF(SUM(p." & POS_VALOR & "), 0) AS FLOAT64) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & joinV & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID" & grpBy & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLSpread = sql
End Function

Private Function SQLTerLookthrough(ws As Worksheet, ByVal ents As String) As String
    Dim sql As String
    sql = "SELECT p.PK_PORTFOLIO_ID," & vbLf & _
          "       CAST(SUM(v." & TER_COL & " * p." & POS_VALOR & ") / NULLIF(SUM(p." & POS_VALOR & "), 0) AS FLOAT64) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & JoinValores & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID" & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLTerLookthrough = sql
End Function

Private Function SQLTerFondo(ws As Worksheet, ByVal ents As String) As String
    Dim sql As String
    sql = "SELECT DISTINCT p.PK_PORTFOLIO_ID, CAST(" & TER_FONDO_EXPR & " AS FLOAT64) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_POS) & " p" & vbLf & _
          "JOIN " & Tbl(DS_PROD, T_FONDOS) & " f ON f.PK_PRODUCTO_DATANOW = p.FK_PRODUCTO_DATANOW" & vbLf & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_POS) & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLTerFondo = sql
End Function

' ---- Construye la SQL segun la metrica del Panel (B8) ----
Public Function ConstruirSQL() As String
    Dim ws As Worksheet, met As String, per As String, ents As String
    Dim colVal As String, colBmk As String, colDif As String, cols As String, sql As String, conBmk As Boolean
    Set ws = Panel()
    met = Trim(CStr(ws.Range("B8").Value))
    per = Trim(CStr(ws.Range("B11").Value))
    ents = ListaEntidades(ws)
    conBmk = (ws.Range("B12").Value = "Con benchmark")

    ' Si se han elegido entidades pero NINGUNA tiene id en 'cartera', no lanzamos
    ' una query sin filtro (traeria toda la tabla). Avisamos que revisen los nombres.
    If Len(ents) = 0 And Len(mAvisoEnt) > 0 Then
        ConstruirSQL = "-- No encuentro el id (pk_portfolio_id) de estas entidades en la hoja 'cartera':" & _
            mAvisoEnt & vbLf & _
            "-- Revisa que B4/B5/B6 coincidan con la columna 'nombre_elemento'." & vbLf & _
            "-- Recuerda: id_elemento = pk_portfolio_id."
        Exit Function
    End If

    If InStr(Fold(met), "duraci") = 1 Then ConstruirSQL = SQLRiesgo(ws, met, ents): Exit Function
    If met = "TIR" Then ConstruirSQL = SQLRiesgo(ws, "", ents, "TIR"): Exit Function
    If met = "Peso" Then ConstruirSQL = SQLComposicion(ws, ents): Exit Function
    If met = "Spread" Then ConstruirSQL = SQLSpread(ws, ents): Exit Function
    If met = "TER" Then ConstruirSQL = SQLTerFondo(ws, ents): Exit Function
    If met = "TER Look-through" Then ConstruirSQL = SQLTerLookthrough(ws, ents): Exit Function

    If Not MapMetrica(met, per, colVal, colBmk, colDif) Then
        ConstruirSQL = "-- Metrica '" & met & "' / periodo '" & per & "': no mapeada." & vbLf & _
                       "-- PER / DividendYield / Liquidez: sin fuente en el diccionario."
        Exit Function
    End If

    cols = "PK_PORTFOLIO_ID, CAST(" & colVal & " AS FLOAT64) AS valor"
    If conBmk And Len(colBmk) > 0 Then cols = cols & ", CAST(" & colBmk & " AS FLOAT64) AS valor_bmk"
    If conBmk And Len(colDif) > 0 Then cols = cols & ", CAST(" & colDif & " AS FLOAT64) AS diferencial"
    sql = "SELECT " & cols & vbLf & _
          "FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
          "WHERE PK_NAV_GNAV = '" & F_NAV_GNAV & "'" & vbLf & _
          "  AND BENCHMARK = '" & F_BENCHMARK & "'"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "QUALIFY ROW_NUMBER() OVER (PARTITION BY PK_PORTFOLIO_ID ORDER BY PK_FECHA_DATOS DESC) = 1" & _
          vbLf & "ORDER BY PK_PORTFOLIO_ID"
    ConstruirSQL = sql
End Function

Public Function ConstruirM() As String
    Dim sql As String: sql = ConstruirSQL()
    If Left(sql, 2) = "--" Then ConstruirM = sql: Exit Function
    ConstruirM = "let" & vbLf & "    Origen = Odbc.Query(""dsn=" & BQ_DSN & """, """ & _
        Replace(sql, vbLf, " ") & """)" & vbLf & "in" & vbLf & "    Origen"
End Function

Public Sub ActualizarSQL()
    On Error Resume Next
    Panel().Range("A41").Value = ConstruirSQL()
End Sub

Public Sub VerSQL()
    ActualizarSQL
    MsgBox ConstruirSQL(), vbInformation, "SQL para los parametros actuales"
End Sub

Public Sub VerM()
    MsgBox ConstruirM(), vbInformation, "Power Query (M)"
End Sub

' =======================  INSTALADOR DE BOTONES  ===========================
' Ejecuta este macro UNA vez (Alt+F8 -> InstalarBotones) y crea los botones en
' las hojas Panel y Tablas con sus macros ya asignadas.
Public Sub InstalarBotones()
    Dim ws As Worksheet
    Set ws = Panel()
    BorrarBotones ws
    CrearBoton ws, "A20", "Cargar carteras (segun B3)", "CargarCarteras"
    CrearBoton ws, "A22", "> Actualizar (BigQuery)", "Actualizar"
    CrearBoton ws, "A24", "Dibujar (ejemplo)", "DibujarGrafico"
    CrearBoton ws, "A26", "Ver SQL", "VerSQL"
    CrearBoton ws, "A28", "> A PowerPoint (Fase 3)", "CopiarAPowerPoint"

    On Error Resume Next
    Dim wt As Worksheet: Set wt = ThisWorkbook.Sheets("Tablas")
    On Error GoTo 0
    If Not wt Is Nothing Then
        BorrarBotones wt
        CrearBoton wt, "D3", "Formatear tabla", "FormatearTablas"
    End If
    MsgBox "Botones creados en 'Panel' y 'Tablas'.", vbInformation, "Instalacion"
End Sub

Private Sub BorrarBotones(ws As Worksheet)
    Dim i As Long
    For i = ws.Buttons.Count To 1 Step -1
        If Left(ws.Buttons(i).Name, 4) = "btn_" Then ws.Buttons(i).Delete
    Next i
End Sub

Private Sub CrearBoton(ws As Worksheet, ByVal ancla As String, ByVal cap As String, ByVal macro As String)
    Dim b As Button, r As Range
    Set r = ws.Range(ancla)
    Set b = ws.Buttons.Add(r.Left, r.Top, 150, 26)
    b.Caption = cap
    b.OnAction = macro
    b.Name = "btn_" & macro
End Sub

' =======================  CARGAR CARTERAS EN LOS DESPLEGABLES  =============
' Lee la hoja de carteras, filtra por el tipo elegido (B3) y rellena los
' desplegables de Entidad (B4/B5/B6) con esos nombres.
Public Sub CargarCarteras()
    Dim ws As Worksheet, wm As Worksheet, we As Worksheet
    Dim colN As Long, colT As Long, tipoSel As String
    Dim r As Long, n As Long, lastR As Long
    Set ws = Panel()
    Set wm = HojaMaestro()
    If wm Is Nothing Then MsgBox "No encuentro la hoja de carteras (cartera/carteras/activos).", vbExclamation: Exit Sub
    colN = ColPorCabecera(wm, ACTIVOS_COL_NOMBRE)
    colT = ColPorCabecera(wm, "tipo_elemento")
    If colN = 0 Then MsgBox "La hoja de carteras no tiene la columna 'nombre_elemento'.", vbExclamation: Exit Sub

    tipoSel = LCase(Trim(CStr(ws.Range("B3").Value)))   ' Fondo/Cartera/Indice -> minusculas
    Set we = HojaAux("_Ent")
    we.Cells.ClearContents
    we.Cells(1, 1).Value = "(ninguna)"
    n = 1
    lastR = wm.Cells(wm.Rows.Count, colN).End(xlUp).Row
    Application.ScreenUpdating = False
    For r = 2 To lastR
        If colT = 0 Or LCase(Trim(CStr(wm.Cells(r, colT).Value))) = tipoSel Then
            n = n + 1
            we.Cells(n, 1).Value = wm.Cells(r, colN).Value
        End If
    Next r
    Application.ScreenUpdating = True

    If n < 2 Then MsgBox "No hay carteras de tipo '" & ws.Range("B3").Value & "' en la hoja.", vbExclamation: Exit Sub
    Application.EnableEvents = False
    PonerDV ws.Range("B4"), "=_Ent!$A$2:$A$" & n           ' Entidad 1: solo nombres
    PonerDV ws.Range("B5"), "=_Ent!$A$1:$A$" & n           ' Entidad 2/3: incluye "(ninguna)"
    PonerDV ws.Range("B6"), "=_Ent!$A$1:$A$" & n
    ws.Range("B4").Value = we.Cells(2, 1).Value
    ws.Range("B5").Value = "(ninguna)"
    ws.Range("B6").Value = "(ninguna)"
    Application.EnableEvents = True
    MsgBox (n - 1) & " carteras cargadas para tipo '" & ws.Range("B3").Value & "'.", vbInformation, "Carteras"
End Sub

' =======================  BOTON UNICO: HACE TODO  ==========================
Public Sub Actualizar()
    ActualizarSQL          ' 1) SQL en A41 (y calcula mAvisoEnt)
    ' Aviso si alguna entidad no se encontro en 'cartera' (pero otras si).
    If Len(mAvisoEnt) > 0 Then
        MsgBox "No encontre estas entidades en la hoja 'cartera' (columna nombre_elemento):" & _
               mAvisoEnt & vbLf & vbLf & _
               "Se han ignorado. Recuerda: id_elemento = pk_portfolio_id.", _
               vbExclamation, "Entidades sin id"
    End If
    RefrescarDatos         ' 2) lanzar + volcar en W + pivotar + dibujar
End Sub

' Lanza la consulta, vuelca el resultado crudo desde la columna W y dibuja.
Public Sub RefrescarDatos()
    Dim cn As Object, rs As Object, ws As Worksheet, sql As String, j As Long
    sql = ConstruirSQL()
    If Left(sql, 2) = "--" Then
        MsgBox "Metrica/periodo no mapeada:" & vbLf & vbLf & sql, vbExclamation, "Fase 2": Exit Sub
    End If
    On Error GoTo fallo
    Set ws = Panel()
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120
    cn.CursorLocation = 3          ' adUseClient: compatible con drivers ODBC de solo lectura (BigQuery)
    cn.Open BQ_CONN
    ' Execute devuelve un recordset de solo avance que el driver si admite.
    Set rs = cn.Execute(sql)

    Application.EnableEvents = False
    ws.Range(ws.Cells(1, 23), ws.Cells(100000, 60)).ClearContents   ' columna W en adelante
    For j = 0 To rs.Fields.Count - 1
        ws.Cells(1, 23 + j).Value = rs.Fields(j).Name
    Next j
    If Not rs.EOF Then ws.Cells(2, 23).CopyFromRecordset rs
    rs.Close: cn.Close
    VolcarResultado ws
    Application.EnableEvents = True

    DibujarGrafico
    Exit Sub
fallo:
    Application.EnableEvents = True
    MsgBox "No se pudo conectar/consultar BigQuery:" & vbLf & Err.Description & _
           vbLf & vbLf & "SQL:" & vbLf & sql, vbExclamation, "Fase 2"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

' Pivota el resultado crudo (desde W) a la tabla del grafico D:H.
Private Sub VolcarResultado(ByVal ws As Worksheet)
    Dim c As Long, hdr As String
    Dim colPort As Long, colCat As Long, colVal As Long, colBmk As Long
    Dim lastData As Long, r As Long, rowOut As Long, rr As Long
    Dim id1 As String, id2 As String, id3 As String, cat As String, pid As String
    Dim cats As Object

    c = 23
    Do While Trim(CStr(ws.Cells(1, c).Value)) <> "" And c < 60
        hdr = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        Select Case hdr
            Case "pk_portfolio_id": colPort = c
            Case "categoria":       colCat = c
            Case "valor":           colVal = c
            Case "valor_bmk":       colBmk = c
        End Select
        c = c + 1
    Loop
    If colVal = 0 Then Exit Sub

    c = colPort: If c = 0 Then c = 23
    r = 2: lastData = 1
    Do While Trim(CStr(ws.Cells(r, c).Value)) <> "" And r < 100000
        lastData = r: r = r + 1
    Loop

    id1 = IdEntidad(Trim(CStr(ws.Range("B4").Value)))
    id2 = IdEntidad(Trim(CStr(ws.Range("B5").Value)))
    id3 = IdEntidad(Trim(CStr(ws.Range("B6").Value)))
    If Trim(CStr(ws.Range("B5").Value)) = "(ninguna)" Then id2 = ""
    If Trim(CStr(ws.Range("B6").Value)) = "(ninguna)" Then id3 = ""

    ws.Range("D3:H402").ClearContents

    If colCat > 0 Then
        Set cats = CreateObject("Scripting.Dictionary")
        rowOut = 3
        For r = 2 To lastData
            cat = Trim(CStr(ws.Cells(r, colCat).Value))
            If cat <> "" And Not cats.Exists(cat) Then
                cats.Add cat, rowOut
                ws.Cells(rowOut, 4).Value = cat
                rowOut = rowOut + 1
                If rowOut > 402 Then Exit For
            End If
        Next r
        For r = 2 To lastData
            pid = Trim(CStr(ws.Cells(r, colPort).Value))
            cat = Trim(CStr(ws.Cells(r, colCat).Value))
            If cats.Exists(cat) Then
                rr = cats(cat)
                If pid = id1 Then ws.Cells(rr, 5).Value = ws.Cells(r, colVal).Value
                If Len(id2) > 0 Then If pid = id2 Then ws.Cells(rr, 6).Value = ws.Cells(r, colVal).Value
                If Len(id3) > 0 Then If pid = id3 Then ws.Cells(rr, 7).Value = ws.Cells(r, colVal).Value
            End If
        Next r
    Else
        ws.Cells(3, 4).Value = Trim(CStr(ws.Range("B8").Value)) & " - " & Trim(CStr(ws.Range("B11").Value))
        For r = 2 To lastData
            pid = Trim(CStr(ws.Cells(r, colPort).Value))
            If pid = id1 Then
                ws.Cells(3, 5).Value = ws.Cells(r, colVal).Value
                If colBmk > 0 Then ws.Cells(3, 8).Value = ws.Cells(r, colBmk).Value
            End If
            If Len(id2) > 0 Then If pid = id2 Then ws.Cells(3, 6).Value = ws.Cells(r, colVal).Value
            If Len(id3) > 0 Then If pid = id3 Then ws.Cells(3, 7).Value = ws.Cells(r, colVal).Value
        Next r
    End If
End Sub

' =======================  DIBUJO DEL GRAFICO  ==============================
Private Sub AjustarSeleccion(ws As Worksheet, ByVal celda As String, ByVal nombreLista As String)
    Dim rng As Range, c As Range, valido As Boolean
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nombreLista).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If c.Value = ws.Range(celda).Value Then valido = True
    Next c
    If Not valido Then ws.Range(celda).Value = rng.Cells(1, 1).Value
End Sub

Private Sub AddEnt(ws As Worksheet, ByVal ch As Chart, ByVal slotCell As String, _
        ByVal colLetter As String, ByVal lastRow As Long)
    Dim s As Series, v As String
    v = Trim(CStr(ws.Range(slotCell).Value))
    If v = "" Or v = "(ninguna)" Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "=Panel!$" & colLetter & "$2"
    s.Values = "=Panel!$" & colLetter & "$3:$" & colLetter & "$" & lastRow
    s.XValues = "=Panel!$D$3:$D$" & lastRow
End Sub

' Redibuja el grafico del Panel con la tabla D:H actual (mock o datos reales).
Public Sub DibujarGrafico()
    Dim ws As Worksheet, ch As Chart, s As Series, conBench As Boolean, tipo As String
    Dim lastRow As Long, benchIdx As Long
    Set ws = Panel()
    AjustarSeleccion ws, "B4", "Ent_" & ws.Range("B3").Value
    AjustarSeleccion ws, "B8", "Grupo_" & ws.Range("B7").Value

    On Error Resume Next
    Set ch = ws.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    conBench = (ws.Range("B12").Value = "Con benchmark")
    tipo = Fold(ws.Range("B14").Value)
    lastRow = 2
    Do While Trim(CStr(ws.Cells(lastRow + 1, 4).Value)) <> "" And lastRow < 402
        lastRow = lastRow + 1
    Loop
    If lastRow < 3 Then lastRow = 3

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    AddEnt ws, ch, "B4", "E", lastRow
    AddEnt ws, ch, "B5", "F", lastRow
    AddEnt ws, ch, "B6", "G", lastRow

    benchIdx = 0
    If conBench And Trim(CStr(ws.Range("B4").Value)) <> "" Then
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "=Panel!$H$2"
        s.Values = "=Panel!$H$3:$H$" & lastRow
        s.XValues = "=Panel!$D$3:$D$" & lastRow
        benchIdx = ch.SeriesCollection.Count
    End If

    Select Case tipo
        Case "barras":            ch.ChartType = xlBarClustered
        Case "lineas":            ch.ChartType = xlLineMarkers
        Case "area":              ch.ChartType = xlArea
        Case "circular":          ch.ChartType = xlPie
        Case "anillo":            ch.ChartType = xlDoughnut
        Case "radar":             ch.ChartType = xlRadarMarkers
        Case "apiladas":          ch.ChartType = xlColumnStacked
        Case "100% apiladas":     ch.ChartType = xlColumnStacked100
        Case Else:                ch.ChartType = xlColumnClustered
    End Select

    If benchIdx > 0 Then
        On Error Resume Next
        Select Case Fold(ws.Range("B13").Value)
            Case "lineas"
                ch.FullSeriesCollection(benchIdx).ChartType = xlLineMarkers
            Case "puntos"
                ch.FullSeriesCollection(benchIdx).ChartType = xlLineMarkers
                ch.FullSeriesCollection(benchIdx).Format.Line.Visible = msoFalse
            Case Else
                ch.FullSeriesCollection(benchIdx).ChartType = xlColumnClustered
        End Select
        On Error GoTo 0
    End If

    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = ws.Range("A19").Value
    If tipo <> "circular" And tipo <> "anillo" And tipo <> "radar" Then
        ch.Axes(xlValue).HasTitle = True
        ch.Axes(xlValue).AxisTitle.Text = ws.Range("B8").Value
        ch.Axes(xlCategory).HasTitle = True
        ch.Axes(xlCategory).AxisTitle.Text = ws.Range("B9").Value
    End If
    On Error GoTo 0
End Sub

' =======================  HOJA "TABLAS"  ==================================
' Ajusta columnas visibles, decimales y estilo de formato condicional.
Public Sub FormatearTablas()
    Const HDR As Long = 17, ROW0 As Long = 18, LASTROW As Long = 21
    Const FIRSTC As Long = 2, LASTC As Long = 13
    Dim ws As Worksheet, body As Range, c As Long, fmt As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Tablas")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set body = ws.Range(ws.Cells(ROW0, FIRSTC), ws.Cells(LASTROW, LASTC))

    Select Case CLng(ws.Range("B10").Value)
        Case 0: fmt = "0"
        Case 1: fmt = "0.0"
        Case Else: fmt = "0.00"
    End Select
    body.NumberFormat = fmt

    For c = FIRSTC + 1 To LASTC
        ws.Columns(c).Hidden = (Trim(CStr(ws.Cells(HDR, c).Value)) = "")
    Next c

    body.FormatConditions.Delete
    Select Case LCase(Trim(ws.Range("B9").Value))
        Case "mapa de calor"
            With body.FormatConditions.AddColorScale(ColorScaleType:=3)
                .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
                .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50
                .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
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
                .Font.Color = RGB(16, 124, 65)
            End With
            With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Font.Color = RGB(192, 0, 0)
            End With
    End Select
End Sub

' =======================  EXPORTAR A POWERPOINT (Fase 3)  =================
Public Sub CopiarAPowerPoint()
    Dim ch As ChartObject
    On Error Resume Next
    Set ch = ThisWorkbook.Sheets("Panel").ChartObjects(1)
    On Error GoTo 0
    If ch Is Nothing Then MsgBox "No encuentro el grafico en 'Panel'.", vbExclamation: Exit Sub

    ch.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    Dim ppt As Object, pres As Object, sld As Object, shp As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    If ppt.Presentations.Count = 0 Then Set pres = ppt.Presentations.Add Else Set pres = ppt.ActivePresentation
    On Error Resume Next
    Set sld = ppt.ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        If pres.Slides.Count = 0 Then Set sld = pres.Slides.Add(1, 12) Else Set sld = pres.Slides(pres.Slides.Count)
    End If
    Set shp = sld.Shapes.PasteSpecial(DataType:=2)
    On Error Resume Next
    shp.Left = (pres.PageSetup.SlideWidth - shp.Width) / 2
    shp.Top = (pres.PageSetup.SlideHeight - shp.Height) / 2
    On Error GoTo 0
End Sub
