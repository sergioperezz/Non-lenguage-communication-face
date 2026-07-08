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
' Escala de los TWR al COMPONER retornos diarios: si vienen como fraccion
' (0.0125 = 1,25%) es "1.0"; si vinieran en % (1.25) pon "100.0". Los datos
' reales vienen en FRACCION (p.ej. 0.0344 = 3,44%), por eso 1.0.
Private Const RET_ESC As String = "1.0"
Private Const T_PERF As String = "CAM_TX_PERFORMANCE_FIGURES_PD"
Private Const T_RISK As String = "CAM_TX_RISK_FIG_AGG_PD"
Private Const RISK_COL_FONDOBMK As String = "PK_TIPOGAMAN1"
' Nivel de agregacion a nivel total de la cartera (columna PK_PORTFOLIO). Las
' sub-carteras/componentes dan valores parciales (~-0.12); queremos el 'Total'.
Private Const RISK_PORTFOLIO As String = "Total"
Private Const T_POS As String = "CAM_TM_PORTFOLIOS_PD"
Private Const T_VALORES As String = "CAM_TM_MSTR_VALORES_PD"
Private Const POS_VALOR As String = "VALUATION_PC"
Private Const SECTOR_COL As String = "CLASSIFICATION_GICS"
Private Const RATING_COL As String = "COMPOSITERATINGSPCOMPOSITE"
Private Const GEO_COL As String = "FCCOUNTRYZONE"       ' zona geografica (todos los activos)
Private Const DIV_COL As String = "CURRENCY"            ' divisa (codigo ISO)
Private Const IND_COL As String = "CLASSIFICATION_GICS" ' industria (GICS; en config puede ser BICS)
Private Const TER_COL As String = "KEYFIGURESTER"
Private Const T_FONDOS As String = "CAM_TM_MSTR_FONDOS_PD"
Private Const TER_FONDO_EXPR As String = "f.COMISION_DE_GESTION_DIRECTA + f.COMISION_DEPOSITARIA_DIRECTA"
' Hoja de carteras (la que copiaras): nombre -> id. Se busca por varios nombres.
Private Const ACTIVOS_SHEET As String = "cartera"          ' nombre principal a buscar
Private Const ACTIVOS_COL_NOMBRE As String = "nombre_elemento"
Private Const ACTIVOS_COL_ID As String = "id_elemento"    ' columna que va a PK_PORTFOLIO_ID
Private Const CACHE_RET As String = "_CacheRet"   ' (compat) hoja antigua de cache; ya no se usa
Private Const CACHE_ANOS As Long = 6              ' anos de historia diaria que se cachean
' --- BLOQUE AMPLIO EN LA HOJA PANEL (a partir de la columna W) --------------
' "Cargar datos (cartera)" baja de UNA vez TODOS los datos de las carteras
' elegidas y los deja en la hoja Panel en tres bloques (columna W en adelante).
' Luego cualquier metrica/dimension/periodo se calcula EN LOCAL (sin re-consultar).
'   RET  (retornos diarios): PK_PORTFOLIO_ID | fecha | twr_1d | twr_1d_bmk
'   RISK (riesgo diario):    fecha | PK_PORTFOLIO_ID | criterio | etiqueta | variable | valor
'   POS  (posiciones hoy):   PK_PORTFOLIO_ID | gics | bics | geo | pais | divisa | rating | activo | valor
'        (bloque LEGADO; hoy Composicion/apilada usan APIL. LocalSpread/LocalTerLT
'         estan DORMIDAS: Spread/Volatilidad van por consulta directa, no por POS.)
Private Const BLK_RET_COL As Long = 23     ' W  (ancho 4)
Private Const BLK_RISK_COL As Long = 29    ' AC (ancho 6)
Private Const BLK_POS_COL As Long = 37     ' AK (ancho 9)
Private Const BLK_MARK_COL As Long = 47    ' AU  (marca "ENTS:'A','B'"; filas 1/2/3/4)
Private Const BLK_APIL_COL As Long = 49    ' AW: bloque de composicion apilada (snapshot mensual)
Private Const BLK_ULTFILA As Long = 200000 ' fila maxima para limpiar los bloques
Private Const TCMP_COL As Long = 4         ' D: matriz de composicion apilada (junto al grafico)
Private Const MAXSER As Long = 18          ' max series (D..V; los bloques empiezan en W=23)
' Columnas auxiliares reutilizables para las TABLAS COMO FORMULAS (solo hay una
' metrica en pantalla; se reescriben en cada calculo). Van despues del bloque APIL.
Private Const HLP_K1 As Long = 60          ' BH: clave = bucket / etiqueta / categoria
Private Const HLP_A1 As Long = 61          ' BI: factor(1+twr) / valor / valoracion / etiqueta(apiladas)
Private Const HLP_A2 As Long = 62          ' BJ: factor benchmark / valoracion (apiladas)
Private Const HLP_PID As Long = 63         ' BK: pid RECORTADO (Trim) -> el "=" de la formula no ignora espacios
Private Const HLP_C1 As Long = 64          ' BL: criterio recortado (riesgo)
Private Const HLP_C2 As Long = 65          ' BM: variable target recortada (riesgo)
' Hoja "Tablas" (matriz entidades x variables) y hoja "Posiciones" (holdings):
' fila de cabecera y primera fila de datos.
' Origen de la tabla de dos niveles (marca/metrica/periodo/datos y columna de
' Entidad). Los fija OrigenTabla segun la hoja: la maestra "Tablas" en A/filas 4-7;
' las hojas generadas "Tabla N" en H/filas 5-8 (config a la izquierda).
Private gCol0 As Long                       ' columna de Entidad (1=A maestra, 8=H generada)
Private gRMark As Long, gRMet As Long, gRHdr As Long, gRData As Long
Private gRCap As Long                        ' fila de CABECERA de presentacion (0 = sin ella)
Private gEstilo As String, gTotal As String, gEstado As String   ' celdas de control
Private Const PS_HDR As Long = 6           ' Posiciones: cabecera de columnas
Private Const PS_ROW0 As Long = 7          ' Posiciones: primera fila de holdings
' ===========================================================================

' Variables de modulo (deben ir aqui arriba, antes de la primera Sub/Function).
' Nombres de entidad que no se pudieron traducir a id (para avisar al usuario).
Private mAvisoEnt As String
' Si esta a True, la SQL de rendimiento se genera SIN columnas de benchmark
' (se activa como reintento cuando el driver dice que la columna _BMK no existe).
Private mForzarSinBmk As Boolean
' Ids (pk_portfolio_id) resueltos de B4/B5/B6, EN ORDEN. Los fija ListaEntidades
' al construir la query y los reutiliza VolcarResultado (asi lo que se consulta y
' lo que se vuelca usan exactamente el mismo id, sin re-traducir).
Private mId1 As String, mId2 As String, mId3 As String
' Marca para que el auto-refresco de la hoja Tablas se dispare solo UNA vez por
' sesion (al entrar en la pestana), no cada vez que se activa.
Private mTablaAuto As Boolean
' Si esta a True, CfgAsOf() devuelve "" (sin tope): fuerza "ultima fecha" aunque
' la Portada tenga un mes elegido. Lo usan los botones "Actualizar" de las hojas
' generadas para traer siempre lo mas reciente.
Private mSinAsOf As Boolean
' Si esta puesto, TODO el motor del Panel (SQL/bloques/dibujo) opera sobre ESA hoja
' en vez de "Panel". Lo usan las hojas generadas con "Crear hoja con grafica" (que
' son copias del Panel) para refrescar/dibujar su propio grafico de forma autonoma.
Private mHojaPanel As Worksheet

Private Function Panel() As Worksheet
    If mHojaPanel Is Nothing Then Set Panel = ThisWorkbook.Sheets("Panel") Else Set Panel = mHojaPanel
End Function

' Prefijo de referencia a una hoja para las series del grafico: "='Hoja'!".
Private Function QHoja(ByVal ws As Worksheet) As String
    QHoja = "='" & ws.Name & "'!"
End Function

' Tipo de hoja generada: "grafica" (copia del Panel) / "tabla" / "". Robusto: por
' marca AD1 o por el prefijo del nombre (Grafica N / Tabla N).
Private Function TipoHoja(ByVal ws As Worksheet) As String
    Dim n As String: n = Fold(ws.Name)
    Dim m As String: m = Fold(CStr(ws.Cells(1, 30).Value))
    If m = "grafica" Or Left(n, 7) = "grafica" Then
        TipoHoja = "grafica"
    ElseIf m = "tabla" Or Left(n, 5) = "tabla" Then
        TipoHoja = "tabla"
    End If
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(CStr(s), "'", "''")
End Function

' Convierte a numero un valor que puede venir como TEXTO. Los valores de las
' consultas llegan formateados con '.' decimal (FORMAT '%.10f') para evitar el
' reescalado del driver ODBC con los NUMERIC. Val() usa siempre '.' como decimal,
' asi que es independiente del idioma. Devuelve "" si esta vacio.
Private Function NumVal(ByVal v As Variant) As Variant
    Dim s As String
    s = Trim(CStr(v))
    If s = "" Then NumVal = "" Else NumVal = Val(Replace(s, ",", "."))
End Function

' Como NumVal pero SIEMPRE devuelve un Double (vacio/no numerico -> 0).
Private Function NumDbl(ByVal v As Variant) As Double
    Dim s As String
    s = Trim(CStr(v))
    If s <> "" Then NumDbl = Val(Replace(s, ",", "."))
End Function

' Formato de numero segun la metrica del panel (B8). Los rendimientos vienen en
' FRACCION (0.003 = 0.3%) -> se muestran como %; la duracion es un numero; el
' resto, numero con 2 decimales. La macro lo aplica sola al volcar/dibujar, asi
' no hay que formatear celdas a mano (y no se rompe al cambiar de metrica).
Private Function FormatoMetrica(ByVal met As String) As String
    Dim m As String: m = Fold(met)
    If InStr(m, "rentab") = 1 Or m = "peso" Then
        FormatoMetrica = "0.00%"
    ElseIf InStr(m, "duraci") = 1 Then
        FormatoMetrica = "0.000"
    ElseIf m = "importe" Then
        FormatoMetrica = "#,##0"            ' valor absoluto (euros, con separador de miles)
    Else
        FormatoMetrica = "0.00"
    End If
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

' ====================  LECTOR DE LA HOJA "config"  =========================
' Lee el valor de una clave en la hoja "config" (columna A = clave, B = valor).
' Si la hoja o la clave no existen (o el valor esta vacio), devuelve el valor por
' defecto (las constantes de arriba). Asi el libro funciona con o sin la hoja.
Private Function Cfg(ByVal clave As String, ByVal defecto As String) As String
    Dim ws As Worksheet, r As Long, lastR As Long, v As String
    Cfg = defecto
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastR
        If Fold(ws.Cells(r, 1).Value) = Fold(clave) Then
            v = Trim(CStr(ws.Cells(r, 2).Value))
            If Len(v) > 0 Then Cfg = v
            Exit Function
        End If
    Next r
End Function

Private Function CfgProject() As String
    CfgProject = Cfg("PROJECT", BQ_PROJECT)
End Function
Private Function CfgConn() As String
    CfgConn = "DSN=" & Cfg("DSN", BQ_DSN) & ";"
End Function
Private Function CfgNav() As String
    CfgNav = Cfg("PK_NAV_GNAV", F_NAV_GNAV)
End Function
Private Function CfgBmk() As String
    CfgBmk = Cfg("BENCHMARK", F_BENCHMARK)
End Function
Private Function CfgFondo() As String
    CfgFondo = Cfg("PK_TIPOGAMAN1", "FONDO")
End Function
Private Function CfgPortfolio() As String
    CfgPortfolio = Cfg("PK_PORTFOLIO", RISK_PORTFOLIO)
End Function
Private Function CfgLtLevel() As String
    CfgLtLevel = Cfg("PK_LTLEVEL", "2")
End Function

' ---- Fecha de referencia GLOBAL del documento (hoja "Portada") ----
' Ultimo dia del mes elegido (Ano B2 + Mes B3). "" si no hay Portada o el mes es
' "(ultimo)"/vacio -> entonces cada consulta usa la ultima fecha disponible.
' Todas las consultas aplican este tope (fecha <= referencia) para que el
' documento entero quede "a cierre" de ese mes.
Private Function CfgAsOf() As String
    If mSinAsOf Then Exit Function                      ' "ultima fecha" forzada
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Portada")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim y As Long, m As Long
    y = CLng(Val(CStr(ws.Range("B2").Value)))
    m = MesNum(CStr(ws.Range("B3").Value))
    If y < 1900 Or m < 1 Then Exit Function            ' "(ultimo)" o vacio -> sin tope
    CfgAsOf = Format(DateSerial(y, m + 1, 0), "yyyy-mm-dd")   ' ultimo dia del mes
End Function

' " AND <col> <= DATE 'YYYY-MM-DD'" (o "" si no hay fecha de referencia).
Private Function AndAsOf(ByVal col As String) As String
    Dim d As String: d = CfgAsOf()
    If Len(d) > 0 Then AndAsOf = " AND " & col & " <= DATE '" & d & "'"
End Function

' " WHERE PK_FECHA_DATOS <= DATE 'YYYY-MM-DD'" para acotar los subselect de MAX.
Private Function MaxAsOf() As String
    Dim d As String: d = CfgAsOf()
    If Len(d) > 0 Then MaxAsOf = " WHERE PK_FECHA_DATOS <= DATE '" & d & "'"
End Function
' Lista (entrecomillada) de PK_CRITERIO_AGREGACION que baja el bloque RISK.
' Duracion y TIR fijos + CMR y VaR (nombres de criterio configurables en 'config').
Private Function CfgRiskCriterios() As String
    CfgRiskCriterios = "'Duracion','TIR'" & _
        ",'" & Esc(Cfg("RISK_CRIT_CMR", "CMR")) & "'" & _
        ",'" & Esc(Cfg("RISK_CRIT_VAR", "VaR")) & "'"
End Function
Private Function CfgRetEsc() As String
    CfgRetEsc = Cfg("RET_ESC", RET_ESC)
End Function
Private Function CfgDataset(ByVal ds As String) As String
    Select Case ds
        Case DS_PROD: CfgDataset = Cfg("DATASET_PROD", DS_PROD)
        Case DS_OPER: CfgDataset = Cfg("DATASET_OPER", DS_OPER)
        Case DS_MERC: CfgDataset = Cfg("DATASET_MERC", DS_MERC)
        Case Else:    CfgDataset = ds
    End Select
End Function

Private Function Tbl(ByVal ds As String, ByVal t As String) As String
    Tbl = "`" & CfgProject() & "." & CfgDataset(ds) & "." & t & "`"
End Function

' Tabla de POSICIONES (holdings) para composicion/spread/TER. Es configurable
' porque puede estar en otra tabla/dataset distinto al de rendimiento. Cuando
' se sepa el nombre real de la tabla con valoracion por valor, se pone en config
' (POS_TABLE y, si aplica, POS_DATASET) sin tocar la macro.
Private Function TblPos() As String
    TblPos = "`" & CfgProject() & "." & Cfg("POS_DATASET", CfgDataset(DS_OPER)) & _
             "." & Cfg("POS_TABLE", "CAM_TX_PORTFOLIOS_COMP_PD") & "`"
End Function
Private Function PosValor() As String
    PosValor = Cfg("POS_VALOR", POS_VALOR)
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
    Dim celdas As Variant, i As Long, v As String, idp As String, out As String
    mAvisoEnt = "": mId1 = "": mId2 = "": mId3 = ""
    celdas = Array("B4", "B5", "B6")
    For i = 0 To 2
        v = Trim(CStr(ws.Range(CStr(celdas(i))).Value))
        idp = ""
        If v <> "" And v <> "(ninguna)" Then
            idp = IdEntidad(v)
            If Len(idp) = 0 Then
                mAvisoEnt = mAvisoEnt & vbLf & "   - " & v
            Else
                If Len(out) > 0 Then out = out & ", "
                out = out & "'" & Esc(idp) & "'"
            End If
        End If
        Select Case i
            Case 0: mId1 = idp
            Case 1: mId2 = idp
            Case 2: mId3 = idp
        End Select
    Next i
    ListaEntidades = out
End Function

Private Function SufijoPeriodo(ByVal p As String) As String
    Select Case UCase(Trim(p))
        Case "MTD": SufijoPeriodo = "MTD"
        Case "QTD": SufijoPeriodo = "QTD"
        Case "YTD": SufijoPeriodo = "YTD"
        Case "WTD": SufijoPeriodo = "WTD"
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
            colVal = "TWR_" & suf
            ' Segun el diccionario homologado, SOLO existe columna de benchmark
            ' (_BMK) para los periodos "hasta la fecha". Para las rentabilidades
            ' rolling (1M/1Y/3Y/5Y) no hay _BMK -> se consulta solo el fondo.
            Select Case suf
                Case "MTD", "QTD", "YTD", "WTD", "1D": colBmk = "TWR_" & suf & "_BMK"
            End Select
            colDif = ""          ' el diferencial no es necesario para el grafico
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

' Devuelve True si la dimension del eje X es temporal (serie en el tiempo).
Private Function DimEsTiempo(ByVal dimen As String) As Boolean
    Select Case Fold(dimen)
        Case "diario", "semanal", "mensual", "trimestral", "semestral", "anual": DimEsTiempo = True
    End Select
End Function

' Mapea el nombre de la metrica del panel al valor real de PK_VARIABLE_TARGET
' (quita acentos y espacios): "Duracion Modificada" -> "DuracionModificada",
' "Duracion Efectiva" -> "DuracionEfectiva".
Private Function VarTarget(ByVal met As String) As String
    Dim s As String
    s = Trim(CStr(met))
    s = Replace(s, ChrW(225), "a"): s = Replace(s, ChrW(233), "e"): s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(243), "o"): s = Replace(s, ChrW(250), "u"): s = Replace(s, ChrW(241), "n")
    s = Replace(s, ChrW(193), "A"): s = Replace(s, ChrW(201), "E"): s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(211), "O"): s = Replace(s, ChrW(218), "U"): s = Replace(s, ChrW(209), "N")
    VarTarget = Replace(s, " ", "")
End Function

' Serie temporal de una metrica PUNTUAL (Duracion/TIR): ultimo valor de cada
' bucket (trimestre/mes/...) dentro de la ventana del periodo. Filtra por
' PK_VARIABLE_TARGET (p.ej. DuracionModificada, NO Macaulay) y, si hay varias
' filas la misma fecha, se queda con el VALOR mayor (la duracion real).
Private Function SQLRiesgoTemporal(ws As Worksheet, ByVal ents As String, _
        ByVal crit As String, ByVal dimen As String, ByVal varTarget As String) As String
    Dim bucket As String, intv As String, sql As String, wVarT As String
    bucket = BucketExpr(dimen)
    intv = IntervaloPeriodo(Trim(CStr(ws.Range("B11").Value)))
    If Len(intv) = 0 Then intv = "INTERVAL 1 YEAR"
    If Len(varTarget) > 0 Then wVarT = " AND PK_VARIABLE_TARGET = '" & Esc(varTarget) & "'"
    wVarT = wVarT & " AND PK_PORTFOLIO = '" & Esc(CfgPortfolio()) & "'"   ' nivel total (no componentes)
    If Len(CfgLtLevel()) > 0 Then wVarT = wVarT & " AND PK_LTLEVEL = " & CfgLtLevel()   ' nivel look-through
    sql = "WITH ult AS (" & vbLf & _
          "  SELECT PK_PORTFOLIO_ID, MAX(PK_FECHA_DATOS) AS dmax" & vbLf & _
          "  FROM " & Tbl(DS_PROD, T_RISK) & vbLf & _
          "  WHERE PK_CRITERIO_AGREGACION = '" & Esc(crit) & "' AND " & RISK_COL_FONDOBMK & " = '" & CfgFondo() & "'" & wVarT
    If Len(ents) > 0 Then sql = sql & " AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & AndAsOf("PK_FECHA_DATOS")
    sql = sql & vbLf & "  GROUP BY PK_PORTFOLIO_ID)" & vbLf & _
          "SELECT p.PK_PORTFOLIO_ID, " & bucket & " AS categoria, FORMAT('%.10f', CAST(p.VALOR AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_RISK) & " p" & vbLf & _
          "JOIN ult ON ult.PK_PORTFOLIO_ID = p.PK_PORTFOLIO_ID" & vbLf & _
          "WHERE p.PK_CRITERIO_AGREGACION = '" & Esc(crit) & "' AND p." & RISK_COL_FONDOBMK & " = '" & CfgFondo() & "'" & _
          Replace(Replace(Replace(wVarT, "PK_VARIABLE_TARGET", "p.PK_VARIABLE_TARGET"), "PK_PORTFOLIO =", "p.PK_PORTFOLIO ="), "PK_LTLEVEL", "p.PK_LTLEVEL") & vbLf & _
          VentanaFechas(Trim(CStr(ws.Range("B11").Value)), dimen) & vbLf & _
          "QUALIFY ROW_NUMBER() OVER (PARTITION BY p.PK_PORTFOLIO_ID, " & bucket & _
          " ORDER BY p.PK_FECHA_DATOS DESC, p.VALOR DESC) = 1" & vbLf & _
          "ORDER BY p.PK_PORTFOLIO_ID, p.PK_FECHA_DATOS"
    SQLRiesgoTemporal = sql
End Function

Private Function SQLRiesgo(ws As Worksheet, ByVal variable As String, ByVal ents As String, _
        Optional ByVal critFijo As String = "") As String
    Dim dimen As String, crit As String, sql As String, wVar As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    ' Dimension temporal: serie en el tiempo (ultimo valor de cada bucket).
    If DimEsTiempo(dimen) Then
        Dim vtT As String
        If Len(critFijo) > 0 Then
            crit = critFijo: vtT = ""
        Else
            crit = "Duracion": vtT = VarTarget(variable)   ' DuracionModificada / DuracionEfectiva
        End If
        SQLRiesgo = SQLRiesgoTemporal(ws, ents, crit, dimen, vtT)
        Exit Function
    End If
    If Len(critFijo) > 0 Then
        crit = critFijo: wVar = ""
    Else
        crit = DimACriterio(dimen)
        If Len(crit) = 0 Then
            SQLRiesgo = "-- Dimension '" & dimen & "' no existe como desglose en CAM_TX_RISK_FIG_AGG_PD." & vbLf & _
                        "-- Criterios: AssetType (Activo), Geo (Geografia), FX (Divisa), Duracion (total)."
            Exit Function
        End If
        ' Distingue la variante de la metrica (p.ej. DuracionModificada, no Macaulay).
        wVar = "  AND PK_VARIABLE_TARGET = '" & Esc(VarTarget(variable)) & "'" & vbLf
    End If
    sql = "SELECT PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION AS categoria, FORMAT('%.10f', CAST(VALOR AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & Tbl(DS_PROD, T_RISK) & vbLf & _
          "WHERE PK_CRITERIO_AGREGACION = '" & Esc(crit) & "'" & vbLf & _
          wVar & "  AND " & RISK_COL_FONDOBMK & " = '" & CfgFondo() & "'" & vbLf & _
          "  AND PK_PORTFOLIO = '" & Esc(CfgPortfolio()) & "'"
    If Len(CfgLtLevel()) > 0 Then sql = sql & vbLf & "  AND PK_LTLEVEL = " & CfgLtLevel()
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "QUALIFY ROW_NUMBER() OVER (PARTITION BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION" & _
          " ORDER BY PK_FECHA_DATOS DESC, VALOR DESC) = 1" & vbLf & "ORDER BY PK_PORTFOLIO_ID, PK_ETIQUETA_AGREGACION"
    SQLRiesgo = sql
End Function

Private Function DimAClasificacion(ByVal dimen As String) As String
    Select Case dimen
        Case "Sector":    DimAClasificacion = "v." & Cfg("SECTOR_COL", SECTOR_COL)
        Case "Industria": DimAClasificacion = "v." & Cfg("IND_COL", IND_COL)
        Case "Rating":    DimAClasificacion = "v." & Cfg("RATING_COL", RATING_COL)
        Case "Continente": DimAClasificacion = "v." & Cfg("GEO_COL", GEO_COL)
        Case "Pais":       DimAClasificacion = "v." & Cfg("PAIS_COL", "FCCOUNTRY")
        Case "Divisa":     DimAClasificacion = "v." & Cfg("DIV_COL", DIV_COL)
        Case "Activo":     DimAClasificacion = "v." & Cfg("ACTIVO_COL", "INSTRUMENT_TYPE")
        Case Else:         DimAClasificacion = ""
    End Select
End Function

' Une posiciones (p) con el maestro de valores (v). Por defecto reproduce la
' consulta que funciona en BigQuery: posiciones = CAM_TX_PORTFOLIOS_COMP_PD
' (operativafinanciera_ds01), maestro = vista V_CAM_TM_MSTR_VALORES_PD
' (informaciondemercado_ds01), unidas por PK_SECURITY_IK (SIN fecha, porque la
' vista es la foto actual del maestro). Todo configurable en la hoja "config":
'   JOIN_KEY_POS / JOIN_KEY_VAL -> columnas de union (def. PK_SECURITY_IK)
'   VALORES_TABLE               -> tabla/vista del maestro (def. V_CAM_TM_MSTR_VALORES_PD)
'   JOIN_FECHA = 1              -> anadir v.PK_FECHA_DATOS = p.PK_FECHA_DATOS (def. no)
Private Function TblValores() As String
    TblValores = Tbl(DS_MERC, Cfg("VALORES_TABLE", "V_CAM_TM_MSTR_VALORES_PD"))
End Function
Private Function JoinValores() As String
    Dim kPos As String, kVal As String
    kPos = Cfg("JOIN_KEY_POS", "PK_SECURITY_IK")
    kVal = Cfg("JOIN_KEY_VAL", "PK_SECURITY_IK")
    JoinValores = "JOIN " & TblValores() & " v" & vbLf & _
                  "  ON v." & kVal & " = p." & kPos
    If Cfg("JOIN_FECHA", "0") = "1" Then _
        JoinValores = JoinValores & " AND v.PK_FECHA_DATOS = p.PK_FECHA_DATOS"
    JoinValores = JoinValores & vbLf
End Function

Private Function SQLComposicion(ws As Worksheet, ByVal ents As String) As String
    Dim dimen As String, grp As String, sql As String
    dimen = Trim(CStr(ws.Range("B9").Value))
    grp = DimAClasificacion(dimen)
    If Len(grp) = 0 Then
        SQLComposicion = "-- Composicion por '" & dimen & "': disponible por Sector y Rating."
        Exit Function
    End If
    sql = "SELECT p.PK_PORTFOLIO_ID, " & grp & " AS categoria, FORMAT('%.10f', CAST(SUM(p." & PosValor() & ") AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & TblPos() & " p" & vbLf & JoinValores & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & ")"
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
          "       FORMAT('%.10f', CAST(SUM(p.SPREAD * p." & PosValor() & ") / NULLIF(SUM(p." & PosValor() & "), 0) AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & TblPos() & " p" & vbLf & joinV & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID" & grpBy & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLSpread = sql
End Function

Private Function SQLTerLookthrough(ws As Worksheet, ByVal ents As String) As String
    Dim sql As String
    sql = "SELECT p.PK_PORTFOLIO_ID," & vbLf & _
          "       FORMAT('%.10f', CAST(SUM(v." & TER_COL & " * p." & PosValor() & ") / NULLIF(SUM(p." & PosValor() & "), 0) AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & TblPos() & " p" & vbLf & JoinValores & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "GROUP BY p.PK_PORTFOLIO_ID" & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLTerLookthrough = sql
End Function

Private Function SQLTerFondo(ws As Worksheet, ByVal ents As String) As String
    Dim sql As String
    sql = "SELECT DISTINCT p.PK_PORTFOLIO_ID, FORMAT('%.10f', CAST(" & TER_FONDO_EXPR & " AS FLOAT64)) AS valor" & vbLf & _
          "FROM " & TblPos() & " p" & vbLf & _
          "JOIN " & Tbl(DS_PROD, T_FONDOS) & " f ON f.PK_PRODUCTO_DATANOW = p.FK_PRODUCTO_DATANOW" & vbLf & _
          "WHERE p.PK_FECHA_DATOS = (SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & ")"
    If Len(ents) > 0 Then sql = sql & vbLf & "  AND p.PK_PORTFOLIO_ID IN (" & ents & ")"
    sql = sql & vbLf & "ORDER BY p.PK_PORTFOLIO_ID"
    SQLTerFondo = sql
End Function

' Intervalo de BigQuery para la ventana del periodo del panel (B11). Parsea el
' numero + unidad: "2A" -> "INTERVAL 2 YEAR", "3M" -> "INTERVAL 3 MONTH". Los
' periodos "a fecha" (MTD/YTD/QTD/WTD) devuelven "" (van por columna directa).
Private Function IntervaloPeriodo(ByVal per As String) As String
    Dim p As String, n As String
    p = UCase(Trim(per))
    If Len(p) < 2 Then Exit Function
    n = Left(p, Len(p) - 1)
    If Not IsNumeric(n) Then Exit Function
    Select Case Right(p, 1)
        Case "M": IntervaloPeriodo = "INTERVAL " & n & " MONTH"
        Case "A": IntervaloPeriodo = "INTERVAL " & n & " YEAR"
    End Select
End Function

' Numero de anos de un periodo tipo "NA" (2A -> 2). 1 por defecto.
Private Function AnyosPeriodo(ByVal per As String) As Long
    Dim p As String, n As String
    p = UCase(Trim(per)): AnyosPeriodo = 1
    If Len(p) >= 2 And Right(p, 1) = "A" Then
        n = Left(p, Len(p) - 1)
        If IsNumeric(n) Then AnyosPeriodo = CLng(n)
    End If
End Function

' Filtro de fechas de la ventana (usa ult.dmax y p.PK_FECHA_DATOS). Para la
' dimension ANUAL se alinea a anos NATURALES (N anos terminando en el actual, YTD),
' asi "3A" da 2024/2025/2026-YTD y no medio 2023. El resto: ventana movil normal.
Private Function VentanaFechas(ByVal per As String, ByVal dimen As String, _
        Optional ByVal udmax As String = "ult.dmax") As String
    Dim intv As String
    ' Calendario cerrado ("... anterior"): ventana [inicio, fin) del periodo anterior,
    ' anclada a udmax con DATE_TRUNC. Tiene prioridad sobre la ventana movil/anual.
    Dim pa As String: pa = PerAnterior(per)
    If pa <> "" Then
        Dim u As String
        Select Case pa
            Case "M": u = "MONTH"
            Case "Q": u = "QUARTER"
            Case "Y": u = "YEAR"
        End Select
        VentanaFechas = "  AND p.PK_FECHA_DATOS >= DATE_SUB(DATE_TRUNC(" & udmax & ", " & u & "), INTERVAL 1 " & u & ")" & vbLf & _
                        "  AND p.PK_FECHA_DATOS <  DATE_TRUNC(" & udmax & ", " & u & ")"
        Exit Function
    End If
    If Fold(dimen) = "anual" Then
        VentanaFechas = "  AND p.PK_FECHA_DATOS >= DATE_TRUNC(DATE_SUB(" & udmax & ", INTERVAL " & _
            (AnyosPeriodo(per) - 1) & " YEAR), YEAR)" & vbLf & _
            "  AND p.PK_FECHA_DATOS <= " & udmax
    Else
        intv = IntervaloPeriodo(per)
        If Len(intv) = 0 Then intv = "INTERVAL 1 YEAR"
        VentanaFechas = "  AND p.PK_FECHA_DATOS >  DATE_SUB(" & udmax & ", " & intv & ")" & vbLf & _
            "  AND p.PK_FECHA_DATOS <= " & udmax
    End If
End Function

' Intervalo de BigQuery para la ventana rolling de un sufijo de periodo.
Private Function IntervaloSuf(ByVal suf As String) As String
    Select Case suf
        Case "1M": IntervaloSuf = "INTERVAL 1 MONTH"
        Case "1Y": IntervaloSuf = "INTERVAL 1 YEAR"
        Case "3Y": IntervaloSuf = "INTERVAL 3 YEAR"
        Case "5Y": IntervaloSuf = "INTERVAL 5 YEAR"
        Case Else: IntervaloSuf = ""
    End Select
End Function

' Expresion SQL que agrupa las fechas en el bucket del eje X (dimension B9):
' Trimestral -> "2025-T3", Mensual -> "2025-07", Semestral -> "2025-S2",
' Anual -> "2025". Devuelve "" si la dimension no es temporal.
Private Function BucketExpr(ByVal dimen As String) As String
    Select Case Fold(dimen)
        Case "diario"
            BucketExpr = "FORMAT_DATE('%Y-%m-%d', p.PK_FECHA_DATOS)"
        Case "semanal"
            BucketExpr = "FORMAT_DATE('%G-W%V', p.PK_FECHA_DATOS)"
        Case "trimestral"
            BucketExpr = "CONCAT(CAST(EXTRACT(YEAR FROM p.PK_FECHA_DATOS) AS STRING), '-T', " & _
                         "CAST(EXTRACT(QUARTER FROM p.PK_FECHA_DATOS) AS STRING))"
        Case "mensual"
            BucketExpr = "FORMAT_DATE('%Y-%m', p.PK_FECHA_DATOS)"
        Case "semestral"
            BucketExpr = "CONCAT(CAST(EXTRACT(YEAR FROM p.PK_FECHA_DATOS) AS STRING), '-S', " & _
                         "CAST(IF(EXTRACT(MONTH FROM p.PK_FECHA_DATOS) <= 6, 1, 2) AS STRING))"
        Case "anual"
            BucketExpr = "CAST(EXTRACT(YEAR FROM p.PK_FECHA_DATOS) AS STRING)"
        Case Else
            BucketExpr = ""
    End Select
End Function

' Rentabilidad rolling COMPUESTA a partir de los retornos diarios, para los
' periodos que no tienen columna propia de benchmark (_BMK): 1M/1Y/3Y/5Y.
' Compone TWR_1D (fondo) y, si se pide, TWR_1D_BMK (benchmark) sobre la ventana
' de fechas: R = EXP(SUM(LN(1 + r))) - 1. Usa SAFE.LN para ignorar dias con
' datos invalidos. Devuelve valor (y valor_bmk) por PK_PORTFOLIO_ID.
Private Function SQLRendimientoDiario(ws As Worksheet, ByVal ents As String, ByVal per As String, _
        ByVal conBmk As Boolean, Optional ByVal bmkId As String = "") As String
    Dim intv As String, sql As String, colB As String, whereBmk As String
    Dim bucket As String, selCat As String, grpCat As String
    Dim dimen As String: dimen = Trim(CStr(ws.Range("B9").Value))
    Dim es As String: es = CfgRetEsc()
    intv = IntervaloPeriodo(per)
    If Len(intv) = 0 And PerAnterior(per) = "" Then Exit Function
    ' Bucket del eje X segun la dimension (B9): un retorno compuesto por trimestre
    ' (o mes/semestre/ano) dentro de la ventana del periodo.
    bucket = BucketExpr(dimen)
    If Len(bucket) > 0 Then
        selCat = "       " & bucket & " AS categoria," & vbLf
        grpCat = ", categoria"
    End If
    ' Benchmark PROPIO de la cartera (TWR_1D_BMK) solo si conBmk y NO se pidio otro indice.
    If conBmk And Len(bmkId) = 0 Then
        colB = "," & vbLf & _
               "       FORMAT('%.10f', (EXP(SUM(SAFE.LN(1 + SAFE_DIVIDE(p.TWR_1D_BMK, " & es & ")))) - 1) * " & es & ") AS valor_bmk"
        whereBmk = vbLf & "  AND p.TWR_1D_BMK IS NOT NULL"
    End If

    Dim ultCte As String
    ultCte = "ult AS (" & vbLf & _
             "  SELECT PK_PORTFOLIO_ID, MAX(PK_FECHA_DATOS) AS dmax" & vbLf & _
             "  FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
             "  WHERE PK_NAV_GNAV = '" & CfgNav() & "' AND BENCHMARK = '" & CfgBmk() & "'"
    If Len(ents) > 0 Then ultCte = ultCte & " AND PK_PORTFOLIO_ID IN (" & ents & ")"
    ultCte = ultCte & AndAsOf("PK_FECHA_DATOS") & vbLf & "  GROUP BY PK_PORTFOLIO_ID)"

    Dim baseSel As String
    baseSel = "SELECT p.PK_PORTFOLIO_ID," & vbLf & selCat & _
          "       FORMAT('%.10f', (EXP(SUM(SAFE.LN(1 + SAFE_DIVIDE(p.TWR_1D, " & es & ")))) - 1) * " & es & ") AS valor" & colB & vbLf & _
          "FROM " & Tbl(DS_PROD, T_PERF) & " p" & vbLf & _
          "JOIN ult ON ult.PK_PORTFOLIO_ID = p.PK_PORTFOLIO_ID" & vbLf & _
          "WHERE p.PK_NAV_GNAV = '" & CfgNav() & "' AND p.BENCHMARK = '" & CfgBmk() & "'" & vbLf & _
          VentanaFechas(per, dimen) & vbLf & _
          "  AND p.TWR_1D IS NOT NULL" & whereBmk & vbLf & _
          "GROUP BY p.PK_PORTFOLIO_ID" & grpCat & vbLf & _
          "ORDER BY p.PK_PORTFOLIO_ID, MIN(p.PK_FECHA_DATOS)"

    If Len(bmkId) = 0 Then
        SQLRendimientoDiario = "WITH " & ultCte & vbLf & baseSel
        Exit Function
    End If

    ' --- Benchmark = OTRO INDICE (bmkId): retorno compuesto del indice por bucket (o
    ' unico), unido a la base por categoria (el mismo indice para todas las carteras).
    Dim idxSelCat As String, idxGrp As String
    If Len(bucket) > 0 Then
        idxSelCat = "       " & bucket & " AS categoria," & vbLf
        idxGrp = vbLf & "  GROUP BY categoria"
    End If
    Dim ultIdx As String
    ultIdx = "ultidx AS (" & vbLf & _
             "  SELECT MAX(PK_FECHA_DATOS) AS dmax" & vbLf & _
             "  FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
             "  WHERE PK_NAV_GNAV = '" & CfgNav() & "' AND BENCHMARK = '" & CfgBmk() & "'" & _
             " AND PK_PORTFOLIO_ID = '" & Esc(bmkId) & "'" & AndAsOf("PK_FECHA_DATOS") & ")"
    Dim idxCte As String
    idxCte = "idx AS (" & vbLf & _
             "  SELECT" & vbLf & idxSelCat & _
             "       FORMAT('%.10f', (EXP(SUM(SAFE.LN(1 + SAFE_DIVIDE(p.TWR_1D, " & es & ")))) - 1) * " & es & ") AS valor_bmk" & vbLf & _
             "  FROM " & Tbl(DS_PROD, T_PERF) & " p, ultidx" & vbLf & _
             "  WHERE p.PK_NAV_GNAV = '" & CfgNav() & "' AND p.BENCHMARK = '" & CfgBmk() & "'" & vbLf & _
             "    AND p.PK_PORTFOLIO_ID = '" & Esc(bmkId) & "'" & vbLf & _
             VentanaFechas(per, dimen, "ultidx.dmax") & vbLf & _
             "    AND p.TWR_1D IS NOT NULL" & idxGrp & ")"
    Dim joinCond As String
    If Len(bucket) > 0 Then joinCond = "USING (categoria)" Else joinCond = "ON TRUE"
    SQLRendimientoDiario = "WITH " & ultCte & "," & vbLf & ultIdx & "," & vbLf & _
          "base AS (" & vbLf & baseSel & vbLf & ")," & vbLf & idxCte & vbLf & _
          "SELECT base.*, idx.valor_bmk" & vbLf & _
          "FROM base LEFT JOIN idx " & joinCond & vbLf & _
          "ORDER BY base.PK_PORTFOLIO_ID"
End Function

' --- Modo de benchmark (celda B14 del Panel) ---------------------------------
' B14: "Sin benchmark" | "Benchmark asociado" (el propio de la cartera) | <nombre
' de un indice> (usar ese indice como benchmark).
Private Function ConBenchmark(ByVal ws As Worksheet) As Boolean
    Dim v As String: v = Trim(CStr(ws.Range("B14").Value))
    ConBenchmark = (Len(v) > 0 And StrComp(v, "Sin benchmark", vbTextCompare) <> 0)
End Function

' Id del indice elegido como benchmark en B14 (si no es "Sin benchmark" ni
' "Benchmark asociado"). "" = benchmark propio de la cartera (o ninguno).
Private Function BmkIndiceId(ByVal ws As Worksheet) As String
    Dim v As String: v = Trim(CStr(ws.Range("B14").Value))
    If Len(v) = 0 Then Exit Function
    If StrComp(v, "Sin benchmark", vbTextCompare) = 0 Then Exit Function
    If StrComp(v, "Benchmark asociado", vbTextCompare) = 0 Then Exit Function
    BmkIndiceId = IdEntidad(v)
End Function

' Etiqueta de la columna H (benchmark) segun B14: el nombre del indice si se eligio
' uno concreto; "Benchmark" para el benchmark propio.
Private Function EtiquetaBenchmark(ByVal ws As Worksheet) As String
    If Len(BmkIndiceId(ws)) > 0 Then EtiquetaBenchmark = Trim(CStr(ws.Range("B14").Value)) Else EtiquetaBenchmark = "Benchmark"
End Function

' Rellena B13 "Benchmark propuesto" con el benchmark asociado a la Entidad 1 (B4).
' DE MOMENTO: comprueba en la tabla de performance que la cartera tiene el slot
' CfgBmk() ("Benchmark 1") con datos y muestra ese nombre; si no, "(sin benchmark)".
' (El nombre real del indice vendra de CAM_TX_BENCHMARK_COMP_PD mas adelante.)
Private Sub ActualizarBenchmarkPropuesto(ByVal ws As Worksheet)
    Dim id As String, prev As Boolean, cn As Object, rs As Object, hay As Boolean
    prev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fin
    id = IdEntidad(CStr(ws.Range("B4").Value))
    If Len(id) = 0 Then ws.Range("B13").Value = "(elige Entidad 1)": GoTo fin
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 30: cn.CursorLocation = 3: cn.Open CfgConn()
    Set rs = cn.Execute("SELECT 1 FROM " & Tbl(DS_PROD, T_PERF) & _
        " WHERE PK_NAV_GNAV='" & CfgNav() & "' AND BENCHMARK='" & CfgBmk() & "'" & _
        " AND PK_PORTFOLIO_ID='" & Esc(id) & "' AND TWR_1D_BMK IS NOT NULL LIMIT 1")
    hay = (Not rs.EOF)
    rs.Close: cn.Close
    If hay Then ws.Range("B13").Value = CfgBmk() Else ws.Range("B13").Value = "(sin benchmark)"
fin:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    Application.EnableEvents = prev
    On Error GoTo 0
End Sub

' Entrada publica (la llama el evento Worksheet_Change de la hoja Panel al cambiar B4).
Public Sub RefrescarBmkPropuesto()
    ActualizarBenchmarkPropuesto Panel()
End Sub

' ---- Construye la SQL segun la metrica del Panel (B8) ----
Public Function ConstruirSQL() As String
    Dim ws As Worksheet, met As String, per As String, ents As String
    Dim colVal As String, colBmk As String, colDif As String, cols As String, sql As String, conBmk As Boolean
    Set ws = Panel()
    met = Trim(CStr(ws.Range("B8").Value))
    per = Trim(CStr(ws.Range("B11").Value))
    ents = ListaEntidades(ws)
    conBmk = ConBenchmark(ws) And Not mForzarSinBmk

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
    If met = "Peso" Or met = "Importe" Then ConstruirSQL = SQLComposicion(ws, ents): Exit Function
    If met = "Spread" Then ConstruirSQL = SQLSpread(ws, ents): Exit Function
    If met = "TER" Then ConstruirSQL = SQLTerFondo(ws, ents): Exit Function
    If met = "TER Look-through" Then ConstruirSQL = SQLTerLookthrough(ws, ents): Exit Function

    ' Rentabilidad rolling (1M/1Y/3Y/5Y): no hay columna de benchmark propia, asi
    ' que componemos los retornos diarios (fondo TWR_1D y benchmark TWR_1D_BMK).
    If (met = "Rentabilidad" Or met = "Rentab. acum.") And _
       (Len(IntervaloPeriodo(per)) > 0 Or PerAnterior(per) <> "") Then
        ConstruirSQL = SQLRendimientoDiario(ws, ents, per, conBmk, BmkIndiceId(ws))
        Exit Function
    End If

    If Not MapMetrica(met, per, colVal, colBmk, colDif) Then
        ConstruirSQL = "-- Metrica '" & met & "' / periodo '" & per & "': no mapeada." & vbLf & _
                       "-- PER / DividendYield / Liquidez: sin fuente en el diccionario."
        Exit Function
    End If

    cols = "PK_PORTFOLIO_ID, FORMAT('%.10f', CAST(" & colVal & " AS FLOAT64)) AS valor"
    If conBmk And Len(colBmk) > 0 Then cols = cols & ", FORMAT('%.10f', CAST(" & colBmk & " AS FLOAT64)) AS valor_bmk"
    If conBmk And Len(colDif) > 0 Then cols = cols & ", FORMAT('%.10f', CAST(" & colDif & " AS FLOAT64)) AS diferencial"
    sql = "SELECT " & cols & vbLf & _
          "FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
          "WHERE PK_NAV_GNAV = '" & CfgNav() & "'" & vbLf & _
          "  AND BENCHMARK = '" & CfgBmk() & "'"
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

' SQL que REALMENTE se lanza para los parametros actuales: si el grupo usa un
' bloque amplio (Rendimiento/Riesgo/Composicion/Composicion apilada), es la SQL
' del bloque; si no, la consulta directa por metrica. Para 'Ver SQL' / A41.
Public Function SQLActual() As String
    Dim ws As Worksheet, gk As String, blkId As String, ents As String
    Set ws = Panel()
    gk = GrupoKey(Trim(CStr(ws.Range("B7").Value)))
    ents = ListaEntidades(ws)
    blkId = BloqueDe(gk, Trim(CStr(ws.Range("B9").Value)))
    If Len(blkId) > 0 And Len(ents) > 0 Then
        SQLActual = "-- Bloque '" & blkId & "' (se descarga una vez y se trocea en local):" & vbLf & BlkSQL(blkId, ents)
    Else
        SQLActual = ConstruirSQL()
    End If
End Function

Public Sub ActualizarSQL()
    On Error Resume Next
    Panel().Range("A41").Value = SQLActual()
End Sub

Public Sub VerSQL()
    ActualizarSQL
    MsgBox SQLActual(), vbInformation, "SQL para los parametros actuales"
End Sub

Public Sub VerM()
    MsgBox ConstruirM(), vbInformation, "Power Query (M)"
End Sub

' Diagnostico del BENCHMARK: para la Entidad 1 (B4) muestra su BMK_GESTION (y
' otros BMK_*) del maestro de carteras, y cuenta cuantas filas hay de ESE id en
' las tablas de posiciones (composicion) y de riesgo. Si son > 0, el benchmark
' se puede traer igual que la cartera (mismas tablas, como un portfolio mas).
Public Sub VerBenchmark()
    Dim ws As Worksheet, id As String, cn As Object, rs As Object, msg As String, bmkG As String
    Set ws = Panel()
    ListaEntidades ws
    id = Trim(mId1)
    If Len(id) = 0 Then MsgBox "Elige una cartera en B4 (y pulsa 'Cargar carteras').", vbExclamation, "Benchmark": Exit Sub
    On Error GoTo fallo
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 60: cn.CursorLocation = 3: cn.Open CfgConn()
    Dim master As String: master = Tbl(DS_PROD, Cfg("PORTF_MASTER", "CAM_TM_PORTFOLIOS_PD"))
    Set rs = cn.Execute("SELECT BMK_GESTION, BMK_FOLLETO, BMK_EVAL FROM " & master & _
                        " WHERE PK_PORTFOLIO_ID = '" & Esc(id) & "' LIMIT 1")
    If Not rs.EOF Then
        bmkG = Trim(CStr(rs.Fields(0).Value))
        msg = "Benchmark de " & id & " (maestro de carteras):" & vbLf & _
              "  BMK_GESTION = " & CStr(rs.Fields(0).Value) & vbLf & _
              "  BMK_FOLLETO = " & CStr(rs.Fields(1).Value) & vbLf & _
              "  BMK_EVAL    = " & CStr(rs.Fields(2).Value) & vbLf
    Else
        msg = "No encontre " & id & " en " & master & "." & vbLf
    End If
    rs.Close
    If Len(bmkG) > 0 Then
        Set rs = cn.Execute("SELECT COUNT(*) FROM " & TblPos() & " WHERE PK_PORTFOLIO_ID = '" & Esc(bmkG) & "'")
        msg = msg & vbLf & "Filas en posiciones con PK_PORTFOLIO_ID='" & bmkG & "': " & CStr(rs.Fields(0).Value)
        rs.Close
        Set rs = cn.Execute("SELECT COUNT(*) FROM " & Tbl(DS_PROD, T_RISK) & " WHERE PK_PORTFOLIO_ID = '" & Esc(bmkG) & "'")
        msg = msg & vbLf & "Filas en riesgo con PK_PORTFOLIO_ID='" & bmkG & "': " & CStr(rs.Fields(0).Value)
        rs.Close
    End If
    cn.Close
    MsgBox msg & vbLf & vbLf & "Si esos contadores son > 0, el benchmark se trae igual que la cartera " & _
           "(mismas tablas). Pasame este resultado y lo conecto.", vbInformation, "Diagnostico benchmark"
    Exit Sub
fallo:
    MsgBox "Error consultando el benchmark:" & vbLf & Err.Description, vbExclamation, "Benchmark"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

' Diagnostico de rentabilidad de UN mes. Vuelca a la hoja "Diag_Rentab" TODOS
' los dias que el panel usa para la entidad de B4 en el mes indicado, marca
' fechas DUPLICADAS (la causa n.1 de un mes inflado) y muestra el compuesto y
' la suma simple para poder cuadrarlo con tus rentabilidades diarias.
Public Sub VerRentabMes()
    Dim ws As Worksheet, c0 As Long, lastR As Long, r As Long
    Set ws = Panel()
    ListaEntidades ws
    If Len(Trim(mId1)) = 0 Then MsgBox "Elige una cartera en B4.", vbExclamation, "Rentab. mes": Exit Sub
    c0 = BLK_RET_COL
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then MsgBox "No hay bloque de rentabilidad descargado. Pulsa 'Actualizar query y grafico' con un grupo de Rendimiento.", vbExclamation, "Rentab. mes": Exit Sub

    ' Mes objetivo: por defecto, el ultimo mes presente en el bloque para B4.
    Dim mesMax As String, mesTmp As String, pid As String, dd As Date
    For r = 2 To lastR
        pid = UCase(Trim(CStr(ws.Cells(r, c0).Value)))
        If SlotDe(pid) = 1 Then
            dd = FechaDe(ws.Cells(r, c0 + 1).Value)
            If dd > 0 Then
                mesTmp = Format(dd, "yyyy-mm")
                If mesTmp > mesMax Then mesMax = mesTmp
            End If
        End If
    Next r
    Dim mes As String
    mes = InputBox("Mes a revisar (yyyy-mm) para la entidad de B4:", "Rentab. mes", mesMax)
    If Len(Trim(mes)) = 0 Then Exit Sub
    mes = Trim(mes)

    ' Hoja de volcado.
    Dim dg As Worksheet
    On Error Resume Next
    Set dg = ThisWorkbook.Worksheets("Diag_Rentab")
    On Error GoTo 0
    If dg Is Nothing Then
        Set dg = ThisWorkbook.Worksheets.Add(After:=ws)
        dg.Name = "Diag_Rentab"
    End If
    dg.Cells.Clear
    dg.Range("A1").Value = "Entidad (B4)": dg.Range("B1").Value = mId1
    dg.Range("A2").Value = "Mes":          dg.Range("B2").Value = mes
    dg.Range("A4").Value = "fecha"
    dg.Range("B4").Value = "twr_1d"
    dg.Range("C4").Value = "1+twr_1d"
    dg.Range("D4").Value = "compuesto acum."
    dg.Range("E4").Value = "aviso"
    dg.Range("A4:E4").Font.Bold = True

    Dim vistos As Object: Set vistos = CreateObject("Scripting.Dictionary")
    Dim prod As Double: prod = 1
    Dim suma As Double: suma = 0
    Dim nDias As Long, nDup As Long, outR As Long: outR = 5
    Dim fkey As String, twr As Double, uno As Double
    For r = 2 To lastR
        pid = UCase(Trim(CStr(ws.Cells(r, c0).Value)))
        If SlotDe(pid) = 1 Then
            dd = FechaDe(ws.Cells(r, c0 + 1).Value)
            If dd > 0 Then
                If Format(dd, "yyyy-mm") = mes Then
                    fkey = Format(dd, "yyyy-mm-dd")
                    twr = NumDbl(ws.Cells(r, c0 + 2).Value)
                    uno = 1 + twr
                    prod = prod * uno
                    suma = suma + twr
                    nDias = nDias + 1
                    dg.Cells(outR, 1).Value = fkey
                    dg.Cells(outR, 2).Value = twr
                    dg.Cells(outR, 3).Value = uno
                    dg.Cells(outR, 4).Value = prod - 1
                    If vistos.Exists(fkey) Then
                        dg.Cells(outR, 5).Value = "FECHA DUPLICADA"
                        nDup = nDup + 1
                    Else
                        vistos.Add fkey, 1
                    End If
                    outR = outR + 1
                End If
            End If
        End If
    Next r
    dg.Range(dg.Cells(5, 2), dg.Cells(outR, 4)).NumberFormat = "0.0000%"
    dg.Columns("A:E").AutoFit

    Dim msg As String
    msg = "Mes " & mes & " - entidad " & mId1 & vbLf & vbLf & _
          "Dias usados por el panel: " & nDias & vbLf & _
          "Fechas DUPLICADAS: " & nDup & vbLf & vbLf & _
          "Rentab. COMPUESTA (lo que muestra el panel): " & Format(prod - 1, "0.0000%") & vbLf & _
          "Suma simple de diarias:                      " & Format(suma, "0.0000%") & vbLf & vbLf
    If nDup > 0 Then
        msg = msg & "*** Hay " & nDup & " fecha(s) duplicada(s): la query trae mas de una fila por dia " & _
              "y el compuesto las multiplica, inflando el mes. Hay que deduplicar el bloque RET. ***"
    ElseIf nDias = 0 Then
        msg = msg & "No hay dias de ese mes en el bloque. Revisa el mes o descarga de nuevo."
    Else
        msg = msg & "Sin duplicados. Compara la columna 'fecha'/'twr_1d' de la hoja Diag_Rentab con tus " & _
              "diarias: si coinciden dia a dia, el 3,33% es correcto; si te falta/sobra algun dia, ahi esta la diferencia."
    End If
    MsgBox msg, vbInformation, "Rentab. mes (diagnostico)"
    dg.Activate
End Sub

' ==============  TABLA CONFIGURABLE (entidades x variables)  ===============
' Hoja "Tablas": FILAS = entidades (columna A, desde la fila 7); COLUMNAS =
' variables (cabecera en la fila 6, desde B). "Rellenar tabla" consulta BigQuery
' para esas entidades y rellena cada celda. Variables soportadas:
'   - Rentabilidad por periodo: MTD/YTD/QTD/WTD/1D/1M/3M/6M/1A/2A/3A (o "Rentab YTD").
'     Compuesta desde el inicio del periodo hasta la ultima fecha.
'   - Ano natural: 2025, 2024...  (rentabilidad del ano completo).
'   - Riesgo (ultimo valor): Duracion Modificada, Duracion Macaulay, TIR, VaR, CMR.
'   (gRHdr/gRData declarados arriba, en la seccion de constantes del modulo.)

' Fija el origen de la tabla segun la hoja: la maestra "Tablas" en columna A y
' filas 4-7 (controles en fila 3); las hojas generadas "Tabla N" en columna H y
' filas 5-8, con los controles en la izquierda (Estilo/Total/estado en C6:C8).
Private Sub OrigenTabla(ByVal ws As Worksheet)
    If StrComp(ws.Name, "Tablas", vbTextCompare) = 0 Then
        gCol0 = 1: gRMark = 4: gRMet = 5: gRHdr = 6: gRCap = 0: gRData = 7
        gEstilo = "B3": gTotal = "D3": gEstado = "H3"
    Else
        ' Generadas: config Metrica/Periodo OCULTA (filas 4-6) + una fila de cabecera
        ' de presentacion (7, visible) + datos (8). Controles a la izquierda.
        gCol0 = 8: gRMark = 4: gRMet = 5: gRHdr = 6: gRCap = 7: gRData = 8
        gEstilo = "C7": gTotal = "C8": gEstado = "C9"
    End If
End Sub

Public Sub RellenarTabla(Optional ByVal quiet As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Tablas")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "No encuentro la hoja 'Tablas'.", vbExclamation, "Tablas": Exit Sub
    RellenarTablaEn ws, quiet
End Sub

' Rellena la tabla de DOS NIVELES (Metrica fila 5 + Periodo fila 6) de una hoja
' cualquiera: la propia "Tablas" o una hoja generada con "Crear hoja con tabla".
' Toda la logica opera sobre 'ws', asi la misma maquinaria sirve para ambas.
Public Sub RellenarTablaEn(ByVal ws As Worksheet, Optional ByVal quiet As Boolean = False)
    OrigenTabla ws                          ' fija origen/celdas segun la hoja
    Dim nv As Long, nc As Long, r As Long, i As Long, j As Long

    ' --- Entidades (columna de Entidad) --- (se leen ANTES para poder expandir) ---
    Dim eName() As String, eRow() As Long, eId() As String, ne As Long
    ReDim eName(1 To 5001): ReDim eRow(1 To 5001): ReDim eId(1 To 5001): ne = 0
    r = gRData
    Do While r <= 5000 And Len(Trim(CStr(ws.Cells(r, gCol0).Value))) > 0
        ne = ne + 1: eName(ne) = Trim(CStr(ws.Cells(r, gCol0).Value)): eRow(ne) = r
        eId(ne) = UCase(Trim(IdEntidad(eName(ne))))
        r = r + 1
    Loop
    If ne = 0 Then MsgBox "Escribe al menos una entidad en la columna A (desde la fila " & gRData & ").", vbExclamation, "Tablas": Exit Sub

    Dim inlist As String, seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To ne
        If Len(eId(i)) > 0 And Not seen.Exists(eId(i)) Then
            seen.Add eId(i), 1
            inlist = inlist & IIf(Len(inlist) > 0, ",", "") & "'" & Esc(eId(i)) & "'"
        End If
    Next i
    If Len(inlist) = 0 Then MsgBox "Ninguna entidad de la columna A esta en la hoja 'cartera' (nombre_elemento/id_elemento).", vbExclamation, "Tablas": Exit Sub

    Dim cn As Object, rs As Object      ' (para el manejador 'fallo')
    On Error GoTo fallo
    ' --- Expande los grupos que "crecen" (Periodo = Meses/Trimestres/Anos (crece)):
    ' cada uno genera sus columnas de periodo concretas a la derecha y empuja las
    ' siguientes. Reescribe las filas 4 (marcas), 5 (Metrica) y 6 (Periodo). ---
    ExpandirTabla ws, inlist

    ' --- Cabecera de dos niveles: fila 5 = Metrica, fila 6 = Periodo. Cada columna
    ' se resuelve sintetizando el texto equivalente que ClasificarVar entiende. ---
    Dim vName() As String, vCol() As Long, vMet() As String, vPer() As String
    ReDim vName(1 To 90): ReDim vCol(1 To 90): ReDim vMet(1 To 90): ReDim vPer(1 To 90): nv = 0
    ' Referencia del ultimo cierre (una sola consulta): ano y mes para los periodos
    ' concretos y los de calendario cerrado ("mes/trimestre/ano anterior").
    Dim ultMM As String: ultMM = UltimoMesPerf(inlist)
    If Len(ultMM) < 7 Then ultMM = Format(Date, "yyyy-mm")
    Dim curYear As Long: curYear = CLng(Val(Left(ultMM, 4)))
    Dim curMes As Long: curMes = CLng(Val(Mid(ultMM, 6, 2)))
    Dim kind() As String, pA() As String, pB() As String
    ReDim kind(1 To 90): ReDim pA(1 To 90): ReDim pB(1 To 90)
    Dim needRet As Boolean, needPos As Boolean, needPatMes As Boolean, needBmk As Boolean
    Dim crits As Object: Set crits = CreateObject("Scripting.Dictionary")
    nc = gCol0 + 1
    Do While nc <= gCol0 + 90
        Dim mMet As String, mPer As String
        mMet = Trim(CStr(ws.Cells(gRMet, nc).Value))
        mPer = Trim(CStr(ws.Cells(gRHdr, nc).Value))
        If Len(mMet) = 0 And Len(mPer) = 0 Then Exit Do
        nv = nv + 1
        vCol(nv) = nc: vMet(nv) = mMet: vPer(nv) = mPer
        vName(nv) = SintHeader(mMet, mPer, curYear, curMes) ' texto combinado interno
        ClasificarVar vName(nv), kind(nv), pA(nv), pB(nv)
        Select Case kind(nv)
            Case "ret", "year", "month", "quarter", "vol", "acum": needRet = True
            Case "te", "exc":                   needRet = True: needBmk = True
            Case "patrim", "peso":              needPos = True
            Case "patmes":                      needPatMes = True
            Case "risk":                        If Not crits.Exists(pA(nv)) Then crits.Add pA(nv), 1
        End Select
        nc = nc + 1
    Loop
    If nv = 0 Then MsgBox "Elige Metrica (fila 5) y Periodo (fila 6) en las columnas.", vbExclamation, "Tablas": Exit Sub
    ' Con fila de Total, se descarga patrimonio aunque no se muestre, para
    ' PONDERAR el Total por patrimonio (si no, seria media simple).
    If Fold(ws.Range(gTotal).Value) = "si" Then needPos = True   ' Fila de Total

    Application.StatusBar = "Tablas: consultando BigQuery..."
    ' --- Descarga (RET/RISK/PATRIM/PATMES) a diccionarios en memoria ---
    Dim retColl As Object, difColl As Object, bmkColl As Object, riskV As Object
    Dim patr As Object, patMes As Object, totPatr As Double, defYear As String
    TablaDescarga inlist, needRet, needBmk, needPos, needPatMes, crits, _
        retColl, difColl, bmkColl, riskV, patr, totPatr, patMes, defYear

    ' --- Rellenar celdas ---
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ' Limpia el bloque de valores (y restos de Total/estilos previos) antes de escribir.
    ws.Range(ws.Cells(gRData, gCol0 + 1), ws.Cells(gRData + 500, gCol0 + nv)).ClearContents
    Dim valM() As Variant: ReDim valM(1 To ne, 1 To nv)
    For i = 1 To ne
        For j = 1 To nv
            Dim cellVal As Variant
            cellVal = CalcCelda(kind(j), pA(j), pB(j), eId(i), retColl, difColl, bmkColl, _
                                riskV, patr, totPatr, patMes, defYear)
            valM(i, j) = cellVal
            ws.Cells(eRow(i), vCol(j)).Value = cellVal
            ws.Cells(eRow(i), vCol(j)).NumberFormat = FormatoVar(kind(j), pA(j))
        Next j
    Next i

    ' Fila de Total (opcional, B4="Si"): patrimonio/peso se SUMAN; el resto es
    ' media PONDERADA por patrimonio (si no hay patrimonio, media simple).
    Dim rTot As Long: rTot = gRData + ne
    If Fold(ws.Range(gTotal).Value) = "si" Then          ' Fila de Total
        ws.Cells(rTot, gCol0).Value = "Total"
        ws.Cells(rTot, gCol0).Font.Bold = True
        For j = 1 To nv
            Dim tv As Variant: tv = ""
            Select Case kind(j)
                Case "patrim", "peso", "patmes"
                    Dim sAcc As Double, any1 As Boolean: sAcc = 0: any1 = False
                    For i = 1 To ne
                        If IsNumeric(valM(i, j)) Then sAcc = sAcc + CDbl(valM(i, j)): any1 = True
                    Next i
                    If any1 Then tv = sAcc
                Case "ret", "year", "month", "quarter", "vol", "risk", "acum", "te", "exc"
                    Dim num As Double, wsum As Double, w As Double: num = 0: wsum = 0
                    For i = 1 To ne
                        If IsNumeric(valM(i, j)) Then
                            w = 1: If patr.Exists(eId(i)) Then w = patr(eId(i))
                            num = num + w * CDbl(valM(i, j)): wsum = wsum + w
                        End If
                    Next i
                    If wsum <> 0 Then tv = num / wsum
            End Select
            ws.Cells(rTot, vCol(j)).Value = tv
            ws.Cells(rTot, vCol(j)).NumberFormat = FormatoVar(kind(j), pA(j))
            ws.Cells(rTot, vCol(j)).Font.Bold = True
        Next j
    End If

    ' Cabecera de presentacion (solo hojas generadas, gRCap>0): una fila con nombre
    ' limpio por columna y la columna de Entidad renombrada por tipo. MIXTA: si la
    ' celda ya tiene texto (auto anterior o editado por ti) se respeta; si esta vacia
    ' se escribe el nombre automatico.
    If gRCap > 0 Then
        If Len(Trim(CStr(ws.Cells(gRCap, gCol0).Value))) = 0 Then _
            ws.Cells(gRCap, gCol0).Value = TituloEntidad(eName, eId, ne)
        For j = 1 To nv
            If Len(Trim(CStr(ws.Cells(gRCap, vCol(j)).Value))) = 0 Then _
                ws.Cells(gRCap, vCol(j)).Value = CaptionCol(vMet(j), vPer(j))
        Next j
    End If

    EstiloTabla ws, kind, vCol, nv, gRData, gRData + ne - 1
    Dim nBad As Long: nBad = MarcarEntidades(ws, eName, eId, eRow, ne)
    EstadoTabla ws, ne & " entidades x " & nv & " variables", nBad
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not quiet Then MsgBox "Tabla rellenada: " & ne & " entidades x " & nv & " variables.", vbInformation, "Tablas"
    Exit Sub
fallo:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    MsgBox "Error al rellenar la tabla:" & vbLf & Err.Description, vbExclamation, "Tablas"
End Sub

' Descarga los bloques (RET/RISK/PATRIM/PATMES) a diccionarios en memoria. La usan
' las dos orientaciones de la hoja Tablas (matriz y series). Abre y cierra su propia
' conexion; los errores propagan al manejador del que la llama.
Private Sub TablaDescarga(ByVal inlist As String, ByVal needRet As Boolean, ByVal needBmk As Boolean, _
        ByVal needPos As Boolean, ByVal needPatMes As Boolean, ByVal crits As Object, _
        ByRef retColl As Object, ByRef difColl As Object, ByRef bmkColl As Object, _
        ByRef riskV As Object, ByRef patr As Object, ByRef totPatr As Double, _
        ByRef patMes As Object, ByRef defYear As String)
    Dim cn As Object, rs As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120: cn.CursorLocation = 3: cn.Open CfgConn()

    ' --- RET en memoria: id -> Collection de Array(dserial, factor) ---
    Set retColl = CreateObject("Scripting.Dictionary")
    Set difColl = CreateObject("Scripting.Dictionary")   ' id -> diarios (twr-bmk) para Tracking Error
    Set bmkColl = CreateObject("Scripting.Dictionary")   ' id -> (dserial, 1+twr_bmk) para Exceso
    If needRet Then
        Dim sqlR As String
        sqlR = "SELECT PK_PORTFOLIO_ID, FORMAT_DATE('%Y-%m-%d', PK_FECHA_DATOS)," & _
               " FORMAT('%.10f', CAST(TWR_1D AS FLOAT64))"
        If needBmk Then sqlR = sqlR & ", FORMAT('%.10f', CAST(TWR_1D_BMK AS FLOAT64))"
        sqlR = sqlR & " FROM " & Tbl(DS_PROD, T_PERF) & _
               " WHERE PK_NAV_GNAV='" & CfgNav() & "' AND BENCHMARK='" & CfgBmk() & "'" & _
               " AND PK_PORTFOLIO_ID IN (" & inlist & ")" & AndAsOf("PK_FECHA_DATOS") & _
               " AND PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_PERF) & _
               MaxAsOf() & "), INTERVAL " & CACHE_ANOS & " YEAR)" & _
               " ORDER BY PK_PORTFOLIO_ID, PK_FECHA_DATOS"
        Set rs = cn.Execute(sqlR)
        Dim idc As String, dser As Double, tw1 As Double, bm1 As Double
        Do While Not rs.EOF
            idc = UCase(Trim(CStr(rs.Fields(0).Value)))
            dser = CDbl(FechaDe(CStr(rs.Fields(1).Value)))
            tw1 = NumDbl(rs.Fields(2).Value)
            If Not retColl.Exists(idc) Then retColl.Add idc, New Collection
            retColl(idc).Add Array(dser, 1 + tw1)
            If needBmk Then
                If Not IsNull(rs.Fields(3).Value) Then
                    bm1 = NumDbl(rs.Fields(3).Value)
                    If Not difColl.Exists(idc) Then difColl.Add idc, New Collection
                    difColl(idc).Add Array(dser, tw1 - bm1)
                    If Not bmkColl.Exists(idc) Then bmkColl.Add idc, New Collection
                    bmkColl(idc).Add Array(dser, 1 + bm1)
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If

    ' --- RISK en memoria: "id|crit|var" -> ultimo valor ---
    Set riskV = CreateObject("Scripting.Dictionary")
    If crits.Count > 0 Then
        Dim inCrit As String, kc As Variant
        For Each kc In crits.Keys
            inCrit = inCrit & IIf(Len(inCrit) > 0, ",", "") & "'" & Esc(CStr(kc)) & "'"
        Next kc
        Dim wq As String
        wq = "WHERE " & RISK_COL_FONDOBMK & "='" & CfgFondo() & "' AND PK_PORTFOLIO='" & Esc(CfgPortfolio()) & "'"
        If Len(CfgLtLevel()) > 0 Then wq = wq & " AND PK_LTLEVEL=" & CfgLtLevel()
        wq = wq & " AND PK_CRITERIO_AGREGACION IN (" & inCrit & ")" & _
             " AND PK_PORTFOLIO_ID IN (" & inlist & ")" & AndAsOf("PK_FECHA_DATOS") & _
             " AND PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_RISK) & _
             MaxAsOf() & "), INTERVAL " & CACHE_ANOS & " YEAR)"
        Set rs = cn.Execute("SELECT PK_PORTFOLIO_ID, PK_CRITERIO_AGREGACION, PK_VARIABLE_TARGET," & _
            " FORMAT('%.10f', CAST(VALOR AS FLOAT64)) FROM " & Tbl(DS_PROD, T_RISK) & " " & wq & _
            " ORDER BY PK_PORTFOLIO_ID, PK_FECHA_DATOS")
        Dim kRisk As String, kIdCrit As String, valK As Double
        Do While Not rs.EOF
            kIdCrit = UCase(Trim(CStr(rs.Fields(0).Value))) & "|" & Trim(CStr(rs.Fields(1).Value)) & "|"
            valK = NumDbl(rs.Fields(3).Value)
            kRisk = kIdCrit & Trim(CStr(rs.Fields(2).Value))
            riskV(kRisk) = valK              ' clave exacta id|criterio|variable
            riskV(kIdCrit) = valK            ' reserva id|criterio| (ultimo del criterio, sin filtrar variable)
            rs.MoveNext                      ' ordenado ASC -> el ultimo (mas reciente) gana en ambas
        Loop
        rs.Close
    End If

    ' --- PATRIMONIO en memoria: id -> valoracion total (ultima foto) ---
    Set patr = CreateObject("Scripting.Dictionary")
    totPatr = 0
    If needPos Then
        Dim sqlP As String
        sqlP = "WITH ult AS (SELECT PK_PORTFOLIO_ID, MAX(PK_FECHA_DATOS) AS f FROM " & TblPos() & _
               " WHERE PK_PORTFOLIO_ID IN (" & inlist & ")" & AndAsOf("PK_FECHA_DATOS") & " GROUP BY PK_PORTFOLIO_ID)" & _
               " SELECT p.PK_PORTFOLIO_ID, FORMAT('%.10f', CAST(SUM(p." & PosValor() & ") AS FLOAT64))" & _
               " FROM " & TblPos() & " p JOIN ult ON ult.PK_PORTFOLIO_ID=p.PK_PORTFOLIO_ID" & _
               " AND p.PK_FECHA_DATOS=ult.f GROUP BY p.PK_PORTFOLIO_ID"
        Set rs = cn.Execute(sqlP)
        Do While Not rs.EOF
            patr(UCase(Trim(CStr(rs.Fields(0).Value)))) = NumDbl(rs.Fields(1).Value)
            rs.MoveNext
        Loop
        rs.Close
        Dim kp As Variant
        For Each kp In patr.Keys: totPatr = totPatr + patr(kp): Next kp
    End If

    ' --- PATRIMONIO MENSUAL en memoria: "id|YYYY-MM" -> valoracion total del mes ---
    Set patMes = CreateObject("Scripting.Dictionary")
    If needPatMes Then
        Dim sqlPM As String
        sqlPM = "WITH base AS (SELECT PK_PORTFOLIO_ID, FORMAT_DATE('%Y-%m', PK_FECHA_DATOS) AS mes," & _
                " PK_FECHA_DATOS, " & PosValor() & " AS valor FROM " & TblPos() & _
                " WHERE PK_PORTFOLIO_ID IN (" & inlist & ")" & AndAsOf("PK_FECHA_DATOS") & _
                " AND PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & _
                "), INTERVAL " & CACHE_ANOS & " YEAR))," & _
                " ult AS (SELECT PK_PORTFOLIO_ID, mes, MAX(PK_FECHA_DATOS) AS f FROM base GROUP BY PK_PORTFOLIO_ID, mes)" & _
                " SELECT b.PK_PORTFOLIO_ID, b.mes, FORMAT('%.10f', CAST(SUM(b.valor) AS FLOAT64))" & _
                " FROM base b JOIN ult ON ult.PK_PORTFOLIO_ID=b.PK_PORTFOLIO_ID AND ult.mes=b.mes" & _
                " AND b.PK_FECHA_DATOS=ult.f GROUP BY b.PK_PORTFOLIO_ID, b.mes"
        Set rs = cn.Execute(sqlPM)
        Do While Not rs.EOF
            patMes(UCase(Trim(CStr(rs.Fields(0).Value))) & "|" & Trim(CStr(rs.Fields(1).Value))) = NumDbl(rs.Fields(2).Value)
            rs.MoveNext
        Loop
        rs.Close
    End If
    ' Ano por defecto para columnas "Patrimonio <mes>" sin ano: el mas reciente.
    defYear = ""
    Dim kpm As Variant
    For Each kpm In patMes.Keys
        Dim yy4 As String: yy4 = Mid(CStr(kpm), InStr(CStr(kpm), "|") + 1, 4)
        If yy4 > defYear Then defYear = yy4
    Next kpm
    cn.Close
End Sub

' Valor de una celda (entidad x variable) a partir de (kind, pA, pB) y los
' diccionarios descargados. La usan las dos orientaciones. "" si no hay dato.
Private Function CalcCelda(ByVal kind As String, ByVal pA As String, ByVal pB As String, _
        ByVal id As String, ByVal retColl As Object, ByVal difColl As Object, _
        ByVal bmkColl As Object, ByVal riskV As Object, ByVal patr As Object, _
        ByVal totPatr As Double, ByVal patMes As Object, ByVal defYear As String) As Variant
    CalcCelda = ""
    If Len(id) = 0 Then Exit Function
    Select Case kind
        Case "ret", "year", "month", "quarter"
            If retColl.Exists(id) Then CalcCelda = ValorRet(retColl(id), kind, pA, pB)
        Case "vol"
            If retColl.Exists(id) Then CalcCelda = ValorVol(retColl(id), pA, pB)
        Case "acum"
            If retColl.Exists(id) Then CalcCelda = ValorRet(retColl(id), "acum", pA, pB)
        Case "te"
            If difColl.Exists(id) Then CalcCelda = ValorTE(difColl(id), pB)
        Case "exc"
            If retColl.Exists(id) And bmkColl.Exists(id) Then
                Dim rf As Variant, rb As Variant
                rf = ValorRet(retColl(id), "month", pA, pB)
                rb = ValorRet(bmkColl(id), "month", pA, pB)
                If IsNumeric(rf) And IsNumeric(rb) Then CalcCelda = CDbl(rf) - CDbl(rb)
            End If
        Case "risk"
            Dim kL As String: kL = id & "|" & pA & "|" & pB
            If riskV.Exists(kL) Then CalcCelda = riskV(kL)
        Case "patrim"
            If patr.Exists(id) Then CalcCelda = patr(id)
        Case "peso"
            If patr.Exists(id) And totPatr <> 0 Then CalcCelda = patr(id) / totPatr
        Case "patmes"
            Dim yr3 As String: yr3 = pB: If Len(yr3) <> 4 Then yr3 = defYear
            If Len(yr3) = 4 Then
                Dim kpat As String: kpat = id & "|" & yr3 & "-" & Right("0" & pA, 2)
                If patMes.Exists(kpat) Then CalcCelda = patMes(kpat)
            End If
    End Select
End Function

' Resalta en rojo claro las entidades de la columna A que NO se encontraron en la
' hoja 'cartera' (id vacio); limpia el resto. Devuelve cuantas no se encontraron.
Private Function MarcarEntidades(ByVal ws As Worksheet, ByRef eName() As String, _
        ByRef eId() As String, ByRef eRow() As Long, ByVal ne As Long) As Long
    Dim i As Long, nBad As Long: nBad = 0
    For i = 1 To ne
        If Len(eName(i)) > 0 And Len(eId(i)) = 0 Then
            ws.Cells(eRow(i), gCol0).Interior.Color = RGB(255, 199, 206)   ' rojo claro: no encontrada
            nBad = nBad + 1
        Else
            ws.Cells(eRow(i), gCol0).Interior.ColorIndex = xlNone
        End If
    Next i
    MarcarEntidades = nBad
End Function

' Escribe la linea de estado (celda gEstado): cuando y que se actualizo (y avisos).
Private Sub EstadoTabla(ByVal ws As Worksheet, ByVal resumen As String, ByVal nBad As Long)
    Dim t As String
    t = "Actualizado " & Format(Now, "dd/mm HH:mm") & "  -  " & resumen
    If nBad > 0 Then t = t & "  -  " & nBad & " entidad(es) no encontrada(s) (en rojo)"
    On Error Resume Next
    ws.Range(gEstado).Value = t
    On Error GoTo 0
End Sub

' Nombre de presentacion de una columna a partir de (Metrica, Periodo): el periodo
' "Ultimo" (redundante) no se muestra; los demas se anexan. Ej: Patrimonio/Ultimo ->
' "Patrimonio"; Rentabilidad/YTD -> "Rentabilidad YTD".
Private Function CaptionCol(ByVal met As String, ByVal per As String) As String
    Dim m As String: m = Trim(met)
    Dim p As String: p = Trim(per)
    If Len(m) = 0 Then CaptionCol = p: Exit Function
    If Len(p) = 0 Or Fold(p) = "ultimo" Then
        CaptionCol = m
    Else
        CaptionCol = m & " " & p
    End If
End Function

' Titulo de la columna de Entidad segun el tipo de las entidades: "Fondo" / "Cartera"
' / "Indice" si todas son del mismo tipo; "Entidad" si estan mezcladas o se desconoce.
Private Function TituloEntidad(ByRef eName() As String, ByRef eId() As String, ByVal ne As Long) As String
    Dim i As Long, t As String, comun As String, mixto As Boolean
    comun = "": mixto = False
    For i = 1 To ne
        If Len(eId(i)) > 0 Then
            t = Fold(TipoEntidad(eName(i)))
            If Len(t) > 0 Then
                If Len(comun) = 0 Then
                    comun = t
                ElseIf comun <> t Then
                    mixto = True
                End If
            End If
        End If
    Next i
    If mixto Or Len(comun) = 0 Then
        TituloEntidad = "Entidad"
    ElseIf InStr(comun, "fondo") > 0 Then
        TituloEntidad = "Fondo"
    ElseIf InStr(comun, "cartera") > 0 Then
        TituloEntidad = "Cartera"
    ElseIf InStr(comun, "indice") > 0 Then
        TituloEntidad = "Indice"
    Else
        TituloEntidad = "Entidad"
    End If
End Function

' tipo_elemento de una entidad (por nombre_elemento o id_elemento) en la hoja
' 'cartera'. "" si no se encuentra o no hay columna de tipo.
Private Function TipoEntidad(ByVal nombre As String) As String
    Dim wa As Worksheet, colN As Long, colI As Long, colT As Long, lastR As Long, r As Long
    Dim clave As String: clave = NormNom(nombre)
    If clave = "" Then Exit Function
    Set wa = HojaMaestro()
    If wa Is Nothing Then Exit Function
    colN = ColPorCabecera(wa, ACTIVOS_COL_NOMBRE)
    colI = ColPorCabecera(wa, ACTIVOS_COL_ID)
    colT = ColPorCabecera(wa, "tipo_elemento")
    If colT = 0 Then Exit Function
    lastR = wa.Cells(wa.Rows.Count, IIf(colN > 0, colN, colT)).End(xlUp).Row
    For r = 2 To lastR
        If colN > 0 Then
            If NormNom(wa.Cells(r, colN).Value) = clave Then TipoEntidad = Trim(CStr(wa.Cells(r, colT).Value)): Exit Function
        End If
        If colI > 0 Then
            If NormNom(wa.Cells(r, colI).Value) = clave Then TipoEntidad = Trim(CStr(wa.Cells(r, colT).Value)): Exit Function
        End If
    Next r
End Function

' Clasifica el texto de una cabecera de variable en (kind, pA, pB):
'  kind="year" pA=ano | kind="ret" pA=periodo | kind="risk" pA=criterio pB=variable
Private Sub ClasificarVar(ByVal v As String, ByRef kind As String, ByRef pA As String, ByRef pB As String)
    Dim s As String, f As String
    s = Trim(v): f = Fold(s): kind = "": pA = "": pB = ""
    If InStr(f, "se actualiza") > 0 Then Exit Sub          ' ancla de un marcador -> sin datos
    ' Trimestre: "T1 2026" (compuesto de los 3 meses del trimestre).
    If Left(f, 1) = "t" And Len(f) >= 2 Then
        If IsNumeric(Mid(f, 2, 1)) Then
            kind = "quarter": pA = Mid(f, 2, 1): pB = ""
            Dim tp() As String, ti As Long: tp = Split(f, " ")
            For ti = 1 To UBound(tp)
                If Len(tp(ti)) = 4 Then If IsNumeric(tp(ti)) Then pB = tp(ti)
            Next ti
            Exit Sub
        End If
    End If
    If Len(s) = 4 And IsNumeric(s) Then kind = "year": pA = s: Exit Sub
    If Left(f, 6) = "patrim" Then          ' "Patrimonio" (ultimo) o "Patrimonio Ene 2026" (mensual)
        Dim pp() As String: pp = Split(f, " ")
        Dim mmes As Long, ayr As String, ii As Long: mmes = 0: ayr = ""
        For ii = 1 To UBound(pp)
            If MesNum(pp(ii)) > 0 Then mmes = MesNum(pp(ii))
            If Len(pp(ii)) = 4 Then If IsNumeric(pp(ii)) Then ayr = pp(ii)
        Next ii
        If mmes > 0 Then kind = "patmes": pA = CStr(mmes): pB = ayr Else kind = "patrim"
        Exit Sub
    End If
    If f = "peso" Then kind = "peso": Exit Sub
    If InStr(f, "tracking") = 1 Then    ' "Tracking Error Ene" (mes) o "Tracking Error 2025" (ano)
        kind = "te"
        Dim tep() As String, tw As Long, tmm As Long, tyy As String
        tep = Split(f, " "): tmm = 0: tyy = ""
        For tw = 1 To UBound(tep)
            If MesNum(tep(tw)) > 0 Then tmm = MesNum(tep(tw))
            If Len(tep(tw)) = 4 Then If IsNumeric(tep(tw)) Then tyy = tep(tw)
        Next tw
        If tmm > 0 Then pB = "M" & tmm Else pB = tyy
        Exit Sub
    End If
    If InStr(f, "acum") = 1 Then         ' "Acum Ene 2026" -> rentabilidad acumulada a cierre de mes
        kind = "acum"
        Dim ap() As String, aw As Long, amm As Long, ayy As String
        ap = Split(f, " "): amm = 0: ayy = ""
        For aw = 1 To UBound(ap)
            If MesNum(ap(aw)) > 0 Then amm = MesNum(ap(aw))
            If Len(ap(aw)) = 4 Then If IsNumeric(ap(aw)) Then ayy = ap(aw)
        Next aw
        pA = CStr(amm): pB = ayy
        Exit Sub
    End If
    If InStr(f, "exceso") = 1 Then       ' "Exceso Ene 2026" -> fondo - benchmark del mes
        kind = "exc"
        Dim ep() As String, ew As Long, emm As Long, eyy As String
        ep = Split(f, " "): emm = 0: eyy = ""
        For ew = 1 To UBound(ep)
            If MesNum(ep(ew)) > 0 Then emm = MesNum(ep(ew))
            If Len(ep(ew)) = 4 Then If IsNumeric(ep(ew)) Then eyy = ep(ew)
        Next ew
        pA = CStr(emm): pB = eyy
        Exit Sub
    End If
    If InStr(f, "duraci") = 1 Then kind = "risk": pA = "Duracion": pB = VarTarget(s): Exit Sub
    If f = "tir" Then kind = "risk": pA = "TIR": pB = "": Exit Sub
    If f = "cmr" Then kind = "risk": pA = Cfg("RISK_CRIT_CMR", "CMR"): pB = Cfg("RISK_VAR_CMR", ""): Exit Sub
    If Left(f, 3) = "var" Then kind = "risk": pA = Cfg("RISK_CRIT_VAR", "VaR"): pB = Cfg("RISK_VAR_VAR", ""): Exit Sub
    If InStr(f, "volatil") = 1 Then     ' "Volatilidad [Anualizada|Diaria]" o "Volatilidad Ene|2025"
        kind = "vol"
        pA = IIf(InStr(f, "diaria") > 0, "dia", "anual")   ' por defecto anualizada
        Dim vp() As String, wi As Long, mtmp As Long, ytmp As String
        vp = Split(f, " "): mtmp = 0: ytmp = ""
        For wi = 1 To UBound(vp)
            If MesNum(vp(wi)) > 0 Then mtmp = MesNum(vp(wi))
            If Len(vp(wi)) = 4 Then If IsNumeric(vp(wi)) Then ytmp = vp(wi)
        Next wi
        If mtmp > 0 Then pB = "M" & mtmp Else pB = ytmp   ' el mes manda si hay ambos
        Exit Sub
    End If
    Dim prt() As String, mn As Long
    prt = Split(f, " "): mn = MesNum(prt(0))
    If mn > 0 Then                                  ' columna por mes: "Ene", "Ene 2025"...
        kind = "month": pA = CStr(mn): pB = ""
        If UBound(prt) >= 1 Then If IsNumeric(prt(1)) Then pB = prt(1)
        Exit Sub
    End If
    Dim tok As String: tok = f
    If InStr(tok, "rent") = 1 Then
        Dim sp As Long: sp = InStr(tok, " ")
        If sp > 0 Then tok = Trim(Mid(tok, sp + 1)) Else tok = ""
    End If
    tok = UCase(Replace(tok, " ", ""))
    If Len(tok) = 4 And IsNumeric(tok) Then kind = "year": pA = tok: Exit Sub
    If EsPeriodoTok(tok) Then kind = "ret": pA = tok
End Sub

Private Function EsPeriodoTok(ByVal t As String) As Boolean
    Select Case t
        Case "MTD", "QTD", "YTD", "WTD", "1D", "DTD": EsPeriodoTok = True
        Case Else
            If Len(t) >= 2 Then
                If IsNumeric(Left(t, Len(t) - 1)) _
                   And (Right(t, 1) = "M" Or Right(t, 1) = "A") Then EsPeriodoTok = True
            End If
    End Select
End Function

' Numero de mes (1-12) a partir del nombre en espanol (abreviado o completo). 0 si no.
Private Function MesNum(ByVal s As String) As Long
    Select Case Left(Fold(s), 3)
        Case "ene": MesNum = 1
        Case "feb": MesNum = 2
        Case "mar": MesNum = 3
        Case "abr": MesNum = 4
        Case "may": MesNum = 5
        Case "jun": MesNum = 6
        Case "jul": MesNum = 7
        Case "ago": MesNum = 8
        Case "sep": MesNum = 9
        Case "oct": MesNum = 10
        Case "nov": MesNum = 11
        Case "dic": MesNum = 12
        Case Else:  MesNum = 0
    End Select
End Function

' Abreviatura de mes en espanol (1->Ene ... 12->Dic).
Private Function MesAbbr(ByVal m As Long) As String
    Dim a As Variant: a = Array("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")
    If m >= 1 And m <= 12 Then MesAbbr = CStr(a(m - 1))
End Function

' Ultimo mes con datos de performance (YYYY-MM) para las entidades dadas. "" si falla.
Private Function UltimoMesPerf(ByVal inlist As String) As String
    Dim cn As Object, rs As Object
    On Error GoTo fin
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 60: cn.CursorLocation = 3: cn.Open CfgConn()
    Set rs = cn.Execute("SELECT FORMAT_DATE('%Y-%m', MAX(PK_FECHA_DATOS)) FROM " & Tbl(DS_PROD, T_PERF) & _
        " WHERE PK_NAV_GNAV='" & CfgNav() & "' AND BENCHMARK='" & CfgBmk() & _
        "' AND PK_PORTFOLIO_ID IN (" & inlist & ")" & AndAsOf("PK_FECHA_DATOS"))
    If Not rs.EOF Then UltimoMesPerf = Trim(CStr(rs.Fields(0).Value))
fin:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Function

' Ano de la ultima fecha con datos de performance (para las cabeceras que llevan ano).
Private Function AnoDatos(ByVal inlist As String) As Long
    Dim mm As String: mm = UltimoMesPerf(inlist)
    If Len(mm) < 7 Then mm = Format(Date, "yyyy-mm")
    AnoDatos = CLng(Val(Left(mm, 4)))
End Function

' Granularidad de un Periodo "que crece": "Meses (crece)"->M, "Trimestres (crece)"->T,
' "Anos (crece)"->A. "" si el periodo no es de los que crecen.
Private Function GranCrece(ByVal per As String) As String
    Dim f As String: f = Fold(per)
    If InStr(f, "crece") = 0 Then Exit Function
    If InStr(f, "trimestr") > 0 Then
        GranCrece = "T"
    ElseIf InStr(f, "ano") > 0 Then
        GranCrece = "A"
    Else
        GranCrece = "M"
    End If
End Function

' Lista de periodos concretos de un grupo que crece, hasta el ultimo con datos.
Private Sub PeriodosCrece(ByVal gran As String, ByVal yData As Long, ByVal mData As Long, _
        ByRef pers() As String, ByRef np As Long)
    ReDim pers(1 To 60): np = 0
    Dim k As Long
    Select Case gran
        Case "M"
            For k = 1 To mData
                np = np + 1: pers(np) = MesAbbr(k)
            Next k
        Case "T"
            Dim curQ As Long: curQ = Int((mData - 1) / 3) + 1
            For k = 1 To curQ
                np = np + 1: pers(np) = "T" & k
            Next k
        Case "A"
            For k = 4 To 0 Step -1
                np = np + 1: pers(np) = CStr(yData - k)
            Next k
    End Select
End Sub

' Expande los grupos de Periodo "que crecen". Cabecera de dos niveles: fila 5 =
' Metrica, fila 6 = Periodo. Si el Periodo de una columna es "Meses/Trimestres/Anos
' (crece)" se convierte en columnas concretas (Ene, Feb.. / T1.. / 2022..) que crecen
' cada mes, empujando las columnas siguientes. La fila 4 (oculta) marca "CRECE:<M|T|A>"
' en la primera columna del grupo y "AUTO" en las generadas, para regenerarlas (crecer)
' en la pasada siguiente sin duplicar ni perder la definicion del grupo.
Private Sub ExpandirTabla(ByVal ws As Worksheet, ByVal inlist As String)
    ' 1) Spec = columnas NO generadas (fila 4 <> "AUTO"): met, per y granularidad-si-crece.
    Dim sMet() As String, sPer() As String, sGran() As String, ns As Long
    ReDim sMet(1 To 90): ReDim sPer(1 To 90): ReDim sGran(1 To 90): ns = 0
    Dim c As Long, mk As String, mt As String, pr As String
    For c = gCol0 + 1 To gCol0 + 90
        mk = Trim(CStr(ws.Cells(gRMark, c).Value))
        mt = Trim(CStr(ws.Cells(gRMet, c).Value))
        pr = Trim(CStr(ws.Cells(gRHdr, c).Value))
        If mk <> "AUTO" Then
            If Len(mt) > 0 Or Len(pr) > 0 Then
                ns = ns + 1: sMet(ns) = mt: sPer(ns) = pr
                Dim g As String: g = GranCrece(pr)
                If Left(mk, 6) = "CRECE:" Then g = Mid(mk, 7, 1)
                sGran(ns) = g
            End If
        End If
    Next c
    If ns = 0 Then Exit Sub

    Dim mm As String: mm = UltimoMesPerf(inlist)
    If Len(mm) < 7 Then mm = Format(Date, "yyyy-mm")
    Dim yData As Long: yData = CLng(Val(Left(mm, 4)))
    Dim mData As Long: mData = CLng(Val(Mid(mm, 6, 2)))

    ' 2) Lista final de columnas: los grupos que crecen se expanden a concretas.
    Dim oMet() As String, oPer() As String, oMk() As String, no As Long
    ReDim oMet(1 To 200): ReDim oPer(1 To 200): ReDim oMk(1 To 200): no = 0
    Dim i As Long, k As Long
    For i = 1 To ns
        If no >= 88 Then Exit For          ' tope de seguridad (arrays 1..200)
        If Len(sGran(i)) > 0 Then
            Dim pers() As String, np As Long
            PeriodosCrece sGran(i), yData, mData, pers, np
            For k = 1 To np
                If no >= 88 Then Exit For
                no = no + 1: oMet(no) = sMet(i): oPer(no) = pers(k)
                oMk(no) = IIf(k = 1, "CRECE:" & sGran(i), "AUTO")
            Next k
        Else
            no = no + 1: oMet(no) = sMet(i): oPer(no) = sPer(i): oMk(no) = ""
        End If
    Next i

    ' 3) Reescribe las filas de marca/metrica/periodo desde la 1a columna de datos.
    Application.EnableEvents = False
    ws.Range(ws.Cells(gRMark, gCol0 + 1), ws.Cells(gRMark, gCol0 + 200)).ClearContents
    ws.Range(ws.Cells(gRMet, gCol0 + 1), ws.Cells(gRMet, gCol0 + 200)).ClearContents
    ws.Range(ws.Cells(gRHdr, gCol0 + 1), ws.Cells(gRHdr, gCol0 + 200)).ClearContents
    For i = 1 To no
        ws.Cells(gRMet, gCol0 + i).Value = oMet(i)
        ws.Cells(gRHdr, gCol0 + i).Value = oPer(i)
        If Len(oMk(i)) > 0 Then ws.Cells(gRMark, gCol0 + i).Value = oMk(i)
    Next i
    Application.EnableEvents = True
End Sub

' Un token de trimestre: "T1".."T4".
Private Function EsTrimTok(ByVal per As String) As Boolean
    Dim f As String: f = Fold(per)
    If Left(f, 1) = "t" And Len(f) >= 2 Then
        If IsNumeric(Mid(f, 2, 1)) Then EsTrimTok = True
    End If
End Function

' Periodo de CALENDARIO CERRADO relativo al ultimo cierre: "M"=mes anterior,
' "Q"=trimestre anterior, "Y"=ano anterior. "" si no es uno de esos.
Private Function PerAnterior(ByVal per As String) As String
    Select Case Fold(per)
        Case "mes anterior":       PerAnterior = "M"
        Case "trimestre anterior": PerAnterior = "Q"
        Case "ano anterior":       PerAnterior = "Y"   ' Fold quita la tilde de "Ano"
    End Select
End Function

' Cabecera (interpretable por ClasificarVar) de un periodo "anterior" resuelto a
' calendario cerrado, relativo al ultimo cierre (curYear/curMes). Lleva el ano
' EXPLICITO para no fallar al cruzar el 1-enero. "" si la metrica no admite ese
' tramo (columna omitida), igual que el resto de combinaciones no aplicables.
Private Function HeaderAnterior(ByVal met As String, ByVal clase As String, _
                                ByVal curYear As Long, ByVal curMes As Long) As String
    Dim m As Long, q As Long, cq As Long, y As Long, fm As String
    fm = Fold(met)
    If curMes < 1 Or curMes > 12 Then curMes = 12
    Select Case clase
        Case "M"
            m = curMes - 1: y = curYear
            If m < 1 Then m = 12: y = y - 1
            Select Case True
                Case fm = "rentabilidad":   HeaderAnterior = MesAbbr(m) & " " & y
                Case fm = "patrimonio":     HeaderAnterior = "Patrimonio " & MesAbbr(m) & " " & y
                Case fm = "acumulada":      HeaderAnterior = "Acum " & MesAbbr(m) & " " & y
                Case fm = "exceso":         HeaderAnterior = "Exceso " & MesAbbr(m) & " " & y
            End Select
        Case "Q"
            cq = Int((curMes - 1) / 3) + 1                ' trimestre en curso (1..4)
            q = cq - 1: y = curYear
            If q < 1 Then q = 4: y = y - 1
            Select Case True
                Case fm = "rentabilidad": HeaderAnterior = "T" & q & " " & y
                Case fm = "patrimonio":   HeaderAnterior = "Patrimonio " & MesAbbr(3 * q) & " " & y
            End Select
        Case "Y"
            y = curYear - 1
            Select Case True
                Case fm = "rentabilidad":   HeaderAnterior = CStr(y)
                Case fm = "patrimonio":     HeaderAnterior = "Patrimonio Dic " & y
                Case fm = "volatilidad":    HeaderAnterior = "Volatilidad " & y
                Case fm = "tracking error": HeaderAnterior = "Tracking Error " & y
            End Select
    End Select
End Function

' Sintetiza el texto de cabecera combinado (el que ClasificarVar entiende) a partir
' de la Metrica (fila 5) y el Periodo (fila 6). "" si la combinacion no aplica.
Private Function SintHeader(ByVal met As String, ByVal per As String, ByVal curYear As Long, ByVal curMes As Long) As String
    If Len(met) = 0 Then Exit Function
    If GranCrece(per) <> "" Then Exit Function        ' marcador sin expandir -> vacio
    ' Periodos de calendario cerrado (mes/trimestre/ano anterior): se resuelven a
    ' una cabecera concreta con ano explicito antes del resto de casos.
    Dim pa As String: pa = PerAnterior(per)
    If pa <> "" Then SintHeader = HeaderAnterior(met, pa, curYear, curMes): Exit Function
    Dim fm As String: fm = Fold(met)
    Dim fp As String: fp = Fold(per)
    Dim tipo As String, val As String
    tipo = "": val = ""
    If Len(fp) = 0 Or fp = "ultimo" Then
        tipo = "fecha"
    ElseIf MesNum(per) > 0 Then
        tipo = "mes": val = MesAbbr(MesNum(per))
    ElseIf EsTrimTok(per) Then
        tipo = "trim": val = Mid(Fold(per), 2, 1)
    ElseIf Len(per) = 4 And IsNumeric(per) Then
        tipo = "ano": val = per
    ElseIf fp = "ytd" Then
        If InStr(fm, "volatil") > 0 Or InStr(fm, "tracking") > 0 Then
            tipo = "fecha"                            ' vol/TE "YTD" = anualizada del ano en curso
        Else
            tipo = "periodo": val = "YTD"
        End If
    Else
        tipo = "periodo": val = UCase(Trim(per))      ' MTD, 1M, 3M, 6M, 1A, 3A...
    End If
    SintHeader = HeaderMetrica(met, tipo, val, curYear)
End Function

' Texto de cabecera para (metrica, tipo de tramo, valor) que ClasificarVar sabe
' interpretar. "" si esa combinacion no aplica (se omite la columna).
Private Function HeaderMetrica(ByVal met As String, ByVal tipo As String, _
                               ByVal val As String, ByVal curYear As Long) As String
    Dim fm As String: fm = Fold(met)
    Select Case True
        Case fm = "rentabilidad"
            Select Case tipo
                Case "mes":     HeaderMetrica = val                       ' "Ene"
                Case "trim":    HeaderMetrica = "T" & val & " " & curYear  ' "T1 2026"
                Case "ano":     HeaderMetrica = val                       ' "2025"
                Case "periodo": HeaderMetrica = val
                Case Else:      HeaderMetrica = "YTD"
            End Select
        Case fm = "volatilidad"
            Select Case tipo
                Case "mes":  HeaderMetrica = "Volatilidad " & val
                Case "ano":  HeaderMetrica = "Volatilidad " & val
                Case "periodo", "trim": HeaderMetrica = ""
                Case Else:   HeaderMetrica = "Volatilidad Anualizada"
            End Select
        Case fm = "patrimonio"
            Select Case tipo
                Case "mes":   HeaderMetrica = "Patrimonio " & val & " " & curYear
                Case "trim":  HeaderMetrica = "Patrimonio " & MesAbbr(3 * CLng(val)) & " " & curYear
                Case "ano":   HeaderMetrica = "Patrimonio Dic " & val
                Case "fecha": HeaderMetrica = "Patrimonio"
            End Select
        Case fm = "peso"
            If tipo = "fecha" Then HeaderMetrica = "Peso"
        Case fm = "tracking error"
            Select Case tipo
                Case "mes":   HeaderMetrica = "Tracking Error " & val
                Case "ano":   HeaderMetrica = "Tracking Error " & val
                Case "fecha": HeaderMetrica = "Tracking Error"   ' anualizado del ano en curso
                Case Else:    HeaderMetrica = ""
            End Select
        Case fm = "acumulada"                                ' rentabilidad acumulada a cierre de mes
            If tipo = "mes" Then HeaderMetrica = "Acum " & val & " " & curYear
        Case fm = "exceso"                                   ' rentabilidad fondo - benchmark del mes
            If tipo = "mes" Then HeaderMetrica = "Exceso " & val & " " & curYear
        Case Else                                            ' riesgo: VaR/TIR/CMR/Duracion
            If tipo = "fecha" Then HeaderMetrica = met       ' por ahora, solo "a la fecha"
    End Select
End Function

' Ultima fecha (serial) de una Collection de (dserial, 1+twr). 0 si vacia.
Private Function MaxSerial(ByVal coll As Collection) As Double
    Dim it As Variant
    For Each it In coll
        If it(0) > MaxSerial Then MaxSerial = it(0)
    Next it
End Function

' Rentabilidad compuesta de una entidad para un periodo/ano/mes, desde su
' Collection de (dserial, 1+twr). Devuelve "" si no hay datos en ese tramo.
Private Function ValorRet(ByVal coll As Collection, ByVal kind As String, _
                          ByVal tok As String, Optional ByVal pB As String = "") As Variant
    Dim prod As Double: prod = 1
    Dim hay As Boolean, it As Variant
    If kind = "year" Then
        Dim y As Long: y = CLng(tok)
        For Each it In coll
            If it(0) > 0 Then If Year(CDate(it(0))) = y Then prod = prod * it(1): hay = True
        Next it
    ElseIf kind = "month" Then
        Dim mn As Long: mn = CLng(tok)
        Dim yy As Long
        If Len(pB) = 4 And IsNumeric(pB) Then yy = CLng(pB) Else yy = Year(CDate(MaxSerial(coll)))
        For Each it In coll
            If it(0) > 0 Then If Year(CDate(it(0))) = yy And Month(CDate(it(0))) = mn Then prod = prod * it(1): hay = True
        Next it
    ElseIf kind = "quarter" Then
        Dim q As Long: q = CLng(tok)
        Dim yq As Long, m1 As Long, m2 As Long
        If Len(pB) = 4 And IsNumeric(pB) Then yq = CLng(pB) Else yq = Year(CDate(MaxSerial(coll)))
        m1 = 3 * q - 2: m2 = 3 * q
        For Each it In coll
            If it(0) > 0 Then If Year(CDate(it(0))) = yq And Month(CDate(it(0))) >= m1 And Month(CDate(it(0))) <= m2 Then prod = prod * it(1): hay = True
        Next it
    ElseIf kind = "acum" Then                     ' acumulada a cierre de mes: meses 1..tok del ano
        Dim am As Long: am = CLng(tok)
        Dim ay As Long
        If Len(pB) = 4 And IsNumeric(pB) Then ay = CLng(pB) Else ay = Year(CDate(MaxSerial(coll)))
        For Each it In coll
            If it(0) > 0 Then If Year(CDate(it(0))) = ay And Month(CDate(it(0))) <= am Then prod = prod * it(1): hay = True
        Next it
    Else
        Dim dMax As Double: dMax = MaxSerial(coll)
        If dMax = 0 Then ValorRet = "": Exit Function
        Dim ini As Double: ini = CDbl(InicioVentana(CDate(dMax), tok, "mensual"))
        Dim anchored As Boolean
        anchored = (tok = "MTD" Or tok = "QTD" Or tok = "YTD" Or tok = "WTD" Or tok = "1D" Or tok = "DTD")
        For Each it In coll
            If it(0) <= dMax And ((anchored And it(0) >= ini) Or ((Not anchored) And it(0) > ini)) Then
                prod = prod * it(1): hay = True
            End If
        Next it
    End If
    If hay Then ValorRet = prod - 1 Else ValorRet = ""
End Function

' Volatilidad de los retornos DIARIOS (desviacion tipica muestral) del ano en
' curso (YTD). modo="anual" -> anualizada (x raiz de 252). "" si <2 datos.
' Volatilidad de los retornos diarios en una ventana: win="" -> ano en curso (YTD);
' win="M<n>" -> mes n del ano en curso; win="YYYY" -> ese ano. modo="anual" -> x raiz(252).
Private Function ValorVol(ByVal coll As Collection, ByVal modo As String, Optional ByVal win As String = "") As Variant
    Dim dMax As Double: dMax = MaxSerial(coll)
    If dMax = 0 Then ValorVol = "": Exit Function
    Dim yy As Long: yy = Year(CDate(dMax))
    Dim n As Long, s As Double, s2 As Double, x As Double, it As Variant, okd As Boolean
    For Each it In coll
        If it(0) > 0 Then
            If Len(win) = 0 Then
                okd = (Year(CDate(it(0))) = yy)
            ElseIf Left(win, 1) = "M" Then
                okd = (Year(CDate(it(0))) = yy And Month(CDate(it(0))) = CLng(Mid(win, 2)))
            Else
                okd = (Year(CDate(it(0))) = CLng(win))
            End If
            If okd Then x = it(1) - 1: n = n + 1: s = s + x: s2 = s2 + x * x
        End If
    Next it
    If n < 2 Then ValorVol = "": Exit Function
    Dim vv As Double: vv = (s2 - s * s / n) / (n - 1)
    If vv < 0 Then vv = 0
    Dim sd As Double: sd = Sqr(vv)
    If Fold(modo) = "anual" Then sd = sd * Sqr(252)
    ValorVol = sd
End Function

' Tracking Error: desviacion tipica muestral de los diferenciales diarios
' (fondo - benchmark) en la ventana, anualizada (x raiz de 252). win="" -> ano en
' curso; win="M<n>" -> mes n del ano en curso; win="YYYY" -> ese ano. "" si <2.
Private Function ValorTE(ByVal coll As Collection, Optional ByVal win As String = "") As Variant
    Dim dMax As Double: dMax = MaxSerial(coll)
    If dMax = 0 Then ValorTE = "": Exit Function
    Dim yy As Long: yy = Year(CDate(dMax))
    Dim n As Long, s As Double, s2 As Double, x As Double, it As Variant, okd As Boolean
    For Each it In coll
        If it(0) > 0 Then
            If Len(win) = 0 Then
                okd = (Year(CDate(it(0))) = yy)
            ElseIf Left(win, 1) = "M" Then
                okd = (Year(CDate(it(0))) = yy And Month(CDate(it(0))) = CLng(Mid(win, 2)))
            Else
                okd = (Year(CDate(it(0))) = CLng(win))
            End If
            If okd Then x = it(1): n = n + 1: s = s + x: s2 = s2 + x * x
        End If
    Next it
    If n < 2 Then ValorTE = "": Exit Function
    Dim vv As Double: vv = (s2 - s * s / n) / (n - 1)
    If vv < 0 Then vv = 0
    ValorTE = Sqr(vv) * Sqr(252)
End Function

Private Function FormatoVar(ByVal kind As String, ByVal pA As String) As String
    Select Case kind
        Case "risk":             FormatoVar = IIf(pA = "Duracion", "0.000", "0.00")
        Case "patrim", "patmes": FormatoVar = "#,##0"
        Case "peso":             FormatoVar = "0.0%"
        Case Else:               FormatoVar = "0.00%"
    End Select
End Function

' Estilo de las columnas de rentabilidad/ano segun B3: "Mapa de calor" (escala
' 3 colores) o "Barras" (barras de datos en celda). Otro valor -> sin estilo.
Private Sub EstiloTabla(ByVal ws As Worksheet, ByRef kind() As String, _
        ByRef vCol() As Long, ByVal nv As Long, ByVal r1 As Long, ByVal r2 As Long)
    Dim est As String: est = Fold(ws.Range(gEstilo).Value)   ' Estilo
    Dim j As Long
    For j = 1 To nv
        Dim rng As Range: Set rng = ws.Range(ws.Cells(r1, vCol(j)), ws.Cells(r2, vCol(j)))
        rng.FormatConditions.Delete
        If kind(j) = "ret" Or kind(j) = "year" Or kind(j) = "month" Or kind(j) = "quarter" _
           Or kind(j) = "acum" Or kind(j) = "exc" Then
            If InStr(est, "calor") > 0 Then
                Dim cs As ColorScale: Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
                cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
                cs.ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
                cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
                cs.ColorScaleCriteria(2).Value = 50
                cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
                cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                cs.ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
            ElseIf InStr(est, "barra") > 0 Then
                Dim db As Databar: Set db = rng.FormatConditions.AddDatabar()
                db.MinPoint.Modify newtype:=xlConditionValueLowestValue
                db.MaxPoint.Modify newtype:=xlConditionValueHighestValue
                db.BarColor.Color = RGB(99, 142, 198)
            End If
        End If
    Next j
End Sub

' ==================  TABLA DE POSICIONES (holdings)  =======================
' Hoja "Posiciones": una foto (ultima fecha) de las posiciones de UN fondo/
' cartera (B3), una fila por valor, con las columnas que pongas en la cabecera
' (fila 6, desde B). Columnas soportadas de fabrica: Peso, Importe, Sector,
' Industria, Pais, Zona, Divisa, Tipo activo, Rating, TER. Y por config (maestro):
' ISIN, Ticker, Nombre, Yield, Duracion, Mercado, Dividendo, Plazo.
' (PS_HDR/PS_ROW0 declarados arriba, en la seccion de constantes del modulo.)

Public Sub RellenarPosiciones(Optional ByVal quiet As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Posiciones")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "No encuentro la hoja 'Posiciones'.", vbExclamation, "Posiciones": Exit Sub

    Dim id As String: id = UCase(Trim(IdEntidad(CStr(ws.Range("B3").Value))))
    If Len(id) = 0 Then MsgBox "Elige un fondo/cartera en B3 (nombre_elemento o id_elemento).", vbExclamation, "Posiciones": Exit Sub

    Dim nv As Long, c As Long, vName() As String, vCol() As Long
    ReDim vName(1 To 80): ReDim vCol(1 To 80): nv = 0
    c = 2
    Do While c <= 80 And Len(Trim(CStr(ws.Cells(PS_HDR, c).Value))) > 0
        nv = nv + 1: vName(nv) = Trim(CStr(ws.Cells(PS_HDR, c).Value)): vCol(nv) = c
        c = c + 1
    Loop
    If nv = 0 Then MsgBox "Pon al menos una columna en la fila " & PS_HDR & " (desde B).", vbExclamation, "Posiciones": Exit Sub

    Dim kindP() As String, gics() As String, fmt() As String, aliasIdx() As Long
    ReDim kindP(1 To nv): ReDim gics(1 To nv): ReDim fmt(1 To nv): ReDim aliasIdx(1 To nv)
    Dim selExtra As String, nAlias As Long: nAlias = 0
    Dim j As Long, expr As String, g As String, ff As String
    For j = 1 To nv
        kindP(j) = MapaPos(vName(j), expr, g, ff)
        gics(j) = g: fmt(j) = ff: aliasIdx(j) = 0
        If (kindP(j) = "attr" Or kindP(j) = "attrnum") And Len(expr) > 0 Then
            nAlias = nAlias + 1: aliasIdx(j) = nAlias
            selExtra = selExtra & ", ANY_VALUE(" & expr & ") AS c" & nAlias
        End If
    Next j

    Dim kPos As String: kPos = Cfg("JOIN_KEY_POS", "PK_SECURITY_IK")
    Dim sql As String
    sql = "WITH ult AS (SELECT MAX(PK_FECHA_DATOS) AS f FROM " & TblPos() & _
          " WHERE PK_PORTFOLIO_ID='" & Esc(id) & "'" & AndAsOf("PK_FECHA_DATOS") & ")" & vbLf & _
          "SELECT FORMAT('%.10f', CAST(SUM(p." & PosValor() & ") AS FLOAT64)) AS valor" & selExtra & vbLf & _
          "FROM " & TblPos() & " p" & vbLf & JoinValores() & _
          "WHERE p.PK_PORTFOLIO_ID='" & Esc(id) & "' AND p.PK_FECHA_DATOS=(SELECT f FROM ult)" & vbLf & _
          "GROUP BY p." & kPos & vbLf & _
          "ORDER BY SUM(p." & PosValor() & ") DESC" & vbLf & "LIMIT 1000"

    Dim cn As Object, rs As Object, data As Variant
    On Error GoTo fallo
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120: cn.CursorLocation = 3: cn.Open CfgConn()
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then data = rs.GetRows()
    rs.Close: cn.Close
    If IsEmpty(data) Then MsgBox "Sin posiciones para '" & id & "'.", vbInformation, "Posiciones": Exit Sub

    Dim nr As Long: nr = UBound(data, 2) + 1
    Dim total As Double, r As Long
    For r = 0 To nr - 1: total = total + NumDbl(data(0, r)): Next r

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(PS_ROW0, 1), ws.Cells(PS_ROW0 + 5000, 1 + nv)).ClearContents
    Dim rr As Long, valp As Double, cellVal As Variant
    For r = 0 To nr - 1
        rr = PS_ROW0 + r: valp = NumDbl(data(0, r))
        For j = 1 To nv
            cellVal = ""
            Select Case kindP(j)
                Case "peso":    If total <> 0 Then cellVal = valp / total
                Case "importe": cellVal = valp
                Case "attr", "attrnum"
                    If aliasIdx(j) > 0 Then
                        cellVal = data(aliasIdx(j), r)
                        If kindP(j) = "attrnum" Then cellVal = NumDbl(cellVal)
                        If gics(j) = "sector" Then cellVal = EtiquetaClasif("Sector", CStr(data(aliasIdx(j), r)))
                        If gics(j) = "ind" Then cellVal = EtiquetaClasif("Industria", CStr(data(aliasIdx(j), r)))
                    End If
            End Select
            ws.Cells(rr, vCol(j)).Value = cellVal
            ws.Cells(rr, vCol(j)).NumberFormat = fmt(j)
        Next j
    Next r

    Dim rt As Long: rt = PS_ROW0 + nr
    ws.Cells(rt, 1).Value = "Total": ws.Cells(rt, 1).Font.Bold = True
    For j = 1 To nv
        Dim tv As Variant: tv = ""
        Select Case kindP(j)
            Case "peso":    tv = 1
            Case "importe": tv = total
            Case "attrnum"
                Dim num As Double, wsum As Double: num = 0: wsum = 0
                If aliasIdx(j) > 0 Then
                    For r = 0 To nr - 1
                        If Len(Trim(CStr(data(aliasIdx(j), r)))) > 0 Then
                            num = num + NumDbl(data(0, r)) * NumDbl(data(aliasIdx(j), r))
                            wsum = wsum + NumDbl(data(0, r))
                        End If
                    Next r
                End If
                If wsum <> 0 Then tv = num / wsum
        End Select
        ws.Cells(rt, vCol(j)).Value = tv
        ws.Cells(rt, vCol(j)).NumberFormat = fmt(j)
        ws.Cells(rt, vCol(j)).Font.Bold = True
    Next j
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not quiet Then MsgBox nr & " posiciones cargadas para '" & id & "'.", vbInformation, "Posiciones"
    Exit Sub
fallo:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    MsgBox "Error al cargar posiciones:" & vbLf & Err.Description & vbLf & vbLf & _
           "Las columnas del maestro (ISIN/Ticker/Yield/Duracion...) se configuran en 'config' (MSTR_*).", _
           vbExclamation, "Posiciones"
End Sub

' Mapea una cabecera de columna de posiciones a (expr SQL, gics, formato) y
' devuelve el tipo: peso/importe/attr/attrnum, o "" si no se reconoce/configura.
Private Function MapaPos(ByVal h As String, ByRef expr As String, ByRef gics As String, ByRef fmt As String) As String
    Dim f As String: f = Fold(h): expr = "": gics = "": fmt = "General"
    Select Case True
        Case f = "peso":                       MapaPos = "peso": fmt = "0.00%"
        Case f = "importe":                    MapaPos = "importe": fmt = "#,##0"
        Case f = "sector" Or f = "segmento":   expr = "v." & Cfg("SECTOR_COL", SECTOR_COL): gics = "sector": MapaPos = "attr"
        Case f = "industria":                  expr = "v." & Cfg("IND_COL", IND_COL): gics = "ind": MapaPos = "attr"
        Case f = "pais" Or f = "geografia":    expr = "v." & Cfg("PAIS_COL", "FCCOUNTRY"): MapaPos = "attr"
        Case f = "zona" Or f = "continente":   expr = "v." & Cfg("GEO_COL", GEO_COL): MapaPos = "attr"
        Case f = "divisa":                     expr = "v." & Cfg("DIV_COL", DIV_COL): MapaPos = "attr"
        Case f = "tipo activo" Or f = "activo" Or f = "tipo de activo": expr = "v." & Cfg("ACTIVO_COL", "INSTRUMENT_TYPE"): MapaPos = "attr"
        Case f = "rating":                     expr = "v." & Cfg("RATING_COL", RATING_COL): MapaPos = "attr"
        Case f = "ter":                        expr = "v." & Cfg("TER_COL", TER_COL): fmt = "0.000": MapaPos = "attrnum"
        Case f = "isin":                       expr = ExprCfg("MSTR_ISIN_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case f = "ticker":                     expr = ExprCfg("MSTR_TICKER_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case f = "nombre":                     expr = ExprCfg("MSTR_NAME_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case f = "yield":                      expr = ExprCfg("MSTR_YIELD_COL"): fmt = "0.00": MapaPos = IIf(Len(expr) > 0, "attrnum", "")
        Case f = "duracion":                   expr = ExprCfg("MSTR_DUR_COL"): fmt = "0.000": MapaPos = IIf(Len(expr) > 0, "attrnum", "")
        Case f = "mercado":                    expr = ExprCfg("MSTR_MKT_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case f = "dividendo":                  expr = ExprCfg("MSTR_DIV_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case f = "plazo" Or f = "vencimiento": expr = ExprCfg("MSTR_PLAZO_COL"): MapaPos = IIf(Len(expr) > 0, "attr", "")
        Case Else:                             MapaPos = ""
    End Select
End Function

Private Function ExprCfg(ByVal clave As String) As String
    Dim col As String: col = Cfg(clave, "")
    If Len(col) > 0 Then ExprCfg = "v." & col Else ExprCfg = ""
End Function

' =======================  INSTALADOR DE BOTONES  ===========================
' Ejecuta este macro UNA vez (Alt+F8 -> InstalarBotones) y crea los botones en
' las hojas Panel y Tablas con sus macros ya asignadas.
' =================  HOJAS AUTONOMAS (Crear hoja con tabla)  ================
' "Crear hoja con tabla" crea una hoja limpia (Tabla N): la TABLA empieza en H6
' (Entidad en col H, metricas a la derecha en I,J..; fondos hacia abajo), la config
' (Estilo/Total) en B6, y la posicion de PowerPoint en E6. Vuelca la configuracion
' del maestro "Tablas" y refresca. Botones: Actualizar / Actualizar a fecha / Exportar.
Public Sub CrearHojaTabla()
    Dim src As Worksheet
    On Error Resume Next
    Set src = ThisWorkbook.Sheets("Tablas")
    On Error GoTo 0
    If src Is Nothing Then MsgBox "No encuentro la hoja 'Tablas'.", vbExclamation: Exit Sub

    Dim nom As String: nom = NombreHojaLibre("Tabla")
    On Error GoTo limp
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nom
    ws.Cells(1, 30).Value = "TABLA"                 ' AD1: marca de tipo de hoja
    OrigenTabla ws                                  ' col H; marca/met/per f4-6, cabecera f7, datos f8; controles C7:C9
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo limp

    ' Titulo + botones (columna A, apilados en filas 1-3, todas visibles).
    ws.Rows(1).RowHeight = 26: ws.Rows(2).RowHeight = 26: ws.Rows(3).RowHeight = 26
    ws.Cells(1, gCol0).Value = "Tabla: " & nom      ' H1
    ws.Cells(1, gCol0).Font.Bold = True
    CrearBoton ws, "A1", ">> ACTUALIZAR", "HojaActualizar"
    CrearBoton ws, "A2", "ACTUALIZAR A FECHA", "HojaActualizarFecha"
    CrearBoton ws, "A3", "EXPORTAR A POWERPOINT", "HojaExportarPPT"

    ' Config (cols B/C, filas 7-9, visibles): estilo, total, estado. Hereda del maestro.
    ws.Cells(7, 2).Value = "Estilo (color)"
    ws.Cells(8, 2).Value = "Fila de Total"
    ws.Cells(9, 2).Value = "Actualizado"
    ws.Range(gEstilo).Value = CStr(src.Range("B3").Value)
    ws.Range(gTotal).Value = CStr(src.Range("D3").Value)
    ws.Range(gEstado).Value = "(sin actualizar)"
    PonerDV ws.Range(gEstilo), "Ninguno,Mapa de calor,Barras"
    PonerDV ws.Range(gTotal), "No,Si"
    ws.Columns(2).ColumnWidth = 15

    ' Posicion PowerPoint (cols E/F, filas 7-12).
    PanelExportar ws

    ' Filas de configuracion de la tabla (marca/metrica/periodo) OCULTAS: la fila de
    ' cabecera (gRCap) queda como unica cabecera visible. La rellena RellenarTablaEn con
    ' nombres limpios (Fondo/Cartera/Indice + metrica) y es editable (mixta).
    ws.Rows(gRMark).Hidden = True                   ' fila de marcas (crece)
    ws.Rows(gRMet).Hidden = True                    ' fila metrica (config)
    ws.Rows(gRHdr).Hidden = True                    ' fila periodo (config)
    ws.Rows(gRCap).RowHeight = 18
    ws.Range(ws.Cells(gRCap, gCol0), ws.Cells(gRCap, gCol0 + 40)).Font.Bold = True

    ' Vuelca del maestro: entidades (col A f7+) -> col H (f8+); columnas Metrica/
    ' Periodo (maestro f5/f6 col B+) -> generada f5/f6 col I+ (filas config ocultas).
    Dim r As Long, c As Long, k As Long
    k = 0: r = 7
    Do While r <= 5000
        If Len(Trim(CStr(src.Cells(r, 1).Value))) = 0 Then Exit Do
        ws.Cells(gRData + k, gCol0).Value = src.Cells(r, 1).Value: k = k + 1: r = r + 1
    Loop
    If k = 0 Then ws.Cells(gRData, gCol0).Value = "Cartera RF Gobierno"
    k = 0: c = 2
    Do While c <= 90
        If Len(Trim(CStr(src.Cells(5, c).Value))) = 0 And Len(Trim(CStr(src.Cells(6, c).Value))) = 0 Then Exit Do
        ws.Cells(gRMet, gCol0 + 1 + k).Value = src.Cells(5, c).Value
        ws.Cells(gRHdr, gCol0 + 1 + k).Value = src.Cells(6, c).Value
        k = k + 1: c = c + 1
    Loop
    If k = 0 Then ws.Cells(gRMet, gCol0 + 1).Value = "Rentabilidad": ws.Cells(gRHdr, gCol0 + 1).Value = "Ultimo"

    ' Desplegables: Metrica (fila metrica, I..), Periodo (fila periodo, I..), Entidad (col H).
    PonerDV ws.Range(ws.Cells(gRMet, gCol0 + 1), ws.Cells(gRMet, gCol0 + 40)), "=MetricasTabla"
    PonerDV ws.Range(ws.Cells(gRHdr, gCol0 + 1), ws.Cells(gRHdr, gCol0 + 40)), "=PeriodosTabla"
    PonerDV ws.Range(ws.Cells(gRData, gCol0), ws.Cells(gRData + 200, gCol0)), "=EntLista"

    ws.Columns(gCol0).ColumnWidth = 26
    Application.ScreenUpdating = True
    RefrescarHoja ws                                ' rellena con el origen generado
    ws.Activate
    MsgBox "Creada la hoja '" & nom & "'. La tabla empieza en la fila de cabecera " & gRCap & _
           " (col " & Left(ws.Cells(1, gCol0).Address(False, False), 1) & "). Para anadir columnas o " & _
           "cambiar metrica/periodo, muestra las filas ocultas " & gRMark & "-" & gRHdr & _
           "; para anadir fondos, escribelos debajo. Pulsa Actualizar.", vbInformation, "Crear hoja"
    Exit Sub
limp:
    Application.ScreenUpdating = True
    MsgBox "No se pudo crear la hoja:" & vbLf & Err.Description, vbExclamation, "Crear hoja"
End Sub

' "Crear hoja con grafica": duplica la hoja PANEL en una hoja autonoma (Grafica N)
' con sus mismos desplegables y su grafico. El grafico se redibuja apuntando a ESTA
' hoja (referencias directas, no a nombres del libro), asi se actualiza solo.
Public Sub CrearHojaGrafica()
    Dim src As Worksheet
    On Error Resume Next
    Set src = ThisWorkbook.Sheets("Panel")
    On Error GoTo 0
    If src Is Nothing Then MsgBox "No encuentro la hoja 'Panel'.", vbExclamation: Exit Sub

    Dim nom As String: nom = NombreHojaLibre("Grafica")
    On Error GoTo limp
    Application.ScreenUpdating = False
    src.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Name = nom
    ws.Visible = xlSheetVisible
    BorrarBotones ws
    ws.Cells(1, 30).Value = "GRAFICA"                ' AD1: marca de tipo de hoja
    PanelExportar ws
    CrearBoton ws, "A20", ">> ACTUALIZAR", "HojaActualizar"
    CrearBoton ws, "A22", "ACTUALIZAR A FECHA", "HojaActualizarFecha"
    CrearBoton ws, "A24", "EXPORTAR A POWERPOINT", "HojaExportarPPT"
    ' Refresca la copia (consulta + dibuja + congela) para que muestre SUS datos
    ' leyendo de ESTA hoja, no del Panel original.
    RefrescarHoja ws
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Creada la hoja '" & nom & "' (copia del Panel). Elige sus desplegables, " & _
           "pulsa Actualizar y exportala a PowerPoint.", vbInformation, "Crear hoja"
    Exit Sub
limp:
    Set mHojaPanel = Nothing
    Application.ScreenUpdating = True
    MsgBox "No se pudo crear la hoja:" & vbLf & Err.Description, vbExclamation, "Crear hoja"
End Sub

' Refresca una hoja generada segun su tipo: "grafica" = motor del Panel redirigido
' a esa hoja; "tabla" = RellenarTablaEn.
Private Sub RefrescarHoja(ByVal ws As Worksheet)
    If TipoHoja(ws) = "grafica" Then
        Set mHojaPanel = ws
        On Error Resume Next
        Actualizar
        ' Congela la matriz del grafico (D..V, ancho maximo de series) a VALORES: el
        ' Panel usa nombres de libro compartidos (f_*), asi que sin congelar todas las
        ' copias mostrarian los datos de la ultima refrescada. Con valores estaticos
        ' cada copia es independiente (incluye composicion apilada, que usa D:V).
        ws.Range(ws.Cells(2, TCMP_COL), ws.Cells(402, TCMP_COL + MAXSER)).Value = _
            ws.Range(ws.Cells(2, TCMP_COL), ws.Cells(402, TCMP_COL + MAXSER)).Value
        On Error GoTo 0
        Set mHojaPanel = Nothing
    Else
        RellenarTablaEn ws, True
    End If
End Sub

' Ultima fila (col Entidad desde gRData) y ultima columna (fila Periodo desde la
' 1a de datos) con datos. Requiere OrigenTabla fijado por el llamante.
Private Sub UltimaCeldaTabla(ByVal ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim r As Long, c As Long
    lastRow = gRData - 1
    r = gRData
    Do While r <= 20000
        If Len(Trim(CStr(ws.Cells(r, gCol0).Value))) = 0 Then Exit Do
        lastRow = r: r = r + 1
    Loop
    lastCol = gCol0
    c = gCol0 + 1
    Do While c <= gCol0 + 200
        If Len(Trim(CStr(ws.Cells(gRHdr, c).Value))) = 0 Then Exit Do
        lastCol = c: c = c + 1
    Loop
End Sub

' Siguiente nombre de hoja libre "Base N".
Private Function NombreHojaLibre(ByVal base As String) As String
    Dim k As Long: k = 1
    Do While k < 1000
        Dim nom As String: nom = base & " " & k
        Dim existe As Boolean, sh As Worksheet: existe = False
        For Each sh In ThisWorkbook.Worksheets
            If StrComp(sh.Name, nom, vbTextCompare) = 0 Then existe = True: Exit For
        Next sh
        If Not existe Then NombreHojaLibre = nom: Exit Function
        k = k + 1
    Loop
    NombreHojaLibre = base & " " & k
End Function

' Celda de los valores del bloque de posicion PPT: columna 'vc' y fila base 'r0'
' (Slide/Izq/Arr/Ancho/Alto en vc(r0+1)..vc(r0+5)). Tabla -> E/F fila 7; grafico -> M/N fila 1.
Private Sub PosVals(ByVal ws As Worksheet, ByRef vc As String, ByRef r0 As Long)
    If TipoHoja(ws) = "grafica" Then
        vc = "N": r0 = 1
    Else
        vc = "F": r0 = 7
    End If
End Sub

' Bloque vertical de exportacion a PowerPoint (etiquetas + valores), con defaults.
Private Sub PanelExportar(ByVal ws As Worksheet)
    Dim vc As String, r0 As Long: PosVals ws, vc, r0
    Dim lc As String: lc = Chr(Asc(vc) - 1)      ' columna de etiquetas (una a la izquierda)
    ws.Range(lc & r0).Value = "Exportar a PowerPoint"
    ws.Range(lc & r0).Font.Bold = True
    ws.Range(lc & (r0 + 1)).Value = "Slide":         ws.Range(vc & (r0 + 1)).Value = 1
    ws.Range(lc & (r0 + 2)).Value = "Izquierda (cm)": ws.Range(vc & (r0 + 2)).Value = 1.5
    ws.Range(lc & (r0 + 3)).Value = "Arriba (cm)":    ws.Range(vc & (r0 + 3)).Value = 3
    ws.Range(lc & (r0 + 4)).Value = "Ancho (cm)":     ws.Range(vc & (r0 + 4)).Value = 24
    ws.Range(lc & (r0 + 5)).Value = "Alto (cm)":      ws.Range(vc & (r0 + 5)).Value = 12
    Dim i As Long
    For i = 1 To 5
        ws.Range(lc & (r0 + i)).Font.Italic = True
        ws.Range(vc & (r0 + i)).Interior.Color = RGB(238, 242, 248)
    Next i
    ws.Columns(lc).ColumnWidth = 15
End Sub

' --- Botones de las hojas generadas (operan sobre la hoja ACTIVA) ---
Public Sub HojaActualizar()          ' ultima fecha disponible (ignora la Portada)
    Dim ws As Worksheet: Set ws = ActiveSheet
    mSinAsOf = True
    On Error GoTo limp
    RefrescarHoja ws
    mSinAsOf = False
    MsgBox "Actualizado a la ultima fecha disponible.", vbInformation, ws.Name
    Exit Sub
limp:
    mSinAsOf = False
    Set mHojaPanel = Nothing
    MsgBox "Error al actualizar:" & vbLf & Err.Description, vbExclamation, ws.Name
End Sub

Public Sub HojaActualizarFecha()     ' a cierre de la fecha global de Portada
    Dim ws As Worksheet: Set ws = ActiveSheet
    mSinAsOf = False
    On Error GoTo limp
    RefrescarHoja ws
    Dim d As String: d = CfgAsOf()
    If Len(d) = 0 Then d = "ultima disponible"
    MsgBox "Actualizado a la fecha de Portada (" & d & ").", vbInformation, ws.Name
    Exit Sub
limp:
    Set mHojaPanel = Nothing
    MsgBox "Error al actualizar:" & vbLf & Err.Description, vbExclamation, ws.Name
End Sub

' Rango imagen de una tabla (cabeceras Metrica/Periodo + datos + columna Entidad).
Private Function RangoTabla(ByVal ws As Worksheet) As Range
    OrigenTabla ws
    Dim lastRow As Long, lastCol As Long
    UltimaCeldaTabla ws, lastRow, lastCol
    If lastRow < gRData Or lastCol < gCol0 + 1 Then Exit Function
    Set RangoTabla = ws.Range(ws.Cells(gRMet, gCol0), ws.Cells(lastRow, lastCol))
End Function

' Pega el objeto (tabla o grafico) de 'ws' en 'pres' en la slide/posicion/tamano
' (cm) de su fila 2. Idempotente: borra su imagen previa (misma hoja) en esa slide.
' Devuelve "" si OK, o el texto del error (para el resumen del lote).
Private Function PegarObjeto(ByVal ws As Worksheet, ByVal pres As Object) As String
    On Error GoTo limp
    Dim esGrafica As Boolean: esGrafica = (TipoHoja(ws) = "grafica")
    Dim rng As Range, co As ChartObject
    If esGrafica Then
        On Error Resume Next
        Set co = ws.ChartObjects(1)
        On Error GoTo limp
        If co Is Nothing Then PegarObjeto = "sin grafico": Exit Function
        co.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Else
        Set rng = RangoTabla(ws)
        If rng Is Nothing Then PegarObjeto = "sin tabla": Exit Function
        rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    End If

    Dim vc As String, r0 As Long: PosVals ws, vc, r0
    Dim slideN As Long, izq As Double, arr As Double, anc As Double, alt As Double
    slideN = CLng(Val(CStr(ws.Range(vc & (r0 + 1)).Value))): If slideN < 1 Then slideN = 1
    izq = NumDbl(ws.Range(vc & (r0 + 2)).Value): arr = NumDbl(ws.Range(vc & (r0 + 3)).Value)
    anc = NumDbl(ws.Range(vc & (r0 + 4)).Value): alt = NumDbl(ws.Range(vc & (r0 + 5)).Value)

    Do While pres.Slides.Count < slideN
        pres.Slides.Add pres.Slides.Count + 1, 12            ' 12 = ppLayoutBlank
    Loop
    Dim sld As Object: Set sld = pres.Slides(slideN)
    Dim nm As String: nm = "XLS_" & ws.Name
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Name = nm Then sld.Shapes(i).Delete
    Next i
    Dim shp As Object: Set shp = sld.Shapes.PasteSpecial(DataType:=2)(1)   ' 2 = EMF
    shp.Name = nm
    Dim KP As Double: KP = 28.3465                            ' cm -> puntos
    shp.LockAspectRatio = False
    shp.Left = izq * KP: shp.Top = arr * KP
    If anc > 0 Then shp.Width = anc * KP
    If alt > 0 Then shp.Height = alt * KP
    PegarObjeto = ""
    Exit Function
limp:
    PegarObjeto = Err.Description
End Function

' PowerPoint + presentacion activa (o nueva). Nothing si no hay PowerPoint.
Private Function PowerPointPres() As Object
    Dim ppt As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Exit Function
    ppt.Visible = True
    If ppt.Presentations.Count = 0 Then Set PowerPointPres = ppt.Presentations.Add Else Set PowerPointPres = ppt.ActivePresentation
End Function

' Ruta absoluta de la plantilla .pptx: 1) celda "Plantilla PowerPoint" de la Portada
' (B19); 2) config PPT_PLANTILLA. Si es un nombre relativo (sin unidad ni "\\") se
' resuelve JUNTO AL LIBRO. "" si no hay o no existe.
Private Function RutaPlantilla() As String
    Dim p As String: p = PortadaTxt("B19")
    If Len(p) = 0 Then p = Trim(Cfg("PPT_PLANTILLA", ""))
    If Len(p) = 0 Then Exit Function
    If InStr(p, ":") = 0 And Left(p, 2) <> "\\" Then
        Dim base As String: base = ThisWorkbook.Path
        If Len(base) > 0 Then
            If Right(base, 1) <> "\" Then base = base & "\"
            p = base & p
        End If
    End If
    On Error Resume Next
    If Len(Dir(p)) > 0 Then RutaPlantilla = p
    On Error GoTo 0
End Function

' Presentacion destino del boton de exportar de una hoja: prioriza la plantilla
' configurada (PPT_PLANTILLA, junto al libro si es relativa): si ya esta abierta la
' reutiliza; si existe en disco la abre. Si no hay plantilla usa la presentacion
' activa o crea una nueva. Nothing si PowerPoint no esta disponible.
Private Function PresentacionDestino() As Object
    Dim ppt As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Exit Function
    ppt.Visible = True

    Dim ruta As String: ruta = RutaPlantilla()
    If Len(ruta) > 0 Then
        Dim pr As Object
        For Each pr In ppt.Presentations           ' 1) ya abierta?
            On Error Resume Next
            If StrComp(pr.FullName, ruta, vbTextCompare) = 0 Then Set PresentacionDestino = pr
            On Error GoTo 0
            If Not PresentacionDestino Is Nothing Then Exit Function
        Next pr
        On Error Resume Next                        ' 2) abrir del disco
        Set PresentacionDestino = ppt.Presentations.Open(ruta)
        On Error GoTo 0
        If Not PresentacionDestino Is Nothing Then Exit Function
    End If

    ' 3) Fallback: activa o nueva.
    If ppt.Presentations.Count = 0 Then Set PresentacionDestino = ppt.Presentations.Add Else Set PresentacionDestino = ppt.ActivePresentation
End Function

' Boton de la hoja generada: exporta su objeto a la plantilla configurada (o, si no
' hay, a la presentacion abierta). Deja el .pptx abierto para que lo revises y guardes.
Public Sub HojaExportarPPT()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim pres As Object: Set pres = PresentacionDestino()
    If pres Is Nothing Then MsgBox "No se pudo abrir PowerPoint.", vbExclamation, ws.Name: Exit Sub
    Dim e As String: e = PegarObjeto(ws, pres)
    Dim vc As String, r0 As Long: PosVals ws, vc, r0
    Dim nomPres As String: nomPres = "presentacion activa"
    On Error Resume Next
    nomPres = pres.Name
    On Error GoTo 0
    If Len(e) = 0 Then
        MsgBox "Exportado a la slide " & CLng(Val(CStr(ws.Range(vc & (r0 + 1)).Value))) & _
               " de '" & nomPres & "'. Guarda la presentacion para conservarlo.", vbInformation, ws.Name
    Else
        MsgBox "No se pudo exportar:" & vbLf & e, vbExclamation, ws.Name
    End If
End Sub

' Texto (recortado) de una celda de la Portada. "" si no hay hoja o esta vacia.
Private Function PortadaTxt(ByVal addr As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Portada")
    If Not ws Is Nothing Then PortadaTxt = Trim(CStr(ws.Range(addr).Value))
    On Error GoTo 0
End Function

' "YYYY-MM" del informe: 1) la celda "Periodo" de la Portada (B17) si es AAAA-MM;
' 2) el mes de cierre (B2/B3); 3) el mes actual.
Private Function MesInforme() As String
    Dim p As String: p = PortadaTxt("B17")
    If p Like "####-##" Then MesInforme = p: Exit Function
    Dim y As Long, m As Long
    y = CLng(Val(PortadaTxt("B2")))
    m = MesNum(PortadaTxt("B3"))
    If y < 1900 Or m < 1 Then MesInforme = Format(Date, "yyyy-mm") Else MesInforme = Format(DateSerial(y, m, 1), "yyyy-mm")
End Function

' Quita de un nombre los caracteres invalidos para carpetas/archivos (los sustituye
' por "-"). Sirve para convertir la "fecha del informe" en un nombre de carpeta.
Private Function SanearNombre(ByVal s As String) As String
    Dim t As String: t = Trim(s)
    Dim bad As Variant, x As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each x In bad
        t = Replace(t, CStr(x), "-")
    Next x
    SanearNombre = t
End Function

' Crea 'ruta' (y sus carpetas padre) si no existe. No falla si ya existe.
Private Sub AsegurarCarpeta(ByVal ruta As String)
    Dim r As String: r = ruta
    If Right(r, 1) = "\" Then r = Left(r, Len(r) - 1)
    If Len(r) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir(r, vbDirectory)) > 0 Then Exit Sub
    On Error GoTo 0
    Dim p As Long: p = InStrRev(r, "\")
    If p > 1 Then
        Dim parent As String: parent = Left(r, p - 1)
        If InStr(parent, "\") > 0 Then AsegurarCarpeta parent
    End If
    On Error Resume Next
    MkDir r
    On Error GoTo 0
End Sub

' ORQUESTADOR MENSUAL (boton de Portada): actualiza todas las hojas generadas a la
' fecha de Portada, guarda una copia .xlsm del mes (auditable) y monta el .pptx del
' mes pegando cada objeto (imagen EMF) en su slide. Robusto: si algo falla sigue y
' lo reporta al final. Degrada sin PowerPoint (la copia .xlsm se guarda igual).
Public Sub GenerarInformeMes()
    Dim carpeta As String, base As String, plantilla As String, mesTxt As String
    ' Carpeta base = la del propio libro. El informe se guarda SIEMPRE junto al Excel.
    carpeta = ThisWorkbook.Path
    base = Cfg("PPT_NOMBRE", "Informe")
    mesTxt = MesInforme()
    If Len(carpeta) = 0 Then MsgBox "Guarda el libro en una carpeta antes de generar el informe.", vbExclamation, "Informe": Exit Sub
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    ' Subcarpeta = "fecha del informe" de la Portada (B18), creada junto al Excel si no existe.
    Dim subc As String: subc = SanearNombre(PortadaTxt("B18"))
    If Len(subc) > 0 Then carpeta = carpeta & subc & "\"
    AsegurarCarpeta carpeta
    ' Plantilla: celda de la Portada (B19) o config PPT_PLANTILLA, resuelta junto al libro.
    plantilla = RutaPlantilla()

    Dim sh As Worksheet, tipo As String
    Dim n As Long, okR As Long, okP As Long, fallos As String: n = 0: okR = 0: okP = 0: fallos = ""

    ' 1) Actualiza todas las hojas generadas a la fecha de Portada.
    Application.StatusBar = "Informe: actualizando hojas..."
    mSinAsOf = False
    For Each sh In ThisWorkbook.Worksheets
        tipo = TipoHoja(sh)
        If tipo = "tabla" Or tipo = "grafica" Then
            n = n + 1
            On Error Resume Next
            Err.Clear
            RefrescarHoja sh
            If Err.Number <> 0 Then fallos = fallos & vbLf & " - " & sh.Name & " (actualizar): " & Err.Description Else okR = okR + 1
            On Error GoTo 0
        End If
    Next sh
    If n = 0 Then
        Application.StatusBar = False
        MsgBox "No hay hojas generadas (Tabla/Grafica) que incluir en el informe.", vbExclamation, "Informe": Exit Sub
    End If

    ' 2) Copia .xlsm del mes (congelada a la Portada actual).
    Dim rutaXls As String: rutaXls = carpeta & base & " " & mesTxt & ".xlsm"
    Dim errXls As String
    On Error Resume Next
    ThisWorkbook.SaveCopyAs rutaXls
    If Err.Number <> 0 Then errXls = Err.Description
    On Error GoTo 0

    ' 3) PPTX del mes (a partir de plantilla si hay).
    Dim rutaPpt As String, errPpt As String
    rutaPpt = carpeta & base & " " & mesTxt & ".pptx"
    Application.StatusBar = "Informe: generando PowerPoint..."
    Dim ppt As Object, pres As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then
        errPpt = "PowerPoint no disponible (se guardo solo el Excel)"
    Else
        ppt.Visible = True
        On Error Resume Next
        If Len(plantilla) > 0 Then
            FileCopy plantilla, rutaPpt
            Set pres = ppt.Presentations.Open(rutaPpt)
        Else
            Set pres = ppt.Presentations.Add
        End If
        On Error GoTo 0
        If pres Is Nothing Then
            errPpt = "no se pudo abrir la presentacion (o la plantilla)"
        Else
            For Each sh In ThisWorkbook.Worksheets
                tipo = TipoHoja(sh)
                If tipo = "tabla" Or tipo = "grafica" Then
                    Dim e As String: e = PegarObjeto(sh, pres)
                    If Len(e) = 0 Then okP = okP + 1 Else fallos = fallos & vbLf & " - " & sh.Name & " (export): " & e
                End If
            Next sh
            On Error Resume Next
            If Len(plantilla) > 0 Then pres.Save Else pres.SaveAs rutaPpt
            If Err.Number <> 0 Then errPpt = "no se pudo guardar: " & Err.Description
            On Error GoTo 0
        End If
    End If
    Application.StatusBar = False

    ' 4) Resumen.
    Dim msg As String
    msg = "Informe " & mesTxt & vbLf & _
          " - Hojas: " & n & " (actualizadas " & okR & ")" & vbLf
    If Len(errXls) = 0 Then msg = msg & " - Excel: " & rutaXls & vbLf Else msg = msg & " - Excel FALLO: " & errXls & vbLf
    If Len(errPpt) = 0 Then msg = msg & " - PowerPoint: " & okP & " objetos -> " & rutaPpt & vbLf Else msg = msg & " - PowerPoint: " & errPpt & vbLf
    If Len(fallos) > 0 Then msg = msg & vbLf & "Incidencias:" & fallos
    MsgBox msg, IIf(Len(fallos) > 0 Or Len(errXls) > 0, vbExclamation, vbInformation), "Generar informe del mes"
End Sub

Public Sub InstalarBotones()
    Dim ws As Worksheet
    Set ws = Panel()
    BorrarBotones ws
    CrearBoton ws, "A20", "Cargar carteras (segun B3)", "CargarCarteras"
    CrearBoton ws, "A22", ">> ACTUALIZAR QUERY Y GRAFICO", "Actualizar"
    CrearBoton ws, "A24", "Dibujar (aplica tipo)", "DibujarGrafico"
    CrearBoton ws, "A26", "Vista previa (dummy)", "VistaPreviaDummy"
    CrearBoton ws, "A28", "Ver SQL", "VerSQL"
    CrearBoton ws, "A30", "Ver columnas (diagnostico)", "VerColumnas"
    CrearBoton ws, "A32", "Ver benchmark (diagnostico)", "VerBenchmark"
    CrearBoton ws, "A34", "Ver rentab. mes (diagnostico)", "VerRentabMes"
    CrearBoton ws, "A36", "> A PowerPoint (Fase 3)", "CopiarAPowerPoint"
    CrearBoton ws, "A38", "+ CREAR HOJA CON GRAFICA", "CrearHojaGrafica"

    On Error Resume Next
    Dim wt As Worksheet: Set wt = ThisWorkbook.Sheets("Tablas")
    On Error GoTo 0
    If Not wt Is Nothing Then
        BorrarBotones wt
        CrearBoton wt, "J1", ">> RELLENAR TABLA", "RellenarTabla"
        CrearBoton wt, "J3", "+ CREAR HOJA CON TABLA", "CrearHojaTabla"
    End If
    On Error Resume Next
    Dim wp As Worksheet: Set wp = ThisWorkbook.Sheets("Posiciones")
    On Error GoTo 0
    If Not wp Is Nothing Then
        BorrarBotones wp
        CrearBoton wp, "D3", ">> RELLENAR POSICIONES", "RellenarPosiciones"
    End If
    On Error Resume Next
    Dim wportada As Worksheet: Set wportada = ThisWorkbook.Sheets("Portada")
    On Error GoTo 0
    If Not wportada Is Nothing Then
        BorrarBotones wportada
        CrearBoton wportada, "A7", "1. Actualizar Excel", "ActualizarTodoExcel"
        CrearBoton wportada, "A9", "2. Actualizar PowerPoint", "ActualizarPowerPoint"
        CrearBoton wportada, "A11", "3. Actualizar Excel y PowerPoint", "ActualizarExcelYPowerPoint"
        CrearBoton wportada, "A13", "4. Generar informe del mes (xlsm + pptx)", "GenerarInformeMes"
    End If
    ' Deja el grafico en modo dinamico (rangos con nombre) desde el principio,
    ' asi se ajusta solo al cambiar periodo/dimension y no deja huecos.
    On Error Resume Next
    DibujarGrafico
    On Error GoTo 0
    ' Cascada automatica: al cambiar el Grupo (B7), Metrica/Dimension se resetean
    ' a valores validos. Se intenta instalar el evento; si no hay acceso al
    ' proyecto VBA, se avisa (Actualizar tambien lo repara como red de seguridad).
    Dim casc As Boolean: casc = InstalarAutoRefresco()
    InstalarAutoTabla   ' auto-refresco opcional de la hoja Tablas (si H3="Si")
    Dim m As String
    m = "Botones creados en 'Panel' y 'Tablas'." & vbLf & vbLf & _
        "Flujo: elige los desplegables y pulsa 'ACTUALIZAR QUERY Y GRAFICO'." & vbLf & _
        "La 1a vez de cada GRUPO y cartera descarga sus datos; despues cambiar " & _
        "metrica/dimension/periodo del mismo grupo es instantaneo."
    If casc Then
        m = m & vbLf & vbLf & "Cascada ACTIVADA: al cambiar el Grupo, Metrica/Dimension se ajustan solas."
    Else
        m = m & vbLf & vbLf & "Para que Metrica/Dimension se ajusten SOLAS al cambiar el Grupo: activa " & _
            "'Confiar en el acceso al modelo de objetos de proyectos de VBA' (Opciones > Centro de " & _
            "confianza) y re-ejecuta InstalarBotones. (Aun sin eso, 'Actualizar' lo corrige.)"
    End If
    MsgBox m, vbInformation, "Instalacion"
End Sub

' Inserta (una vez) el evento Worksheet_Change en el modulo de la hoja Panel
' para la CASCADA automatica al cambiar un desplegable (sin consultar BigQuery):
'   B3 (tipo de entidad Fondo/Cartera/Indice) -> repuebla Entidad 1/2/3.
'   B7 (grupo) -> resetea Metrica (B8) y Dimension (B9) a valores validos.
'   B3:B15 (cualquier parametro) -> recalcula tabla+grafico desde los datos ya
'                                   descargados (AutoLocal); si no hay, no hace nada.
' Requiere acceso al modelo de objetos VBA; devuelve False si no se pudo instalar.
Public Function InstalarAutoRefresco() As Boolean
    Dim cm As Object, cn As String, txt As String, s As String
    On Error GoTo sinacceso
    cn = Panel().CodeName
    Set cm = ThisWorkbook.VBProject.VBComponents(cn).CodeModule
    If cm.CountOfLines > 0 Then txt = cm.Lines(1, cm.CountOfLines)
    If InStr(txt, "Worksheet_Change") > 0 Then
        ' Si ya es la version nueva (dispara AutoLocal solo en B7:B15), nada.
        If InStr(txt, "B7:B15") > 0 Then InstalarAutoRefresco = True: Exit Function
        ' Version antigua -> la borramos y reinstalamos la nueva.
        Dim pl As Long, pc As Long
        pl = cm.ProcStartLine("Worksheet_Change", 0)   ' 0 = vbext_pk_Proc
        pc = cm.ProcCountLines("Worksheet_Change", 0)
        If pc > 0 Then cm.DeleteLines pl, pc
    End If
    ' Nota: el auto-recalculo (AutoLocal) SOLO se dispara al cambiar B7:B15
    ' (metrica/dimension/periodo/grafico/benchmark). Elegir entidad (B4:B6) NO recalcula
    ' -> elegir carteras/fondos es instantaneo (luego pulsas 'Actualizar').
    s = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    If Application.EnableEvents = False Then Exit Sub" & vbCrLf & _
        "    If Intersect(Target, Me.Range(""B3:B15"")) Is Nothing Then Exit Sub" & vbCrLf & _
        "    Application.EnableEvents = False" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    If Not Intersect(Target, Me.Range(""B3"")) Is Nothing Then CargarCarteras True" & vbCrLf & _
        "    If Not Intersect(Target, Me.Range(""B4"")) Is Nothing Then RefrescarBmkPropuesto" & vbCrLf & _
        "    If Not Intersect(Target, Me.Range(""B7"")) Is Nothing Then ReiniciarMetricaDim" & vbCrLf & _
        "    If Not Intersect(Target, Me.Range(""B7:B15"")) Is Nothing Then AutoLocal" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "    Application.EnableEvents = True" & vbCrLf & _
        "End Sub"
    cm.AddFromString s
    InstalarAutoRefresco = True
    Exit Function
sinacceso:
    InstalarAutoRefresco = False
End Function

' Auto-refresco de la hoja Tablas: lo llama el evento Worksheet_Activate. Solo
' actua si E4 = "Si" y SOLO una vez por sesion (para no re-consultar cada vez que
' se pincha la pestana). Silencioso (sin el MsgBox final de RellenarTabla).
Public Sub AutoRellenarTabla()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Tablas")
    If ws Is Nothing Then Exit Sub
    If mTablaAuto Then Exit Sub
    If Fold(CStr(ws.Range("F3").Value)) <> "si" Then Exit Sub   ' Auto al abrir
    mTablaAuto = True
    RellenarTabla True
End Sub

' ===================  BOTONES DE LA PORTADA (documento)  ===================
' Refresca TODO el Excel a la fecha de referencia global (Portada). Recalcula el
' Panel y rellena Tablas y Posiciones (silenciosas). El tope de fecha lo aplican
' las propias consultas via CfgAsOf().
Public Sub ActualizarTodoExcel(Optional ByVal quiet As Boolean = False)
    On Error Resume Next
    Actualizar                                   ' Panel (con la config actual)
    If Not ThisWorkbook.Sheets("Tablas") Is Nothing Then RellenarTabla True
    If Not ThisWorkbook.Sheets("Posiciones") Is Nothing Then RellenarPosiciones True
    On Error GoTo 0
    If Not quiet Then
        Dim d As String: d = CfgAsOf()
        MsgBox "Excel actualizado (Panel, Tablas y Posiciones) a fecha " & _
               IIf(Len(d) > 0, d, "ultima disponible") & ".", vbInformation, "Actualizar Excel"
    End If
End Sub

' Actualiza el PowerPoint (por ahora, copia el grafico del Panel; el deck completo
' es Fase 3). Usa el grafico ya actualizado a la fecha de referencia.
Public Sub ActualizarPowerPoint()
    CopiarAPowerPoint
End Sub

' Los dos: primero el Excel a la fecha de referencia, luego el PowerPoint.
Public Sub ActualizarExcelYPowerPoint()
    ActualizarTodoExcel True
    ActualizarPowerPoint
End Sub

' Inyecta el evento Worksheet_Activate en la hoja Tablas (auto-refresco al entrar
' en la pestana si H3="Si"). Requiere acceso al modelo de objetos VBA.
Public Function InstalarAutoTabla() As Boolean
    Dim cm As Object, cn As String, txt As String, s As String, wt As Worksheet
    On Error GoTo sinacceso
    Set wt = ThisWorkbook.Sheets("Tablas")
    cn = wt.CodeName
    Set cm = ThisWorkbook.VBProject.VBComponents(cn).CodeModule
    If cm.CountOfLines > 0 Then txt = cm.Lines(1, cm.CountOfLines)
    If InStr(txt, "Worksheet_Activate") > 0 Then InstalarAutoTabla = True: Exit Function
    s = "Private Sub Worksheet_Activate()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    AutoRellenarTabla" & vbCrLf & _
        "End Sub"
    cm.AddFromString s
    InstalarAutoTabla = True
    Exit Function
sinacceso:
    InstalarAutoTabla = False
End Function

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
' Lee la hoja de carteras y rellena los desplegables de Entidad. Entidad 1 (B4) se
' filtra por el tipo elegido (B3); Entidad 2 y 3 (B5/B6) ofrecen TODAS las entidades
' (cualquier tipo) para poder comparar, p.ej., una cartera con su indice o un fondo.
Public Sub CargarCarteras(Optional ByVal quiet As Boolean = False)
    Dim ws As Worksheet, wm As Worksheet, we As Worksheet
    Dim colN As Long, colT As Long, tipoSel As String
    Dim r As Long, n As Long, lastR As Long, prevE As Boolean
    Set ws = Panel()
    Set wm = HojaMaestro()
    If wm Is Nothing Then
        If Not quiet Then MsgBox "No encuentro la hoja de carteras (cartera/carteras/activos).", vbExclamation
        Exit Sub
    End If
    colN = ColPorCabecera(wm, ACTIVOS_COL_NOMBRE)
    colT = ColPorCabecera(wm, "tipo_elemento")
    If colN = 0 Then
        If Not quiet Then MsgBox "La hoja de carteras no tiene la columna 'nombre_elemento'.", vbExclamation
        Exit Sub
    End If

    tipoSel = LCase(Trim(CStr(ws.Range("B3").Value)))   ' Fondo/Cartera/Indice -> minusculas
    Set we = HojaAux("_Ent")
    we.Columns(1).ClearContents
    we.Cells(1, 1).Value = "(ninguna)"
    n = 1
    lastR = wm.Cells(wm.Rows.Count, colN).End(xlUp).Row

    ' RAPIDO: lee toda la hoja de carteras a un array de UNA vez y construye en memoria
    ' DOS listas: la filtrada por tipo (col A de _Ent, para Entidad 1) y la completa de
    ' TODAS las entidades (col B de _Ent, para Entidad 2/3). Se escriben de una vez.
    Dim r3 As Long, r4 As Long                            ' col B = todas; col C = benchmark
    r3 = 0: r4 = 0
    If lastR >= 2 Then
        Dim lastCol As Long, datos As Variant, filt() As Variant, todo() As Variant, bmk() As Variant, r2 As Long
        lastCol = colN: If colT > lastCol Then lastCol = colT
        If lastCol < 2 Then lastCol = 2
        datos = wm.Range(wm.Cells(2, 1), wm.Cells(lastR, lastCol)).Value   ' array 2D
        ReDim filt(1 To UBound(datos, 1), 1 To 1)
        ReDim todo(1 To UBound(datos, 1) + 1, 1 To 1)     ' +1 por "(ninguna)"
        ReDim bmk(1 To UBound(datos, 1) + 2, 1 To 1)      ' +2 por las 2 opciones fijas
        todo(1, 1) = "(ninguna)": r3 = 1
        bmk(1, 1) = "Sin benchmark": bmk(2, 1) = "Benchmark asociado": r4 = 2
        Dim okTipo As Boolean, nom As String, tp As String
        For r = 1 To UBound(datos, 1)
            nom = Trim(CStr(datos(r, colN)))
            If Len(nom) > 0 Then
                r3 = r3 + 1: todo(r3, 1) = nom            ' lista completa (cualquier tipo)
                tp = ""
                If colT > 0 Then tp = LCase(Trim(CStr(datos(r, colT))))
                If tp = "indice" Then r4 = r4 + 1: bmk(r4, 1) = nom   ' benchmark: solo indices
                If colT = 0 Then okTipo = True Else okTipo = (tp = tipoSel)
                If okTipo Then n = n + 1: r2 = r2 + 1: filt(r2, 1) = nom
            End If
        Next r
        If r2 > 0 Then we.Range(we.Cells(2, 1), we.Cells(r2 + 1, 1)).Value = filt
        we.Columns(2).ClearContents: we.Columns(3).ClearContents
        If r3 > 0 Then we.Range(we.Cells(1, 2), we.Cells(r3, 2)).Value = todo
        If r4 > 0 Then we.Range(we.Cells(1, 3), we.Cells(r4, 3)).Value = bmk
    End If

    ' Guarda/restaura EnableEvents (si viene del evento, ya esta False; no re-activar).
    prevE = Application.EnableEvents
    Application.EnableEvents = False
    ' Entidad 2/3: SIEMPRE la lista completa (todas las entidades, con "(ninguna)").
    If r3 >= 1 Then
        PonerDV ws.Range("B5"), "=_Ent!$B$1:$B$" & r3
        PonerDV ws.Range("B6"), "=_Ent!$B$1:$B$" & r3
    End If
    ' Benchmark (B14): "Sin benchmark", "Benchmark asociado" + la lista de indices.
    If r4 >= 1 Then PonerDV ws.Range("B14"), "=_Ent!$C$1:$C$" & r4
    If n >= 2 Then
        PonerDV ws.Range("B4"), "=_Ent!$A$2:$A$" & n      ' Entidad 1: solo el tipo de B3
        ws.Range("B4").Value = we.Cells(2, 1).Value
    End If
    ws.Range("B5").Value = "(ninguna)"
    ws.Range("B6").Value = "(ninguna)"
    Application.EnableEvents = prevE
    ActualizarBenchmarkPropuesto ws          ' refresca B13 con el benchmark de la Entidad 1

    If n < 2 Then
        If Not quiet Then MsgBox "No hay entidades de tipo '" & ws.Range("B3").Value & _
            "' para Entidad 1 (Entidad 2/3 si ofrecen todas).", vbExclamation
        Exit Sub
    End If
    If Not quiet Then MsgBox (n - 1) & " entidades de tipo '" & ws.Range("B3").Value & _
        "' en Entidad 1; " & (r3 - 1) & " (todas) en Entidad 2/3.", vbInformation, "Carteras"
End Sub

' =======================  BOTON UNICO: HACE TODO  ==========================
' Reinicia Metrica (B8) y Dimension (B9) al primer valor valido del Grupo (B7),
' leyendo los rangos con nombre DIRECTAMENTE (sin Evaluate/INDIRECT, que con
' nombres definidos puede devolver #REF). Se puede llamar desde el evento de hoja.
' Clave interna del grupo (los rangos con nombre no admiten espacios): el grupo
' que se MUESTRA como "Composicion apilada" se guarda como "Apiladas".
Private Function GrupoKey(ByVal g As String) As String
    If Fold(g) = "composicion apilada" Then GrupoKey = "Apiladas" Else GrupoKey = Trim(g)
End Function

Public Sub ReiniciarMetricaDim()
    Dim ws As Worksheet, g As String, r As Range
    Set ws = Panel(): g = GrupoKey(Trim(CStr(ws.Range("B7").Value)))
    On Error Resume Next
    Set r = Nothing: Set r = ThisWorkbook.Names("Grupo_" & g).RefersToRange
    If Not r Is Nothing Then ws.Range("B8").Value = r.Cells(1, 1).Value
    Set r = Nothing: Set r = ThisWorkbook.Names("Dim_" & g).RefersToRange
    If Not r Is Nothing Then ws.Range("B9").Value = r.Cells(1, 1).Value
    On Error GoTo 0
End Sub

' True si 'val' esta entre los valores del rango con nombre (o si no se puede
' comprobar). Sirve para validar B8/B9 contra el grupo actual.
Private Function EnRango(ByVal val As String, ByVal nombre As String) As Boolean
    Dim r As Range, c As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(nombre).RefersToRange
    On Error GoTo 0
    If r Is Nothing Then EnRango = True: Exit Function
    For Each c In r.Cells
        If Trim(CStr(c.Value)) = Trim(val) Then EnRango = True: Exit Function
    Next c
End Function

' BOTON UNICO "Actualizar query y grafico":
'  - Descarga (una vez) TODOS los datos del GRUPO de metrica elegido (B7) para
'    las carteras elegidas, con TODOS los periodos, y los guarda en la hoja Panel.
'  - Si ya estan descargados (misma cartera y grupo), NO vuelve a consultar: usa
'    los datos que hay y solo recalcula/dibuja en local.
' Asi la primera vez de cada grupo/cartera consulta; el resto es instantaneo.
Public Sub Actualizar()
    Dim ws As Worksheet, ents As String, grupo As String, dimen As String, blkId As String, msg As String
    Set ws = Panel()
    ' Red de seguridad de la cascada: si B8/B9 quedaron en #REF o con un valor
    ' que NO es valido para el grupo actual (p.ej. "Pais" en Composicion), se
    ' resetean al primer valor valido del grupo.
    Dim gk As String: gk = GrupoKey(Trim(CStr(ws.Range("B7").Value)))
    If IsError(ws.Range("B8").Value) Or IsError(ws.Range("B9").Value) _
       Or Not EnRango(CStr(ws.Range("B8").Value), "Grupo_" & gk) _
       Or Not EnRango(CStr(ws.Range("B9").Value), "Dim_" & gk) Then ReiniciarMetricaDim
    ActualizarSQL          ' deja la SQL en A41 (para 'Ver SQL') y calcula mAvisoEnt
    If Len(mAvisoEnt) > 0 Then
        MsgBox "No encontre estas entidades en la hoja 'cartera' (columna nombre_elemento):" & _
               mAvisoEnt & vbLf & vbLf & _
               "Se han ignorado. Recuerda: id_elemento = pk_portfolio_id.", _
               vbExclamation, "Entidades sin id"
    End If
    ents = ListaEntidades(ws)                 ' fija mId1/2/3
    ActualizarBenchmarkPropuesto ws           ' refresca B13 con el benchmark de la Entidad 1
    grupo = GrupoKey(Trim(CStr(ws.Range("B7").Value)))   ' "Composicion apilada" -> "Apiladas"
    dimen = Trim(CStr(ws.Range("B9").Value))
    blkId = BloqueDe(grupo, dimen)

    ' Grupo "Composicion apilada" = evolucion en el tiempo para la Entidad 1.
    ' Igual que las demas: descarga (una vez) su bloque amplio despues del resto de
    ' columnas y de ahi construye la tabla D2 que alimenta el grafico.
    If Fold(grupo) = "apiladas" Then
        If Len(ents) = 0 Then MsgBox "Elige una cartera en B4.", vbExclamation, "Composicion apilada": Exit Sub
        If AsegurarBloque(ws, "APIL", ents, msg) Then
            LocalApiladas ws
        ElseIf InStr(msg, "not found") > 0 Or InStr(msg, "Unrecognized name") > 0 Then
            MsgBox "La composicion apilada usa la tabla de posiciones (config POS_TABLE) unida al " & _
                   "maestro. Revisa POS_TABLE/JOIN_KEY en 'config'." & vbLf & vbLf & msg, _
                   vbExclamation, "Composicion apilada"
        ElseIf Len(msg) > 0 Then
            MsgBox "No se pudo descargar la composicion apilada:" & vbLf & msg, vbExclamation, "Composicion apilada"
        End If
        Exit Sub
    End If

    ' Grupos con bloque amplio (Rendimiento/Riesgo/Composicion): cache por grupo.
    If Len(ents) > 0 And Len(blkId) > 0 Then
        If AsegurarBloque(ws, blkId, ents, msg) Then
            If ResolverLocal(ws) Then DibujarGrafico: Exit Sub
            ' El bloque esta pero esta dimension/metrica no sale de el -> directa.
        ElseIf InStr(msg, "not found inside p") > 0 Or InStr(msg, "Unrecognized name") > 0 Then
            MsgBox "Composicion/Spread/TER necesita una tabla de POSICIONES (valoracion por " & _
                   "valor: PK_PORTFOLIO_ID + " & PosValor() & " + clave de valor) para unirla al maestro." & vbLf & vbLf & _
                   "La tabla configurada ('" & Cfg("POS_TABLE", T_POS) & "') NO tiene esos campos: " & _
                   "solo trae columnas de nivel cartera. Es un tema de DISPONIBILIDAD del dato." & vbLf & vbLf & _
                   "Cuando tengas la tabla real de posiciones, ponla en la hoja 'config':" & vbLf & _
                   "  POS_TABLE (y POS_DATASET si esta en otro dataset), POS_VALOR," & vbLf & _
                   "  JOIN_KEY_POS / JOIN_KEY_VAL (clave con el maestro)." & vbLf & vbLf & _
                   "Detalle: " & msg, vbExclamation, "Composicion: falta tabla de posiciones": Exit Sub
        ElseIf Len(msg) > 0 Then
            MsgBox "No se pudo descargar el grupo de datos:" & vbLf & msg, vbExclamation, "Actualizar": Exit Sub
        End If
    End If

    ' Grupos sin bloque (Costes/Liquidez/Valoracion) o casos sueltos: consulta directa.
    RefrescarDatos
End Sub

' Grupo (B7) -> bloque amplio que lo cubre ("" si no tiene). Riesgo y Composicion
' son cosas distintas y van por queries distintas:
'   Rendimiento -> RET (retornos)   Riesgo -> RISK (riesgo agregado: Duracion/TIR)
'   Composicion -> POS (posiciones + maestro: Sector/Industria/Continente/Pais/Divisa)
Private Function BloqueDe(ByVal grupo As String, ByVal dimen As String) As String
    Select Case Fold(grupo)
        Case "rendimiento": BloqueDe = "RET"
        Case "riesgo":      BloqueDe = "RISK"
        Case "composicion": BloqueDe = "APIL"   ' MISMO bloque que la apilada (comparten datos)
        Case "apiladas":    BloqueDe = "APIL"
        Case Else:          BloqueDe = ""
    End Select
End Function

Private Function BlkBaseCol(ByVal blkId As String) As Long
    Select Case blkId
        Case "RET":  BlkBaseCol = BLK_RET_COL
        Case "RISK": BlkBaseCol = BLK_RISK_COL
        Case "POS":  BlkBaseCol = BLK_POS_COL
        Case "APIL": BlkBaseCol = BLK_APIL_COL
    End Select
End Function

Private Function BlkAncho(ByVal blkId As String) As Long
    Select Case blkId
        Case "RET":  BlkAncho = 4
        Case "RISK": BlkAncho = 6
        Case "POS":  BlkAncho = 9   ' PID + gics/bics/geo/pais/divisa/rating/activo + valor
        Case "APIL": BlkAncho = 9   ' PID + mes + gics/bics/geo/pais/divisa/activo + valor
    End Select
End Function

Private Function BlkMarcaFila(ByVal blkId As String) As Long
    Select Case blkId
        Case "RET":  BlkMarcaFila = 1
        Case "RISK": BlkMarcaFila = 2
        Case "POS":  BlkMarcaFila = 3
        Case "APIL": BlkMarcaFila = 4
    End Select
End Function

Private Function BlkSQL(ByVal blkId As String, ByVal ents As String) As String
    Select Case blkId
        Case "RET":  BlkSQL = SQLBloqueRet(ents)
        Case "RISK": BlkSQL = SQLBloqueRisk(ents)
        Case "POS":  BlkSQL = SQLBloquePos(ents)
        Case "APIL": BlkSQL = SQLBloqueApil(ents)
    End Select
End Function

' True si el bloque 'blkId' ya esta descargado y cubre las carteras elegidas.
Private Function BloqueCubreId(ByVal ws As Worksheet, ByVal blkId As String) As Boolean
    Dim marca As String
    marca = CStr(ws.Cells(BlkMarcaFila(blkId), BLK_MARK_COL).Value)
    If Len(marca) = 0 Then Exit Function
    Dim ok As Boolean: ok = True
    If Len(Trim(mId1)) > 0 Then If InStr(marca, "'" & UCase(mId1) & "'") = 0 Then ok = False
    If Len(Trim(mId2)) > 0 Then If InStr(marca, "'" & UCase(mId2) & "'") = 0 Then ok = False
    If Len(Trim(mId3)) > 0 Then If InStr(marca, "'" & UCase(mId3) & "'") = 0 Then ok = False
    BloqueCubreId = ok
End Function

' Garantiza que el bloque del grupo esta descargado para estas carteras. Si ya
' esta (cache), no consulta. Si no, ejecuta la query y lo guarda. False + msg si
' la descarga fallo.
Private Function AsegurarBloque(ByVal ws As Worksheet, ByVal blkId As String, _
        ByVal ents As String, ByRef msg As String) As Boolean
    If BloqueCubreId(ws, blkId) Then AsegurarBloque = True: Exit Function
    Dim base As Long, anch As Long
    base = BlkBaseCol(blkId): anch = BlkAncho(blkId)
    Application.EnableEvents = False
    Application.Cursor = xlWait
    ws.Range(ws.Cells(1, base), ws.Cells(BLK_ULTFILA, base + anch - 1)).ClearContents
    Dim ok As Boolean: ok = DescargarBloque(ws, BlkSQL(blkId, ents), base, anch, msg)
    If ok Then
        ws.Cells(BlkMarcaFila(blkId), BLK_MARK_COL).NumberFormat = "@"
        ws.Cells(BlkMarcaFila(blkId), BLK_MARK_COL).Value = "ENTS:" & UCase(ents)
    End If
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    AsegurarBloque = ok
End Function

' Lanza la consulta, vuelca el resultado crudo desde la columna W y dibuja.
' Si la primera consulta falla porque el driver no reconoce una columna de
' benchmark (p.ej. TWR_1Y_BMK no existe), reintenta SIN benchmark para que al
' menos lleguen los datos de la cartera.
Public Sub RefrescarDatos()
    Dim ws As Worksheet, sql As String, msg As String, intento As Long
    Set ws = Panel()

    ' MODELO "solo lo necesario": consulta a BigQuery exactamente los datos de la
    ' seleccion actual (metrica/dimension/periodo/carteras) y dibuja. Sin bloques
    ' amplios ni troceo local.
    mForzarSinBmk = False
    For intento = 1 To 2
        sql = ConstruirSQL()
        If Left(sql, 2) = "--" Then
            MsgBox "Metrica/periodo no mapeada:" & vbLf & vbLf & sql, vbExclamation, "Fase 2": Exit Sub
        End If
        If EjecutarYVolcar(ws, sql, msg) Then
            DibujarGrafico
            If mForzarSinBmk Then
                MsgBox "La columna de benchmark de ese periodo no existe en la tabla, " & _
                       "asi que se ha traido SOLO la serie de la cartera (sin benchmark)." & vbLf & vbLf & _
                       "Ejecuta la macro 'VerColumnas' (Alt+F8) para ver los nombres reales " & _
                       "de las columnas de benchmark y ajustarlos.", vbInformation, "Benchmark no disponible"
            End If
            Exit Sub
        End If
        ' Reintento: si habia benchmark y el error es de columna no reconocida, quitamos benchmark.
        If intento = 1 And ConBenchmark(ws) And Not mForzarSinBmk _
           And InStr(msg, "Unrecognized name") > 0 Then
            mForzarSinBmk = True
        ElseIf InStr(msg, "not found inside p") > 0 Then
            ' La clave del JOIN posiciones->maestro no existe con ese nombre en la
            ' tabla de posiciones. Aviso claro (sin volcar toda la SQL).
            MsgBox "La composicion/spread/TER necesita unir posiciones con el maestro de " & _
                   "valores, pero la columna de union no existe con ese nombre en la tabla " & _
                   "de posiciones." & vbLf & vbLf & _
                   "1) Ejecuta 'VerColumnas' (Alt+F8) para ver los nombres reales de las " & _
                   "columnas de posiciones y del maestro." & vbLf & _
                   "2) En la hoja 'config' pon JOIN_KEY_POS (columna en posiciones) y " & _
                   "JOIN_KEY_VAL (columna en el maestro) con la clave correcta." & vbLf & vbLf & _
                   "Detalle: " & msg, vbExclamation, "Composicion: falta la clave de union"
            Exit Sub
        Else
            MsgBox "No se pudo conectar/consultar BigQuery:" & vbLf & msg & _
                   vbLf & vbLf & "SQL:" & vbLf & sql, vbExclamation, "Fase 2"
            Exit Sub
        End If
    Next intento
End Sub

' Auto-refresco LOCAL: lo llama el evento Worksheet_Change al cambiar un
' desplegable. Recalcula el grafico SOLO desde los datos ya descargados (sin ir
' a BigQuery). Si aun no hay datos de esa cartera, no hace nada (hay que pulsar
' 'Cargar datos'); nunca lanza una consulta por si solo.
Public Sub AutoLocal()
    Dim ws As Worksheet: Set ws = Panel()
    On Error Resume Next
    If IsError(ws.Range("B8").Value) Or IsError(ws.Range("B9").Value) Then Exit Sub
    ListaEntidades ws                 ' fija mId1/2/3 para B4/B5/B6
    Dim grupo As String, dimen As String, blkId As String
    grupo = GrupoKey(Trim(CStr(ws.Range("B7").Value)))
    dimen = Trim(CStr(ws.Range("B9").Value))
    blkId = BloqueDe(grupo, dimen)
    ' Solo recalcula si el bloque de ese grupo YA esta descargado para esta cartera.
    ' Si no (o el grupo no tiene bloque), no hace nada: hay que pulsar 'Actualizar'
    ' (nunca consulta BigQuery por su cuenta).
    If Len(blkId) = 0 Then Exit Sub
    If Not BloqueCubreId(ws, blkId) Then Exit Sub
    If Fold(grupo) = "apiladas" Then
        LocalApiladas ws
    Else
        If ResolverLocal(ws) Then DibujarGrafico
    End If
End Sub

' =====================  BLOQUE AMPLIO + TROCEO LOCAL  ======================
' Los bloques (retornos diarios, riesgo diario, posiciones de hoy) los descarga
' 'AsegurarBloque' bajo demanda desde el boton unico, por GRUPO de metrica, con
' TODOS los periodos. Luego el troceo (periodo/dimension) se hace en local.

' --- SQL de los tres bloques amplios -------------------------------------
Private Function SQLBloqueRet(ByVal ents As String) As String
    SQLBloqueRet = _
        "SELECT PK_PORTFOLIO_ID," & vbLf & _
        "       FORMAT_DATE('%Y-%m-%d', PK_FECHA_DATOS) AS fecha," & vbLf & _
        "       FORMAT('%.10f', CAST(TWR_1D AS FLOAT64)) AS twr_1d," & vbLf & _
        "       FORMAT('%.10f', CAST(TWR_1D_BMK AS FLOAT64)) AS twr_1d_bmk" & vbLf & _
        "FROM " & Tbl(DS_PROD, T_PERF) & vbLf & _
        "WHERE PK_NAV_GNAV = '" & CfgNav() & "' AND BENCHMARK = '" & CfgBmk() & "'" & vbLf & _
        "  AND PK_PORTFOLIO_ID IN (" & ents & ")" & AndAsOf("PK_FECHA_DATOS") & vbLf & _
        "  AND PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_PERF) & _
        MaxAsOf() & "), INTERVAL " & CACHE_ANOS & " YEAR)" & vbLf & _
        "ORDER BY PK_PORTFOLIO_ID, PK_FECHA_DATOS"
End Function

Private Function SQLBloqueRisk(ByVal ents As String) As String
    Dim w As String
    w = "WHERE " & RISK_COL_FONDOBMK & " = '" & CfgFondo() & "'" & vbLf & _
        "  AND PK_PORTFOLIO = '" & Esc(CfgPortfolio()) & "'"
    If Len(CfgLtLevel()) > 0 Then w = w & vbLf & "  AND PK_LTLEVEL = " & CfgLtLevel()
    w = w & vbLf & "  AND PK_CRITERIO_AGREGACION IN (" & CfgRiskCriterios() & ")"
    w = w & vbLf & "  AND PK_PORTFOLIO_ID IN (" & ents & ")" & AndAsOf("PK_FECHA_DATOS") & vbLf & _
        "  AND PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & Tbl(DS_PROD, T_RISK) & _
        MaxAsOf() & "), INTERVAL " & CACHE_ANOS & " YEAR)"
    SQLBloqueRisk = _
        "SELECT FORMAT_DATE('%Y-%m-%d', PK_FECHA_DATOS) AS fecha," & vbLf & _
        "       PK_PORTFOLIO_ID," & vbLf & _
        "       PK_CRITERIO_AGREGACION AS criterio," & vbLf & _
        "       PK_ETIQUETA_AGREGACION AS etiqueta," & vbLf & _
        "       PK_VARIABLE_TARGET AS variable," & vbLf & _
        "       FORMAT('%.10f', CAST(VALOR AS FLOAT64)) AS valor" & vbLf & _
        "FROM " & Tbl(DS_PROD, T_RISK) & vbLf & w & vbLf & _
        "ORDER BY PK_PORTFOLIO_ID, PK_FECHA_DATOS"
End Function

Private Function SQLBloquePos(ByVal ents As String) As String
    Dim gics As String, bics As String, geo As String, pais As String
    Dim divc As String, rat As String, act As String
    gics = Cfg("SECTOR_COL", SECTOR_COL)
    bics = "CLASSIFICATION_BICS"
    geo = Cfg("GEO_COL", GEO_COL)
    pais = Cfg("PAIS_COL", "FCCOUNTRY")
    divc = Cfg("DIV_COL", DIV_COL)
    rat = Cfg("RATING_COL", RATING_COL)
    act = Cfg("ACTIVO_COL", "INSTRUMENT_TYPE")
    ' Solo lo necesario para composicion (Peso): clasificaciones + valoracion.
    ' Columnas: PID | gics | bics | geo(zona) | pais | divisa | rating | activo | valor
    ' Fecha por cartera (MAX por portfolio), como en la consulta que funciona.
    SQLBloquePos = _
        "SELECT p.PK_PORTFOLIO_ID," & vbLf & _
        "       v." & gics & " AS gics, v." & bics & " AS bics," & vbLf & _
        "       v." & geo & " AS geo, v." & pais & " AS pais, v." & divc & " AS divisa," & vbLf & _
        "       v." & rat & " AS rating, v." & act & " AS activo," & vbLf & _
        "       FORMAT('%.10f', CAST(p." & PosValor() & " AS FLOAT64)) AS valor" & vbLf & _
        "FROM " & TblPos() & " p" & vbLf & JoinValores & _
        "WHERE p.PK_PORTFOLIO_ID IN (" & ents & ")" & vbLf & _
        "  AND p.PK_FECHA_DATOS = (SELECT MAX(sub.PK_FECHA_DATOS) FROM " & TblPos() & " sub" & vbLf & _
        "                          WHERE sub.PK_PORTFOLIO_ID = p.PK_PORTFOLIO_ID" & AndAsOf("sub.PK_FECHA_DATOS") & ")" & vbLf & _
        "ORDER BY p.PK_PORTFOLIO_ID"
End Function

' Ejecuta una SQL y escribe el resultado (cabeceras en fila 1, datos desde
' fila 2) en la hoja Panel a partir de la columna 'baseCol', como TEXTO.
' True si fue bien; deja el error en msg.
Private Function DescargarBloque(ByVal ws As Worksheet, ByVal sql As String, _
        ByVal baseCol As Long, ByVal ancho As Long, ByRef msg As String) As Boolean
    Dim cn As Object, rs As Object, j As Long
    On Error GoTo fallo
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 180
    cn.CursorLocation = 3
    cn.Open CfgConn()
    Set rs = cn.Execute(sql)
    ws.Range(ws.Cells(1, baseCol), ws.Cells(BLK_ULTFILA, baseCol + ancho - 1)).NumberFormat = "@"
    For j = 0 To rs.Fields.Count - 1
        ws.Cells(1, baseCol + j).Value = rs.Fields(j).Name
    Next j
    If Not rs.EOF Then ws.Cells(2, baseCol).CopyFromRecordset rs
    rs.Close: cn.Close
    DescargarBloque = True
    Exit Function
fallo:
    msg = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    On Error GoTo 0
    DescargarBloque = False
End Function

' True si los bloques amplios existen y cubren TODAS las entidades elegidas.
Private Function BloqueCubre(ByVal ws As Worksheet) As Boolean
    Dim marca As String
    marca = CStr(ws.Cells(1, BLK_MARK_COL).Value)
    If Len(marca) = 0 Then Exit Function
    Dim ok As Boolean: ok = True
    If Len(Trim(mId1)) > 0 Then If InStr(marca, "'" & UCase(mId1) & "'") = 0 Then ok = False
    If Len(Trim(mId2)) > 0 Then If InStr(marca, "'" & UCase(mId2) & "'") = 0 Then ok = False
    If Len(Trim(mId3)) > 0 Then If InStr(marca, "'" & UCase(mId3) & "'") = 0 Then ok = False
    BloqueCubre = ok
End Function

' Ultima fila con datos de un bloque (por su primera columna).
Private Function UltFilaBloque(ByVal ws As Worksheet, ByVal baseCol As Long) As Long
    UltFilaBloque = ws.Cells(BLK_ULTFILA, baseCol).End(xlUp).Row
End Function

' Slot (1/2/3) de un pk_portfolio_id segun B4/B5/B6; 0 si no es ninguno.
Private Function SlotDe(ByVal pid As String) As Long
    If IgualId(pid, mId1) Then
        SlotDe = 1
    ElseIf IgualId(pid, mId2) Then
        SlotDe = 2
    ElseIf IgualId(pid, mId3) Then
        SlotDe = 3
    End If
End Function

' Convierte "yyyy-mm-dd" en Date (independiente del idioma).
Private Function FechaDe(ByVal s As String) As Date
    s = Trim(CStr(s))
    If Len(s) < 10 Then Exit Function
    FechaDe = DateSerial(CLng(Left(s, 4)), CLng(Mid(s, 6, 2)), CLng(Mid(s, 9, 2)))
End Function

' Etiqueta de bucket de una fecha segun la dimension (igual que BucketExpr en SQL).
Private Function BucketLocal(ByVal d As Date, ByVal dimen As String) As String
    Dim y As Long, m As Long
    y = Year(d): m = Month(d)
    Select Case Fold(dimen)
        Case "diario":     BucketLocal = Format(d, "yyyy-mm-dd")
        Case "semanal":    BucketLocal = y & "-W" & Format(DatePart("ww", d, vbMonday, vbFirstFourDays), "00")
        Case "mensual":    BucketLocal = Format(d, "yyyy-mm")
        Case "trimestral": BucketLocal = y & "-T" & (Int((m - 1) / 3) + 1)
        Case "semestral":  BucketLocal = y & "-S" & IIf(m <= 6, 1, 2)
        Case "anual":      BucketLocal = CStr(y)
        Case Else:         BucketLocal = Format(d, "yyyy-mm")
    End Select
End Function

' Fecha de inicio de la ventana (Anual alineada a ano natural; resto, movil).
' Primer dia del periodo de CALENDARIO CERRADO anterior (mes/trimestre/ano) a dMax.
Private Function InicioAnterior(ByVal dMax As Date, ByVal clase As String) As Date
    Select Case clase
        Case "M": InicioAnterior = DateSerial(Year(dMax), Month(dMax) - 1, 1)
        Case "Q": InicioAnterior = DateSerial(Year(dMax), (Int((Month(dMax) - 1) / 3)) * 3 - 2, 1)
        Case "Y": InicioAnterior = DateSerial(Year(dMax) - 1, 1, 1)
        Case Else: InicioAnterior = dMax
    End Select
End Function

' Ultimo dia de la ventana: dMax en periodos normales; el ultimo dia del periodo
' cerrado anterior para "mes/trimestre/ano anterior".
Private Function FinVentana(ByVal dMax As Date, ByVal per As String) As Date
    Dim pa As String: pa = PerAnterior(per)
    Select Case pa
        Case "M": FinVentana = DateSerial(Year(dMax), Month(dMax), 1) - 1
        Case "Q": FinVentana = DateSerial(Year(dMax), (Int((Month(dMax) - 1) / 3)) * 3 + 1, 1) - 1
        Case "Y": FinVentana = DateSerial(Year(dMax), 1, 1) - 1
        Case Else: FinVentana = dMax
    End Select
End Function

Private Function InicioVentana(ByVal dMax As Date, ByVal per As String, ByVal dimen As String) As Date
    Dim p As String, n As Long
    Dim pa As String: pa = PerAnterior(per)
    If pa <> "" Then InicioVentana = InicioAnterior(dMax, pa): Exit Function
    p = UCase(Trim(per))
    ' Periodos "a fecha" (independientes de la dimension): desde el inicio del
    ' mes/trimestre/ano/semana en curso hasta hoy. YTD -> desde el 1-ene del ano
    ' de dMax (NO un ano movil, que era el bug: YTD daba 12 meses hacia atras).
    Select Case p
        Case "MTD":      InicioVentana = DateSerial(Year(dMax), Month(dMax), 1): Exit Function
        Case "QTD":      InicioVentana = DateSerial(Year(dMax), Int((Month(dMax) - 1) / 3) * 3 + 1, 1): Exit Function
        Case "YTD":      InicioVentana = DateSerial(Year(dMax), 1, 1): Exit Function
        Case "WTD":      InicioVentana = dMax - (Weekday(dMax, vbMonday) - 1): Exit Function
        Case "1D", "DTD": InicioVentana = dMax: Exit Function
    End Select
    If Fold(dimen) = "anual" Then
        InicioVentana = DateSerial(Year(dMax) - (AnyosPeriodo(per) - 1), 1, 1): Exit Function
    End If
    If Len(p) >= 2 And IsNumeric(Left(p, Len(p) - 1)) Then
        n = CLng(Left(p, Len(p) - 1))
        If Right(p, 1) = "A" Then InicioVentana = DateAdd("yyyy", -n, dMax): Exit Function
        If Right(p, 1) = "M" Then InicioVentana = DateAdd("m", -n, dMax): Exit Function
    End If
    InicioVentana = DateAdd("yyyy", -1, dMax)
End Function

' Primer dia del bucket (mes/trimestre/etc.) que contiene la fecha d. Se usa para
' alinear el inicio de la ventana al comienzo del periodo y NO cortar el primer
' bucket por la mitad (si no, el mes inicial saldria parcial y con retorno erroneo).
Private Function InicioBucket(ByVal d As Date, ByVal dimen As String) As Date
    Dim y As Long, m As Long
    y = Year(d): m = Month(d)
    Select Case Fold(dimen)
        Case "diario":     InicioBucket = d
        Case "semanal":    InicioBucket = d - (Weekday(d, vbMonday) - 1)
        Case "mensual":    InicioBucket = DateSerial(y, m, 1)
        Case "trimestral": InicioBucket = DateSerial(y, Int((m - 1) / 3) * 3 + 1, 1)
        Case "semestral":  InicioBucket = DateSerial(y, IIf(m <= 6, 1, 7), 1)
        Case "anual":      InicioBucket = DateSerial(y, 1, 1)
        Case Else:         InicioBucket = DateSerial(y, m, 1)
    End Select
End Function

' ================  TABLAS COMO FORMULAS (calculo auditable)  ================
' Las tablas D:H (y la matriz apilada) se escriben como FORMULAS de Excel que
' leen el bloque descargado de BigQuery, para que el usuario VEA el calculo en
' la celda y se recalcule solo. Solo hay UNA metrica en pantalla, asi que se
' reutiliza una region de columnas auxiliares (se reescribe en cada calculo) y
' unos rangos con nombre (f_pid, f_k1, f_a1, f_a2, f_crit, f_var, f_mes) que se
' re-apuntan al bloque activo. Las columnas auxiliares HLP_* estan declaradas
' arriba, en la seccion de constantes del modulo.

' Literal de cadena entrecomillado para meter dentro de una formula.
Private Function Q(ByVal s As String) As String
    Q = Chr(34) & Replace(CStr(s), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

' (Re)define un nombre de libro apuntando a <HojaActiva>!col$2:col$lastR (la hoja
' es Panel() -> el Panel real o la copia que se este refrescando).
Private Sub FijarNombre(ByVal nm As String, ByVal col As Long, ByVal lastR As Long)
    Dim ws As Worksheet: Set ws = Panel()
    Dim r2 As Long: r2 = lastR: If r2 < 2 Then r2 = 2
    Dim ref As String
    ref = "='" & ws.Name & "'!" & ws.Cells(2, col).Address(True, True) & ":" & ws.Cells(r2, col).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=ref
End Sub

' Claves de un diccionario ordenadas ASCENDENTE (texto). Sirve para poner los
' buckets temporales en orden cronologico (las etiquetas "yyyy-mm", "yyyy-Tn",
' "yyyy" ordenan bien alfabeticamente).
Private Function ClavesOrdenadas(ByVal d As Object) As Variant
    Dim a() As String, i As Long, j As Long, t As String, k As Variant
    ReDim a(0 To d.Count - 1)
    i = 0
    For Each k In d.Keys: a(i) = CStr(k): i = i + 1: Next k
    For i = 0 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    ClavesOrdenadas = a
End Function

' Escribe una columna de la tabla (E/F/G/H) con la formula de rentabilidad para
' una entidad. acum -> compuesto acumulado desde el primer bucket; esBmk -> usa
' el factor de benchmark (f_a2).
Private Sub EscribirColRentab(ByVal ws As Worksheet, ByVal col As Long, ByVal id As String, _
                              ByVal nb As Long, ByVal acum As Boolean, ByVal esBmk As Boolean)
    If Len(id) = 0 Or nb < 1 Then Exit Sub
    Dim fac As String: fac = IIf(esBmk, "f_a2", "f_a1")
    Dim qid As String: qid = Q(UCase(Trim(id)))
    Dim f As String
    If acum Then
        f = "=IFERROR(EXP(SUMPRODUCT((f_k1>=$D$3)*(f_k1<=$D3)*(f_pid=" & qid & _
            ")*LN(" & fac & ")))-1," & Q("") & ")"
    Else
        f = "=IFERROR(EXP(SUMPRODUCT((f_k1=$D3)*(f_pid=" & qid & _
            ")*LN(" & fac & ")))-1," & Q("") & ")"
    End If
    ws.Range(ws.Cells(3, col), ws.Cells(2 + nb, col)).Formula = f
End Sub

' Escribe una columna (E/F/G) con la formula de riesgo (ultimo valor del bucket/
' etiqueta): LOOKUP(2,1/cond,valor) devuelve el ultimo match = fecha mas reciente
' (el bloque viene ORDER BY pid, fecha).
Private Sub EscribirColRiesgo(ByVal ws As Worksheet, ByVal col As Long, ByVal id As String, _
                              ByVal crit As String, ByVal varT As String, ByVal nb As Long)
    If Len(id) = 0 Or nb < 1 Then Exit Sub
    Dim cond As String
    cond = "(f_k1=$D3)*(f_pid=" & Q(UCase(Trim(id))) & ")*(f_crit=" & Q(Trim(crit)) & ")"
    If Len(varT) > 0 Then cond = cond & "*(f_var=" & Q(Trim(varT)) & ")"
    ws.Range(ws.Cells(3, col), ws.Cells(2 + nb, col)).Formula = _
        "=IFERROR(LOOKUP(2,1/(" & cond & "),f_a1)," & Q("") & ")"
End Sub

' Escribe una columna (E/F/G) con la formula de composicion. importe=False -> %
' (peso = valor de la categoria / total de la cartera en su ultimo mes).
Private Sub EscribirColComp(ByVal ws As Worksheet, ByVal col As Long, ByVal id As String, _
                            ByVal mesLit As String, ByVal nb As Long, ByVal importe As Boolean)
    If Len(id) = 0 Or nb < 1 Then Exit Sub
    Dim qid As String: qid = Q(UCase(Trim(id)))
    Dim qm As String: qm = Q(Trim(mesLit))
    Dim num As String
    num = "SUMIFS(f_a1,f_pid," & qid & ",f_mes," & qm & ",f_k1,$D3)"
    Dim f As String
    If importe Then
        f = "=IFERROR(" & num & "," & Q("") & ")"
    Else
        f = "=IFERROR(" & num & "/SUMIFS(f_a1,f_pid," & qid & ",f_mes," & qm & ")," & Q("") & ")"
    End If
    ws.Range(ws.Cells(3, col), ws.Cells(2 + nb, col)).Formula = f
End Sub

' Calcula Rentabilidad (compuesta por bucket) EN LOCAL desde el bloque RET de
' la hoja Panel (columna W) y la vuelca a D:H COMO FORMULAS. False si no hay datos.
Private Function LocalRentabilidad(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, i As Long, n As Long
    c0 = BLK_RET_COL
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function

    Dim per As String, dimen As String, conBmk As Boolean, acum As Boolean
    per = Trim(CStr(ws.Range("B11").Value))
    dimen = Trim(CStr(ws.Range("B9").Value))
    conBmk = ConBenchmark(ws)
    acum = (Trim(CStr(ws.Range("B8").Value)) = "Rentab. acum.")

    Dim blk As Variant
    blk = ws.Range(ws.Cells(2, c0), ws.Cells(lastR, c0 + 3)).Value   ' pid|fecha|twr|twr_bmk
    n = lastR - 1

    ' Fecha maxima y ventana alineada al inicio del bucket (mes/trim/etc.).
    Dim dMax As Date, dd As Date
    For i = 1 To n
        dd = FechaDe(CStr(blk(i, 2)))
        If dd > dMax Then dMax = dd
    Next i
    If dMax = 0 Then Exit Function
    Dim ini As Date: ini = InicioBucket(InicioVentana(dMax, per, dimen), dimen)
    Dim dEnd As Date: dEnd = FinVentana(dMax, per)   ' = dMax salvo en "... anterior"

    ' Columnas auxiliares (TODAS las filas): clave=bucket, factor(1+twr) y factor
    ' benchmark. La formula compone el bucket ENTERO (por eso no hay que deduplicar
    ' ni recortar aqui: SUMPRODUCT suma LN de todos los dias del bucket).
    ReDim h1(1 To n, 1 To 1) As Variant
    ReDim a1(1 To n, 1 To 1) As Variant
    ReDim a2(1 To n, 1 To 1) As Variant
    ReDim pidC(1 To n, 1 To 1) As Variant
    Dim shown As Object: Set shown = CreateObject("Scripting.Dictionary")
    Dim bkt As String
    For i = 1 To n
        pidC(i, 1) = UCase(Trim(CStr(blk(i, 1))))
        dd = FechaDe(CStr(blk(i, 2)))
        If dd = 0 Then
            h1(i, 1) = "": a1(i, 1) = 1#: a2(i, 1) = 1#
        Else
            bkt = BucketLocal(dd, dimen)
            h1(i, 1) = bkt
            a1(i, 1) = 1# + NumDbl(blk(i, 3))
            a2(i, 1) = 1# + NumDbl(blk(i, 4))
            ' Factor <= 0 (un dia de -100% o dato corrupto) romperia LN() y dejaria
            ' TODA la tabla en blanco: lo acotamos a un positivo minusculo.
            If a1(i, 1) <= 0 Then a1(i, 1) = 0.000001
            If a2(i, 1) <= 0 Then a2(i, 1) = 0.000001
            If dd >= ini And dd <= dEnd Then
                If Not shown.Exists(bkt) Then shown.Add bkt, 1
            End If
        End If
    Next i
    If shown.Count = 0 Then Exit Function

    ws.Range(ws.Cells(2, HLP_K1), ws.Cells(lastR, HLP_K1)).Value = h1
    ws.Range(ws.Cells(2, HLP_A1), ws.Cells(lastR, HLP_A1)).Value = a1
    ws.Range(ws.Cells(2, HLP_A2), ws.Cells(lastR, HLP_A2)).Value = a2
    ws.Range(ws.Cells(2, HLP_PID), ws.Cells(lastR, HLP_PID)).Value = pidC
    FijarNombre "f_pid", HLP_PID, lastR
    FijarNombre "f_k1", HLP_K1, lastR
    FijarNombre "f_a1", HLP_A1, lastR
    FijarNombre "f_a2", HLP_A2, lastR

    Dim labs As Variant: labs = ClavesOrdenadas(shown)
    Dim nb As Long: nb = UBound(labs) - LBound(labs) + 1

    GuardarPreviewSiNoExiste ws
    Application.EnableEvents = False
    LimpiarTablaNormal ws
    Dim dcol() As Variant: ReDim dcol(1 To nb, 1 To 1)
    For i = 1 To nb: dcol(i, 1) = labs(i - 1): Next i
    ws.Range(ws.Cells(3, 4), ws.Cells(2 + nb, 4)).Value = dcol
    EscribirColRentab ws, 5, mId1, nb, acum, False
    EscribirColRentab ws, 6, mId2, nb, acum, False
    EscribirColRentab ws, 7, mId3, nb, acum, False
    If conBmk Then
        EscribirColRentab ws, 8, mId1, nb, acum, True
        ws.Range("H2").Value = EtiquetaBenchmark(ws)
    End If
    ws.Range("E3:H402").NumberFormat = FormatoMetrica(Trim(CStr(ws.Range("B8").Value)))
    Application.EnableEvents = True
    LocalRentabilidad = True
End Function

' Restaura la tabla del grafico (D:H) y borra los restos de la vista apilada
' (columnas E..V y sus cabeceras). Se llama antes de cada volcado "normal".
Private Sub LimpiarTablaNormal(ByVal ws As Worksheet)
    ws.Range(ws.Cells(3, 4), ws.Cells(402, TCMP_COL + MAXSER)).ClearContents   ' D3:V402
    ws.Range(ws.Cells(2, 9), ws.Cells(2, TCMP_COL + MAXSER)).ClearContents      ' I2:V2
    ws.Range("D2").Value = "Categoria"
    ws.Range("E2").Formula = "=B4"
    ws.Range("F2").Formula = "=IF(B5=""(ninguna)"","""",B5)"
    ws.Range("G2").Formula = "=IF(B6=""(ninguna)"","""",B6)"
    ' La cabecera "Benchmark" (col H) SOLO la escribe quien vuelca datos de benchmark
    ' (rentabilidad con benchmark). Metricas sin benchmark (Composicion/Duracion/TIR)
    ' la dejan vacia para no mostrar una columna "Benchmark" fantasma.
    ws.Range("H2").Value = ""
End Sub

' Escribe en la tabla del grafico (D:H) una serie categoria -> valores por slot.
' 'cats' es el diccionario categoria->fila; 'vals' tiene claves "slot|cat".
Private Sub VolcarLocal(ByVal ws As Worksheet, ByVal cats As Object, ByVal vals As Object)
    GuardarPreviewSiNoExiste ws
    Application.EnableEvents = False
    LimpiarTablaNormal ws
    Dim vc As Variant, rr As Long
    For Each vc In cats.Keys
        rr = cats(vc)
        ws.Cells(rr, 4).Value = vc
        If vals.Exists("1|" & vc) Then ws.Cells(rr, 5).Value = vals("1|" & vc)
        If vals.Exists("2|" & vc) Then ws.Cells(rr, 6).Value = vals("2|" & vc)
        If vals.Exists("3|" & vc) Then ws.Cells(rr, 7).Value = vals("3|" & vc)
    Next vc
    ws.Range("E3:H402").NumberFormat = FormatoMetrica(Trim(CStr(ws.Range("B8").Value)))
    Application.EnableEvents = True
End Sub

' Calcula Duracion/TIR EN LOCAL desde el bloque RISK de la hoja Panel.
' Temporal (Diario..Anual): ultimo valor de cada bucket. Dimensional (Activo/
' Geografia/Divisa): ultimo valor por etiqueta a la fecha mas reciente.
Private Function LocalRiesgo(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, i As Long, n As Long
    c0 = BLK_RISK_COL                       ' fecha|PID|criterio|etiqueta|variable|valor
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function

    Dim met As String, dimen As String, esTiempo As Boolean
    Dim crit As String, varT As String
    met = Trim(CStr(ws.Range("B8").Value))
    dimen = Trim(CStr(ws.Range("B9").Value))
    esTiempo = DimEsTiempo(dimen)
    If InStr(Fold(met), "duraci") = 1 Then
        varT = VarTarget(met)               ' DuracionModificada / DuracionMacaulay...
        crit = IIf(esTiempo, "Duracion", DimACriterio(dimen))
    ElseIf met = "TIR" Then
        varT = ""                           ' TIR es criterio, sin variable target
        crit = "TIR"
    ElseIf Fold(met) = "cmr" Then
        crit = Cfg("RISK_CRIT_CMR", "CMR")
        varT = Cfg("RISK_VAR_CMR", "")
    ElseIf Left(Fold(met), 3) = "var" Then  ' VaR (y variantes VaR 95 / VaR 99...)
        crit = Cfg("RISK_CRIT_VAR", "VaR")
        ' Variable target: si la metrica trae un sufijo (VaR 95) se busca en config
        ' RISK_VAR_<sufijo>; si no, el generico RISK_VAR_VAR (por defecto sin filtro).
        Dim suf As String: suf = Replace(Fold(met), "var", "", 1, 1)
        suf = Trim(Replace(Replace(suf, "%", ""), " ", ""))
        If Len(suf) > 0 Then varT = Cfg("RISK_VAR_VAR_" & suf, "") Else varT = Cfg("RISK_VAR_VAR", "")
    Else
        Exit Function
    End If
    If Len(crit) = 0 Then Exit Function     ' dimension sin desglose -> que lo resuelva la consulta

    Dim per As String: per = Trim(CStr(ws.Range("B11").Value))
    Dim blk As Variant
    blk = ws.Range(ws.Cells(2, c0), ws.Cells(lastR, c0 + 5)).Value   ' fecha|pid|crit|etiq|var|valor
    n = lastR - 1

    ' Fecha maxima (de las filas de este criterio/variable) y ventana alineada.
    Dim dMax As Date, dd As Date
    For i = 1 To n
        If Fold(CStr(blk(i, 3))) = Fold(crit) Then
            If Len(varT) = 0 Or Fold(CStr(blk(i, 5))) = Fold(varT) Then
                dd = FechaDe(CStr(blk(i, 1)))
                If dd > dMax Then dMax = dd
            End If
        End If
    Next i
    If dMax = 0 Then Exit Function
    Dim ini As Date: ini = InicioBucket(InicioVentana(dMax, per, dimen), dimen)

    ' Auxiliares: clave (bucket temporal o etiqueta) y valor numerico, para TODAS
    ' las filas. criterio/variable se comparan en la formula (rangos f_crit/f_var).
    ReDim h1(1 To n, 1 To 1) As Variant
    ReDim a1(1 To n, 1 To 1) As Variant
    ReDim pidC(1 To n, 1 To 1) As Variant
    ReDim critC(1 To n, 1 To 1) As Variant
    ReDim varC(1 To n, 1 To 1) As Variant
    Dim shown As Object: Set shown = CreateObject("Scripting.Dictionary")
    Dim cat As String, esta As Boolean
    For i = 1 To n
        a1(i, 1) = NumDbl(blk(i, 6))
        pidC(i, 1) = UCase(Trim(CStr(blk(i, 2))))
        critC(i, 1) = Trim(CStr(blk(i, 3)))
        varC(i, 1) = Trim(CStr(blk(i, 5)))
        esta = (Fold(CStr(blk(i, 3))) = Fold(crit))
        If esta And Len(varT) > 0 Then esta = (Fold(CStr(blk(i, 5))) = Fold(varT))
        If esTiempo Then
            dd = FechaDe(CStr(blk(i, 1)))
            If dd = 0 Then h1(i, 1) = "" Else h1(i, 1) = BucketLocal(dd, dimen)
            If esta And dd >= ini And dd <= dMax And Len(CStr(h1(i, 1))) > 0 Then
                If Not shown.Exists(CStr(h1(i, 1))) Then shown.Add CStr(h1(i, 1)), 1
            End If
        Else
            cat = Trim(CStr(blk(i, 4)))
            h1(i, 1) = cat
            If esta And Len(cat) > 0 Then
                If Not shown.Exists(cat) Then shown.Add cat, 1
            End If
        End If
    Next i
    If shown.Count = 0 Then Exit Function

    ws.Range(ws.Cells(2, HLP_K1), ws.Cells(lastR, HLP_K1)).Value = h1
    ws.Range(ws.Cells(2, HLP_A1), ws.Cells(lastR, HLP_A1)).Value = a1
    ws.Range(ws.Cells(2, HLP_PID), ws.Cells(lastR, HLP_PID)).Value = pidC
    ws.Range(ws.Cells(2, HLP_C1), ws.Cells(lastR, HLP_C1)).Value = critC
    ws.Range(ws.Cells(2, HLP_C2), ws.Cells(lastR, HLP_C2)).Value = varC
    FijarNombre "f_pid", HLP_PID, lastR
    FijarNombre "f_crit", HLP_C1, lastR
    FijarNombre "f_var", HLP_C2, lastR
    FijarNombre "f_k1", HLP_K1, lastR
    FijarNombre "f_a1", HLP_A1, lastR

    Dim labs As Variant, nb As Long, kk As Variant
    If esTiempo Then
        labs = ClavesOrdenadas(shown)          ' cronologico
    Else
        ReDim labs(0 To shown.Count - 1)        ' orden de aparicion
        i = 0
        For Each kk In shown.Keys: labs(i) = CStr(kk): i = i + 1: Next kk
    End If
    nb = UBound(labs) - LBound(labs) + 1

    GuardarPreviewSiNoExiste ws
    Application.EnableEvents = False
    LimpiarTablaNormal ws
    Dim dcol() As Variant: ReDim dcol(1 To nb, 1 To 1)
    For i = 1 To nb: dcol(i, 1) = labs(i - 1): Next i
    ws.Range(ws.Cells(3, 4), ws.Cells(2 + nb, 4)).Value = dcol
    EscribirColRiesgo ws, 5, mId1, crit, varT, nb
    EscribirColRiesgo ws, 6, mId2, crit, varT, nb
    EscribirColRiesgo ws, 7, mId3, crit, varT, nb
    ws.Range("E3:H402").NumberFormat = FormatoMetrica(met)
    Application.EnableEvents = True
    LocalRiesgo = True
End Function

' Devuelve la columna del bloque POS que corresponde a la clasificacion elegida.
Private Function ColClasifPos(ByVal dimen As String) As Long
    ' POS: PID(+0) gics(+1) bics(+2) geo(+3) pais(+4) divisa(+5) rating(+6) activo(+7) valor(+8)
    Select Case dimen
        Case "Sector":     ColClasifPos = BLK_POS_COL + IIf(InStr(UCase(Cfg("SECTOR_COL", SECTOR_COL)), "BICS") > 0, 2, 1)
        Case "Industria":  ColClasifPos = BLK_POS_COL + IIf(InStr(UCase(Cfg("IND_COL", IND_COL)), "BICS") > 0, 2, 1)
        Case "Continente": ColClasifPos = BLK_POS_COL + 3   ' FCCOUNTRYZONE (zona)
        Case "Pais":       ColClasifPos = BLK_POS_COL + 4   ' FCCOUNTRY (pais)
        Case "Divisa":     ColClasifPos = BLK_POS_COL + 5
        Case "Rating":     ColClasifPos = BLK_POS_COL + 6
        Case "Activo":     ColClasifPos = BLK_POS_COL + 7   ' tipo de activo (maestro)
        Case Else:         ColClasifPos = 0
    End Select
End Function

' Composicion (Peso) EN LOCAL desde el MISMO bloque APIL que la composicion
' apilada (comparten datos): toma el ULTIMO mes de cada cartera y reparte el
' PESO (%) por la clasificacion elegida (B9), por slot (E/F/G).
Private Function LocalComposicion(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, i As Long, n As Long, colCat As Long, clas As String
    c0 = BLK_APIL_COL                         ' PID | mes | gics|bics|geo|pais|divisa|activo | valor
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function
    clas = Trim(CStr(ws.Range("B9").Value))   ' en Composicion, B9 = la clasificacion
    colCat = ColClasifApil(clas)
    If colCat = 0 Then Exit Function
    Dim iCat As Long: iCat = colCat - c0 + 1  ' indice (1-based) de la clasificacion en el array
    Dim importe As Boolean: importe = (Fold(Trim(CStr(ws.Range("B8").Value))) = "importe")

    Dim blk As Variant
    blk = ws.Range(ws.Cells(2, c0), ws.Cells(lastR, c0 + 8)).Value
    n = lastR - 1

    ' Ultimo mes disponible por slot (los meses "YYYY-MM" ordenan cronologicamente).
    Dim mesMax(1 To 3) As String, slot As Long, mes As String
    For i = 1 To n
        slot = SlotDe(UCase(Trim(CStr(blk(i, 1)))))
        If slot > 0 Then
            mes = Trim(CStr(blk(i, 2)))
            If mes > mesMax(slot) Then mesMax(slot) = mes
        End If
    Next i

    ' Auxiliares: etiqueta de la categoria y valoracion (numerica), TODAS las filas.
    ReDim h1(1 To n, 1 To 1) As Variant
    ReDim a1(1 To n, 1 To 1) As Variant
    ReDim pidC(1 To n, 1 To 1) As Variant
    ReDim mesC(1 To n, 1 To 1) As Variant
    Dim shown As Object: Set shown = CreateObject("Scripting.Dictionary")
    Dim cat As String
    For i = 1 To n
        cat = EtiquetaClasif(clas, CStr(blk(i, iCat)))
        h1(i, 1) = cat
        a1(i, 1) = NumDbl(blk(i, 9))
        pidC(i, 1) = UCase(Trim(CStr(blk(i, 1))))
        mesC(i, 1) = Trim(CStr(blk(i, 2)))
        slot = SlotDe(UCase(Trim(CStr(blk(i, 1)))))
        If slot > 0 Then
            If Trim(CStr(blk(i, 2))) = mesMax(slot) And Len(cat) > 0 Then
                If Not shown.Exists(cat) Then shown.Add cat, 1
            End If
        End If
    Next i
    If shown.Count = 0 Then Exit Function

    ws.Range(ws.Cells(2, HLP_K1), ws.Cells(lastR, HLP_K1)).Value = h1
    ws.Range(ws.Cells(2, HLP_A1), ws.Cells(lastR, HLP_A1)).Value = a1
    ws.Range(ws.Cells(2, HLP_PID), ws.Cells(lastR, HLP_PID)).Value = pidC
    ws.Range(ws.Cells(2, HLP_C1), ws.Cells(lastR, HLP_C1)).Value = mesC
    FijarNombre "f_pid", HLP_PID, lastR
    FijarNombre "f_mes", HLP_C1, lastR
    FijarNombre "f_k1", HLP_K1, lastR
    FijarNombre "f_a1", HLP_A1, lastR

    Dim labs() As String, nb As Long, kk As Variant
    ReDim labs(0 To shown.Count - 1)
    nb = 0
    For Each kk In shown.Keys: labs(nb) = CStr(kk): nb = nb + 1: Next kk

    GuardarPreviewSiNoExiste ws
    Application.EnableEvents = False
    LimpiarTablaNormal ws
    Dim dcol() As Variant: ReDim dcol(1 To nb, 1 To 1)
    For i = 1 To nb: dcol(i, 1) = labs(i - 1): Next i
    ws.Range(ws.Cells(3, 4), ws.Cells(2 + nb, 4)).Value = dcol
    EscribirColComp ws, 5, mId1, mesMax(1), nb, importe
    EscribirColComp ws, 6, mId2, mesMax(2), nb, importe
    EscribirColComp ws, 7, mId3, mesMax(3), nb, importe
    ws.Range("E3:H402").NumberFormat = FormatoMetrica(Trim(CStr(ws.Range("B8").Value)))
    Application.EnableEvents = True
    LocalComposicion = True
End Function

' =====================  COMPOSICION APILADA EN EL TIEMPO  ==================
' Nombre legible del sector GICS (codigo de 2 digitos). "" si no lo reconoce.
Private Function GicsSector(ByVal code2 As String) As String
    Select Case Trim(code2)
        Case "10": GicsSector = "Energia"
        Case "15": GicsSector = "Materiales"
        Case "20": GicsSector = "Industria"
        Case "25": GicsSector = "Consumo discrecional"
        Case "30": GicsSector = "Consumo basico"
        Case "35": GicsSector = "Salud"
        Case "40": GicsSector = "Financiero"
        Case "45": GicsSector = "Tecnologia"
        Case "50": GicsSector = "Comunicaciones"
        Case "55": GicsSector = "Utilities"
        Case "60": GicsSector = "Inmobiliario"
        Case Else: GicsSector = ""
    End Select
End Function

' Nombre legible del grupo de industria GICS (codigo de 4 digitos). "" si no lo
' reconoce (p.ej. si la columna es BICS, con otra numeracion).
Private Function GicsIndustria(ByVal code4 As String) As String
    Select Case Trim(code4)
        Case "1010": GicsIndustria = "Energia"
        Case "1510": GicsIndustria = "Materiales"
        Case "2010": GicsIndustria = "Bienes de equipo"
        Case "2020": GicsIndustria = "Servicios comerciales"
        Case "2030": GicsIndustria = "Transporte"
        Case "2510": GicsIndustria = "Automocion"
        Case "2520": GicsIndustria = "Consumo duradero y textil"
        Case "2530": GicsIndustria = "Servicios al consumo"
        Case "2550": GicsIndustria = "Distribucion consumo discrecional"
        Case "3010": GicsIndustria = "Distribucion consumo basico"
        Case "3020": GicsIndustria = "Alimentacion, bebidas y tabaco"
        Case "3030": GicsIndustria = "Hogar y cuidado personal"
        Case "3510": GicsIndustria = "Equipos y servicios de salud"
        Case "3520": GicsIndustria = "Farmacia y biotecnologia"
        Case "4010": GicsIndustria = "Bancos"
        Case "4020": GicsIndustria = "Servicios financieros"
        Case "4030": GicsIndustria = "Seguros"
        Case "4510": GicsIndustria = "Software y servicios"
        Case "4520": GicsIndustria = "Hardware tecnologico"
        Case "4530": GicsIndustria = "Semiconductores"
        Case "5010": GicsIndustria = "Telecomunicaciones"
        Case "5020": GicsIndustria = "Medios y entretenimiento"
        Case "5510": GicsIndustria = "Utilities"
        Case "6010": GicsIndustria = "Inmobiliario"
        Case Else:   GicsIndustria = ""
    End Select
End Function

' Etiqueta legible de una categoria de clasificacion: agrupa GICS a nivel sector
' (2 digitos) o industria (4 digitos) segun config y le pone NOMBRE. Asi el eje/
' leyenda no muestra codigos de 8 digitos. Si no reconoce el codigo (p.ej. BICS,
' o un activo sin GICS como liquidez), deja el codigo tal cual.
Private Function EtiquetaClasif(ByVal clas As String, ByVal raw As String) As String
    Dim s As String, n As Long, nm As String
    s = Trim(CStr(raw))
    Select Case clas
        Case "Sector"
            n = CLng(Val(Cfg("SECTOR_DIG", "2")))
            If n > 0 And Len(s) >= n Then s = Left(s, n)
            nm = GicsSector(s): If Len(nm) > 0 Then s = nm
        Case "Industria"
            n = CLng(Val(Cfg("IND_DIG", "4")))
            If n > 0 And Len(s) >= n Then s = Left(s, n)
            nm = GicsIndustria(s): If Len(nm) > 0 Then s = nm
    End Select
    If Len(s) = 0 Then s = "(sin dato)"
    EtiquetaClasif = s
End Function

' SQL del BLOQUE de composicion apilada: snapshot MENSUAL (ultima foto de cada
' mes) por cartera, con todas las clasificaciones + valoracion, para los ultimos
' CACHE_ANOS anos. Se descarga UNA vez (como RET/RISK/POS) y luego se trocea en
' local por periodo/granularidad y clasificacion, sin volver a consultar.
' Columnas: PID | mes | gics | bics | geo | pais | divisa | activo | valor
Private Function SQLBloqueApil(ByVal ents As String) As String
    Dim gics As String, bics As String, geo As String, pais As String
    Dim divc As String, act As String, kPos As String, kVal As String
    gics = Cfg("SECTOR_COL", SECTOR_COL): bics = "CLASSIFICATION_BICS"
    geo = Cfg("GEO_COL", GEO_COL): pais = Cfg("PAIS_COL", "FCCOUNTRY")
    divc = Cfg("DIV_COL", DIV_COL): act = Cfg("ACTIVO_COL", "INSTRUMENT_TYPE")
    kPos = Cfg("JOIN_KEY_POS", "PK_SECURITY_IK"): kVal = Cfg("JOIN_KEY_VAL", "PK_SECURITY_IK")
    SQLBloqueApil = _
        "WITH base AS (" & vbLf & _
        "  SELECT p.PK_PORTFOLIO_ID, p.PK_FECHA_DATOS, FORMAT_DATE('%Y-%m', p.PK_FECHA_DATOS) AS mes," & vbLf & _
        "         p." & kPos & " AS seckey, p." & PosValor() & " AS valor" & vbLf & _
        "  FROM " & TblPos() & " p" & vbLf & _
        "  WHERE p.PK_PORTFOLIO_ID IN (" & ents & ")" & AndAsOf("p.PK_FECHA_DATOS") & vbLf & _
        "    AND p.PK_FECHA_DATOS > DATE_SUB((SELECT MAX(PK_FECHA_DATOS) FROM " & TblPos() & MaxAsOf() & "), INTERVAL " & CACHE_ANOS & " YEAR))," & vbLf & _
        " ult AS (SELECT PK_PORTFOLIO_ID, mes, MAX(PK_FECHA_DATOS) AS f FROM base GROUP BY PK_PORTFOLIO_ID, mes)" & vbLf & _
        "SELECT b.PK_PORTFOLIO_ID, b.mes," & vbLf & _
        "       v." & gics & " AS gics, v." & bics & " AS bics, v." & geo & " AS geo," & vbLf & _
        "       v." & pais & " AS pais, v." & divc & " AS divisa, v." & act & " AS activo," & vbLf & _
        "       FORMAT('%.10f', CAST(SUM(b.valor) AS FLOAT64)) AS valor" & vbLf & _
        "FROM base b" & vbLf & _
        "JOIN ult ON ult.PK_PORTFOLIO_ID = b.PK_PORTFOLIO_ID AND ult.mes = b.mes AND b.PK_FECHA_DATOS = ult.f" & vbLf & _
        "JOIN " & TblValores() & " v ON v." & kVal & " = b.seckey" & vbLf & _
        "GROUP BY b.PK_PORTFOLIO_ID, b.mes, gics, bics, geo, pais, divisa, activo" & vbLf & _
        "ORDER BY b.PK_PORTFOLIO_ID, b.mes"
End Function

' Columna del bloque APIL para la clasificacion elegida (B8).
Private Function ColClasifApil(ByVal clas As String) As Long
    ' APIL: PID(+0) mes(+1) gics(+2) bics(+3) geo(+4) pais(+5) divisa(+6) activo(+7) valor(+8)
    Select Case clas
        Case "Sector":     ColClasifApil = BLK_APIL_COL + IIf(InStr(UCase(Cfg("SECTOR_COL", SECTOR_COL)), "BICS") > 0, 3, 2)
        Case "Industria":  ColClasifApil = BLK_APIL_COL + IIf(InStr(UCase(Cfg("IND_COL", IND_COL)), "BICS") > 0, 3, 2)
        Case "Continente": ColClasifApil = BLK_APIL_COL + 4
        Case "Pais":       ColClasifApil = BLK_APIL_COL + 5
        Case "Divisa":     ColClasifApil = BLK_APIL_COL + 6
        Case "Activo":     ColClasifApil = BLK_APIL_COL + 7
        Case Else:         ColClasifApil = 0
    End Select
End Function

' Bucket temporal de un mes "YYYY-MM" segun la granularidad (Mensual/Trimestral/
' Semestral/Anual). El bloque es mensual; granularidades finas caen a Mensual.
Private Function BucketMes(ByVal mes As String, ByVal gran As String) As String
    Dim y As String, m As Long
    y = Left(mes, 4): m = CLng(Val(Mid(mes, 6, 2)))
    Select Case Fold(gran)
        Case "trimestral": BucketMes = y & "-T" & (Int((m - 1) / 3) + 1)
        Case "semestral":  BucketMes = y & "-S" & IIf(m <= 6, 1, 2)
        Case "anual":      BucketMes = y
        Case Else:         BucketMes = mes   ' mensual (y granularidades finas)
    End Select
End Function

Private Function FechaDeMes(ByVal mes As String) As Date
    If Len(Trim(mes)) < 7 Then Exit Function
    FechaDeMes = DateSerial(CLng(Val(Left(mes, 4))), CLng(Val(Mid(mes, 6, 2))), 1)
End Function

' Composicion apilada EN LOCAL desde el bloque APIL: trocea por periodo/
' granularidad (B11/B9) y agrupa por la clasificacion (B8), para la Entidad 1.
' Escribe la matriz en D2 y dibuja el grafico apilado. False si no hay datos.
Private Function LocalApiladas(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, i As Long, n As Long, colCat As Long
    c0 = BLK_APIL_COL
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function
    Dim clas As String, gran As String, per As String
    clas = Trim(CStr(ws.Range("B8").Value))
    gran = Trim(CStr(ws.Range("B9").Value))
    If Fold(gran) = "diario" Or Fold(gran) = "semanal" Or Not DimEsTiempo(gran) Then gran = "Mensual"
    colCat = ColClasifApil(clas)
    If colCat = 0 Then Exit Function
    Dim iCat As Long: iCat = colCat - c0 + 1
    per = Trim(CStr(ws.Range("B11").Value))

    Dim blk As Variant
    blk = ws.Range(ws.Cells(2, c0), ws.Cells(lastR, c0 + 8)).Value
    n = lastR - 1

    Dim dMax As Date, d As Date
    For i = 1 To n
        If IgualId(UCase(Trim(CStr(blk(i, 1)))), mId1) Then
            d = FechaDeMes(CStr(blk(i, 2)))
            If d > dMax Then dMax = d
        End If
    Next i
    If dMax = 0 Then Exit Function
    Dim ini As Date: ini = InicioBucket(InicioVentana(dMax, per, gran), gran)

    ' Auxiliares (TODAS las filas): bucket (bmes), etiqueta (lbl) y valoracion.
    ReDim h1(1 To n, 1 To 1) As Variant   ' bmes
    ReDim a1(1 To n, 1 To 1) As Variant   ' lbl
    ReDim a2(1 To n, 1 To 1) As Variant   ' valoracion
    ReDim pidC(1 To n, 1 To 1) As Variant
    Dim bkD As Object: Set bkD = CreateObject("Scripting.Dictionary")   ' buckets a mostrar
    Dim seD As Object: Set seD = CreateObject("Scripting.Dictionary")   ' series (orden aparicion)
    Dim bk As String, se As String
    For i = 1 To n
        h1(i, 1) = BucketMes(CStr(blk(i, 2)), gran)
        a1(i, 1) = EtiquetaClasif(clas, CStr(blk(i, iCat)))
        a2(i, 1) = NumDbl(blk(i, 9))
        pidC(i, 1) = UCase(Trim(CStr(blk(i, 1))))
        If IgualId(UCase(Trim(CStr(blk(i, 1)))), mId1) Then
            d = FechaDeMes(CStr(blk(i, 2)))
            If d >= ini And d <= dMax And d > 0 Then
                bk = CStr(h1(i, 1)): se = CStr(a1(i, 1))
                If Not bkD.Exists(bk) Then bkD.Add bk, 1
                If Len(se) > 0 And Not seD.Exists(se) Then seD.Add se, 1
            End If
        End If
    Next i
    If bkD.Count = 0 Or seD.Count = 0 Then Exit Function

    ws.Range(ws.Cells(2, HLP_K1), ws.Cells(lastR, HLP_K1)).Value = h1
    ws.Range(ws.Cells(2, HLP_A1), ws.Cells(lastR, HLP_A1)).Value = a1
    ws.Range(ws.Cells(2, HLP_A2), ws.Cells(lastR, HLP_A2)).Value = a2
    ws.Range(ws.Cells(2, HLP_PID), ws.Cells(lastR, HLP_PID)).Value = pidC
    FijarNombre "f_pid", HLP_PID, lastR
    FijarNombre "f_k1", HLP_K1, lastR
    FijarNombre "f_a1", HLP_A1, lastR
    FijarNombre "f_a2", HLP_A2, lastR

    ' Buckets en orden cronologico; series en orden de aparicion (con tope MAXSER;
    ' el resto se agrupa en una columna residual "Otros" = 1 - suma del resto).
    Dim bArr As Variant: bArr = ClavesOrdenadas(bkD)
    Dim nb As Long: nb = UBound(bArr) - LBound(bArr) + 1
    Dim sArr() As String, nSer As Long, hayOtros As Boolean, kk As Variant
    Dim tope As Long: tope = seD.Count
    If tope > MAXSER Then tope = MAXSER: hayOtros = True
    ReDim sArr(0 To tope - 1)
    nSer = 0
    For Each kk In seD.Keys
        If nSer < IIf(hayOtros, tope - 1, tope) Then sArr(nSer) = CStr(kk): nSer = nSer + 1
    Next kk
    If hayOtros Then sArr(nSer) = "Otros": nSer = nSer + 1

    EscribirApiladas ws, bArr, nb, sArr, nSer, hayOtros
    LocalApiladas = True
End Function

' Escribe la matriz apilada (fechas x categorias, en %) COMO FORMULAS desde D2 y
' dibuja el grafico apilado. Cada celda = valoracion(categoria)/valoracion(bucket)
' de la Entidad 1. La ultima serie puede ser "Otros" (residual = 1 - resto).
Private Sub EscribirApiladas(ByVal ws As Worksheet, ByVal bArr As Variant, ByVal nb As Long, _
        ByRef sArr() As String, ByVal nSer As Long, ByVal hayOtros As Boolean)
    GuardarPreviewSiNoExiste ws
    Application.EnableEvents = False
    ws.Range(ws.Cells(2, TCMP_COL), ws.Cells(402, TCMP_COL + MAXSER)).ClearContents
    ws.Cells(2, TCMP_COL).Value = "Fecha"
    Dim i As Long, cc As Long, r0 As Long, r1 As Long
    r0 = 3: r1 = 2 + nb
    ' Etiquetas de fecha (columna D).
    Dim dcol() As Variant: ReDim dcol(1 To nb, 1 To 1)
    For i = 1 To nb: dcol(i, 1) = bArr(i - 1): Next i
    ws.Range(ws.Cells(r0, TCMP_COL), ws.Cells(r1, TCMP_COL)).Value = dcol
    ' Cabeceras de serie + formulas por columna.
    Dim hdr As String, f As String, ultCol As Long
    Dim qid As String: qid = Q(UCase(Trim(mId1)))
    For i = 0 To nSer - 1
        cc = TCMP_COL + 1 + i
        ws.Cells(2, cc).Value = sArr(i)
        If hayOtros And i = nSer - 1 Then
            ' Residual: 1 - suma de las series explicitas de esa fila.
            f = "=IFERROR(1-SUM(" & ws.Cells(r0, TCMP_COL + 1).Address(False, True) & ":" & _
                ws.Cells(r0, cc - 1).Address(False, True) & ")," & Q("") & ")"
        Else
            hdr = ws.Cells(2, cc).Address(True, True)   ' cabecera de ESTA serie (categoria)
            f = "=IFERROR(SUMIFS(f_a2,f_pid," & qid & ",f_k1,$D" & r0 & ",f_a1," & hdr & ")/" & _
                "SUMIFS(f_a2,f_pid," & qid & ",f_k1,$D" & r0 & ")," & Q("") & ")"
        End If
        ws.Range(ws.Cells(r0, cc), ws.Cells(r1, cc)).Formula = f
        ultCol = cc
    Next i
    ws.Range(ws.Cells(r0, TCMP_COL + 1), ws.Cells(r1, ultCol)).NumberFormat = "0.00%"
    Application.EnableEvents = True
    DibujarApiladas ws, nb, nSer
End Sub

' Dibuja columnas/barras apiladas: una serie por cada columna de categoria
' (BB..), eje X = fechas (BA). 100% apiladas si B12 (tipo de grafico) lo indica.
Private Sub DibujarApiladas(ByVal ws As Worksheet, ByVal nRows As Long, ByVal nSer As Long)
    Dim ch As Chart, s As Series, c As Long, r0 As Long, r1 As Long
    On Error Resume Next
    Set ch = ws.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub
    r0 = 3: r1 = 2 + nRows
    Do While ch.SeriesCollection.Count > 0: ch.SeriesCollection(1).Delete: Loop
    Dim q As String: q = QHoja(ws)
    For c = TCMP_COL + 1 To TCMP_COL + nSer
        Set s = ch.SeriesCollection.NewSeries
        s.Name = q & ws.Cells(2, c).Address
        s.Values = q & ws.Range(ws.Cells(r0, c), ws.Cells(r1, c)).Address
        s.XValues = q & ws.Range(ws.Cells(r0, TCMP_COL), ws.Cells(r1, TCMP_COL)).Address
    Next c
    Dim t As String: t = Fold(ws.Range("B12").Value)
    If InStr(t, "100%") > 0 Then
        ch.ChartType = IIf(InStr(t, "barra") > 0, xlBarStacked100, xlColumnStacked100)
    Else
        ch.ChartType = IIf(InStr(t, "barra") > 0, xlBarStacked, xlColumnStacked)
    End If
    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = ws.Range("A19").Value & " - evolucion"
    ch.HasLegend = True
    ' Ejes: % en el de valores, granularidad en el de categorias (borra el titulo
    ' antiguo, p.ej. "Duracion Modificada", que quedaba de un dibujo anterior).
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "Peso (%)"
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = ws.Range("B9").Value
    On Error GoTo 0
End Sub

' Spread medio (ponderado por valoracion) EN LOCAL desde el bloque POS. Si la
' dimension es una clasificacion, desglosa por ella; si no, un valor por slot.
Private Function LocalSpread(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, r As Long, colCat As Long
    c0 = BLK_POS_COL
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function
    colCat = ColClasifPos(Trim(CStr(ws.Range("B9").Value)))   ' 0 = sin desglose

    Dim cats As Object, num As Object, den As Object
    Set cats = CreateObject("Scripting.Dictionary")
    Set num = CreateObject("Scripting.Dictionary")
    Set den = CreateObject("Scripting.Dictionary")
    Dim rowOut As Long: rowOut = 3
    Dim pid As String, slot As Long, cat As String, k As String, w As Double, sp As Double
    For r = 2 To lastR
        pid = UCase(Trim(CStr(ws.Cells(r, c0).Value)))
        slot = SlotDe(pid)
        If slot = 0 Then GoTo seguir
        If colCat = 0 Then
            cat = "Total"
        Else
            cat = Trim(CStr(ws.Cells(r, colCat).Value))
            If Len(cat) = 0 Then cat = "(sin dato)"
        End If
        If Not cats.Exists(cat) Then
            If rowOut > 402 Then GoTo seguir
            cats.Add cat, rowOut: rowOut = rowOut + 1
        End If
        w = NumDbl(ws.Cells(r, c0 + 6).Value)    ' valoracion (peso)
        sp = NumDbl(ws.Cells(r, c0 + 7).Value)   ' spread
        k = slot & "|" & cat
        If num.Exists(k) Then num(k) = num(k) + sp * w Else num(k) = sp * w
        If den.Exists(k) Then den(k) = den(k) + w Else den(k) = w
seguir:
    Next r
    If cats.Count = 0 Then Exit Function
    Dim vals As Object: Set vals = CreateObject("Scripting.Dictionary")
    Dim kk As Variant
    For Each kk In num.Keys
        If den(kk) <> 0 Then vals(kk) = num(kk) / den(kk)
    Next kk
    VolcarLocal ws, cats, vals
    LocalSpread = True
End Function

' TER look-through (ponderado por valoracion) EN LOCAL desde el bloque POS.
Private Function LocalTerLT(ByVal ws As Worksheet) As Boolean
    Dim c0 As Long, lastR As Long, r As Long
    c0 = BLK_POS_COL
    lastR = UltFilaBloque(ws, c0)
    If lastR < 2 Then Exit Function

    Dim cats As Object, num As Object, den As Object
    Set cats = CreateObject("Scripting.Dictionary")
    Set num = CreateObject("Scripting.Dictionary")
    Set den = CreateObject("Scripting.Dictionary")
    cats.Add "Total", 3
    Dim pid As String, slot As Long, k As String, w As Double, te As Double
    For r = 2 To lastR
        pid = UCase(Trim(CStr(ws.Cells(r, c0).Value)))
        slot = SlotDe(pid)
        If slot = 0 Then GoTo seguir
        w = NumDbl(ws.Cells(r, c0 + 6).Value)    ' valoracion
        te = NumDbl(ws.Cells(r, c0 + 8).Value)   ' ter
        k = slot & "|Total"
        If num.Exists(k) Then num(k) = num(k) + te * w Else num(k) = te * w
        If den.Exists(k) Then den(k) = den(k) + w Else den(k) = w
seguir:
    Next r
    Dim vals As Object: Set vals = CreateObject("Scripting.Dictionary")
    Dim kk As Variant, hay As Boolean
    For Each kk In num.Keys
        If den(kk) <> 0 Then vals(kk) = num(kk) / den(kk): hay = True
    Next kk
    If Not hay Then Exit Function
    VolcarLocal ws, cats, vals
    LocalTerLT = True
End Function

' Enruta la metrica actual a su calculo LOCAL desde los bloques amplios.
' Devuelve True si la resolvio en local (y ya dibujo); False para que
' RefrescarDatos caiga a la consulta directa de BigQuery.
Private Function ResolverLocal(ByVal ws As Worksheet) As Boolean
    Dim met As String: met = Trim(CStr(ws.Range("B8").Value))
    If (met = "Rentabilidad" Or met = "Rentab. acum.") Then
        ' Los periodos de calendario cerrado ("... anterior") se resuelven SIEMPRE por
        ' SQL: la ventana anclada de VentanaFechas devuelve EXACTAMENTE el periodo
        ' anterior, mientras que el recalculo local compone por bucket completo y, con
        ' dimensiones mas gruesas que el periodo, abarcaria de mas. -> ResolverLocal=False.
        If PerAnterior(CStr(ws.Range("B11").Value)) <> "" Then Exit Function
        ' Benchmark = otro indice: la serie del indice se calcula en SQL (el bloque
        ' local solo cachea el benchmark propio de la cartera). -> ResolverLocal=False.
        If Len(BmkIndiceId(ws)) > 0 Then Exit Function
        ResolverLocal = LocalRentabilidad(ws)
    ElseIf InStr(Fold(met), "duraci") = 1 Or met = "TIR" _
           Or Fold(met) = "cmr" Or Left(Fold(met), 3) = "var" Then
        ResolverLocal = LocalRiesgo(ws)
    ElseIf met = "Peso" Or met = "Importe" Then
        ResolverLocal = LocalComposicion(ws)      ' composicion (% o valor absoluto)
    ' Spread / Volatilidad: no salen de estos bloques; van por consulta directa
    ' (ResolverLocal = False).
    End If
End Function

' Ejecuta la SQL de una metrica concreta (via directa, sin bloque amplio),
' vuelca el crudo a la hoja oculta _Volcado y pivota a D:H. La columna W de la
' hoja Panel queda LIBRE para los bloques amplios (no se toca aqui).
' Devuelve True si fue bien; si no, deja el mensaje de error en msg.
Private Function EjecutarYVolcar(ByVal ws As Worksheet, ByVal sql As String, ByRef msg As String) As Boolean
    Dim cn As Object, rs As Object, j As Long, wv As Worksheet
    On Error GoTo fallo
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 120
    cn.CursorLocation = 3          ' adUseClient: compatible con drivers ODBC de solo lectura (BigQuery)
    cn.Open CfgConn()
    Set rs = cn.Execute(sql)       ' recordset de solo avance (el driver si lo admite)

    Application.EnableEvents = False
    Set wv = HojaAux("_Volcado")   ' hoja oculta de trabajo (crudo de la consulta directa)
    wv.Cells.Clear
    ' Todo como TEXTO: los valores llegan formateados con '.' decimal; asi Excel
    ' no los auto-convierte (ni reescala) y VolcarResultado los parsea con NumVal.
    wv.Cells.NumberFormat = "@"
    For j = 0 To rs.Fields.Count - 1
        wv.Cells(1, 1 + j).Value = rs.Fields(j).Name
    Next j
    If Not rs.EOF Then wv.Cells(2, 1).CopyFromRecordset rs
    rs.Close: cn.Close
    VolcarResultado ws, wv
    Application.EnableEvents = True
    EjecutarYVolcar = True
    Exit Function
fallo:
    msg = Err.Description
    Application.EnableEvents = True
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    On Error GoTo 0
    EjecutarYVolcar = False
End Function

' Diagnostico: lista los nombres de columna de las tablas clave (rendimiento,
' POSICIONES y MAESTRO de valores) en la hoja visible '_Columnas', una tabla por
' columna. Sirve para ver los nombres reales (p.ej. la clave de union
' posiciones->maestro para la composicion, o las columnas de benchmark).
Public Sub VerColumnas()
    Dim wc As Worksheet
    Set wc = HojaAux("_Columnas")
    wc.Visible = xlSheetVisible
    wc.Cells.ClearContents
    Dim col As Long, nf As Long
    col = 1
    nf = nf + VolcarColumnasTabla(wc, col, T_PERF, Tbl(DS_PROD, T_PERF))
    col = col + 1
    nf = nf + VolcarColumnasTabla(wc, col, T_POS, TblPos())
    col = col + 1
    nf = nf + VolcarColumnasTabla(wc, col, T_VALORES, Tbl(DS_MERC, T_VALORES))
    wc.Activate
    MsgBox "Columnas volcadas en la hoja '_Columnas' (una tabla por columna):" & vbLf & _
           "  A = " & T_PERF & " (rendimiento/benchmark)" & vbLf & _
           "  B = " & T_POS & " (posiciones)" & vbLf & _
           "  C = " & T_VALORES & " (maestro de valores)" & vbLf & vbLf & _
           "Para la COMPOSICION: busca en A y B una columna comun para unir " & _
           "posiciones y maestro (p.ej. algun *SECURITY* o *ACTIVO*), y ponla en " & _
           "la hoja 'config' (JOIN_KEY_POS y JOIN_KEY_VAL).", vbInformation, "Columnas"
End Sub

' Vuelca en la columna 'col' de wc los nombres de campo de una tabla. Devuelve
' el numero de campos (0 si fallo). No corta el resto del diagnostico si una
' tabla concreta falla.
Private Function VolcarColumnasTabla(ByVal wc As Worksheet, ByVal col As Long, _
        ByVal titulo As String, ByVal tblRef As String) As Long
    Dim cn As Object, rs As Object, j As Long
    On Error GoTo fallo
    Set cn = CreateObject("ADODB.Connection")
    cn.CommandTimeout = 60
    cn.CursorLocation = 3
    cn.Open CfgConn()
    Set rs = cn.Execute("SELECT * FROM " & tblRef & " LIMIT 1")
    wc.Cells(1, col).Value = titulo & " (" & rs.Fields.Count & ")"
    For j = 0 To rs.Fields.Count - 1
        wc.Cells(j + 2, col).Value = rs.Fields(j).Name
    Next j
    VolcarColumnasTabla = rs.Fields.Count
    rs.Close: cn.Close
    Exit Function
fallo:
    wc.Cells(1, col).Value = titulo & " (ERROR)"
    wc.Cells(2, col).Value = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Function

' Guarda una copia de las formulas dummy (D3:H402) en la hoja oculta _Prev,
' pero SOLO la primera vez (cuando aun estan las formulas de previsualizacion).
Private Sub GuardarPreviewSiNoExiste(ByVal ws As Worksheet)
    Dim wp As Worksheet
    Set wp = HojaAux("_Prev")
    If Trim(CStr(wp.Range("A1").Value)) = "GUARDADO" Then Exit Sub
    wp.Cells.Clear
    wp.Range("A1").Value = "GUARDADO"
    ws.Range("D3:H402").Copy
    wp.Range("D3").PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Sub

' Restaura la vista previa (datos dummy) sobre la tabla del grafico y redibuja.
' Util para volver a jugar con parametros/tipo de grafico tras un Actualizar.
Public Sub VistaPreviaDummy()
    Dim ws As Worksheet, wp As Worksheet
    Set ws = Panel()
    Set wp = HojaAux("_Prev")
    If Trim(CStr(wp.Range("A1").Value)) <> "GUARDADO" Then
        MsgBox "Todavia no hay vista previa guardada." & vbLf & _
               "Se guarda automaticamente la primera vez que pulsas Actualizar.", _
               vbInformation, "Vista previa"
        Exit Sub
    End If
    Application.EnableEvents = False
    LimpiarTablaNormal ws
    ws.Range("E3:H402").NumberFormat = "0.00"    ' los datos dummy no vienen en fraccion
    wp.Range("D3:H402").Copy
    ws.Range("D3").PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    Application.EnableEvents = True
    DibujarGrafico
End Sub

' Pivota el resultado crudo (hoja 'src', desde la columna 1) a la tabla del
' grafico D:H de la hoja Panel (ws).
Private Sub VolcarResultado(ByVal ws As Worksheet, ByVal src As Worksheet)
    Dim c As Long, hdr As String
    Dim colPort As Long, colCat As Long, colVal As Long, colBmk As Long
    Dim lastData As Long, r As Long, rowOut As Long, rr As Long
    Dim id1 As String, id2 As String, id3 As String, cat As String, pid As String
    Dim cats As Object

    c = 1
    Do While Trim(CStr(src.Cells(1, c).Value)) <> "" And c < 40
        hdr = LCase(Trim(CStr(src.Cells(1, c).Value)))
        Select Case hdr
            Case "pk_portfolio_id": colPort = c
            Case "categoria":       colCat = c
            Case "valor":           colVal = c
            Case "valor_bmk":       colBmk = c
        End Select
        c = c + 1
    Loop
    If colVal = 0 Then Exit Sub

    c = colPort: If c = 0 Then c = 1
    r = 2: lastData = 1
    Do While Trim(CStr(src.Cells(r, c).Value)) <> "" And r < 200000
        lastData = r: r = r + 1
    Loop

    ' Mapa robusto pid -> columna destino (5=E/6=F/7=G). Primero empareja cada
    ' id del resultado con el slot cuyo id resuelto coincide (IgualId); los ids
    ' del resultado que no casen se asignan POSICIONALMENTE a los slots activos
    ' libres (la query ya filtro exactamente esas carteras, asi que no dependemos
    ' de que la traduccion nombre->id sea identica byte a byte).
    Dim mapPid As Object, dist As Object, ky As Variant, kk As Long
    Dim idsv(1 To 3) As String, colv(1 To 3) As Long, usado(1 To 3) As Boolean
    idsv(1) = mId1: idsv(2) = mId2: idsv(3) = mId3
    colv(1) = 5: colv(2) = 6: colv(3) = 7
    Set mapPid = CreateObject("Scripting.Dictionary")
    Set dist = CreateObject("Scripting.Dictionary")
    For r = 2 To lastData
        pid = UCase(Trim(CStr(src.Cells(r, colPort).Value)))
        If pid <> "" And Not dist.Exists(pid) Then dist.Add pid, 0
    Next r
    For Each ky In dist.Keys                       ' 1) match por id
        For kk = 1 To 3
            If Not usado(kk) And Not mapPid.Exists(ky) And IgualId(CStr(ky), idsv(kk)) Then
                mapPid(ky) = colv(kk): usado(kk) = True
            End If
        Next kk
    Next ky
    For Each ky In dist.Keys                       ' 2) fallback posicional
        If Not mapPid.Exists(ky) Then
            For kk = 1 To 3
                If Not usado(kk) And Len(Trim(idsv(kk))) > 0 Then
                    mapPid(ky) = colv(kk): usado(kk) = True: Exit For
                End If
            Next kk
        End If
    Next ky

    ' Guarda las formulas dummy la PRIMERA vez, para poder volver a la vista
    ' previa despues (boton "Vista previa"). Luego ya sobrescribimos con lo real.
    GuardarPreviewSiNoExiste ws
    LimpiarTablaNormal ws

    Dim tc As Long
    If colCat > 0 Then
        Set cats = CreateObject("Scripting.Dictionary")
        rowOut = 3
        For r = 2 To lastData
            cat = Trim(CStr(src.Cells(r, colCat).Value))
            If cat <> "" And Not cats.Exists(cat) Then
                cats.Add cat, rowOut
                ws.Cells(rowOut, 4).Value = cat
                rowOut = rowOut + 1
                If rowOut > 402 Then Exit For
            End If
        Next r
        For r = 2 To lastData
            pid = UCase(Trim(CStr(src.Cells(r, colPort).Value)))
            cat = Trim(CStr(src.Cells(r, colCat).Value))
            If cats.Exists(cat) And mapPid.Exists(pid) Then
                rr = cats(cat): tc = mapPid(pid)
                ws.Cells(rr, tc).Value = NumVal(src.Cells(r, colVal).Value)
                If tc = 5 And colBmk > 0 Then ws.Cells(rr, 8).Value = NumVal(src.Cells(r, colBmk).Value)
            End If
        Next r
    Else
        ws.Cells(3, 4).Value = Trim(CStr(ws.Range("B8").Value)) & " - " & Trim(CStr(ws.Range("B11").Value))
        For r = 2 To lastData
            pid = UCase(Trim(CStr(src.Cells(r, colPort).Value)))
            If mapPid.Exists(pid) Then
                tc = mapPid(pid)
                ws.Cells(3, tc).Value = NumVal(src.Cells(r, colVal).Value)
                If tc = 5 And colBmk > 0 Then ws.Cells(3, 8).Value = NumVal(src.Cells(r, colBmk).Value)
            End If
        Next r
    End If

    ' Cabecera del benchmark solo si el resultado trajo columna valor_bmk.
    If colBmk > 0 Then ws.Range("H2").Value = EtiquetaBenchmark(ws)

    ' Formato adecuado a la metrica (% para rendimientos, numero para el resto).
    ws.Range("E3:H402").NumberFormat = FormatoMetrica(Trim(CStr(ws.Range("B8").Value)))
End Sub

' Compara dos ids de portfolio de forma robusta (sin mayus/espacios). Devuelve
' False si el id de referencia esta vacio (para no cuadrar slots no usados).
Private Function IgualId(ByVal a As String, ByVal b As String) As Boolean
    IgualId = (Len(Trim(b)) > 0) And (UCase(Trim(a)) = UCase(Trim(b)))
End Function

' True si la columna (E/F/G/H) tiene algun valor en las filas de datos (3..402).
Private Function HayDatosCol(ByVal ws As Worksheet, ByVal colLetter As String) As Boolean
    Dim r As Long
    For r = 3 To 402
        If Trim(CStr(ws.Range(colLetter & r).Value)) <> "" Then HayDatosCol = True: Exit Function
    Next r
End Function

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

' Referencia a un rango con nombre del propio libro (para las series del grafico).
Private Function RefNombre(ByVal nm As String) As String
    RefNombre = "='" & ThisWorkbook.Name & "'!" & nm
End Function

' Define (o redefine) un rango con nombre dinamico a nivel de libro.
Private Sub DefNombre(ByVal nm As String, ByVal ref As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=ref
End Sub

' Crea rangos con nombre que se AJUSTAN SOLOS al numero de categorias visibles
' (cuenta las celdas no vacias de D3:D402). Asi el grafico no deja huecos al
' cambiar de periodo/dimension: la altura del rango sigue a los datos.
Private Sub AsegurarNombresGrafico()
    Dim cnt As String
    cnt = "SUMPRODUCT(--(Panel!$D$3:$D$402<>""""))"
    DefNombre "ChCats", "=OFFSET(Panel!$D$3,0,0,MAX(1," & cnt & "),1)"
    DefNombre "ChE", "=OFFSET(Panel!$E$3,0,0,MAX(1," & cnt & "),1)"
    DefNombre "ChF", "=OFFSET(Panel!$F$3,0,0,MAX(1," & cnt & "),1)"
    DefNombre "ChG", "=OFFSET(Panel!$G$3,0,0,MAX(1," & cnt & "),1)"
    DefNombre "ChH", "=OFFSET(Panel!$H$3,0,0,MAX(1," & cnt & "),1)"
End Sub

Private Sub AddEnt(ws As Worksheet, ByVal ch As Chart, ByVal slotCell As String, _
        ByVal colLetter As String, ByVal lastRow As Long)
    Dim s As Series, v As String
    v = Trim(CStr(ws.Range(slotCell).Value))
    If v = "" Or v = "(ninguna)" Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    Dim q As String: q = QHoja(ws)         ' referencia DIRECTA a la hoja (no a nombres del libro)
    s.Name = q & "$" & colLetter & "$2"
    s.Values = q & "$" & colLetter & "$3:$" & colLetter & "$" & lastRow
    s.XValues = q & "$D$3:$D$" & lastRow
End Sub

' Redibuja el grafico del Panel con la tabla D:H actual (mock o datos reales).
Public Sub DibujarGrafico()
    Dim ws As Worksheet, ch As Chart, s As Series, conBench As Boolean, tipo As String
    Dim lastRow As Long, benchIdx As Long
    Set ws = Panel()
    ' NO tocar B4: el usuario elige su cartera real (no la lista de ejemplo). Antes
    ' AjustarSeleccion reseteaba B4 a la demo si no estaba en Ent_<tipo>; eliminado.

    On Error Resume Next
    Set ch = ws.ChartObjects(1).Chart
    On Error GoTo 0
    If ch Is Nothing Then Exit Sub

    conBench = ConBenchmark(ws)
    tipo = Fold(ws.Range("B12").Value)

    ' Ultima fila con categoria (col D) para acotar las series a la hoja ACTUAL
    ' (asi el grafico de una copia lee de SU hoja, no del Panel original).
    lastRow = 2
    Dim rr As Long
    For rr = 3 To 402
        If Len(Trim(CStr(ws.Cells(rr, TCMP_COL).Value))) > 0 Then lastRow = rr
    Next rr
    If lastRow < 3 Then lastRow = 3

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    AddEnt ws, ch, "B4", "E", lastRow
    AddEnt ws, ch, "B5", "F", lastRow
    AddEnt ws, ch, "B6", "G", lastRow

    ' Solo dibuja la serie de benchmark si REALMENTE hay datos en H (columna del
    ' benchmark). Metricas sin benchmark (Duracion/TIR/Composicion) dejan H vacia,
    ' asi que no se pinta una linea de ceros pegada abajo.
    benchIdx = 0
    If conBench And Trim(CStr(ws.Range("B4").Value)) <> "" And HayDatosCol(ws, "H") Then
        Set s = ch.SeriesCollection.NewSeries
        Dim qb As String: qb = QHoja(ws)
        s.Name = qb & "$H$2"
        s.Values = qb & "$H$3:$H$" & lastRow
        s.XValues = qb & "$D$3:$D$" & lastRow
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
        Case "barras apiladas":   ch.ChartType = xlBarStacked
        Case Else:                ch.ChartType = xlColumnClustered
    End Select

    If benchIdx > 0 Then
        On Error Resume Next
        Select Case Fold(ws.Range("B15").Value)
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
        ch.Axes(xlValue).TickLabels.NumberFormat = FormatoMetrica(Trim(CStr(ws.Range("B8").Value)))
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
