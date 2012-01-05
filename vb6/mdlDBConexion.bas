Attribute VB_Name = "mdlDBConexion"
'*******************************************************************************
' la base de datos tiene por defecto las tablas AUTHOR, FILE_TYPE, GENRE, PARENT
' STORAGE, STORAGE_TYPE y SUB_GENRE con el indice 0 asignado a cadena vacia
' toda insercion en dichas tablas comienza de 1.
'*******************************************************************************
Option Explicit
Option Base 1

'*******************************************************************************
' Constantes DB
'*******************************************************************************
Public Const DB_MAX_LEN_FILE_TYPE = 15
Public Const DB_MAX_LEN_GENRE = 30
Public Const DB_MAX_LEN_AUTHOR = 50
Public Const DB_MAX_LEN_FILE_NAME = 160
Public Const DB_MAX_LEN_FILE_SYS_NAME = 255
Public Const DB_MAX_LEN_FILE_PARENT = 50
Public Const DB_MAX_LEN_FILE_TYPE_QUALITY = 8
Public Const DB_MAX_LEN_FILE_COMMENT = 50
Public Const DB_MAX_LEN_FILE_TYPE_COMMENT = 20
Public Const DB_MAX_LEN_STORAGE_NAME = 25
Public Const DB_MAX_LEN_STORAGE_LABEL = 20
Public Const DB_MAX_LEN_STORAGE_SERIAL = 10
Public Const DB_MAX_LEN_STORAGE_TYPE_LEN = 12
Public Const DB_MAX_LEN_STORAGE_COMMENT = 50

Public Const DB_EDIT_SLEEP = 0.5

Public Enum DB_SearchStyle
   db_Todos = 0
   db_Con = 1
   db_ConPalabra = 2
   db_QueComience = 3
   db_QueTermine = 4
End Enum

Public Enum DB_CompareStyle
   db_Menor = 0
   db_MenorIgual = 1
   db_Igual = 2
   db_MayorIgual = 3
   db_Mayor = 4
End Enum

Public Const DB_SQL_STR_AND = " && "
Public Const DB_SQL_STR_OR = " || "

'**************************************************
' variables globales para acceso a datos
'**************************************************
Public cn As ADODB.Connection
Public cnTransaction As ADODB.Connection
Public query As String

Public gs_DSN As String
Public gs_Pwd As String

'**************************************************
' opciones de busqueda de la DB
'**************************************************
Public gb_DBConexionOK As Boolean
Public gb_DBFormatOK As Boolean

Public gb_DBPertenciaPorAlmacenamiento As Boolean

Public gb_DBCampoAuxiliarPorGenero As Boolean

Public gs_DBOrdenarCampo(5) As String
Public gb_DBOrdenarAsc(5) As Boolean
Public gb_DBOrdenarEnabled(5) As Boolean

Public gb_DBNameSearchStyleActive As Boolean
Public gt_DBNameSearchStyle As DB_SearchStyle

Public gb_DBAuthorSearchStyleActive As Boolean
Public gt_DBAuthorSearchStyle As DB_SearchStyle

Public gb_DBConPrioridad As Boolean
Public gs_DBConPrioridadDe As String
Public gb_DBPrioritySearchStyleActive As Boolean
Public gt_DBPrioritySearchStyle As DB_CompareStyle

Public gb_DBConCalidad As Boolean
Public gs_DBConCalidadDe As String
Public gb_DBQualitySearchStyleActive As Boolean
Public gt_DBQualitySearchStyle As DB_CompareStyle

Public gb_DBConTipo As Boolean
Public gl_DBConTipoDe As Long
Public gl_DBFileTypeIndex As Long

Public gb_DBConFecha As Boolean
Public gd_DBConFechaDe As Date
Public gb_DBDateSearchStyleActive As Boolean
Public gt_DBDateSearchStyle As DB_CompareStyle

Public gb_DBConTamanyo As Boolean
Public gs_DBConTamanyoDe As String
Public gb_DBFileSizeSearchStyleActive As Boolean
Public gt_DBFileSizeSearchStyle As DB_CompareStyle

Public gs_DBConCampoDe As String
Public gn_DBFieldIndex As Integer

Public gb_DBNameFromFile As Boolean
Public gb_DBShowPathInFileName As Boolean

Public gb_DBShowHiddenFiles As Boolean

'**************************************************
' opciones de edicion de la DB
'**************************************************
Public gb_DBAddRegister As Boolean
Public gl_DB_IDRegistroModificar As Long

Public gb_AddRegisterStarted As Boolean

Public gl_DBRegIDNew As Long
Public gs_DBRegName As String
Public gs_DBRegAuthor As String
Public gs_DBRegStorage As String
Public gs_DBRegParent As String
Public gs_DBRegGenre As String
Public gs_DBRegPriority As String
Public gs_DBRegQuality As String
Public gs_DBRegFileSize As String
Public gd_DBRegFileDate As Date
Public gs_DBRegFileType As String

'**************************************************
' API ODBC
'**************************************************
Declare Function SQLAllocEnv Lib "ODBC32.DLL" (env As Long) As Integer
Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Long, ByVal fdir As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDesc As String, ByVal cbDescMax As Integer, pcbDesc As Integer) As Integer
Declare Function SQLManageDataSources Lib "ODBCCP32.DLL" (ByVal hWnd As Long) As Long
Declare Function SQLCreateDataSource Lib "ODBCCP32.DLL" (ByVal hWnd As Long, ByVal lpszDS As String) As Long


Public Const SQL_SUCCESS As Long = 0
Public Const SQL_FETCH_NEXT As Long = 1
Public Const SQL_FETCH_FIRST As Long = 2
Public Const SQL_FETCH_FIRST_USER As Long = 31
Public Const SQL_FETCH_FIRST_SYSTEM As Long = 32    'don't easy to guess!


'**************************************************
' FUNCIONES DE CONEXION BD
'**************************************************
Public Function gfnc_CrearConexion(sDSN As String, sPwd As String) As Boolean
    
    On Error GoTo Handler

    gfnc_CrearConexion = False
    
    Set cn = New ADODB.Connection
    cn.Open "DSN=" & Trim(sDSN) & ";pwd=" & Trim(sPwd)

    gfnc_CrearConexion = True
    
    Exit Function

Handler:
    Select Case Err.Number
        Case -2147467259
            gfnc_CrearConexion = False
        Case Else
            MsgBox Err.Description, vbCritical, "gfnc_CrearConexion()"
    End Select
End Function

Public Sub gsub_CerrarConexionBaseDatos()
    On Error Resume Next
    cn.Close
    Set cn = Nothing
End Sub

Public Sub gsub_ShowMessageFailedConection()
    MsgBox "No se pudo entablar conexión con el DSN." & vbCrLf & "Verifique el DSN y la contraseña. También es" & vbCrLf & "posible que la BD ya se encuentre abierta en" & vbCrLf & "modo exclusivo por otro usuario o que tenga" & vbCrLf & "atributos de sólo lectura.", vbExclamation, "Error estableciendo conexión"
End Sub

Public Sub gsub_ShowMessageNoConection()
    MsgBox "No está creada la conexión con la base de datos" & vbCrLf & "Para hacerlo use el menú:" & vbCrLf & "[Datos] -> [Conectar a la BD]." & vbCrLf & "De no estar creado el DSN de su base de datos" & vbCrLf & "deberá crearlo primero usando el menú:" & vbCrLf & "[Datos] -> [Crear DSN]." & vbCrLf & "Consulte la ayuda acerca de la instalación", vbExclamation, "Conexión no establecida"
End Sub

Public Sub gsub_ShowMessageWrongDB()
    MsgBox "Se ha establecido conexión con el DSN, pero" & vbCrLf & "la BD no corresponde con el formato esperado." & vbCrLf & "También es posible que se encuentre dañada." & vbCrLf & "Sustituya la BD por su copia de seguridad.", vbExclamation, "Error verificando base de datos"
End Sub

Public Function gfnc_CrearConexionTransaccion(sDSN As String, sPwd As String) As Boolean
    
    On Error GoTo Handler

    gfnc_CrearConexionTransaccion = False
    
    Set cnTransaction = New ADODB.Connection
    cnTransaction.Open "DSN=" & Trim(sDSN) & ";pwd=" & Trim(sPwd)

    gfnc_CrearConexionTransaccion = True
    
    Exit Function

Handler:
    Select Case Err.Number
        Case -2147467259
            'MsgBox "No se pudo establecer conexion con el servidor o no coincide la contraseña", vbExclamation, "Error"
            gfnc_CrearConexionTransaccion = False
        Case Else
            MsgBox Err.Description, vbCritical, "gfnc_CrearConexionTransaccion()"
    End Select
End Function

Public Sub gsub_CerrarConexionTransaccion()
    On Error Resume Next
    cnTransaction.Close
    Set cnTransaction = Nothing
End Sub

Public Function gfnc_ParseString(ByVal s_in As String, ByRef s_out As String) As Boolean

    Dim sz As String
    Dim st As String
    Dim n As Long
    On Error GoTo Handler
    
    sz = s_in
    st = ""
    '-------------------------------------
    ' buscar apostrofe
    Do
        n = InStrRev(sz, "'")
        If n = 0 Then
            st = sz & st
            Exit Do
        Else
            ' duplicar apostrofe
            st = "'" & Mid(sz, n) & st
            
            sz = Mid(sz, 1, n - 1)
        End If
    Loop
    
    s_out = st
    gfnc_ParseString = True
    Exit Function
    
Handler:
    s_out = s_in
    gfnc_ParseString = False
End Function

Public Function gfnc_ValidateDB() As Boolean

    Dim rs As ADODB.Recordset
    Dim clx As clsCrypto
    Dim cad As String
    On Error GoTo Handler
    
    gfnc_ValidateDB = False
    query = "SELECT autor FROM r_porter"
            
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Set clx = New clsCrypto
    clx.SetCod 2004
    '---------------------------------------------------
    'cad = clx.Encrypt("Laurens Esaú Rodriguez Oscanoa")
    '---------------------------------------------------
    cad = clx.Decrypt("bL.ulU}~*}Lú~3]du-|.lV~z}\LU]L")
    Set clx = Nothing

    If cad = rs!autor Then
        gfnc_ValidateDB = True
    End If
    rs.Close
    
    Exit Function
    
Handler:
    'Error verificando BD
    gfnc_ValidateDB = False
End Function

Public Function gfnc_GetLogicalParts(ByRef strSQL As String, ByRef zsSqlParts() As String, ByRef zsSqlOperators() As String) As Boolean
    
    Dim n As Long
    Dim p As Long
    On Error GoTo Handler
    
    gfnc_GetLogicalParts = False
    
    If (strSQL = "") Then Exit Function
    
    n = 0   ' posicion donde se encontro cadena de operador
    p = 1   ' posicion desde la cual buscar cadena
    
    '-------------------------------------
    ' buscar separadores
    Do
        n = InStr(p, strSQL, DB_SQL_STR_AND)
        
        If n <= 0 Then
            If gfnc_GetLogicalParts Then
                ' se encontro elementos: agregar parte que falta
                ReDim Preserve zsSqlParts(UBound(zsSqlParts) + 1)
                zsSqlParts(UBound(zsSqlParts)) = Mid(strSQL, p)
            End If
            Exit Do
        Else
            gfnc_GetLogicalParts = True
            
            ' separar cadenas
            ReDim Preserve zsSqlParts(UBound(zsSqlParts) + 1)
            ReDim Preserve zsSqlOperators(UBound(zsSqlOperators) + 1)
            
JMP_ARRAYS_INITIALIZED:
            
            zsSqlParts(UBound(zsSqlParts)) = Mid(strSQL, p, n - p)
            zsSqlOperators(UBound(zsSqlOperators)) = "AND"
            
            p = n + Len(DB_SQL_STR_AND)
            
        End If
    Loop
    
    Exit Function
    
Handler:
    If Err.Number = 9 Then
        ' El subíndice está fuera del intervalo (cuando llega el array vacio)
        ' error lanzado por la funcion UBound()
        ReDim Preserve zsSqlParts(1)
        ReDim Preserve zsSqlOperators(1)
        Resume JMP_ARRAYS_INITIALIZED
    Else
        MsgBox Err.Description, vbCritical, "gfnc_GetLogicalParts"
    End If
End Function

Public Function gfnc_getTypeDataAdoRecordset(ByRef nType As ADODB.DataTypeEnum, ByRef isText As Boolean, ByRef isDate As Boolean) As String
    Select Case nType
        Case ADODB.DataTypeEnum.adArray
            gfnc_getTypeDataAdoRecordset = "adArray"
            
        Case ADODB.DataTypeEnum.adBigInt
            gfnc_getTypeDataAdoRecordset = "adBigInt"
            
        Case ADODB.DataTypeEnum.adBinary
            gfnc_getTypeDataAdoRecordset = "adBinary"
            
        Case ADODB.DataTypeEnum.adBoolean
            gfnc_getTypeDataAdoRecordset = "adBoolean"
            
        Case ADODB.DataTypeEnum.adBSTR
            gfnc_getTypeDataAdoRecordset = "adBSTR"
            isText = True
            
        Case ADODB.DataTypeEnum.adChapter
            gfnc_getTypeDataAdoRecordset = "adChapter"
            
        Case ADODB.DataTypeEnum.adChar
            gfnc_getTypeDataAdoRecordset = "adChar"
            isText = True
            
        Case ADODB.DataTypeEnum.adCurrency
            gfnc_getTypeDataAdoRecordset = "adCurrency"
            
        Case ADODB.DataTypeEnum.adDate
            gfnc_getTypeDataAdoRecordset = "adDate"
            isDate = True
            
        Case ADODB.DataTypeEnum.adDBDate
            gfnc_getTypeDataAdoRecordset = "adDBDate"
            isDate = True
            
        Case ADODB.DataTypeEnum.adDBTime
            gfnc_getTypeDataAdoRecordset = "adDBTime"
            isDate = True
            
        Case ADODB.DataTypeEnum.adDBTimeStamp
            gfnc_getTypeDataAdoRecordset = "adDBTimeStamp"
            isDate = True
            
        Case ADODB.DataTypeEnum.adDecimal
            gfnc_getTypeDataAdoRecordset = "adDecimal"
            
        Case ADODB.DataTypeEnum.adDouble
            gfnc_getTypeDataAdoRecordset = "adDouble"
            
        Case ADODB.DataTypeEnum.adEmpty
            gfnc_getTypeDataAdoRecordset = "adEmpty"
            
        Case ADODB.DataTypeEnum.adError
            gfnc_getTypeDataAdoRecordset = "adError"
            
        Case ADODB.DataTypeEnum.adFileTime
            gfnc_getTypeDataAdoRecordset = "adFileTime"
            isDate = True
            
        Case ADODB.DataTypeEnum.adGUID
            gfnc_getTypeDataAdoRecordset = "adGUID"
            
        Case ADODB.DataTypeEnum.adIDispatch
            gfnc_getTypeDataAdoRecordset = "adIDispatch"
            
        Case ADODB.DataTypeEnum.adInteger
            gfnc_getTypeDataAdoRecordset = "adInteger"
            
        Case ADODB.DataTypeEnum.adInteger
            gfnc_getTypeDataAdoRecordset = "adInteger"
            
        Case ADODB.DataTypeEnum.adIUnknown
            gfnc_getTypeDataAdoRecordset = "adIUnknown"
            
        Case ADODB.DataTypeEnum.adLongVarBinary
            gfnc_getTypeDataAdoRecordset = "adLongVarBinary"
            
        Case ADODB.DataTypeEnum.adLongVarChar
            gfnc_getTypeDataAdoRecordset = "adLongVarChar"
            isText = True
            
        Case ADODB.DataTypeEnum.adLongVarWChar
            gfnc_getTypeDataAdoRecordset = "adLongVarWChar"
            isText = True
            
        Case ADODB.DataTypeEnum.adNumeric
            gfnc_getTypeDataAdoRecordset = "adNumeric"
            
        Case ADODB.DataTypeEnum.adPropVariant
            gfnc_getTypeDataAdoRecordset = "adPropVariant"
            
        Case ADODB.DataTypeEnum.adSingle
            gfnc_getTypeDataAdoRecordset = "adSingle"
            
        Case ADODB.DataTypeEnum.adSmallInt
            gfnc_getTypeDataAdoRecordset = "adSmallInt"
            
        Case ADODB.DataTypeEnum.adTinyInt
            gfnc_getTypeDataAdoRecordset = "adTinyInt"
            
        Case ADODB.DataTypeEnum.adUnsignedBigInt
            gfnc_getTypeDataAdoRecordset = "adUnsignedBigInt"
            
        Case ADODB.DataTypeEnum.adUnsignedInt
            gfnc_getTypeDataAdoRecordset = "adUnsignedInt"
            
        Case ADODB.DataTypeEnum.adUnsignedSmallInt
            gfnc_getTypeDataAdoRecordset = "adUnsignedSmallInt"
            
        Case ADODB.DataTypeEnum.adUnsignedTinyInt
            gfnc_getTypeDataAdoRecordset = "adUnsignedTinyInt"
            
        Case ADODB.DataTypeEnum.adUserDefined
            gfnc_getTypeDataAdoRecordset = "adUserDefined"
            
        Case ADODB.DataTypeEnum.adVarBinary
            gfnc_getTypeDataAdoRecordset = "adVarBinary"
            
        Case ADODB.DataTypeEnum.adVarChar
            gfnc_getTypeDataAdoRecordset = "adVarChar"
            isText = True
            
        Case ADODB.DataTypeEnum.adVariant
            gfnc_getTypeDataAdoRecordset = "adVariant"
            
        Case ADODB.DataTypeEnum.adVarNumeric
            gfnc_getTypeDataAdoRecordset = "adVarNumeric"
            
        Case ADODB.DataTypeEnum.adVarWChar
            gfnc_getTypeDataAdoRecordset = "adVarWChar"
            isText = True
            
        Case ADODB.DataTypeEnum.adWChar
            gfnc_getTypeDataAdoRecordset = "adWChar"
            isText = True
            
        Case Else
            gfnc_getTypeDataAdoRecordset = ""
    End Select
End Function
