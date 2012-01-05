Attribute VB_Name = "mdlMain"
Option Explicit


Public Const guidMYTOOL$ = "_E_X__A_D_D_I_N_"

'**************************************************
' Constantes del API
'**************************************************

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

'**************************************************
' API
'**************************************************
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'**************************************************
' Variables globales del AddIn
'**************************************************
Global gVBInstance As VBIDE.VBE         'la variable VBInstance se establece a la instancia actual de Visual Basic.
Global gwinWindow As VBIDE.Window       'se usa para asegurarse de que sólo se ejecuta una instancia
Global gdoc_usrdoc As Object            'objeto documento de usuario
Public gb_PathFind As Boolean
Public gs_Path As String

'**************************************************
' variables globales para insercion de formulario
'**************************************************
Public gs_Form As String

'**************************************************
' variables globales para acceso a datos
'**************************************************
Public cn As ADODB.Connection
Public query As String

Public gs_ex_DSN As String
Public gs_ex_PWD As String

Public gb_DBConexionOK As Boolean

'**************************************************
' API ODBC
'**************************************************
Declare Function SQLAllocEnv Lib "ODBC32.DLL" (env As Long) As Integer
Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Long, ByVal fdir As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDesc As String, ByVal cbDescMax As Integer, pcbDesc As Integer) As Integer

Public Const SQL_SUCCESS As Long = 0
Public Const SQL_FETCH_NEXT As Long = 1
Public Const SQL_FETCH_FIRST As Long = 2
Public Const SQL_FETCH_FIRST_USER As Long = 31
Public Const SQL_FETCH_FIRST_SYSTEM As Long = 32    'don't easy to guess!

'**************************************************
' FUNCIONES DEL ADDIN
'**************************************************
Public Sub AddToINI()
Dim lng As Long
    
    lng = WritePrivateProfileString("Add-Ins32", "exAddInTable.AddInClass", "0", "VBADDIN.INI")
    MsgBox "Se ha registrado el complemento en el " & "archivo VBADDIN.INI"
    
End Sub

Public Sub Show()
  
  On Error GoTo Handler
  
  gwinWindow.Visible = True
  Exit Sub
  
Handler:

  MsgBox Err.Description, vbCritical, "Show()"
End Sub

'**************************************************
' FUNCIONES DE CONEXION BD
'**************************************************
Public Function gf_CrearConexion(sDSN As String, sPwd As String) As Boolean
    
    On Error GoTo Handler

    gf_CrearConexion = False
    
    Set cn = New ADODB.Connection
    cn.Open "DSN=" & Trim(sDSN) & ";pwd=" & Trim(sPwd)

    gf_CrearConexion = True
    
    Exit Function

Handler:

    Select Case Err.Number
    
        Case -2147467259
            'MsgBox "No se pudo establecer conexion con el servidor o no coincide la contraseña", vbExclamation, "Error"
            gf_CrearConexion = False
        Case Else
            MsgBox Err.Description, vbCritical, "gf_CrearConexion()"
        
    End Select

End Function

Public Sub gp_CerrarConexionBaseDatos()
    
    On Error Resume Next
    cn.Close
    Set cn = Nothing
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                               ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function
