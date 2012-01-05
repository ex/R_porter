Attribute VB_Name = "mdlPlugin"
Option Explicit

Public frmMainHost As Object

Public cn As ADODB.Connection
Public query As String
Public gs_DNS As String
Public gs_Pwd As String

'**************************************************
' MAIN
'**************************************************
Public Sub Main()
    
    On Error GoTo Handler
    
    gs_DNS = frmMainHost.plg_GetDNS
    gs_Pwd = frmMainHost.plg_GetPWD
    
    If False = gfnc_CrearConexion(gs_DNS, gs_Pwd) Then
        gsub_ShowMessageNoConection
    Else
        Load frmDataExplorer
        frmDataExplorer.Show vbModeless
    End If
   
    Exit Sub
    
Handler:    MsgBox Err.Description, vbCritical, "Error de inicializacion"
End Sub

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

Public Sub gsub_ShowMessageNoConection()
    MsgBox "No está creada la conexión con la base de datos" & vbCrLf & "Para hacerlo use el menú:" & vbCrLf & "[Datos] -> [Conectar a la BD]." & vbCrLf & "De no estar creado el DSN de su base de datos" & vbCrLf & "deberá crearlo primero usando el menú:" & vbCrLf & "[Datos] -> [Crear DSN]." & vbCrLf & "Consulte la ayuda acerca de la instalación", vbExclamation, "Conexión no establecida"
End Sub

'**************************************************
' OTRAS FUNCIONES
'**************************************************
Public Function gfnc_GetFileNameWithoutPath(File_Name) As String
    On Error Resume Next
    gfnc_GetFileNameWithoutPath = File_Name
    If InStr(1, File_Name, "\") > 0 Then
       gfnc_GetFileNameWithoutPath = Mid(File_Name, InStrRev(File_Name, "\") + 1)
    End If
End Function

