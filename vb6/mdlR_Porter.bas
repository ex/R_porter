Attribute VB_Name = "mdlR_Porter"
'*******************************************************************************
' MAIN MODULE
' Reviewed:     Laurens Rodriguez Oscanoa - September 2006
'*******************************************************************************
Option Explicit

Public Const EX_VERSION = "1.2.0"

'*******************************************************************************
' Declaracion de constantes publicas
'*******************************************************************************
Public Const MAX_TAM_DRIVES = 120
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const SW_NORMAL = 1
Public Const EM_SETTABSTOPS = &HCB

'*******************************************************************************
' Declaracion de funciones de la API
'*******************************************************************************
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal filename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'*******************************************************************************
' mostrar icono asociado
' [NOTA] no se usa pues el MSHflexgrid no dibuja bien los iconos y no tiene HDC
' Los dibuja grandes y con un borde negro (XP)
'*******************************************************************************
Public Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Public Type CLSID
    id(16) As Byte
End Type

Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long

'*******************************************************************************
' Declaracion de variables globales
'*******************************************************************************
Public NumDrivesSel As Integer
Public stbarDir As Boolean
Public stbarDrives As Boolean
Public ExplorarDir As Boolean
Public strExplorarU As String
Public strExplorarD As String

Public gs_Dev As String
Public gs_About As String

' ruta donde se estan guardando los archivos del menu [Guardar como...] del frmDataControl
Public gs_CopyPath As String

' drive opcional de donde se estan intentara reproducir los archivos
Public gs_OptionalDrive As String

' se activaran o no el script
Public gb_ScriptActivated As Boolean
Public gs_ScriptFile As String

' indica si esta cargado el formulario frmR_porter (plugins)
Public gb_FrmPluginsActive As Boolean

'*******************************************************************************
' FUNCIONES GLOBALES
'*******************************************************************************

'*******************************************************************************
' Funcion inicial del programa
'*******************************************************************************
Public Sub Main()
    
    On Error GoTo Handler
    
    '*****************************
    'inicializacion de opciones
    '*****************************
    gl_ColorDirNormal = RGB(0, 128, 255)
    gl_ColorDirReadOnly = RGB(128, 0, 255)
    gl_ColorDirHidden = RGB(128, 128, 255)
    gl_ColorDirOther = RGB(128, 0, 128)
    gl_ColorFileNormal = RGB(0, 0, 0)
    gl_ColorFileReadOnly = RGB(80, 60, 232)
    gl_ColorFileHidden = RGB(111, 111, 111)
    gl_ColorFileOther = RGB(0, 0, 132)
    
    gl_File1 = RGB(0, 80, 198)
    gl_File2 = RGB(0, 128, 192)
    gl_File3 = RGB(111, 111, 111)
    gl_File4 = RGB(0, 128, 128)
    gl_File5 = RGB(64, 0, 128)
    gl_File6 = RGB(128, 64, 128)
    gl_File7 = RGB(128, 0, 192)
    gl_File8 = RGB(0, 0, 132)

    gb_ColoresEnReporte = True
    gb_LeyendaEnReporte = True
    gb_TodosLosArchivos = True
    
    gb_ExportToDB = False
    gb_SetGenreBySubdir = False
    gb_SetTypeByExtent = True
    gb_ConsiderFileAuthorName = False
    gb_SetQualityMP3 = False
    gb_SetInfoMPEG = False
    
    gb_AddRegisterStarted = False
    
    gs_DSN = "r_porter"
    gs_Pwd = "ex"
    gs_MediaName = "msk"
    gs_MediaComment = ""
    gn_DefaultPriority = 1
    gl_IndexStorageType = 1        ' primero no vacio CD
    gs_StorageCategory = "Música"  ' por defecto buscar musica
    
    gt_DBAuthorSearchStyle = db_Con
    gt_DBNameSearchStyle = db_Con
    
    gb_DBPertenciaPorAlmacenamiento = True
    
    gb_DBCampoAuxiliarPorGenero = False         ' por defecto se mostrara el TIPO (2005)
    
    gb_DBOrdenarAsc(1) = True
    gb_DBOrdenarAsc(2) = True
    gb_DBOrdenarAsc(3) = True
    gb_DBOrdenarAsc(4) = True
    gb_DBOrdenarAsc(5) = True
    gb_DBOrdenarEnabled(1) = True
    gb_DBOrdenarEnabled(2) = True
    gb_DBOrdenarEnabled(3) = False
    gb_DBOrdenarEnabled(4) = False
    gb_DBOrdenarEnabled(5) = False
    gs_DBOrdenarCampo(1) = "Nombre"
    gs_DBOrdenarCampo(2) = "Parent"
    gs_DBOrdenarCampo(3) = "Aux"
    gs_DBOrdenarCampo(4) = "Autor"
    gs_DBOrdenarCampo(5) = "Cantidad"
    
    gb_DBConPrioridad = False
    gt_DBPrioritySearchStyle = db_Igual
    gs_DBConPrioridadDe = "0"
    
    gb_DBConCalidad = False
    gt_DBQualitySearchStyle = db_Igual
    gs_DBConCalidadDe = "0"
    
    gb_DBConTipo = False
    gl_DBConTipoDe = 0
    gl_DBFileTypeIndex = 0
    
    gb_DBConTamanyo = False
    gt_DBFileSizeSearchStyle = db_Igual
    gs_DBConTamanyoDe = "0"
    
    gb_DBConFecha = False
    gt_DBDateSearchStyle = db_Igual
    gd_DBConFechaDe = Now
    
    '[WARN]: ligado a la inicializacion de frmOpcBusqueda.cmbResultados (ver [COD-001})
    gs_DBConCampoDe = "Tamaño"
    gn_DBFieldIndex = 0
    
    gb_DBNameFromFile = False
    gb_DBShowPathInFileName = False
    
    gb_DBShowHiddenFiles = False
    
    gn_DefaultBufferLen = 4096
    gb_CheckBitrateVariable = True
    
    gs_CopyPath = App.Path
    
    gl_IndexStorageExistent = 0
    gb_AddToStorageExistent = False
    
    gs_OptionalDrive = "E"
    
    gb_ExportDirToDB = True
    
    gs_ScriptFile = ""
    gb_ScriptActivated = False
    
    gb_IncluirSubdirectorios = True
    
    gs_NameAuthorSeparator = " - "
    
    gb_ActivateDirDepthLimit = False
    gn_DirDepthLimit = 1
        
    gb_ActivateDirFileLimit = False
    gn_DirFileLimit = 3
        
    If False = gfnc_CrearConexion(gs_DSN, gs_Pwd) Then
        gsub_ShowMessageNoConection
        gb_DBConexionOK = False
        gb_ExportToDB = False
        gb_DBFormatOK = False
    Else
        gb_DBConexionOK = True
        
        ' verificar validez de la base de datos
        If False = gfnc_ValidateDB() Then
            gsub_ShowMessageWrongDB
            gb_DBFormatOK = False
        Else
            gb_DBFormatOK = True
        End If
    End If
   
    If ("DATA" = Command()) Then
        Load frmDataControl
        frmDataControl.Show
        frmDataControl.Refresh
    ElseIf ("EXPLORER" = Command()) Then
        Load frmDataExplorer
        frmDataExplorer.Show
        frmDataExplorer.Refresh
    ElseIf ("SCRIPT" = Command()) Then
        Load frmScript
        frmScript.Show
        frmScript.Refresh
    Else
        frmSplash.Show
        frmSplash.Refresh

        Load frmR_Porter
        frmR_Porter.Show
        frmR_Porter.Refresh

        ' quitar splash
        Unload frmSplash
    End If
   
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "Error de inicializacion"
    End
End Sub

Public Function gfnc_ZGetStrings(ByVal str_buffer As String, ByVal str_len As Integer) As String()
    '===================================================
    Dim k As Long
    Dim char As Integer
    Dim chars As String
    Dim mystr As String
    Dim mz_str() As String
    '===================================================
    k = 0
    Do
        mystr = ""
        Do
            k = k + 1
            chars = Mid$(str_buffer, k, 1)
            char = Asc(chars)
            If char = 0 Then
                Exit Do
            End If
            mystr = mystr & chars
        Loop
        
        gfnc_ZPush mz_str, mystr
        
        chars = Mid(str_buffer, k + 1, 1)
        char = Asc(chars)
        If (char = 0) Or (k >= str_len) Then
            Exit Do
        End If
    Loop
    gfnc_ZGetStrings = mz_str()
    
End Function

Public Sub gfnc_ZPush(ZBuffer, value)
    '===================================================
    Dim k As Integer
    '===================================================
    On Error GoTo Handler
    
    k = UBound(ZBuffer) ' <- throws Error If Not initalized
    ReDim Preserve ZBuffer(k + 1)
    ZBuffer(UBound(ZBuffer)) = value
    Exit Sub
    
Handler:
    ReDim ZBuffer(0): ZBuffer(0) = value
End Sub

Public Function gfnc_ZIsEmpty(ZBuffer) As Boolean
    '===================================================
    Dim k As Integer
    '===================================================
    On Error GoTo Handler
    
    k = UBound(ZBuffer) ' <- throws Error If Not initalized
    gfnc_ZIsEmpty = False
    Exit Function
    
Handler:
    gfnc_ZIsEmpty = True
End Function

Public Function gfnc_GetFileNameWithoutExt(File_Name) As String
    On Error Resume Next
    gfnc_GetFileNameWithoutExt = File_Name
    If InStr(1, File_Name, ".") > 0 Then
       gfnc_GetFileNameWithoutExt = Mid(File_Name, 1, InStrRev(File_Name, ".") - 1)
    End If
End Function

Public Function gfnc_GetFileNameWithoutPath(File_Name) As String
    Dim pos As Integer
    On Error Resume Next
    gfnc_GetFileNameWithoutPath = File_Name
    pos = InStrRev(File_Name, "\")
    If pos > 0 Then
       gfnc_GetFileNameWithoutPath = Mid(File_Name, pos + 1)
    End If
End Function

Public Sub gsub_SetRichTabs(ByVal hWnd As Long, ByVal num As Integer)
    SendMessage hWnd, EM_SETTABSTOPS, 0&, vbNullString
    SendMessage hWnd, EM_SETTABSTOPS, 1, num * 4
End Sub

'-------------------------------------------------------------------------------
' NOTA: Sub procedimiento para registrar el complemento exAddinTable
' desde la ventana inmediato (Ctrl+G) escribir:     gsub_AddToVBADDIN
Sub gsub_AddToVBADDIN()
    Dim ret As Long
    ret = WritePrivateProfileString("Add-Ins32", "exAddInTAble.AddInClass", "0", "VBADDIN.INI")
    MsgBox "El complemento está ahora en el" & "archivo VBADDIN.INI."
End Sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' NOTA: Sub procedimiento para crear registros de plugins
' desde la ventana inmediato (Ctrl+G) escribir:     gsub_Create_R_porter_ADDIN
Sub gsub_Create_R_porter_ADDIN()
    Dim ret As Long
    ret = WritePrivateProfileSection("Add-Ins32", "", App.Path & "\R_porter.ini")
    If ret <> 0 Then
        MsgBox "Se creo el archivo R_porter.ini", vbExclamation
    End If
End Sub
'-------------------------------------------------------------------------------

