Attribute VB_Name = "mdlExplorar"
'*******************************************************************************
' Modulo de variables, constantes y funciones globales para frmexplorar
'*******************************************************************************
' Revisado:     Esau (Agosto 2003)
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Constantes
'*******************************************************************************
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const ERROR_NO_MORE_FILES = 18&
Public Const NL = vbNewLine
Public Const E_INVALID_H = vbObjectError + 1
Public Const PM_REMOVE = &H1
Public Const VK_ESCAPE = &H1B
Public Const UNLEN = 256 + 1    '+1 para el caracter nulo
Public Const MAX_COMPUTERNAME_LENGTH = 15 + 1   '+1 para el caracter nulo

'*******************************************************************************
' Constantes MPEG Audio
'*******************************************************************************
Public Const EX_MPEG_2_5 = 0
Public Const EX_MPEG_FAIL = 1
Public Const EX_MPEG_2 = 2
Public Const EX_MPEG_1 = 3

Public Const EX_LAYER_FAIL = 0
Public Const EX_LAYER_III = 1
Public Const EX_LAYER_II = 2
Public Const EX_LAYER_I = 3

Public Const EX_MODE_STEREO = 0
Public Const EX_MODE_JOINT_STEREO = 1
Public Const EX_MODE_DUAL_CHANNEL = 2
Public Const EX_MODE_SINGLE_CHANNEL = 3

'*******************************************************************************
' Variables globales
'*******************************************************************************
Public strExplorar As String
Public IncluirSubDir As Boolean

'*******************************************************************************
' Definicion de tipos de Datos
'*******************************************************************************
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type MSG
        hWnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type FileExtentions
        num As Integer
        exten(1 To 8) As String * 4
End Type

'*******************************************************************************
' Colores del reporte
'*******************************************************************************
Public gl_ColorDirNormal As Variant
Public gl_ColorDirHidden As Variant
Public gl_ColorDirReadOnly As Variant
Public gl_ColorDirOther As Variant
Public gl_ColorFileNormal As Variant
Public gl_ColorFileHidden As Variant
Public gl_ColorFileReadOnly As Variant
Public gl_ColorFileOther As Variant

Public gl_File1 As Variant
Public gl_File2 As Variant
Public gl_File3 As Variant
Public gl_File4 As Variant
Public gl_File5 As Variant
Public gl_File6 As Variant
Public gl_File7 As Variant
Public gl_File8 As Variant

'*******************************************************************************
' Opciones del reporte
'*******************************************************************************
Public gb_ColoresEnReporte As Boolean
Public gb_LeyendaEnReporte As Boolean
Public gb_TodosLosArchivos As Boolean
Public gt_Extensiones As FileExtentions
Public gb_IncluirSubdirectorios As Boolean

'*******************************************************************************
' Opciones de exportacion a DB
'*******************************************************************************
Public gb_ExportToDB As Boolean
Public gb_SetGenreBySubdir As Boolean
Public gb_SetTypeByExtent As Boolean
Public gb_ConsiderFileAuthorName As Boolean
Public gb_SetQualityMP3 As Boolean
Public gb_SetInfoMPEG As Boolean
Public gs_MediaName As String
Public gs_MediaComment As String
Public gn_DefaultPriority As Byte
Public gn_DefaultBufferLen As Integer
Public gb_CheckBitrateVariable As Boolean
Public gl_StorageType As Long
Public gl_IndexStorageType As Long
Public gl_StorageCategory As Long
Public gs_StorageCategory As String

Public gl_IndexStorageExistent As Long
Public gb_AddToStorageExistent As Long

Public gb_ExportDirToDB As Boolean

Public gs_NameAuthorSeparator As String

Public gb_ActivateDirDepthLimit As Boolean
Public gn_DirDepthLimit As Integer
Public gb_ActivateDirFileLimit As Boolean
Public gn_DirFileLimit As Integer

'*******************************************************************************
' Llamadas al API
'*******************************************************************************
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function PeekMessage Lib "User32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ClientToScreen Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, ByRef PNT As POINTAPI) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long


