VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "exPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************************
'   CLASE PLUGIN PARA R_PORTER
'-------------------------------------------------------------------------------------------
'   Programador:    Esau Rodriguez Oscanoa
'   URL:            (http://www.geocities.com/planeta_dev/spa/r_porter)
'   Notas:          - Debe llamarse exPlugin para poder ser reconocida por el host
'                   - Editar entre las zonas marcadas para edicion
'*******************************************************************************************
Option Explicit

'*******************************************************************************************
'   FUNCION DE REGISTRO EN EL HOST
'-------------------------------------------------------------------------------------------
'Function RegisterPlugin(intMenu As Integer,             <-- parametro reservado (cero)
'                        strMenuName As String,          <-- texto del menu
'                        strDescription As String,       <-- pequeña descripcion
'                        intStartupArgument As Integer,  <-- argumento para llamar al plugin
'                        strPluginName As String)        <-- Nombre del plugin
'-------------------------------------------------------------------------------------------

Private Const cn_PluginMenuParent = 0   ' Actualmente este valor solo puede ser cero

Private mn_PluginNumExportedFunctions As Integer
Private mzs_PluginMenuCaption() As String
Private mzs_PluginDescription() As String

Public PluginAuthor As String

Const KN_END_SCAN_FILES = -1

Const KN_START_DELETE_BLOCK_TEXT = -2
Const KN_SEARCH_DELETE_BLOCK_TEXT = -3

Const KN_START_DELETE_LINE_TEXT = -4
Const KN_SEARCH_DELETE_LINE_TEXT = -5

Const KN_START_FIND_PHRASE = -6
Const KN_SEARCH_FIND_PHRASE = -7

Const KN_START_SCAN_VB = -8
Const KN_SEARCH_SCAN_VB = -9
Const KN_END_SCAN_VB = -10

Const KN_START_SCAN_CPP = -11
Const KN_SEARCH_SCAN_CPP = -12
Const KN_END_SCAN_CPP = -13

Public Sub SetHost(ByRef HostRef As Object, ByVal PluginName As String)
    '===================================================
    Dim k As Integer
    '===================================================
    On Error Resume Next
    
    Set frmMainHost = HostRef
    
    For k = 1 To mn_PluginNumExportedFunctions
        frmMainHost.RegisterPlugin cn_PluginMenuParent, mzs_PluginMenuCaption(k), mzs_PluginDescription(k), k, PluginName
    Next k
    
End Sub

Public Function StartUp(ByVal intArg As Integer, Optional ArrayParam As Variant = Nothing) As Boolean
    On Error Resume Next
    
    StartUp = False
     
    Select Case intArg
    '************************************************************************
    '   INICIO ZONA EDITABLE
    '------------------------------------------------------------------------
    ' (TODO)    Llamar aqui a la funcion principal
    Case 1:
        StartUp = gfnc_exPLGInit(ArrayParam)
        
    Case KN_START_SCAN_VB:
        StartUp = gfnc_exOnStart_ScanVB(ArrayParam)
        
    Case KN_SEARCH_SCAN_VB:
        StartUp = gfnc_exOnSearch_ScanVB(ArrayParam)
        
    Case KN_END_SCAN_VB:
        StartUp = gfnc_exOnEnd_ScanVB(ArrayParam)
        
    Case KN_START_SCAN_CPP:
        StartUp = gfnc_exOnStart_ScanCPP(ArrayParam)
        
    Case KN_SEARCH_SCAN_CPP:
        StartUp = gfnc_exOnSearch_ScanCPP(ArrayParam)
        
    Case KN_END_SCAN_CPP:
        StartUp = gfnc_exOnEnd_ScanCPP(ArrayParam)
    
    Case KN_END_SCAN_FILES:
        StartUp = gfnc_exOnEnd_ScanFiles(ArrayParam)
        
    Case KN_START_DELETE_BLOCK_TEXT:
        StartUp = gfnc_exOnStart_DeleteBlockText(ArrayParam)
        
    Case KN_SEARCH_DELETE_BLOCK_TEXT:
        StartUp = gfnc_exOnSearch_DeleteBlockText(ArrayParam)
        
    Case KN_START_DELETE_LINE_TEXT:
        StartUp = gfnc_exOnStart_DeleteLineText(ArrayParam)
        
    Case KN_SEARCH_DELETE_LINE_TEXT:
        StartUp = gfnc_exOnSearch_DeleteLineText(ArrayParam)
        
    Case KN_START_FIND_PHRASE:
        StartUp = gfnc_exOnStart_FindPhrase(ArrayParam)
        
    Case KN_SEARCH_FIND_PHRASE:
        StartUp = gfnc_exOnSearch_FindPhrase(ArrayParam)
        
    '------------------------------------------------------------------------
    '   FIN ZONA EDITABLE
    '************************************************************************
    End Select
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    
    '************************************************************************
    '   INICIO ZONA EDITABLE
    '------------------------------------------------------------------------
    ' (TODO)    Modificar por quien corresponda
    PluginAuthor = "[ex]"
    
    '------------------------------------------------------------------------
    ' (TODO)    Modificar limites superiores para agregar funciones al plugin
    ' Por ejemplo:  ReDim mzs_PluginMenuCaption(1 To 3), etc..
    '
    ReDim mzs_PluginMenuCaption(1 To 1)
    ReDim mzs_PluginDescription(1 To 1)
    
    mzs_PluginMenuCaption(1) = "Probar scripts en IDE VB"
    mzs_PluginDescription(1) = "Sirve para el desarrollo de scripts usando el IDE VB."
    
    '------------------------------------------------------------------------
    ' (TODO)    Luego agregar el nombre del menu, descripción y autor
    '
    'mzs_PluginMenuCaption(2) = "Otra funcion..."
    'mzs_PluginDescription(2) = "Sirve para etc..."
    'mzs_PluginMenuCaption(3) = "Otra funcion..."
    'mzs_PluginDescription(3) = "Sirve para etc..."
    '------------------------------------------------------------------------
    '   FIN ZONA EDITABLE
    '************************************************************************
    
    mn_PluginNumExportedFunctions = UBound(mzs_PluginMenuCaption)
End Sub

Private Sub Class_Terminate()
    '************************************************************************
    '   INICIO ZONA EDITABLE
    '------------------------------------------------------------------------
    ' (TODO)    Agregar codigo de finalizacion aqui
    If gb_exFormVisible Then
        Unload frmScripter
    End If
    '------------------------------------------------------------------------
    '   FIN ZONA EDITABLE
    '************************************************************************
End Sub

