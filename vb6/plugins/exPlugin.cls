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

Public Sub SetHost(ByVal HostRef As Object, ByVal PluginName As String)
    '===================================================
    Dim k As Integer
    '===================================================
    On Error Resume Next
    
    Set frmMainHost = HostRef
    
    For k = 1 To mn_PluginNumExportedFunctions
        frmMainHost.RegisterPlugin cn_PluginMenuParent, mzs_PluginMenuCaption(k), mzs_PluginDescription(k), k, PluginName
    Next k
    
End Sub

Public Sub StartUp(ByVal intArg As Integer)
    On Error Resume Next
    
    Select Case intArg
        Case 1:
            Main
    End Select
End Sub

Private Sub Class_Initialize()

    On Error Resume Next
    
    PluginAuthor = "[ex]"
    
    ReDim mzs_PluginMenuCaption(1 To 1)
    ReDim mzs_PluginDescription(1 To 1)
    
    mzs_PluginMenuCaption(1) = "Explorar DB"
    mzs_PluginDescription(1) = "Muestra la estructura de archivos de un medio explorado usando R_porter"
    
    mn_PluginNumExportedFunctions = UBound(mzs_PluginMenuCaption)
    
End Sub

Private Sub Class_Terminate()
    Unload frmDataExplorer
End Sub
