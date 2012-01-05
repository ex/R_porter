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

Private frmMainHost As Object

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


Public Function StartUp(ByVal intArg As Integer, Optional ArrayParam As Variant = Nothing) As Boolean
    On Error Resume Next
    
    StartUp = False
     
    Select Case intArg
    '************************************************************************
    '   INICIO ZONA EDITABLE
    '------------------------------------------------------------------------
    ' (TODO)    Llamar aqui a la funcion principal
    Case 1:
        MsgBox "Hola R_porter", vbExclamation, "Plugin Hola"
    '------------------------------------------------------------------------
    ' (TODO)    Agregar aqui las llamadas a las otras funciones:
    '
    'Case 2:
    '    Funcion2
    'Case 3:
    '    Funcion3
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
    
    mzs_PluginMenuCaption(1) = "Plugin de ejemplo"
    mzs_PluginDescription(1) = "Sirve como ejemplo para crear plugins."
    
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
    '------------------------------------------------------------------------
    '   FIN ZONA EDITABLE
    '************************************************************************
End Sub

