VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strSearchPath As String

Public strFilePath As String

Public strFileName As String

Public bolPreSearchByDir As Boolean     ' VERDADERO si la busqueda es por Directorios, FALSO si es por unidades

Public bolInSearchIsDir As Boolean      ' VERDADERO si el archivo encontrado es un directorio, FALSO de lo contrario

Public bolCancelReport As Boolean       ' VERDADERO cancelara el reporte original del programa, FALSO reporte normal del programa

Private rchtxtForScript As RichTextBox


Public Property Get rchtxtResults() As Object
    Set rchtxtResults = rchtxtForScript
End Property

Public Property Set rchtxtResults(newObject As Object)
    Set rchtxtForScript = newObject
End Property

Private Sub Class_Initialize()
    bolPreSearchByDir = False
    bolInSearchIsDir = False
    bolCancelReport = False
End Sub

'-----------------------------------------------------------------------------------------------------
' Ejecuta el plugin cargado que tenga como nombre "strPluginName"
'
Public Function ExecutePlugin(ByVal strPluginName As String, ByVal intParam As Integer, Optional ArrayParam As Variant = Nothing) As Boolean

    If gb_FrmPluginsActive = True Then
        ExecutePlugin = frmR_Porter.gfnc_ScriptExecutePlugin(strPluginName, intParam, ArrayParam)
    Else
        ExecutePlugin = False
    End If
End Function

