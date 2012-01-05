VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public objScript As MSScriptControl.ScriptControl


Private Sub Class_Initialize()

    Set objScript = New MSScriptControl.ScriptControl
    
    objScript.Language = "VBScript"
    objScript.Timeout = 120000
    objScript.AllowUI = True
    objScript.UseSafeSubset = False
    
End Sub
