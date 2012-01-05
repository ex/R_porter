Attribute VB_Name = "mdlMain"
Option Explicit

Public frmMainHost As Object

Public gb_exSCriptTestActive As Boolean
Public gb_exFormVisible As Boolean

'**************************************************
' INIT
'**************************************************
Public Function gfnc_exPLGInit(ArrayParam As Variant) As Boolean
    
    On Error GoTo Handler
    
    gb_exSCriptTestActive = False
    Load frmScripter
    frmScripter.Show vbModeless
    gb_exFormVisible = True

    Exit Function
    
Handler:    MsgBox Err.Description, vbCritical, "Error de inicializacion"
End Function

