Attribute VB_Name = "mdlTest"
'*******************************************************************************************
' EJEMPLO NASICO
'-------------------------------------------------------------------------------------------
'   Programador:    Esau Rodriguez Oscanoa
'*******************************************************************************************
Option Explicit

Private mn_NumFiles
Private mn_NumDirs
Private mn_NumTotal

Public Function gfnc_exTestOnStart(ArrayParam As Variant) As Boolean
    
    On Error GoTo Handler

    If gb_exSCriptTestActive Then
        frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
        frmScripter.rchtxtResults.SelText = vbTab & vbTab & "---  INICIO SCRIPT  ---" & vbCrLf
        mn_NumFiles = 0
        mn_NumDirs = 0
        mn_NumTotal = 0
        If vbYes = MsgBox("Quieres cancelar el reporte por defecto?", vbYesNo) Then
            ArrayParam(0) = True
        Else
            ArrayParam(0) = False
        End If
    End If
    
    gfnc_exTestOnStart = True
    
    Exit Function
    
Handler:    MsgBox Err.Description, vbCritical, "gfnc_exPLGOnStart"
End Function

Public Function gfnc_exTestOnSearch(ArrayParam As Variant) As Boolean
    
    On Error GoTo Handler
    
    If gb_exSCriptTestActive Then
        If ArrayParam(1) = False Then
            mn_NumFiles = mn_NumFiles + 1
            frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
            frmScripter.rchtxtResults.SelText = mn_NumFiles & ") " & ArrayParam(2) & "\" & ArrayParam(3) & vbCrLf
        Else
            mn_NumDirs = mn_NumDirs + 1
            frmScripter.rchtxtResults.SelBold = True
            frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
            frmScripter.rchtxtResults.SelText = mn_NumDirs & ") " & ArrayParam(2) & "\" & ArrayParam(3) & vbCrLf
            frmScripter.rchtxtResults.SelBold = False
        End If
        mn_NumTotal = mn_NumTotal + 1
    End If
    
    Exit Function
    
Handler:    MsgBox Err.Description, vbCritical, "gfnc_exPLGOnSearch"
End Function

Public Function gfnc_exTestOnEnd(ArrayParam As Variant) As Boolean
    
    On Error GoTo Handler
    
    If gb_exSCriptTestActive Then
        frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
        frmScripter.rchtxtResults.SelText = "Total Archivos: " & mn_NumFiles & vbCrLf
        frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
        frmScripter.rchtxtResults.SelText = "Total Directorios: " & mn_NumDirs & vbCrLf
        frmScripter.rchtxtResults.SelColor = RGB(255, 0, 0)
        frmScripter.rchtxtResults.SelText = vbTab & vbTab & "---   FIN SCRIPT  ---" & vbCrLf
    End If
    
    Exit Function
    
Handler:    MsgBox Err.Description, vbCritical, "gfnc_exPLGOnEnd"
End Function

