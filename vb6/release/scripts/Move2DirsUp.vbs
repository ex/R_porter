'===========================================================
' MOVE2DIRSUP.VBS - (Esau Rodriguez O.)
' Mueve todos los archivos dos directorios mas arriba
' ADVERTENCIA: Este script es potencialmente peligroso
'-----------------------------------------------------------
Option Explicit

'===========================================================
' Interfaz publica para el script --> IScript (Objeto)
'-----------------------------------------------------------
' Public strSearchPath As String
' Public strFilePath As String
' Public strFileName As String
' Public rchtxtResults() As Object      ' EL objeto control RichTextBox del formulario de exploracion de R_porter
' Public bolPreSearchByDir As Boolean   ' VERDADERO si la busqueda es por Directorios, FALSO si es por unidades
' Public bolInSearchIsDir As Boolean    ' VERDADERO si el archivo encontrado es un directorio, FALSO de lo contrario
' Public bolCancelReport As Boolean     ' VERDADERO cancelara el reporte original del programa, FALSO reporte normal del programa
'___________________________________________________________

Private mn_NumFiles

Private mo_fso
Private mo_TextStream
Private mb_Cancel
Private mb_Continue

Const const_ForReading = 1, const_ForWriting = 2, const_ForAppending = 8

'===========================================================
' Esta funcion es llamada antes de iniciar la busqueda
Public Sub gsub_exPreSearch()
    On Error Resume Next
        
    mb_Cancel = False
    Set mo_fso = CreateObject("Scripting.FileSystemObject")
    Set mo_TextStream = mo_fso.OpenTextFile("C:\r_porter.log", const_ForWriting, True)
        
    If Err.Number <> 0 Then
        mb_Cancel = True
        MsgBox "Error creando r_porter.log", vbCritical, "Move2DirsUp"
        Exit Sub
    End If

    msub_WriteInReport vbTab & vbTab & "---  INICIO SCRIPT  ---" & vbCrLf
    mn_NumFiles = 0
    IScript.bolCancelReport = True  ' cancelar reporte por defecto
End Sub
'___________________________________________________________

'===========================================================
' Esta funcion es llamada durante la busqueda
Public Sub gsub_exInSearch()
    On Error Resume Next
    
    If mb_Cancel Then Exit Sub

    If IScript.bolInSearchIsDir = False Then
        mn_NumFiles = mn_NumFiles + 1
        msub_WriteInReport mn_NumFiles & ") " & IScript.strFilePath & "\" & IScript.strFileName
        mo_TextStream.WriteLine IScript.strFilePath & "\" & IScript.strFileName
    Else
        ' nothing
    End If
End Sub
'___________________________________________________________

'===========================================================
' Esta funcion es llamada al finalizar la busqueda
Public Sub gsub_exEndSearch()
    On Error Resume Next
    
    If mb_Cancel Then Exit Sub
    
    mo_TextStream.Close
    
    If vbNo = MsgBox("Quieres mover " & mn_NumFiles & " archivos?", vbYesNo) Then
        Exit Sub
    End If
    
    msub_WriteInReport vbTab & vbTab & "---   PROCESANDO   ---" 
    
    msub_ProcesarArchivos
    
    msub_WriteInReport vbTab & vbTab & "---   FIN SCRIPT  ---"
End Sub
'___________________________________________________________

Private Sub msub_ProcesarArchivos()
    Dim strFileOld
    Dim strFileNew
    Dim lngFiles
    On Error Resume Next

    Set mo_TextStream = mo_fso.OpenTextFile("C:\r_porter.log", const_ForReading)
    lngFiles = 0

    If Err.Number <> 0 Then
        MsgBox "Error leyendo r_porter.log", vbCritical, "Move2DirsUp"
        Exit Sub
    End If

    msub_WriteInReport "...Moviendo archivos" 

    mb_Continue = True
    Do While (mb_Continue)	'<- sale del bucle cuando se lanza error fin de archivo...

        strFileOld = mo_TextStream.ReadLine

        If (Trim(strFileOld) = "") Or (Err.Number <> 0) Then Exit Do

        strFileNew = mfnc_NewFile(strFileOld)

        mo_fso.MoveFile strFileOld, strFileNew

        If (Err.Number <> 0) Then Exit Do

        lngFiles = lngFiles + 1
        msub_WriteInReport lngFiles & ") " & strFileNew 

    Loop

    If Err.Number = 62 Then
        ' Fin de archivo
        mo_TextStream.Close
    Else
        MsgBox "Sucedio un error:" & vbCrLf & Err.Description, vbCritical, "Move2DirsUp"
    End If
End Sub

Private Function mfnc_NewFile(ByVal strFileName)
    Dim strFile
    Dim strDir

    Dim intPos
    On Error Resume Next
    
    mb_Continue = False
    intPos = InStrRev(strFileName, "\", -1)
    If (intPos > 0) Then
        strFile = Mid(strFileName, intPos + 1)
        intPos = InStrRev(strFileName, "\", intPos - 1)
        If (intPos > 0) Then
            intPos = InStrRev(strFileName, "\", intPos - 1)
            If (intPos > 0) Then
                strDir = Mid(strFileName, 1, intPos)
                mb_Continue = True
                mfnc_NewFile = strDir & strFile
            End If
        End If
    End If
End Function

Private Sub msub_WriteInReport(ByVal strCadena)
    IScript.rchtxtResults.selcolor = RGB(255, 0, 0)
    IScript.rchtxtResults.SelText = strCadena & vbCrLf
End Sub
