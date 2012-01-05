'=======================================================================================
' Renombra todos los directorios (Usa archivo temporal y crea log)
' Esau Rodriguez O.
'_______________________________________________________________________________________
Option Explicit

Private mn_NumFiles				' almacena el numero de archivos procesados
Private mo_FSO					' objeto para manejo de archivos
Private mo_TextStream			' handle al archivo auxiliar
Private mo_LogStream			' handle al archivo de log
Private mb_Cancel				' flag para cancelar procesamiento script
Private	ms_RenamedFilesFolder	' 
Private	ms_OptionalCad			' 

Const kn_ForReading = 1, kn_ForWriting = 2
Const ks_ScriptName = "RenameDirs"
Const ks_TempFileName = "C:\r_porter.$$$"
Const ks_LogFileName = "C:\exRename.log"

'=======================================================================================
' Esta funcion es llamada antes de iniciar la busqueda
Public Sub gsub_exPreSearch()
    On Error Resume Next
        
    mb_Cancel = False
    Set mo_FSO = CreateObject("Scripting.FileSystemObject")
    Set mo_TextStream = mo_FSO.OpenTextFile (ks_TempFileName, kn_ForWriting, True)

    If Err.Number <> 0 Then
        mb_Cancel = True
        MsgBox "Error creando " & ks_TempFileName & vbCrLf & Err.Description, vbCritical, ks_ScriptName
        Exit Sub
    End If

	' Este script renombra los directorios raiz de la carpeta (pero no debe mover los subdirectorios)
	MsgBox "Verifica en opciones de busqueda" & vbCrLf & "QUE NO SE INCLUYAN SUBDIRECTORIOS", vbExclamation, ks_ScriptName

	ms_RenamedFilesFolder = InputBox ("Ingresa el nombre de la carpeta existente donde se guardaran los " & _
									  "directorios renombrados.", , IScript.strSearchPath)

    msub_WriteInReportRGB vbTab & vbTab & "---  INICIO SCRIPT  ---" & vbCrLf, 255, 0, 0
    mn_NumFiles = 0
    IScript.bolCancelReport = True  ' cancelar reporte por defecto
End Sub

'=======================================================================================
' Esta funcion es llamada durante la busqueda
Public Sub gsub_exInSearch()
    On Error Resume Next

    If mb_Cancel Then Exit Sub

    If IScript.bolInSearchIsDir = True Then
        
		mn_NumFiles = mn_NumFiles + 1
        mo_TextStream.WriteLine IScript.strFilePath & "\" & IScript.strFileName

		If Err.Number <> 0 Then
			mb_Cancel = True
			MsgBox "Error escribiendo " & IScript.strFileName & vbCrLf & Err.Description, vbCritical, ks_ScriptName
		Else
	        msub_WriteInReportRGB mn_NumFiles & ") " & IScript.strFilePath & "\" & IScript.strFileName, 0, 0, 128
		End If
    Else
        ' nothing
    End If
End Sub

'=======================================================================================
' Esta funcion es llamada al finalizar la busqueda
Public Sub gsub_exEndSearch()
    On Error Resume Next
    
    If mb_Cancel Then Exit Sub
    
    mo_TextStream.Close
    
	If mn_NumFiles <= 0 Then Exit Sub

    If vbNo = MsgBox("Quieres renombrar " & mn_NumFiles & " directorios?", vbYesNo) Then
        Exit Sub
    End If
    
    msub_WriteInReportRGB vbTab & vbTab & "---   PROCESANDO   ---", 255, 0, 0
    
	msub_ProcesarArchivos

    msub_WriteInReportRGB vbTab & vbTab & "---   FIN SCRIPT  ---", 255, 0, 0
End Sub

'---------------------------------------------------------------------------------------
' Renombrar archivos
Private Sub msub_ProcesarArchivos()
	'=====================================
	Dim strFileOld
    Dim strFileNew
    Dim lngFiles
	'=====================================
    On Error Resume Next

	ms_OptionalCad = InputBox ("Ingresa cadena opcional a agregar a los nombres.", , "")

	Set mo_TextStream = mo_FSO.OpenTextFile (ks_TempFileName, kn_ForReading)

	If Err.Number <> 0 Then
        MsgBox "Error leyendo " & ks_TempFileName & vbCrLf & Err.Description, vbCritical, ks_ScriptName
        Exit Sub
    End If

	Set mo_LogStream = mo_FSO.OpenTextFile (ks_LogFileName, kn_ForWriting, True)

	If Err.Number <> 0 Then
        MsgBox "Error creando " & ks_LogFileName & vbCrLf & Err.Description, vbCritical, ks_ScriptName
        Exit Sub
    End If

    msub_WriteInReportRGB "[Renombrando directorios]", 128, 128, 128

    lngFiles = 1

    Do While (True)

        strFileOld = mo_TextStream.ReadLine		'<- sale del bucle cuando se lanza error fin de archivo...

        If (Trim(strFileOld) = "") Or (Err.Number <> 0) Then Exit Do

		strFileNew = ms_RenamedFilesFolder & "\" & ms_OptionalCad & mfnc_strNumber (mn_NumFiles, lngFiles)

        mo_FSO.MoveFolder strFileOld, strFileNew

        If (Err.Number <> 0) Then Exit Do

		' agregar al log
		mo_LogStream.WriteLine strFileOld & " >>> " & strFileNew

        msub_WriteInReportRGB lngFiles & ") " & strFileNew, 64, 0, 64
        lngFiles = lngFiles + 1
    Loop

    If Err.Number = 62 Then
        ' Fin de archivo
    Else
        MsgBox "Error moviendo archivos." & vbCrLf & Err.Description, vbCritical, ks_ScriptName
    End If

	mo_TextStream.Close

	mo_LogStream.Close

End Sub

'---------------------------------------------------------------------------------------
Private Function mfnc_strNumber (ByVal intTotFiles, Byval intNumFile)
    Dim strNumber
	Dim intNumCifras
    On Error Resume Next
    
	strNumber = Trim (CStr (intNumFile))
	intNumCifras = (Log (intTotFiles) / Log (10)) + 1

	If (Len (strNumber) < intNumCifras) Then
		mfnc_strNumber = String (intNumCifras - Len (strNumber), "0") & strNumber
	Else
		mfnc_strNumber = strNumber
	End If
End Function

'---------------------------------------------------------------------------------------
Private Sub msub_WriteInReportRGB (ByVal strCadena, Byval m_Red, Byval m_Green, Byval m_Blue)
    IScript.rchtxtResults.selcolor = RGB (m_Red, m_Green, m_Blue)
    IScript.rchtxtResults.SelText = strCadena & vbCrLf
End Sub
