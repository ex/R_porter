'=======================================================================================
' Elimina un bloque de texto
' Esau Rodriguez O.
'_______________________________________________________________________________________
Option Explicit

Const kn_ForReading = 1, kn_ForWriting = 2, kn_ForAppending = 8
Const ks_ScriptName = "DeleteBlockOfText.vbs"

Const ks_StartText2Delete = "<INI BLOCK>"
Const ks_EndText2Delete = "<END BLOCK>"
Const kn_MaxLines2Delete = 3
Const kb_DeleteExactNumOfLines = false

'--------------------------------------------------------
' Esta funcion es llamada antes de iniciar la busqueda
Public Sub gsub_exPreSearch()

	On Error Resume Next

	msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
	mn_NumFiles = 0
	mn_NumFilesProcessed = 0
	mn_NumFilesModified = 0
	
	If vbYes = MsgBox("Este script eliminará un bloque de texto" & vbCrLf & _
					  "de los archivos explorados. ¿Quieres continuar?", vbYesNo) Then
				   
		Set mo_FSO = CreateObject("Scripting.FileSystemObject")
		
		If Err.Number <> 0 Then
			msub_WriteInReportRGB "Error creando FSO" & vbCrLf & Err.Description, 255, 0, 0
			gb_exSCriptTestActive = False
		End If
		
		gs_StartText2Delete = ArrayParam(0)
		gs_EndText2Delete = ArrayParam(1)
		gn_MaxLines2Delete = ArrayParam(2)
		
		If Err.Number <> 0 Then
			msub_WriteInReportRGB "Error estableciendo parametros" & vbCrLf & Err.Description, 255, 0, 0
			gb_exSCriptTestActive = False
		End If
		
	End If

End Sub

'--------------------------------------------------------
' Esta funcion es llamada durante la busqueda
Public Sub gsub_exInSearch()

	Dim strFilePath
    Dim strLine
    Dim bolStartFound
    Dim bolEndFound
    Dim bolLookingForStart
    Dim bolErrScanning
    Dim intLines
    
    On Error Resume Next
    
    gfnc_exOnSearch_DeleteBlockText = False
    
    If gb_exSCriptTestActive Then
        If ArrayParam(0) = False Then
            
            bolStartFound = False
            bolEndFound = False
            bolLookingForStart = True
            bolErrScanning = False
            
            ' abrir archivo a explorar en modo lectura
            strFilePath = ArrayParam(1) & "\" & ArrayParam(2)
            Set mo_FileStream = mo_FSO.OpenTextFile(strFilePath, kn_ForReading)
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error abriendo archivo:" & vbCrLf & strFilePath & vbCrLf & Err.Description, 255, 0, 0
                Exit Function
            End If
            
            ' crear archivo temporal
            Set mo_TempStream = mo_FSO.OpenTextFile(ks_TempFile, kn_ForWriting, True)
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error creando archivo temporal:" & vbCrLf & ks_TempFile & vbCrLf & Err.Description, 255, 0, 0
                mo_FileStream.Close
                Exit Function
            End If
            
            ' buscar tag
            Do While (Not bolErrScanning)
            
                ' leer linea del archivo
                strLine = mo_FileStream.ReadLine     '<- sale del bucle cuando se lanza error fin de archivo...
                
                If (Err.Number <> 0) Then Exit Do
                
                If Not bolEndFound Then
                
                    If bolLookingForStart Then
                    
                        If strLine = gs_StartText2Delete Then
                            ' se encontro linea a borrar
                            bolStartFound = True
                            bolLookingForStart = False
                            intLines = 1
                        End If
                    Else
                    
                        intLines = intLines + 1
                        
                        If strLine = gs_EndText2Delete Then
                            
                            If intLines = gn_MaxLines2Delete Then
                                ' se encontro cadena final buscada
                                bolEndFound = True
                            Else
                                Exit Do                 '<- salir de la busqueda cuando falla primer intento
                            End If
                        End If
                    End If
                    
                    If bolLookingForStart And Not bolEndFound Then
                        ' copiar linea a archivo temporal
                        mo_TempStream.WriteLine strLine
                        If (Err.Number <> 0) Then Exit Do
                    End If
                    
                Else
                    ' copiar linea a archivo temporal
                    mo_TempStream.WriteLine strLine
                    If (Err.Number <> 0) Then Exit Do
                End If
                
            Loop
            
            If Err.Number = 62 Or Err.Number = 0 Then
                ' Fin de archivo (o salir cuando falla primer intento)
                Err.Clear
            Else
                msub_WriteInReportRGB "Sucedio un error:" & vbCrLf & Err.Description, 255, 0, 0
                bolErrScanning = True
            End If
            
            ' cerrar archivos
            mo_FileStream.Close
            mo_TempStream.Close
            
            If Not bolErrScanning Then
            
                mn_NumFiles = mn_NumFiles + 1
                
                If Not bolEndFound Then
                    msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> TAG no encontrado", 100, 0, 255
                Else
                    ' remplazar archivo original por copia sin texto eliminado
                    mo_FSO.CopyFile ks_TempFile, strFilePath, True
            
                    If Err.Number <> 0 Then
                        msub_WriteInReportRGB "Error remplazando archivo original" & vbCrLf & Err.Description, 255, 0, 0
                        Exit Function
                    Else
                        msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> TAG eliminado", 0, 0, 0
                        mn_NumFilesModified = mn_NumFilesModified + 1
                    End If
                End If
                
                gfnc_exOnSearch_DeleteBlockText = True
                
            End If
        Else
            msub_WriteInReportRGB "----------------------------------------------------------------------", 200, 200, 255
            gfnc_exOnSearch_DeleteBlockText = True
        End If
        
    End If

End Sub

'--------------------------------------------------------
' Esta funcion es llamada al finalizar la busqueda
Public Sub gsub_exEndSearch()

	On Error Resume Next
    
    gfnc_exOnEnd_ScanFiles = False
    
    If gb_exSCriptTestActive Then
        msub_WriteInReportRGB vbCrLf & "Total de archivos: " & mn_NumFiles, 0, 0, 255
        If mn_NumFilesModified > 0 Then
            msub_WriteInReportRGB "Total de archivos modificados: " & mn_NumFilesModified & vbCrLf, 0, 0, 255
        End If
        If mn_NumFilesProcessed > 0 Then
            msub_WriteInReportRGB "Total de archivos procesados: " & mn_NumFilesProcessed & vbCrLf, 0, 0, 255
        End If
        msub_WriteInReportRGB "[FIN SCRIPT]", 150, 150, 255
        gfnc_exOnEnd_ScanFiles = True
    End If

End Sub
