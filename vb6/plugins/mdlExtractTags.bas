Attribute VB_Name = "mdlExtractTags"
'*******************************************************************************************
' SIRVE PARA ELIMINAR UN TEXTO ESPECIFICO DEL ARCHIVO
'-------------------------------------------------------------------------------------------
'   Programador:    Esau Rodriguez Oscanoa
'*******************************************************************************************
Option Explicit

Private mn_NumFiles             ' almacena el numero de archivos por procesar
Private mn_NumFilesProcessed    ' almacena el numero de archivos correctamente procesados
Private mn_NumFilesModified     ' almacena el numero de archivos modificados
Private mo_FSO                  ' objeto para manejo de archivos
Private mo_FileStream           ' handle al archivo de texto a explorar
Private mo_TempStream           ' handle al archivo temporal

Private mn_NumLines
Private mn_PureLines
Private mn_NumBlankLines
Private mn_NumCommentLines
Private mn_NumFunctions
Private mn_NumSubs
Private mn_NumLinesWithComments

Private ms_Line             ' linea para proceso
Private mb_LinewithCode     ' la linea tenia codigo antes del inicio de comentario en bloque

Private gs_StartText2Delete
Private gs_EndText2Delete
Private gn_MaxLines2Delete

Private gs_LineText2Delete

Private gs_Word2Search
Private gb_SearchCaseSensitive

Const kn_ForReading = 1
Const kn_ForWriting = 2
Const kn_ForAppending = 8
Const ks_ScriptName = "UtilsScripts.vbs"
Const ks_TempFile = "C:\temp.000"

'============================================================================================================
' ELIMINAR BLOQUE EXACTO DE TEXTO
'============================================================================================================
Public Function gfnc_exOnStart_DeleteBlockText(ArrayParam As Variant) As Boolean
    
    On Error Resume Next

    gfnc_exOnStart_DeleteBlockText = False
    
    If gb_exSCriptTestActive Then
        
        msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
        mn_NumFiles = 0
        mn_NumFilesProcessed = 0
        mn_NumFilesModified = 0
        
        If vbYes = MsgBox("Este script eliminará un bloque de texto exacto" & vbCrLf & _
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
            
            gfnc_exOnStart_DeleteBlockText = True
        End If
    End If
    
End Function

Public Function gfnc_exOnSearch_DeleteBlockText(ArrayParam As Variant) As Boolean
    
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
    
End Function

'============================================================================================================
' ELIMINAR LINEA DE TEXTO
'============================================================================================================
Public Function gfnc_exOnStart_DeleteLineText(ArrayParam As Variant) As Boolean
    
    On Error Resume Next

    gfnc_exOnStart_DeleteLineText = False
    
    If gb_exSCriptTestActive Then
        
        msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
        mn_NumFiles = 0
        mn_NumFilesProcessed = 0
        mn_NumFilesModified = 0
        
        If vbYes = MsgBox("Este script eliminará lineas de texto de los archivos" & vbCrLf & _
                          "de la ruta de exploración. ¿Quieres continuar?", vbYesNo) Then
                       
            Set mo_FSO = CreateObject("Scripting.FileSystemObject")
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error creando FSO" & vbCrLf & Err.Description, 255, 0, 0
                gb_exSCriptTestActive = False
            End If
            
            gs_LineText2Delete = ArrayParam(0)
            
            gfnc_exOnStart_DeleteLineText = True
        End If
    End If
    
End Function

Public Function gfnc_exOnSearch_DeleteLineText(ArrayParam As Variant) As Boolean
    
    Dim strFilePath
    Dim strLine
    Dim bolLineFound
    Dim bolSkipLine
    Dim bolErrScanning
    
    On Error Resume Next
    
    gfnc_exOnSearch_DeleteLineText = False
    
    If gb_exSCriptTestActive Then
        If ArrayParam(0) = False Then
            
            bolLineFound = False
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
            
            ' buscar linea a eliminar
            Do While (Not bolErrScanning)
            
                ' leer linea del archivo
                strLine = mo_FileStream.ReadLine     '<- sale del bucle cuando se lanza error fin de archivo...
                
                If (Err.Number <> 0) Then Exit Do
                
                bolSkipLine = False
                
                If strLine = gs_LineText2Delete Then
                    ' se encontro linea a borrar
                    bolLineFound = True
                    bolSkipLine = True
                End If
                
                If Not bolSkipLine Then
                    ' copiar linea a archivo temporal
                    mo_TempStream.WriteLine strLine
                    If (Err.Number <> 0) Then Exit Do
                End If
                    
            Loop
            
            If Err.Number = 62 Then
                ' Fin de archivo
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
                
                If Not bolLineFound Then
                    msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> LINEA no encontrada", 100, 0, 255
                Else
                    ' remplazar archivo original por copia sin linea eliminada
                    mo_FSO.CopyFile ks_TempFile, strFilePath, True
            
                    If Err.Number <> 0 Then
                        msub_WriteInReportRGB "Error remplazando archivo original" & vbCrLf & Err.Description, 255, 0, 0
                        Exit Function
                    Else
                        msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> LINEA eliminada", 0, 0, 0
                        mn_NumFilesModified = mn_NumFilesModified + 1
                    End If
                End If
                
                gfnc_exOnSearch_DeleteLineText = True
                
            End If
            
        Else
            msub_WriteInReportRGB "----------------------------------------------------------------------", 200, 200, 255
            gfnc_exOnSearch_DeleteLineText = True
        End If
        
    End If
    
End Function

'============================================================================================================
' BUSCAR CADENA DE TEXTO
'============================================================================================================
Public Function gfnc_exOnStart_FindPhrase(ArrayParam As Variant) As Boolean
    
    On Error Resume Next

    gfnc_exOnStart_FindPhrase = False
    
    If gb_exSCriptTestActive Then
        
        msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
        mn_NumFiles = 0
        mn_NumFilesProcessed = 0
        mn_NumFilesModified = 0
        
        If vbYes = MsgBox("Este script buscará cadenas de texto en los archivos" & vbCrLf & _
                          "de la ruta de exploración. ¿Quieres continuar?", vbYesNo) Then
                       
            Set mo_FSO = CreateObject("Scripting.FileSystemObject")
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error creando FSO" & vbCrLf & Err.Description, 255, 0, 0
                gb_exSCriptTestActive = False
            End If
            
            gs_Word2Search = ArrayParam(0)
            gb_SearchCaseSensitive = ArrayParam(1)
            
            gfnc_exOnStart_FindPhrase = True
        End If
    End If
    
End Function

Public Function gfnc_exOnSearch_FindPhrase(ArrayParam As Variant) As Boolean
    
    Dim strFilePath
    Dim strLine
    Dim intNumWords
    Dim lngPos
    Dim bolErrScanning
    
    On Error Resume Next
    
    gfnc_exOnSearch_FindPhrase = False
    
    If gb_exSCriptTestActive Then
        If ArrayParam(0) = False Then
            
            intNumWords = 0
            bolErrScanning = False
            
            ' abrir archivo a explorar en modo lectura
            strFilePath = ArrayParam(1) & "\" & ArrayParam(2)
            Set mo_FileStream = mo_FSO.OpenTextFile(strFilePath, kn_ForReading)
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error abriendo archivo:" & vbCrLf & strFilePath & vbCrLf & Err.Description, 255, 0, 0
                Exit Function
            End If
            
            ' buscar palabra
            Do While (Not bolErrScanning)
            
                ' leer linea del archivo
                strLine = mo_FileStream.ReadLine     '<- sale del bucle cuando se lanza error fin de archivo...
                
                If (Err.Number <> 0) Then Exit Do
                
                lngPos = 0
                
                If gb_SearchCaseSensitive Then
                    lngPos = InStr(strLine, gs_Word2Search)
                Else
                    lngPos = InStr(UCase(strLine), UCase(gs_Word2Search))
                End If
                
                If lngPos > 0 Then
                    ' se encontro palabra
                    intNumWords = intNumWords + 1
                End If
                
            Loop
            
            If Err.Number = 62 Then
                ' Fin de archivo
                Err.Clear
            Else
                msub_WriteInReportRGB "Sucedio un error:" & vbCrLf & Err.Description, 255, 0, 0
                bolErrScanning = True
            End If
            
            ' cerrar archivos
            mo_FileStream.Close
            
            If Not bolErrScanning Then
            
                mn_NumFiles = mn_NumFiles + 1
                
                If intNumWords <= 0 Then
                    msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> No encontrada", 128, 0, 255
                Else
                    msub_WriteInReportRGB mn_NumFiles & ") " & strFilePath & " --> (" & intNumWords & ") encontrados", 0, 0, 0
                    mn_NumFilesProcessed = mn_NumFilesProcessed + 1
                End If
                
                gfnc_exOnSearch_FindPhrase = True
                
            End If
        Else
            msub_WriteInReportRGB strFilePath, 100, 100, 155
            gfnc_exOnSearch_FindPhrase = True
        End If
        
    End If
    
End Function

'============================================================================================================
' SCAN FILES END
'============================================================================================================
Public Function gfnc_exOnEnd_ScanFiles(ArrayParam As Variant) As Boolean
    
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
End Function

'============================================================================================================
' SCAN VB
'============================================================================================================
Public Function gfnc_exOnStart_ScanVB(ArrayParam As Variant) As Boolean
    
    On Error Resume Next

    gfnc_exOnStart_ScanVB = False
    
    If gb_exSCriptTestActive Then
        
        msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
        mn_NumFiles = 0
        mn_NumFilesProcessed = 0
        mn_NumBlankLines = 0
        mn_NumCommentLines = 0
        mn_NumFunctions = 0
        mn_NumLines = 0
        mn_NumSubs = 0
        mn_NumLinesWithComments = 0
        
        If vbYes = MsgBox("Este script analizara el codigo de tu proyecto VB" & vbCrLf & _
                          "¿Quieres continuar?", vbYesNo) Then
                       
            Set mo_FSO = CreateObject("Scripting.FileSystemObject")
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error creando FSO" & vbCrLf & Err.Description, 255, 0, 0
                gb_exSCriptTestActive = False
            End If
            
            gfnc_exOnStart_ScanVB = True
        End If
    End If
    
End Function

Public Function gfnc_exOnSearch_ScanVB(ArrayParam As Variant) As Boolean
    
    Dim strFilePath
    Dim strLine
    Dim lngFunctions
    Dim lngSubs
    Dim intCommentLines
    Dim intBlankLines
    Dim intLines
    Dim intLinesWithComments
    Dim lngPos
    Dim lngPosComm
    Dim bolErrScanning
    
    On Error Resume Next
    
    gfnc_exOnSearch_ScanVB = False
    
    If gb_exSCriptTestActive Then
        If ArrayParam(0) = False Then
            
            intCommentLines = 0
            lngFunctions = 0
            lngSubs = 0
            intBlankLines = 0
            intLines = 0
            intLinesWithComments = 0
            bolErrScanning = False
            
            ' abrir archivo a explorar en modo lectura
            strFilePath = ArrayParam(1) & "\" & ArrayParam(2)
            Set mo_FileStream = mo_FSO.OpenTextFile(strFilePath, kn_ForReading)
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error abriendo archivo:" & vbCrLf & strFilePath & vbCrLf & Err.Description, 255, 0, 0
                Exit Function
            End If
                
            mn_NumFiles = mn_NumFiles + 1
            
            frmScripter.rchtxtResults.SelBold = True
            strLine = mn_NumFiles & ")" & strFilePath
            msub_WriteInReportRGB strLine, 128, 0, 255
            frmScripter.rchtxtResults.SelBold = True
            msub_WriteInReportRGB String(Len(strLine), "-"), 128, 0, 255
            frmScripter.rchtxtResults.SelBold = False
            
            ' buscar comentarios y lineas en blanco
            Do While (Not bolErrScanning)
            
                ' leer linea del archivo
                strLine = mo_FileStream.ReadLine     '<- sale del bucle cuando se lanza error fin de archivo...
                
                If (Err.Number <> 0) Then Exit Do
                
                intLines = intLines + 1
                
                strLine = Trim(Replace(strLine, vbTab, " "))    ' remplazar tabs!! (no usar LTrim)
                
                If strLine = "" Then
                    ' si es linea en blanco
                    intBlankLines = intBlankLines + 1
                Else
                    
                    If Left(strLine, 1) = "'" Then
                        ' si es linea de puro comentario
                        intCommentLines = intCommentLines + 1
                        
                    Else
                        ' verificar si linea tiene comentarios
                        lngPosComm = 0
                        lngPosComm = InStr(strLine, "'")
                        If lngPosComm > 0 Then
                            intLinesWithComments = intLinesWithComments + 1
                        End If
                        
                        ' contabilizar funciones por "End Sub" y "End Function" hallados
                        lngPos = 0
                        lngPos = InStr(UCase(strLine), "END SUB")
                        
                        If lngPos > 0 Then
                            ' se encontro fin de subprocedimiento
                            ' verificar que no esta comentado
                            If (lngPosComm = 0) Or (lngPos < lngPosComm) Then
                                lngSubs = lngSubs + 1
                            End If
                        End If
                    
                        lngPos = 0
                        lngPos = InStr(UCase(strLine), "END FUNCTION")
                        
                        If lngPos > 0 Then
                            ' se encontro fin de funcion
                            ' verificar que no esta comentado
                            If (lngPosComm = 0) Or (lngPos < lngPosComm) Then
                                lngFunctions = lngFunctions + 1
                            End If
                        End If
                        
                    End If
                End If
                
            Loop
            
            If Err.Number = 62 Then
                ' Fin de archivo
                Err.Clear
            Else
                msub_WriteInReportRGB "Sucedio un error:" & vbCrLf & Err.Description, 255, 0, 0
                bolErrScanning = True
            End If
            
            ' cerrar archivos
            mo_FileStream.Close
            
            If Not bolErrScanning Then
                
                msub_WriteInReportRGB vbTab & "LINEAS EN BLANCO       : " & intBlankLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS DE COMENTARIOS  : " & intCommentLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS CON COMENTARIOS : " & intLinesWithComments, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS DE CODIGO       : " & intLines - intBlankLines - intCommentLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "TOTAL DE LINEAS        : " & intLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "SUBPROCEDIMIENTOS      : " & lngSubs, 180, 0, 255
                msub_WriteInReportRGB vbTab & "FUNCIONES              : " & lngFunctions & vbCrLf, 180, 0, 255
                
                mn_NumLines = mn_NumLines + intLines
                mn_NumBlankLines = mn_NumBlankLines + intBlankLines
                mn_NumCommentLines = mn_NumCommentLines + intCommentLines
                mn_NumFunctions = mn_NumFunctions + lngFunctions
                mn_NumSubs = mn_NumSubs + lngSubs
                mn_NumLinesWithComments = mn_NumLinesWithComments + intLinesWithComments
                
                mn_NumFilesProcessed = mn_NumFilesProcessed + 1
                
                gfnc_exOnSearch_ScanVB = True
                
            End If
        Else
            gfnc_exOnSearch_ScanVB = True
        End If
        
    End If
    
End Function

Public Function gfnc_exOnEnd_ScanVB(ArrayParam As Variant) As Boolean
    
    On Error Resume Next
    
    gfnc_exOnEnd_ScanVB = False
    
    If gb_exSCriptTestActive Then
        msub_WriteInReportRGBB vbTab & "----------------------------------------------------", 100, 100, 100, True
        msub_WriteInReportRGBB vbTab & "TOTAL ARCHIVOS                 : " & mn_NumFiles, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL ARCHIVOS PROCESADOS      : " & mn_NumFilesProcessed & vbCrLf, 0, 0, 0, True
        
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS DE CODIGO      : " & mn_NumLines - mn_NumBlankLines - mn_NumCommentLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS VACIAS         : " & mn_NumBlankLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS DE COMENTARIO  : " & mn_NumCommentLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS CON COMENTARIO : " & mn_NumLinesWithComments, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS                : " & mn_NumLines & vbCrLf, 0, 0, 0, True
        
        Dim prcActiv
        ' porcentaje de lineas que no estan vacias
        prcActiv = 100 * (mn_NumLines - mn_NumBlankLines) / mn_NumLines
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE LINEAS ACTIVAS   : " & FormatNumber(prcActiv, 2) & " %", 255, 0, 0, True
        
        Dim prcEfect
        ' porcentaje de lineas activas que no son puro comentario
        prcEfect = 100 * (mn_NumLines - mn_NumBlankLines - mn_NumCommentLines) / (mn_NumLines - mn_NumBlankLines)
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE LINEAS EFECTIVAS : " & FormatNumber(prcEfect, 2) & " %", 255, 0, 0, True
        
        Dim prcComms
        ' porcentaje de lineas que tienen o son comentarios de las lineas utiles (no vacias)
        prcComms = 100 * (mn_NumLinesWithComments + mn_NumCommentLines) / (mn_NumLines - mn_NumBlankLines)
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE COMENTARIOS      : " & FormatNumber(prcComms, 2) & " %" & vbCrLf, 255, 0, 0, True
        
        msub_WriteInReportRGBB vbTab & "TOTAL DE FUNCIONES             : " & mn_NumFunctions, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE SUBPROCEDIMIENTOS     : " & mn_NumSubs, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "----------------------------------------------------", 100, 100, 100, True
        gfnc_exOnEnd_ScanVB = True
    End If
End Function

'============================================================================================================
' SCAN CPP
'============================================================================================================
Public Function gfnc_exOnStart_ScanCPP(ArrayParam As Variant) As Boolean
    
    On Error Resume Next

    gfnc_exOnStart_ScanCPP = False
    
    If gb_exSCriptTestActive Then
        
        msub_WriteInReportRGB "[INICIO SCRIPT]" & vbCrLf, 150, 150, 255
        mn_NumFiles = 0
        mn_NumFilesProcessed = 0
        mn_NumBlankLines = 0
        mn_NumCommentLines = 0
        mn_NumFunctions = 0
        mn_NumLines = 0
        mn_NumSubs = 0
        mn_NumLinesWithComments = 0
        
        If vbYes = MsgBox("Este script analizara el codigo de tu proyecto (C/C++)" & vbCrLf & _
                          "¿Quieres continuar?", vbYesNo) Then
                       
            Set mo_FSO = CreateObject("Scripting.FileSystemObject")
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error creando FSO" & vbCrLf & Err.Description, 255, 0, 0
                gb_exSCriptTestActive = False
            End If
            
            gfnc_exOnStart_ScanCPP = True
        End If
    End If
    
End Function

Public Sub msub_ProcessComments_CPP(ByRef bolInsideBlock, ByRef numCommentLines, ByRef numLinesWithComments, ByRef numPureLines, ByVal bolCallBack)

    Dim lngPosComm
    Dim lngPosBlockComm

    ' verificar si linea tiene comentarios
    lngPosComm = 0
    lngPosComm = InStr(ms_Line, "//")
    
    If Not bolInsideBlock Then
    '-------------------------------------------------------------
    ' ESTAMOS FUERA DE BLOQUE
    '-------------------------------------------------------------
        lngPosBlockComm = 0
        lngPosBlockComm = InStr(ms_Line, "/*")
        
        If lngPosComm > 0 Then
            If lngPosBlockComm > 0 Then
            
                If lngPosComm < lngPosBlockComm Then
                    ' el // esta antes que el /* (bloque ignorado)
                    '-------------------------------------------------------------
                    If lngPosComm = 1 Then
                        numCommentLines = numCommentLines + 1
                    Else
                        numLinesWithComments = numLinesWithComments + 1
                    End If
                Else
                    ' el /* esta antes que el //
                    '-------------------------------------------------------------
                    ' [STRIP] quitar parte de la linea comentada
                    ms_Line = Trim(Mid(ms_Line, lngPosBlockComm + 2))
                    ' guardar el hecho de que la linea tenia codigo previo al bloque
                    If lngPosBlockComm > 1 Then
                        mb_LinewithCode = True
                    End If
                    ' se esta dentro de bloque
                    bolInsideBlock = True
                    ' seguir procesando, buscar fin de bloque
                    msub_ProcessComments_CPP bolInsideBlock, numCommentLines, numLinesWithComments, numPureLines, True
                End If
            Else
                ' se encontro // pero no /*
                '-------------------------------------------------------------
                If lngPosComm = 1 Then
                    numCommentLines = numCommentLines + 1
                Else
                    numLinesWithComments = numLinesWithComments + 1
                End If
            End If
        Else
            If lngPosBlockComm > 0 Then
                ' se encontro /* pero no //
                '-------------------------------------------------------------
                ' [STRIP] quitar parte de la linea comentada
                ms_Line = Trim(Mid(ms_Line, lngPosBlockComm + 2))
                ' guardar el hecho de que la linea tenia codigo previo al bloque
                If lngPosBlockComm > 1 Then
                    mb_LinewithCode = True
                End If
                ' se esta dentro de bloque
                bolInsideBlock = True
                ' seguir procesando, buscar fin de bloque
                msub_ProcessComments_CPP bolInsideBlock, numCommentLines, numLinesWithComments, numPureLines, True
            Else
                ' no se encontro ningun comienzo de comentario
                '-------------------------------------------------------------
                If bolCallBack Then
                    ' la funxion ha sido llamada recursivamente
                    If Trim(ms_Line) = "" Then
                        ' la linea es de puro comentario
                        numCommentLines = numCommentLines + 1
                    Else
                        ' la linea no es de puro comentario
                        numLinesWithComments = numLinesWithComments + 1
                    End If
                Else
                    ' contabilizar la linea como de puro codigo
                    numPureLines = numPureLines + 1
                End If
            End If
        End If
        
    Else
    '-------------------------------------------------------------
    ' ESTAMOS DENTRO DE BLOQUE
    '-------------------------------------------------------------
        lngPosBlockComm = 0
        lngPosBlockComm = InStr(ms_Line, "*/")
        
        If lngPosBlockComm > 0 Then
            ' se encontro */
            '-------------------------------------------------------------
            ' [STRIP] quitar parte de la linea comentada
            ms_Line = Trim(Mid(ms_Line, lngPosBlockComm + 2))
            ' se salio de bloque
            bolInsideBlock = False
            ' seguir procesando, buscar inicio de bloque
            msub_ProcessComments_CPP bolInsideBlock, numCommentLines, numLinesWithComments, numPureLines, True
        Else
            ' no se encontro fin de bloque de comentario
            '-------------------------------------------------------------
            If bolCallBack Then
                If mb_LinewithCode Then
                    ' la linea no es de puro comentario
                    numLinesWithComments = numLinesWithComments + 1
                Else
                    ' contabilizar la linea como de puro comentario
                    numCommentLines = numCommentLines + 1
                End If
            Else
                ' contabilizar la linea como de puro comentario
                numCommentLines = numCommentLines + 1
            End If
        End If
        
    End If
    
End Sub

Public Function gfnc_exOnSearch_ScanCPP(ArrayParam As Variant) As Boolean
    
    Dim strFilePath
    Dim intCommentLines
    Dim intBlankLines
    Dim intLines
    Dim intPureLines
    Dim intLinesWithComments
    Dim lngPos
    Dim lngPosComm
    Dim lngPosBlockComm
    Dim bolErrScanning
    Dim bolInsideBlockComments
    
    On Error Resume Next
    
    gfnc_exOnSearch_ScanCPP = False
    
    If gb_exSCriptTestActive Then
        If ArrayParam(0) = False Then
            
            intCommentLines = 0
            intBlankLines = 0
            intLines = 0
            intPureLines = 0
            intLinesWithComments = 0
            bolErrScanning = False
            bolInsideBlockComments = False
            
            ' abrir archivo a explorar en modo lectura
            strFilePath = ArrayParam(1) & "\" & ArrayParam(2)
            Set mo_FileStream = mo_FSO.OpenTextFile(strFilePath, kn_ForReading)
            
            If Err.Number <> 0 Then
                msub_WriteInReportRGB "Error abriendo archivo:" & vbCrLf & strFilePath & vbCrLf & Err.Description, 255, 0, 0
                Exit Function
            End If
                
            mn_NumFiles = mn_NumFiles + 1
            
            frmScripter.rchtxtResults.SelBold = True
            ms_Line = mn_NumFiles & ")" & strFilePath
            msub_WriteInReportRGB ms_Line, 128, 0, 255
            frmScripter.rchtxtResults.SelBold = True
            msub_WriteInReportRGB String(Len(ms_Line), "-"), 128, 0, 255
            frmScripter.rchtxtResults.SelBold = False
            
            ' buscar comentarios y lineas en blanco
            Do While (Not bolErrScanning)
            
                ' leer linea del archivo
                ms_Line = mo_FileStream.ReadLine     '<- sale del bucle cuando se lanza error fin de archivo...
                
                If (Err.Number <> 0) Then Exit Do
                
                intLines = intLines + 1
                
                ms_Line = Trim(Replace(ms_Line, vbTab, " "))    ' remplazar tabs!! (no usar LTrim)
                
                If ms_Line = "" Then
                    ' si es linea en blanco
                    intBlankLines = intBlankLines + 1
                Else
                    mb_LinewithCode = False
                    
                    msub_ProcessComments_CPP bolInsideBlockComments, intCommentLines, intLinesWithComments, intPureLines, False
                    
                End If
            Loop
            
            If Err.Number = 62 Then
                ' Fin de archivo
                Err.Clear
            Else
                msub_WriteInReportRGB "Sucedio un error:" & vbCrLf & Err.Description, 255, 0, 0
                bolErrScanning = True
            End If
            
            ' cerrar archivos
            mo_FileStream.Close
            
            If Not bolErrScanning Then
                
                msub_WriteInReportRGB vbTab & "LINEAS EN BLANCO       : " & intBlankLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS DE COMENTARIOS  : " & intCommentLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS CON COMENTARIOS : " & intLinesWithComments, 0, 0, 0
                msub_WriteInReportRGB vbTab & "LINEAS DE CODIGO       : " & intPureLines, 0, 0, 0
                msub_WriteInReportRGB vbTab & "TOTAL DE LINEAS        : " & intLines, 0, 0, 0
                
                mn_NumLines = mn_NumLines + intLines
                mn_PureLines = mn_PureLines + intPureLines
                mn_NumBlankLines = mn_NumBlankLines + intBlankLines
                mn_NumCommentLines = mn_NumCommentLines + intCommentLines
                mn_NumLinesWithComments = mn_NumLinesWithComments + intLinesWithComments
                
                mn_NumFilesProcessed = mn_NumFilesProcessed + 1
                
                gfnc_exOnSearch_ScanCPP = True
                
            End If
        Else
            gfnc_exOnSearch_ScanCPP = True
        End If
        
    End If
    
End Function

Public Function gfnc_exOnEnd_ScanCPP(ArrayParam As Variant) As Boolean
    
    On Error Resume Next
    
    gfnc_exOnEnd_ScanCPP = False
    
    If gb_exSCriptTestActive Then
        msub_WriteInReportRGBB vbTab & "----------------------------------------------------", 100, 100, 100, True
        msub_WriteInReportRGBB vbTab & "TOTAL ARCHIVOS                 : " & mn_NumFiles, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL ARCHIVOS PROCESADOS      : " & mn_NumFilesProcessed & vbCrLf, 0, 0, 0, True
        
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS DE PURO CODIGO : " & mn_PureLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS VACIAS         : " & mn_NumBlankLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS DE COMENTARIO  : " & mn_NumCommentLines, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS CON COMENTARIO : " & mn_NumLinesWithComments, 0, 0, 0, True
        msub_WriteInReportRGBB vbTab & "TOTAL DE LINEAS                : " & mn_NumLines & vbCrLf, 0, 0, 0, True
        
        Dim prcActiv
        ' porcentaje de lineas que no estan vacias
        prcActiv = 100 * (mn_NumLines - mn_NumBlankLines) / mn_NumLines
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE LINEAS ACTIVAS   : " & FormatNumber(prcActiv, 2) & " %", 255, 0, 0, True
        
        Dim prcEfect
        ' porcentaje de lineas activas que no son puro comentario
        prcEfect = 100 * (mn_NumLines - mn_NumBlankLines - mn_NumCommentLines) / (mn_NumLines - mn_NumBlankLines)
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE LINEAS EFECTIVAS : " & FormatNumber(prcEfect, 2) & " %", 255, 0, 0, True
        
        Dim prcComms
        ' porcentaje de lineas que tienen o son comentarios de las lineas utiles (no vacias)
        prcComms = 100 * (mn_NumLinesWithComments + mn_NumCommentLines) / (mn_NumLines - mn_NumBlankLines)
        msub_WriteInReportRGBB vbTab & "PORCENTAJE DE COMENTARIOS      : " & FormatNumber(prcComms, 2) & " %", 255, 0, 0, True
        msub_WriteInReportRGBB vbTab & "----------------------------------------------------", 100, 100, 100, True
        gfnc_exOnEnd_ScanCPP = True
    End If
End Function

'============================================================================================================
'
'============================================================================================================
Private Sub msub_WriteInReportRGB(ByVal strCadena, ByVal m_Red, ByVal m_Green, ByVal m_Blue)
    frmScripter.rchtxtResults.SelColor = RGB(m_Red, m_Green, m_Blue)
    frmScripter.rchtxtResults.SelText = strCadena & vbCrLf
End Sub

Private Sub msub_WriteInReportRGBB(ByVal strCadena, ByVal m_Red, ByVal m_Green, ByVal m_Blue, ByVal bolBold)
    frmScripter.rchtxtResults.SelColor = RGB(m_Red, m_Green, m_Blue)
    If bolBold = True Then frmScripter.rchtxtResults.SelBold = True
    frmScripter.rchtxtResults.SelText = strCadena & vbCrLf
    If bolBold = True Then frmScripter.rchtxtResults.SelBold = False
End Sub


