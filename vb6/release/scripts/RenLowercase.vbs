'=======================================================================================
' Este script renombra todos lo que se encuentre dentro del subdirectorio a minúsculas
' Author:		Christian d'Heureuse (www.source-code.biz)
' Modified by:	Esau Rodriguez (R_porter 1.1)
'_______________________________________________________________________________________
Option Explicit

Private mn_FilesRenamed
Private mn_FilesSkipped
Private mn_FoldersRenamed
Private mn_FoldersSkipped

Private ms_ToLowerFolder
Private mo_FSO

Const ks_ScriptName = "RenLowercase"

'=======================================================================================
Sub Main()
	On Error Resume Next

	mn_FilesRenamed = 0
	mn_FilesSkipped = 0
	mn_FoldersRenamed = 0
	mn_FoldersSkipped = 0
	Set mo_FSO = CreateObject ("Scripting.FileSystemObject")

	ms_ToLowerFolder = InputBox ("Ingresa el nombre de la carpeta" & vbCrLf & _
								 "a renombrar a minúsculas", "Ingresa folder", "")

	Dim CurrentFolder
	Set CurrentFolder = mo_FSO.GetFolder (ms_ToLowerFolder)

    If Err.Number <> 0 Then
        MsgBox "No se pudo abrir carpeta """ & ms_ToLowerFolder & _
			   """ " & vbCrLf & Err.Description, vbCritical, ks_ScriptName
		Exit Sub
    End If

	If vbNo = MsgBox ("ADVERTENCIA: todos los archivos y subdirectorios dentro del directorio """ & _
					  ms_ToLowerFolder & """ serán renombrados a minúsculas." & "¿Continuar?", _
					  vbYesNo + vbExclamation, "Advertencia") Then
		Exit Sub
	End If

	ProcessFolder CurrentFolder

	MsgBox mn_FilesRenamed & " Files and " & mn_FoldersRenamed & " Folders renombrados a minuscula." & vbCrLf & _
		   mn_FilesSkipped & " Files and " & mn_FoldersSkipped & " Folders estaban ya en minúsculas."

End Sub

'=======================================================================================
Sub ProcessFolder (ByVal Folder)
	Dim Files
	Dim File
	On Error Resume Next

	Set Files = Folder.Files

	For Each File In Files
		'If File.Name <> Trim(File.Name) Then
		'	File.Move "E:\nhk\" & Trim(File.Name)
		'	mn_FilesRenamed = mn_FilesRenamed + 1
		'Else
		'	mn_FilesSkipped = mn_FilesSkipped + 1
		'End If
		If File.Name <> LCase(File.Name) Then
			File.Move LCase(File.Path)
			mn_FilesRenamed = mn_FilesRenamed + 1
		Else
			mn_FilesSkipped = mn_FilesSkipped + 1
		End If
	Next

	Dim SubFolders: Set SubFolders = Folder.SubFolders
	Dim SubFolder

	For Each SubFolder In SubFolders
		If SubFolder.Name <> LCase(SubFolder.Name) Then
			SubFolder.Move LCase(SubFolder.Path)
			mn_FoldersRenamed = mn_FoldersRenamed + 1
		Else

			mn_FoldersSkipped = mn_FoldersSkipped + 1
		End If
		ProcessFolder SubFolder
	Next
End Sub
