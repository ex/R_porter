'=======================================================================================
' EJEMPLO DE COMO CARGAR PLUGINS DESDE EL SCRIPT
' Esau Rodriguez O.
'_______________________________________________________________________________________
Option Explicit

Private	mb_PluginLoaded		' Flag para cancelar procesamiento
							' (si no se carga correctamente el plugin)

Private mz_Params(2)		' Matriz para pasar parametros
							' (El plugin puede modificar sus valores)

Const KN_END_SCAN_FILES = -1

Const KN_START_DELETE_BLOCK_TEXT = -2
Const KN_SEARCH_DELETE_BLOCK_TEXT = -3

Const KN_START_DELETE_LINE_TEXT = -4
Const KN_SEARCH_DELETE_LINE_TEXT = -5

Const KN_START_FIND_PHRASE = -6
Const KN_SEARCH_FIND_PHRASE = -7

Const KN_START_SCAN_VB = -8
Const KN_SEARCH_SCAN_VB = -9
Const KN_END_SCAN_VB = -10

Const KN_START_SCAN_CPP = -11
Const KN_SEARCH_SCAN_CPP = -12
Const KN_END_SCAN_CPP = -13

Const ks_PluginName = "exPlugin_Scripter.dll"

Const ks_StartText2Delete = "<INI BLOCK>"
Const ks_EndText2Delete = "<END BLOCK>"
Const kn_MaxLines2Delete = 3

Const ks_LineText2Delete = "<LINE TO DEL>"

Const ks_Word2Search = "<FIND THIS PHRASE>"
Const kb_SearchCaseSensitive = False

'=======================================================================================
' CONTABILIZAR LINEAS DE CODIGO DE PROYECTO CPP
'_______________________________________________________________________________________
Public Sub gsub_exPreSearch()
	
	mb_PluginLoaded = False
	mb_PluginLoaded = IScript.ExecutePlugin (ks_PluginName, KN_START_SCAN_CPP, mz_Params)

	If Not mb_PluginLoaded Then
		MsgBox "No se pudo cargar correctamente el plugin", vbExclamation
	Else
		IScript.bolCancelReport = True
	End If
End Sub

Public Sub gsub_exInSearch()
	Dim ret
	If mb_PluginLoaded Then

		mz_Params(0) = IScript.bolInSearchIsDir
		mz_Params(1) = IScript.strFilePath
		mz_Params(2) = IScript.strFileName

		ret = IScript.ExecutePlugin(ks_PluginName, KN_SEARCH_SCAN_CPP, mz_Params)

		If ret = False Then
			If vbYes = MsgBox("El valor de retorno es de error." & vbCrLf & _
							  "¿Deseas cancelar las siguientes llamadas?", vbYesNo) Then 
				mb_PluginLoaded = False
			End If
		End If
	End If
End Sub

Public Sub gsub_exEndSearch()
	If mb_PluginLoaded Then
		IScript.ExecutePlugin ks_PluginName, KN_END_SCAN_CPP 
	End If
End Sub

''=======================================================================================
'' BUSCAR CIERTO TEXTO EN LOS ARCHIVOS
''_______________________________________________________________________________________
'Public Sub gsub_exPreSearch()
'	
'	mb_PluginLoaded = False
'	mz_Params(0) = ks_Word2Search
'	mz_Params(1) = kb_SearchCaseSensitive
'
'	mb_PluginLoaded = IScript.ExecutePlugin (ks_PluginName, KN_START_FIND_PHRASE, mz_Params)
'
'	If Not mb_PluginLoaded Then
'		MsgBox "No se pudo cargar correctamente el plugin", vbExclamation
'	Else
'		IScript.bolCancelReport = True
'	End If
'End Sub
'
'Public Sub gsub_exInSearch()
'	Dim ret
'	If mb_PluginLoaded Then
'
'		mz_Params(0) = IScript.bolInSearchIsDir
'		mz_Params(1) = IScript.strFilePath
'		mz_Params(2) = IScript.strFileName
'
'		ret = IScript.ExecutePlugin(ks_PluginName, KN_SEARCH_FIND_PHRASE, mz_Params)
'
'		If ret = False Then
'			If vbYes = MsgBox("El valor de retorno es de error." & vbCrLf & _
'							  "¿Deseas cancelar las siguientes llamadas?", vbYesNo) Then 
'				mb_PluginLoaded = False
'			End If
'		End If
'	End If
'End Sub
'
'Public Sub gsub_exEndSearch()
'	If mb_PluginLoaded Then
'		IScript.ExecutePlugin ks_PluginName, KN_END_SCAN_FILES 
'	End If
'End Sub

''=======================================================================================
'' ELIMINAR TODAS LAS LINEAS DE CIERTO TEXTO QUE ENCUENTRE
''_______________________________________________________________________________________
'Public Sub gsub_exPreSearch()
'	
'	mb_PluginLoaded = False
'	mz_Params(0) = ks_LineText2Delete
'
'	mb_PluginLoaded = IScript.ExecutePlugin (ks_PluginName, KN_START_DELETE_LINE_TEXT, mz_Params)
'
'	If Not mb_PluginLoaded Then
'		MsgBox "No se pudo cargar correctamente el plugin", vbExclamation
'	Else
'		IScript.bolCancelReport = True
'	End If
'End Sub
'
'Public Sub gsub_exInSearch()
'	Dim ret
'	If mb_PluginLoaded Then
'
'		mz_Params(0) = IScript.bolInSearchIsDir
'		mz_Params(1) = IScript.strFilePath
'		mz_Params(2) = IScript.strFileName
'
'		ret = IScript.ExecutePlugin(ks_PluginName, KN_SEARCH_DELETE_LINE_TEXT, mz_Params)
'
'		If ret = False Then
'			If vbYes = MsgBox("El valor de retorno es de error." & vbCrLf & _
'							  "¿Deseas cancelar las siguientes llamadas?", vbYesNo) Then 
'				mb_PluginLoaded = False
'			End If
'		End If
'	End If
'End Sub
'
'Public Sub gsub_exEndSearch()
'	If mb_PluginLoaded Then
'		IScript.ExecutePlugin ks_PluginName, KN_END_SCAN_FILES 
'	End If
'End Sub

''=======================================================================================
'' ELIMINAR UN BLOQUE DE TEXTO
''_______________________________________________________________________________________
'Public Sub gsub_exPreSearch()
'	
'	mb_PluginLoaded = False
'	mz_Params(0) = ks_StartText2Delete
'	mz_Params(1) = ks_EndText2Delete
'	mz_Params(2) = kn_MaxLines2Delete
'
'	mb_PluginLoaded = IScript.ExecutePlugin (ks_PluginName, KN_START_DELETE_BLOCK_TEXT, mz_Params)
'
'	If Not mb_PluginLoaded Then
'		MsgBox "No se pudo cargar correctamente el plugin", vbExclamation
'	Else
'		IScript.bolCancelReport = True
'	End If
'End Sub
'
'Public Sub gsub_exInSearch()
'	Dim ret
'	If mb_PluginLoaded Then
'
'		mz_Params(0) = IScript.bolInSearchIsDir
'		mz_Params(1) = IScript.strFilePath
'		mz_Params(2) = IScript.strFileName
'
'		ret = IScript.ExecutePlugin(ks_PluginName, KN_SEARCH_DELETE_BLOCK_TEXT, mz_Params)
'
'		If ret = False Then
'			If vbYes = MsgBox("El valor de retorno es de error." & vbCrLf & _
'							  "¿Deseas cancelar las siguientes llamadas?", vbYesNo) Then 
'				mb_PluginLoaded = False
'			End If
'		End If
'	End If
'End Sub
'

'Public Sub gsub_exEndSearch()
'	If mb_PluginLoaded Then
'		IScript.ExecutePlugin ks_PluginName, KN_END_SCAN_FILES 
'	End If
'End Sub
