'--------------------------------------------------------
' EJEMPLO BASICO PARA USAR SCRIPTS CON R_PORTER
' Esau Rodriguez O.
'--------------------------------------------------------
Option Explicit

'========================================================
' Interfaz publica para el script --> IScript (Objeto)
'--------------------------------------------------------
' Public strSearchPath As String		
' Public strFilePath As String
' Public strFileName As String
' Public rchtxtResults() As Object		' EL objeto control RichTextBox del formulario de exploracion de R_porter
' Public bolPreSearchByDir As Boolean	' VERDADERO si la busqueda es por Directorios, FALSO si es por unidades
' Public bolInSearchIsDir As Boolean	' VERDADERO si el archivo encontrado es un directorio, FALSO de lo contrario
' Public bolCancelReport As Boolean	' VERDADERO cancelara el reporte original del programa, FALSO reporte normal del programa
'--------------------------------------------------------

Private mn_NumFiles 
Private mn_NumDirs
Private mn_NumTotal

'--------------------------------------------------------
' Esta funcion es llamada antes de iniciar la busqueda
Public Sub gsub_exPreSearch()

	IScript.rchtxtResults.selcolor = RGB(255,0,0)
	IScript.rchtxtResults.SelText = vbTab & vbTab & "---  INICIO SCRIPT  ---" & vbCrLf
	mn_NumFiles = 0
	mn_NumDirs = 0
	mn_NumTotal = 0
	If vbYes = MsgBox ("Quieres cancelar el reporte por defecto?",vbYesNo) Then
		IScript.bolCancelReport = True
	End If
End Sub

'--------------------------------------------------------
' Esta funcion es llamada durante la busqueda
Public Sub gsub_exInSearch()
	If IScript.bolInSearchIsDir = False Then
		mn_NumFiles = mn_NumFiles + 1
		IScript.rchtxtResults.Selcolor = RGB(255,0,0)
		IScript.rchtxtResults.SelText = mn_NumFiles & ") " & IScript.strFilePath & "\" & IScript.strFileName & vbCrLf
	Else
		mn_NumDirs = mn_NumDirs + 1
		IScript.rchtxtResults.SelBold = True
		IScript.rchtxtResults.Selcolor = RGB(255,0,0)
		IScript.rchtxtResults.SelText = mn_NumDirs & ") " & IScript.strFilePath & "\" & IScript.strFileName & vbCrLf
		IScript.rchtxtResults.SelBold = False
	End If
	mn_NumTotal = mn_NumTotal + 1
End Sub

'--------------------------------------------------------
' Esta funcion es llamada al finalizar la busqueda
Public Sub gsub_exEndSearch()
	IScript.rchtxtResults.selcolor = RGB(255,0,0)
	IScript.rchtxtResults.SelText = "Total Archivos: " & mn_NumFiles  & vbCrLf
	IScript.rchtxtResults.selcolor = RGB(255,0,0)
	IScript.rchtxtResults.SelText = "Total Directorios: " & mn_NumDirs & vbCrLf
	IScript.rchtxtResults.selcolor = RGB(255,0,0)
	IScript.rchtxtResults.SelText = vbTab & vbTab & "---   FIN SCRIPT  ---" & vbCrLf
End Sub
