Attribute VB_Name = "mdlFlexExport"
Option Explicit

'--------------------------------------------------------------------------------------------
'   Muestra el cuadro [Guardar como...] para la funcion de exportacion del contenido de un
'   flexgrid a un archivo de texto o Excel
'
Public Sub gsub_FlxShowSaveAsDialog(cmmdlg As CommonDialog, flxResults As MSHFlexGrid, Optional ByVal sql_str As String = "")
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Exportar los resultados como:"
        .Filter = "Archivos de texto(*.txt)|*.txt|Archivos Excel(*.xls)|*.xls|Todos los Archivos(*.*)|*.*"
        'necesario para controlar la extension con que se salvaran los archivos
        'sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .DefaultExt = ""
        .InitDir = App.Path
        'tipo predefinido TXT
        .FilterIndex = 1
        'nombre del reporte inicial
        .filename = "SQL_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now)
        .ShowSave
        If .filename <> "" Then
            '.FilterIndex devuelve la extension seleccionada en el cuadro guardar como
            If .FilterIndex = 1 Then
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea TXT
                If UCase(Right(.filename, 4)) <> ".TXT" Then
                    .filename = .filename & ".txt"
                End If
                '---------------------------------------------
                ' salvar con formato texto
                gfnc_Flx2Txt .filename, flxResults
            Else
            'en otro caso guardar como excel
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea XLS
                If UCase(Right(.filename, 4)) <> ".XLS" Then
                    .filename = .filename & ".xls"
                End If
                '---------------------------------------------
                ' salvar con formato excel
                gfnc_Flx2Xls .filename, flxResults, sql_str
            End If
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "gsub_FlxShowSaveAsDialog()"
    End If

End Sub


'--------------------------------------------------------------------------------------------
'   Exporta el contenido de un flexgrid a un archivo de texto cuyo nombre es pasado como
'   como parametro. Aplica formato, elimina saltos de linea y separa por tabuladores.
'   NOTA: No guarda en el archivo los datos de columnas de ancho cero.
'
Public Function gfnc_Flx2Txt(ByVal filename As String, flxResults As MSHFlexGrid) As Boolean
    '===================================================
    Dim k As Long
    Dim p As Long
    Dim num_rows As Long
    Dim num_cols As Long
    Dim str_put As String
    Dim str_cell As String
    Dim b_progress As Boolean
    Dim b_fileopen As Boolean
    Dim step As Integer
    Dim zLen() As Long
    '===================================================

    On Error GoTo Handler:
    
    num_rows = flxResults.Rows
    num_cols = flxResults.Cols
    
    If (num_cols < 1) Or (num_rows < 1) Then
        gfnc_Flx2Txt = False
        Exit Function
    End If
    
    b_progress = False
    b_fileopen = False
    
    Screen.MousePointer = vbHourglass
    flxResults.MousePointer = flexHourglass
    
    flxResults.Redraw = False
    flxResults.Row = 0
    
    Open filename For Output As #1
    b_fileopen = True
    
    '----------------------------------------------
    ' crear progress bar
    '
    b_progress = True
    frmProgress.mb_Active = True
    frmProgress.lblMessage.Caption = "Espere mientras el programa exporta los registros a texto."
    frmProgress.Icon = frmProgress.ImageList.ListImages.Item("icoToText").Picture
    
    step = mfnc_GetStep(num_rows)
    
    If num_rows < 32767 Then
        frmProgress.ProgressBar.Max = num_rows
    Else
        frmProgress.ProgressBar.Max = 32767
    End If
    frmProgress.Visible = True
    '----------------------------------------------
    
    '----------------------------------------------
    ' buscar longitud maxima del texto por columna
    '
    ReDim zLen(1 To num_cols)
    
    For k = 0 To num_rows - 1

        flxResults.Row = k
        
        For p = 0 To num_cols - 1
        
            flxResults.Col = p
        
            If zLen(p + 1) < Len(flxResults.text) Then
                zLen(p + 1) = Len(flxResults.text)
            End If
        Next p
        
        If (k Mod step) = 0 Then
            DoEvents
            If frmProgress.Visible Then
                frmProgress.ProgressBar.value = (k Mod 32767)
                frmProgress.Refresh
            Else
                If vbYes = MsgBox("¿Estás seguro de cancelar?", vbExclamation + vbYesNo, "Cancelar") Then
                    GoTo FINISH_EXPORT
                Else
                    frmProgress.Visible = True
                    frmProgress.ProgressBar.value = (k Mod 32767)
                    frmProgress.Refresh
                End If
            End If
        End If
        
    Next k
    
    '----------------------------------------------
    ' grabar en archivo
    '
    str_put = ""
    flxResults.Row = 0
    For p = 0 To num_cols - 1
        flxResults.Col = p
        If flxResults.ColWidth(flxResults.Col) > 0 Then
            str_cell = Space$(zLen(p + 1) - Len(flxResults.text))
            str_put = str_put & gfnc_ParseChar(flxResults.text, vbCrLf, " ") & str_cell
            If p < (num_cols - 1) Then
                str_put = str_put & vbTab
            End If
        End If
    Next p
    Print #1, str_put
    
    str_put = ""
    For p = 0 To num_cols - 1
        flxResults.Col = p
        If flxResults.ColWidth(flxResults.Col) > 0 Then
            str_cell = String$(zLen(p + 1), "-")
            str_put = str_put & str_cell
            If p < (num_cols - 1) Then
                str_put = str_put & vbTab
            End If
        End If
    Next p
    Print #1, str_put

    For k = 1 To num_rows - 1

        flxResults.Row = k
        str_put = ""
        
        If flxResults.RowHeight(flxResults.Row) > 0 Then
            
            For p = 0 To num_cols - 1
                flxResults.Col = p
                If flxResults.ColWidth(flxResults.Col) > 0 Then
                    str_cell = Space$(zLen(p + 1) - Len(flxResults.text))
                    str_put = str_put & gfnc_ParseChar(flxResults.text, vbCrLf, " ") & str_cell
                    If p < (num_cols - 1) Then
                        str_put = str_put & vbTab
                    End If
                End If
            Next p
            Print #1, str_put
            
        End If
        
        If (k Mod step) = 0 Then
            DoEvents
            If frmProgress.Visible Then
                frmProgress.ProgressBar.value = (k Mod 32767)
                frmProgress.Refresh
            Else
                If vbYes = MsgBox("¿Estás seguro de cancelar?", vbExclamation + vbYesNo, "Cancelar") Then
                    Exit For
                Else
                    frmProgress.Visible = True
                    frmProgress.ProgressBar.value = (k Mod 32767)
                    frmProgress.Refresh
                End If
            End If
        End If
        
    Next k
    
FINISH_EXPORT:
    
    frmProgress.mb_Active = False
    Unload frmProgress
    
    flxResults.Row = 1
    flxResults.Col = 0
    flxResults.ColSel = flxResults.Cols - 1
    flxResults.Redraw = True
    
    ' salvar archivo
    Close #1
    b_fileopen = False
    
    Screen.MousePointer = vbDefault
    flxResults.MousePointer = vbDefault
    
    If vbYes = MsgBox("El reporte se ha guardado como:" & vbCrLf & filename & vbCrLf & "¿Deseas abrirlo?", vbInformation + vbYesNo, "Reporte guardado") Then
        ' abrir archivo texto
        ShellExecute vbNull, "open", filename, vbNull, vbNull, SW_NORMAL
    End If
        
    gfnc_Flx2Txt = True
    
    Exit Function
    
Handler:

    Select Case Err.Number
        Case 76
            '-----------------------------------------
            ' no se pudo abrir archivo
            MsgBox "No se pudo encontrar la ruta del archivo o el" & vbCrLf & "archivo esta siendo usado por otra aplicación." & vbCrLf & "Verique la ruta y vuelva a intentarlo.", vbExclamation, "Error"
        
        Case Else
            MsgBox Err.Description, vbExclamation, "gfnc_Flx2Txt()"
    End Select

    If b_progress Then
        frmProgress.mb_Active = False
        Unload frmProgress
    End If
    
    If b_fileopen Then
        Close #1
    End If
    
    Screen.MousePointer = vbDefault
    flxResults.MousePointer = vbDefault
    
    flxResults.Row = 1
    flxResults.Col = 0
    flxResults.ColSel = flxResults.Cols - 1
    flxResults.Redraw = True
    
    gfnc_Flx2Txt = False

End Function

'--------------------------------------------------------------------------------------------
'   Exporta el contenido de un flexgrid a un archivo excel cuyo nombre es pasado como
'   como parametro. Elimina saltos de linea.
'   NOTA: No guarda en el archivo los datos de columnas de ancho cero.
'
Public Function gfnc_Flx2Xls(ByVal filename As String, flxResults As MSHFlexGrid, Optional sql_str As String = "") As Boolean
    '===================================================
    Dim xlObject As Object
    Dim wrkBook As Object
    Dim celRange As Object
    Dim num_rows As Long
    Dim num_cols As Long
    Dim step As Integer
    Dim nMin As Integer
    Dim k As Long
    Dim p As Long
    Dim b_progress As Boolean
    '===================================================

    On Error GoTo Handler:
    
    num_rows = flxResults.Rows
    num_cols = flxResults.Cols
    
    If (num_cols < 1) Or (num_rows < 1) Then
        gfnc_Flx2Xls = False
        Exit Function
    End If
    
    b_progress = False
    
    Screen.MousePointer = vbHourglass
    flxResults.MousePointer = flexHourglass
    
    Set xlObject = CreateObject("Excel.Application")
    Set wrkBook = xlObject.Workbooks.Add
    
    With xlObject
        
        '----------------------------------------------
        ' crear progress bar
        '
        b_progress = True
        frmProgress.mb_Active = True
        frmProgress.lblMessage.Caption = "Espere mientras el programa exporta los registros a Excel."
        frmProgress.Icon = frmProgress.ImageList.ListImages.Item("icoToExcel").Picture
        
        step = mfnc_GetStep(num_rows)
        
        If num_rows < 32767 Then
            frmProgress.ProgressBar.Max = num_rows
        Else
            frmProgress.ProgressBar.Max = 32767
        End If
        frmProgress.Visible = True
        '----------------------------------------------
        
        If Trim(sql_str) <> "" Then
            .Range("A1").Select
            .ActiveCell.FormulaR1C1 = "SQL"
            .ActiveCell.Font.Color = RGB(100, 170, 255)
                
            .Range("B1").Select
            .ActiveCell.FormulaR1C1 = Trim(sql_str)
            .ActiveCell.Font.Color = RGB(130, 130, 130)
            
            nMin = 3
        Else
            nMin = 1
        End If
        
        flxResults.Redraw = False
        
        flxResults.Row = 0
        
        For p = 0 To num_cols - 1
        
            flxResults.Col = p
            
            If flxResults.ColWidth(flxResults.Col) > 0 Then
                Set celRange = .Cells.Item(nMin, p + 1)
                celRange.Select
                .ActiveCell.FormulaR1C1 = flxResults.text
                .ActiveCell.Font.Bold = True
                .ActiveCell.Font.Color = RGB(255, 255, 255)
                .ActiveCell.Interior.Color = RGB(63, 150, 255)
            End If
        
        Next p
        
        
        For k = 1 To num_rows - 1

            flxResults.Row = k
            
            If flxResults.RowHeight(flxResults.Row) > 0 Then
                
                For p = 0 To num_cols - 1
                
                    flxResults.Col = p
                    
                    If flxResults.ColWidth(flxResults.Col) > 0 Then
                        '----------------------------------------------
                        ' me fue dificil descubrir como usar el objeto
                        ' Range... :-(
                        '
                        Set celRange = .Cells.Item(k + nMin, p + 1)
                        celRange.Select
                        .ActiveCell.FormulaR1C1 = gfnc_ParseChar(flxResults.text, vbCrLf, " ")
                        
                        If 0 = p Then
                            .ActiveCell.Font.Color = RGB(100, 170, 255)
                        End If
                    End If
                
                Next p
                
            End If
            
            If (k Mod step) = 0 Then
                DoEvents
                If frmProgress.Visible Then
                    frmProgress.ProgressBar.value = (k Mod 32767)
                    frmProgress.Refresh
                Else
                    If vbYes = MsgBox("¿Estás seguro de cancelar?", vbExclamation + vbYesNo, "Cancelar") Then
                        Exit For
                    Else
                        frmProgress.Visible = True
                        frmProgress.ProgressBar.value = (k Mod 32767)
                        frmProgress.Refresh
                    End If
                End If
            End If
            
        Next k
        
        '-----------------------------------------
        ' todas las celdas con fuente de tamaño 8
        '
        .Cells.Font.size = 8
        
        .Range("A1").Select
    
    End With
    
    frmProgress.mb_Active = False
    Unload frmProgress
    
    flxResults.Row = 1
    flxResults.Col = 0
    flxResults.ColSel = flxResults.Cols - 1
    flxResults.Redraw = True
    b_progress = False
    
    wrkBook.SaveAs filename
    
    xlObject.Visible = True
    
    Screen.MousePointer = vbDefault
    flxResults.MousePointer = vbDefault
    
    If vbYes = MsgBox("El reporte se ha guardado como:" & vbCrLf & filename & vbCrLf & "¿Deseas cerrar Excel?", vbInformation + vbYesNo, "Reporte guardado") Then
        
        'cerrar libro excel
        wrkBook.Close
        Set wrkBook = Nothing
        
        If xlObject.Workbooks.Count = 0 Then
            xlObject.Quit
            Set xlObject = Nothing
        End If
        
    End If
        
    gfnc_Flx2Xls = True
        
    Exit Function
    
Handler:

    Select Case Err.Number
        Case 429
            '-----------------------------------------
            ' no se pudo crear objeto excel
            MsgBox "Para poder exportar en este formato" & vbCrLf & "necesitas tener instalado Excel.", vbExclamation, "Excel no encontrado"
        
        Case 1004
            '-----------------------------------------
            ' error enviado por objeto excel
            MsgBox "No se pudo encontrar la ruta del archivo o el" & vbCrLf & "archivo esta siendo usado por otra aplicación" & vbCrLf & "o sucedió algun error interno de Excel." & vbCrLf & "Verique la ruta y vuelva a intentarlo.", vbExclamation, "Error Excel"
            ' para evitar que excel pregunte si guarda archivo
            xlObject.DisplayAlerts = False
            xlObject.Quit
            Set xlObject = Nothing
        
        Case 91, 462, -2147417848
            '-----------------------------------------
            ' se cerro la aplicacion excel o el libro
            xlObject.Quit
            Set xlObject = Nothing
        
        Case -2147023170
            MsgBox "No se puede cerrar Excel cuando se esta en vista de impresion" & vbCrLf & "Evita dejarlo en ese estado", vbExclamation, "Error"
        
        Case Else
            ' algun error jodido
            xlObject.DisplayAlerts = False
            xlObject.Quit
            Set xlObject = Nothing
            MsgBox Err.Description, vbExclamation, "gfnc_Flx2Xls()"
    End Select

    If b_progress Then
        frmProgress.mb_Active = False
        Unload frmProgress
    End If
    
    Screen.MousePointer = vbDefault
    flxResults.MousePointer = vbDefault
    
    flxResults.Row = 1
    flxResults.Col = 0
    flxResults.ColSel = flxResults.Cols - 1
    flxResults.Redraw = True

    gfnc_Flx2Xls = False

End Function

'--------------------------------------------------------------------------------------------
'   Calcular paso de orden 2 logaritmico
'   21 <    = 1
'   100     = 5
'   1000    = 15
'   10000   = 30
'   100000  = 50
'
Private Function mfnc_GetStep(ByVal num_rows As Long) As Long
    Dim fLog As Double
    If num_rows <= 21 Then
        mfnc_GetStep = 1
    Else
        fLog = Log(num_rows) / Log(10) - 1#
        mfnc_GetStep = CLng((1 + fLog) * fLog * 2.5) ' a little math...
    End If
End Function

Public Function gfnc_ParseChar(ByVal s_in As String, ByVal s_chr As String, ByVal s_rmp As String) As String

    Dim sz As String
    Dim st As String
    Dim n As Integer
    On Error GoTo Handler
    
    sz = s_in
    st = ""
    
    ' buscar s_char
    Do
        n = InStrRev(sz, s_chr)
        If n = 0 Then
            st = sz & st
            Exit Do
        Else
            ' remplazar por s_rmp
            If s_chr = vbCrLf Then
                st = s_rmp & Mid(sz, n + 2) & st    ' son dos caracteres
            Else
                st = s_rmp & Mid(sz, n + 1) & st
            End If
            sz = Mid(sz, 1, n - 1)
        End If
    Loop
    
    gfnc_ParseChar = st
    Exit Function
    
Handler: gfnc_ParseChar = ""
End Function

'--------------------------------------------------------------------------------------------
'   Editing FlexGrid with textbox
Public Sub gsub_startEditFlexGrid(MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 0 To 32
            ' Un espacio significa modificar el texto actual.
            Edt.text = MSHFlexGrid.text
            Edt.SelStart = 0
            Edt.SelLength = Len(Edt.text)
        Case Else
            ' Otro carácter reemplaza el texto actual.
            Edt = chr(KeyAscii)
            Edt.SelStart = 1
    End Select
    
    Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
        MSHFlexGrid.Top + MSHFlexGrid.CellTop - 30, _
        MSHFlexGrid.CellWidth, MSHFlexGrid.CellHeight
    Edt.Visible = True
    Edt.SetFocus
End Sub

Public Sub gsub_editFlexGrid(MSHFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case 27     ' ESC
        Edt.Visible = False
        MSHFlexGrid.SetFocus
    Case 13     ' RETURN
        MSHFlexGrid.SetFocus
    Case 38     ' UP
        MSHFlexGrid.SetFocus
        DoEvents
        If MSHFlexGrid.Row > MSHFlexGrid.FixedRows Then
            MSHFlexGrid.Row = MSHFlexGrid.Row - 1
            MSHFlexGrid.ColSel = MSHFlexGrid.Cols - 1
        End If
    Case 40     ' DOWN
        MSHFlexGrid.SetFocus
        DoEvents
        If MSHFlexGrid.Row < MSHFlexGrid.Rows - 1 Then
            MSHFlexGrid.Row = MSHFlexGrid.Row + 1
            MSHFlexGrid.ColSel = MSHFlexGrid.Cols - 1
        End If
    End Select
End Sub


