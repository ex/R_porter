VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGenre 
   Caption         =   "Género"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5190
   Icon            =   "frmGenre.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFlex 
      Height          =   285
      Left            =   1035
      TabIndex        =   1
      Top             =   285
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   4920
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   8678
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   19
      Cols            =   4
      BackColorFixed  =   16750143
      ForeColorFixed  =   16777215
      BackColorSel    =   16775910
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuActualizar 
         Caption         =   "Actualizar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Registros"
      Begin VB.Menu mnuOrdenar 
         Caption         =   "&Ordenar columna"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditar 
         Caption         =   "&Editar campo"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "E&liminar registro"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfirmEdit 
         Caption         =   "Pedir &confirmacion en edición"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuConfirmInsert 
         Caption         =   "Pedir con&firmacion en inserción"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSort 
         Caption         =   "&Ordenar columna"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Editar campo"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&Nuevo registro"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "E&liminar registro"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmGenre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************
' CONSTANTES
'***************************************

Private Const FLX_NUM_FORECOLOR = &HFFAA64


'***************************************
' VARIABLES DE MODULO
'***************************************

Private m_BackColorSel As Long
Private m_FormWidth As Integer
Private m_FormHeight As Integer
Private m_FlexWidth As Integer
Private m_FlexHeight As Integer


Private mb_ConfirmarNuevo As Boolean
Private mb_ConfirmarEdicion As Boolean

'***************************************
' MENUS
'***************************************
Private Sub mnuActualizar_Click()
    Actualizar_Registros
End Sub

Private Sub mnuBorrar_Click()
    Eliminar_Registro
End Sub

Private Sub mnuConfirmEdit_Click()
    If Me.mnuConfirmEdit.Checked = True Then
        mb_ConfirmarEdicion = False
        Me.mnuConfirmEdit.Checked = False
    Else
        mb_ConfirmarEdicion = True
        Me.mnuConfirmEdit.Checked = True
    End If
End Sub

Private Sub mnuConfirmInsert_Click()
    If Me.mnuConfirmInsert.Checked = True Then
        mb_ConfirmarEdicion = False
        Me.mnuConfirmInsert.Checked = False
    Else
        mb_ConfirmarNuevo = True
        Me.mnuConfirmInsert.Checked = True
    End If
End Sub

Private Sub mnuDelete_Click()
    mnuBorrar_Click
End Sub

Private Sub mnuEdit_Click()
    mnuEditar_Click
End Sub

Private Sub mnuEditar_Click()
    Editar_Campo
End Sub

Private Sub mnuExit_Click()
    mnuSalir_Click
End Sub

Private Sub mnuNew_Click()
    mnuNuevo_Click
End Sub

Private Sub mnuNuevo_Click()
    Agregar_Registro
End Sub

Private Sub mnuSort_Click()
    mnuOrdenar_Click
End Sub

Private Sub mnuOrdenar_Click()
Dim k As Integer
Dim oldRow As Long
Dim oldCol As Long

    With Me.flxResults
        
        oldRow = .Row
        oldCol = .Col
        
'        If 1 <> .Col Then
            
        If 0 <> .Col Then
        
            '************************************************
            'quitar fila para nuevo registro
            '************************************************
            .RemoveItem (.Rows - 1)
            
            If 0 = .ColData(.Col) Then
                .Sort = flexSortGenericDescending
                .ColData(.Col) = 1
            Else
                .Sort = flexSortGenericAscending
                .ColData(.Col) = 0
            End If
            
            .Redraw = False
            
'            .Col = 1
            .Col = 0
            For k = 1 To .Rows - 1
                .Row = k
                .Text = k
            Next
            
            '************************************************
            'agregar fila para nuevo registro
            '************************************************
            .Rows = .Rows + 1
            .Row = .Rows - 1
            'forzar visible
            .RowHeight(.Row) = -1
            
'            .Col = 1
'            .CellForeColor = FLX_NUM_FORECOLOR
'            .CellAlignment = flexAlignRightCenter
'            .Text = "[+]"
            
            .Col = 0
            .CellForeColor = FLX_NUM_FORECOLOR
            .CellAlignment = flexAlignRightCenter
            .Text = "[+]"
            
            .Redraw = True
            
            .Row = oldRow
            .Col = oldCol
        
        End If
        
    End With


End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

'***************************************
' FORMULARIO
'***************************************
Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim b_Redraw As Byte
Dim k As Integer

    On Error GoTo Handler

    Screen.MousePointer = vbHourglass
    
    m_FormWidth = Me.Width
    m_FormHeight = Me.Height
    
    mb_ConfirmarEdicion = True
    mb_ConfirmarNuevo = True
    
    With Me.flxResults
        
        m_BackColorSel = .BackColorSel
        m_FlexWidth = .Width
        m_FlexHeight = .Height
        
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        
        'extension
        .Rows = 2
        .FixedRows = 1
        .Cols = 3
        .FixedCols = 0
        
        'ancho celdas
'        .ColWidth(0) = 0
'        .ColWidth(1) = 650
'        .ColWidth(2) = 3360
'        .ColWidth(3) = 780
        
        .ColWidth(0) = 650
        .ColWidth(1) = 3360
        .ColWidth(2) = 780
        
        
        'titulos columna
        .Row = 0
'        .Col = 1
'        .CellFontBold = True
'        .CellAlignment = flexAlignCenterCenter
'        .Text = "Nº"
'        .Col = 2
'        .CellFontBold = True
'        .CellAlignment = flexAlignCenterCenter
'        .Text = "Género"
'        .Col = 3
'        .CellFontBold = True
'        .CellAlignment = flexAlignCenterCenter
'        .Text = "Activo"
        
        .Col = 0
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .Text = "Nº"
        .Col = 1
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .Text = "Género"
        .Col = 2
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .Text = "Activo"
        
        Actualizar_Registros
                
    End With
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

Handler:
    
    If Err.Number = 0 Then
        '
    Else
        MsgBox Err.Description, vbCritical, "Form_Load()"
        Me.flxResults.Redraw = True
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Form_Resize()

    On Error GoTo Handler

    If Me.Width < 3000 Then
        Me.Width = 3000
    End If

    Me.flxResults.Width = Me.Width - m_FormWidth + m_FlexWidth
    
    If Me.Height < 3000 Then
        Me.Height = 3000
    End If
    
    Me.flxResults.Height = Me.Height - m_FormHeight + m_FlexHeight
    
Handler:

End Sub

'***************************************
' FLEXEDIT
'***************************************
Sub FlexGridEdit(MSFLexGrid As MSHFlexGrid, Edt As TextBox, KeyAscii As Integer)
    
    On Error GoTo Handler
    
    If 0 = MSFLexGrid.Col Then
        Exit Sub
    End If
    
    Select Case KeyAscii
    
        Case 0 To 12, 14 To 31
            Exit Sub
        Case 13, 32
            'teclea [ESP] o [ENTER] para modificar el texto actual.
            Edt.Text = MSFLexGrid.Text
        Case Else
            'cualquier otro carácter reemplaza el texto actual.
            Edt.Text = Chr(KeyAscii)
            Edt.SelStart = 1
    
    End Select
    
    Edt.Move MSFLexGrid.Left + MSFLexGrid.CellLeft, MSFLexGrid.Top + MSFLexGrid.CellTop, MSFLexGrid.CellWidth, MSFLexGrid.CellHeight
    Edt.Visible = True
    'do it [SelStar] after visible!
    Edt.SelStart = Len(Edt.Text)
    Edt.SetFocus
    
    Exit Sub
    
Handler:
    
    MsgBox Err.Description, vbCritical, "FlexGridEdit()"
End Sub

Sub EditKeyCode(MSFLexGrid As MSHFlexGrid, Edt As TextBox, KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Handler
    
    Select Case KeyCode
    
        Case 27
            '-------------------------------------------------
            ' ESC: ocultar, devuelve el enfoque a MSFlexGrid.
            '-------------------------------------------------
            MSFLexGrid.SetFocus
            
        Case 13
            '-------------------------------------------------
            ' ENTRAR acepta la entrada
            '-------------------------------------------------
            MSFLexGrid.Text = Edt.Text
            MSFLexGrid.SetFocus
            
        Case vbKeyUp
            '-------------------------------------------------
            ' Arriba.
            '-------------------------------------------------
            MSFLexGrid.SetFocus
            DoEvents
            If MSFLexGrid.Row > MSFLexGrid.FixedRows Then
                MSFLexGrid.Row = MSFLexGrid.Row - 1
            End If

        Case vbKeyDown
            '-------------------------------------------------
            ' Abajo.
            '-------------------------------------------------
            MSFLexGrid.SetFocus
            DoEvents
            If MSFLexGrid.Row < MSFLexGrid.Rows - 1 Then
                MSFLexGrid.Row = MSFLexGrid.Row + 1
            End If
        
    End Select
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "EditKeycode()"
End Sub

Private Sub flxResults_DblClick()
    'doble click para editar la celda actual
    FlexGridEdit flxResults, txtFlex, 32
End Sub

Private Sub flxResults_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = vbKeyUp) Or (KeyCode = vbKeyPageUp)) And (Me.flxResults.Row = 1) Then
        flxResults_SelChange
    End If

    If ((KeyCode = vbKeyDown) Or (KeyCode = vbKeyPageDown)) And (Me.flxResults.Row = Me.flxResults.Rows - 1) Then
        flxResults_SelChange
    End If

'    If ((KeyCode = vbKeyLeft) Or (KeyCode = vbKeyHome)) And (Me.flxResults.Col = 1) Then
    If ((KeyCode = vbKeyLeft) Or (KeyCode = vbKeyHome)) And (Me.flxResults.Col = 0) Then
        flxResults_SelChange
    End If

End Sub

Private Sub flxResults_KeyPress(KeyAscii As Integer)
    FlexGridEdit flxResults, txtFlex, KeyAscii
End Sub

Private Sub flxResults_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            Eliminar_Registro
    End Select
End Sub

Private Sub flxResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        If Me.flxResults.Row <> 0 And Me.flxResults.RowHeight(Me.flxResults.Row) <> 0 Then
            PopupMenu Me.mnupopup, 2
        End If
    Else
        Me.flxResults.BackColorSel = Me.flxResults.BackColor
    End If
    
End Sub

Private Sub flxResults_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.flxResults.BackColorSel = m_BackColorSel
End Sub

Private Sub flxResults_SelChange()
    
    'para evitar multiple seleccion de filas
    If (flxResults.RowSel - flxResults.Row) <> 0 Then
        flxResults.RowSel = flxResults.Row
    End If
    
'    If (1 = flxResults.Col) And ((Me.flxResults.Cols - 1) <> Me.flxResults.ColSel) Then
'        Me.flxResults.ColSel = Me.flxResults.Cols - 1
'    Else
'        If (1 <> flxResults.Col) Then
'            'para evitar multiple seleccion de columnas
'            If (flxResults.ColSel - flxResults.Col) <> 0 Then
'                flxResults.ColSel = flxResults.Col
'            End If
'        End If
'    End If
    
    If (0 = flxResults.Col) And ((Me.flxResults.Cols - 1) <> Me.flxResults.ColSel) Then
        Me.flxResults.ColSel = Me.flxResults.Cols - 1
    Else
        If (0 <> flxResults.Col) Then
            'para evitar multiple seleccion de columnas
            If (flxResults.ColSel - flxResults.Col) <> 0 Then
                flxResults.ColSel = flxResults.Col
            End If
        End If
    End If
    
    
End Sub

Private Sub flxResults_Scroll()
    Me.txtFlex.Visible = False
End Sub

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode flxResults, txtFlex, KeyCode, Shift
End Sub

Private Sub txtFlex_LostFocus()
    txtFlex.Visible = False
End Sub

'***************************************
' MANEJO DE REGISTROS
'***************************************
Private Sub Actualizar_Registros()
Dim rs As ADODB.Recordset
Dim b_Redraw As Byte
Dim k As Integer

    On Error GoTo Handler

    With Me.flxResults
        
        .Rows = 2
        .RowHeight(1) = 0
        
        query = "SELECT * FROM genre WHERE (id_genre > 0) ORDER BY genre"
        Set rs = New ADODB.Recordset
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

        
        If rs.EOF = True Then
            
            rs.Close
            Exit Sub
            
        Else
            
            .Redraw = False
            
            k = 0
            b_Redraw = 1
            
            While rs.EOF = False
               
                k = k + 1
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                'forzar visible
                .RowHeight(.Row) = -1
                
                .RowData(.Row) = rs!id_genre
'                .Col = 1
'                .CellForeColor = FLX_NUM_FORECOLOR
'                .Text = k
'                .Col = 2
'                .CellAlignment = flexAlignLeftCenter
'                .Text = rs!genre
'                .Col = 3
'                .Text = rs!active
                
                .Col = 0
                .CellForeColor = FLX_NUM_FORECOLOR
                .Text = k
                .Col = 1
                .CellAlignment = flexAlignLeftCenter
                .Text = rs!genre
                .Col = 2
                .Text = rs!active
                
                rs.MoveNext
                
                If b_Redraw = 1 Then
                    If k >= CInt(((.Height - .RowHeight(0)) / .RowHeight(.Row))) + 1 Then
                        .Redraw = True
                        .Refresh
                        .Redraw = False
                        b_Redraw = 0
                    End If
                End If
               
            Wend
            
            rs.Close
            
            '************************************************
            'agregar fila para nuevo registro
            '************************************************
            .Rows = .Rows + 1
            .Row = .Rows - 1
            'forzar visible
            .RowHeight(.Row) = -1
            
'            .Col = 1
'            .CellForeColor = FLX_NUM_FORECOLOR
'            .CellAlignment = flexAlignRightCenter
'            .Text = "[+]"
            
            .Col = 0
            .CellForeColor = FLX_NUM_FORECOLOR
            .CellAlignment = flexAlignRightCenter
            .Text = "[+]"
            
            
            '************************************************
            'eliminar la primera fila invisible
            '************************************************
            .RemoveItem (1)
            
            
            'seleccionar el primero
            .Row = 1
            .ColSel = .Cols - 1
            .Redraw = True
            
        End If
        
    End With
    
    Exit Sub

Handler:
    
    If Err.Number = 0 Then
        '
    Else
        MsgBox Err.Description, vbCritical, "Actualizar_Registros()"
        Me.flxResults.Redraw = True
    End If
    
End Sub

Private Sub Eliminar_Registro()
Dim id_registro As Long
Dim fila_eliminada As Long
Dim nombre_fila As String
Dim cd As ADODB.Command
Dim oldRow As Long

    On Error GoTo Handler

    'eliminar registro seleccionado
    With Me.flxResults
            
        If (.Row <> 0) And (.RowHeight(.Row) <> 0) Then
    
            .Redraw = False
            
            oldRow = .TopRow
            
            id_registro = .RowData(.Row)
'            .Col = 1
'            fila_eliminada = .Text
'            .Col = 2
'            nombre_fila = .Text
'
'            .Col = 1
'            .ColSel = .Cols - 1
            
            .Col = 0
            fila_eliminada = .Text
            .Col = 1
            nombre_fila = .Text
            
            .Col = 0
            .ColSel = .Cols - 1
            
            .Redraw = True
            
            'don't remove
            .TopRow = oldRow
            
            If vbYes = MsgBox("Estas seguro de eliminar el registro [" & Trim(str(fila_eliminada)) & "]:" & vbCrLf & nombre_fila & "?", vbExclamation + vbYesNo, "Confirmar eliminacion") Then
        
                Set cd = New ADODB.Command
                Set cd.ActiveConnection = cn
                
                cd.CommandText = "DELETE FROM genre WHERE id_genre=" & id_registro
                cd.Execute
                
                If .Rows > 2 Then
                    oldRow = .Row
                    .RemoveItem (.Row)
                    
                    'seleccionar la fila anterior
                    If oldRow > .Rows - 1 Then
                        .Row = oldRow - 1
                    Else
                        .Row = oldRow
                    End If
                Else
                    'volver invisible la primera fila
                    .RowHeight(1) = 0
                    Exit Sub
                End If
                
'                .Col = 1
'                .ColSel = .Cols - 1
                
                .Col = 0
                .ColSel = .Cols - 1
                
                
            End If
                
            'verificar si la fila esta visible
            If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.Height - .RowHeight(0)) / .RowHeight(.Row)))) Then
                .TopRow = .Row
            End If

            .SetFocus
        
        End If
            
    End With
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "Eliminar_Registro()"
    Me.flxResults.Redraw = True
    
End Sub

Private Sub Editar_Campo()
Dim id_registro As Long
Dim col_editada As Long
Dim valor_campo As String
Dim cd As ADODB.Command
Dim oldRow As Long

    On Error GoTo Handler

    'editar campo del registro seleccionado
    With Me.flxResults
            
        If (.Row <> 0) And (.RowHeight(.Row) <> 0) And (.Col > 1) Then
    
            .Redraw = False
            
            .Col = 0
            id_registro = CLng(.Text)
'            .Col = 1
'            fila_eliminada = .Text
'            .Col = 2
'            nombre_fila = .Text
'
'            .Col = 1
'            .ColSel = .Cols - 1
'
            .Redraw = True
'
'            If vbYes = MsgBox("Estas seguro de eliminar el registro [" & Trim(str(fila_eliminada)) & "]:" & vbCrLf & nombre_fila & "?", vbExclamation + vbYesNo, "Confirmar eliminacion") Then
'
'                Set cd = New ADODB.Command
'                Set cd.ActiveConnection = cn
'
'                cd.CommandText = "DELETE FROM genre WHERE id_genre=" & id_registro
'                cd.Execute
'
'                If .Rows > 2 Then
'                    oldRow = .Row
'                    .RemoveItem (.Row)
'
'                    'seleccionar la fila anterior
'                    If oldRow > .Rows - 1 Then
'                        .Row = oldRow - 1
'                    Else
'                        .Row = oldRow
'                    End If
'                Else
'                    'volver invisible la primera fila
'                    .RowHeight(1) = 0
'                    Exit Sub
'                End If
'
'                .Col = 1
'                .ColSel = .Cols - 1
'
'            End If
                
            .SetFocus
        
        End If
            
    End With
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "Eliminar_Registro()"
    Me.flxResults.Redraw = True
    
End Sub

Private Sub Agregar_Registro()
Dim scampo As String

    On Error GoTo Handler

    With Me.flxResults
    
        .Redraw = False
        
'        .Row = .Rows - 1
'        .Col = 2
'        scampo = .Text
'
'        .Col = 1
'        .ColSel = .Cols - 1
        
        .Row = .Rows - 1
        .Col = 1
        scampo = .Text
        
        .Col = 0
        .ColSel = .Cols - 1
        
        .Redraw = True
        
        .TopRow = .Row
        
        'confirmacion
        If mb_ConfirmarNuevo = True Then
            If vbNo = MsgBox("¿Deseas ingresar el nuevo registro:" & vbCrLf & "( " & scampo & " )?", vbExclamation + vbYesNo, "Confirmar inserción") Then
                Exit Sub
            End If
        End If
        
        'insercion
        
    End With
    
    Exit Sub

Handler:

    If Err.Number = 0 Then
    '
    Else
        MsgBox Err.Description, vbCritical, "Agregar_Registro()"
    End If

End Sub
