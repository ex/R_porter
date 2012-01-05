VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGenre 
   Caption         =   "Género"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5190
   Icon            =   "Spanish.frx":0000
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
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Registros"
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
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
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
Option Explicit

Private m_BackColorSel As Long
Private m_FormWidth As Integer
Private m_FormHeight As Integer
Private m_FlexWidth As Integer
Private m_FlexHeight As Integer


'***************************************
' MENUS
'***************************************
Private Sub mnuActualizar_Click()
    Actualizar_Registros
End Sub

Private Sub mnuBorrar_Click()
    Eliminar_Registro
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

Private Sub mnuSalir_Click()
    Unload Me
End Sub

'***************************************
' FORMULARIO
'***************************************
Private Sub Form_Load()
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
    
    If 1 = MSFLexGrid.Col Then
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

    If ((KeyCode = vbKeyLeft) Or (KeyCode = vbKeyHome)) And (Me.flxResults.Col = 1) Then
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

Private Sub flxResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        If Me.flxResults.Row <> 0 And Me.flxResults.RowHeight(Me.flxResults.Row) <> 0 Then
            PopupMenu Me.mnuPopUp, 2
        End If
    Else
        Me.flxResults.BackColorSel = Me.flxResults.BackColor
    End If
    
End Sub

Private Sub flxResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.flxResults.BackColorSel = m_BackColorSel
End Sub

Private Sub flxResults_SelChange()
    
    'para evitar multiple seleccion de filas
    If (flxResults.RowSel - flxResults.Row) <> 0 Then
        flxResults.RowSel = flxResults.Row
    End If
    
    If (1 = flxResults.Col) And ((Me.flxResults.Cols - 1) <> Me.flxResults.ColSel) Then
        Me.flxResults.ColSel = Me.flxResults.Cols - 1
    Else
        If (1 <> flxResults.Col) Then
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
End Sub

Private Sub Eliminar_Registro()
End Sub

Private Sub Editar_Campo()
End Sub

Private Sub Agregar_Registro()
End Sub
