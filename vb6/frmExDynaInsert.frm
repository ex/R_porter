VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmExDynaInsert 
   ClientHeight    =   3855
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5055
   Icon            =   "frmExDynaInsert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   990
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16750143
      ForeColorFixed  =   16777215
      BackColorSel    =   16775910
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuRegistro 
      Caption         =   "&Registro"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Editar valor"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Insertar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancelar"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmExDynaInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public gs_SQL As String
Public go_parent As clsExDynaTable

Private m_FormWidth As Integer
Private m_FormHeight As Integer
Private m_FlexWidth As Integer
Private m_FlexHeight As Integer

Private Sub flxResults_DblClick()
    startEditingFlexgrid 32 ' Simula un espacio
End Sub

Private Sub flxResults_KeyPress(KeyAscii As Integer)
    startEditingFlexgrid KeyAscii
End Sub

Private Sub mnuEdit_Click()
    startEditingFlexgrid 32 ' Simula un espacio
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    gsub_editFlexGrid flxResults, txtEdit, KeyCode, Shift
End Sub

Private Sub flxResults_GotFocus()
    
    On Error GoTo Handler:
    ' Copy text in FlexGrid
    If txtEdit.Visible = False Then Exit Sub
    flxResults.text = txtEdit.text
    txtEdit.Visible = False
    
    ' Update associated data
    ' [WARNING] This code only woks for numeric indexes! (the only used in R_porter database)
    flxResults.Col = 0
    If flxResults.text <> "" Then
        If IsNumeric(txtEdit.text) Then
            query = flxResults.text & "=" & txtEdit.text
            flxResults.Col = 2
            
            Dim rs As ADODB.Recordset
            Set rs = New ADODB.Recordset
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            flxResults.Col = 3
            If Not rs.EOF Then
                flxResults.CellForeColor = RGB(0, 0, 0)
                flxResults.text = rs.Fields(0)
            Else
                flxResults.CellForeColor = RGB(255, 0, 0)
                flxResults.text = "[ERROR] Indice no encontrado"
            End If
        Else
            flxResults.Col = 3
            flxResults.CellForeColor = RGB(255, 0, 0)
            flxResults.text = "[ERROR] Indice no valido"
        End If
    End If
    flxResults.Col = 2
    flxResults.ColSel = flxResults.Cols - 1
    Exit Sub

Handler: MsgBox Err.Number & ": " & Err.Description, vbCritical, "flxResults_GotFocus"
End Sub

Private Sub flxResults_LeaveCell()
    ' Copy text in EditBox (double click)
    If txtEdit.Visible = False Then Exit Sub
    flxResults.text = txtEdit.text
    txtEdit.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    m_FormWidth = Me.ScaleWidth
    m_FormHeight = Me.ScaleHeight
    With flxResults
        m_FlexWidth = .width
        m_FlexHeight = .height
    End With
    Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.width < 3000 Then
        Me.width = 3000
    End If
    flxResults.width = Me.ScaleWidth - m_FormWidth + m_FlexWidth
    If Me.height < 1500 Then
        Me.height = 1500
    End If
    flxResults.height = Me.ScaleHeight - m_FormHeight + m_FlexHeight
End Sub

Private Sub mnuCancel_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    
    On Error GoTo Handler
    If vbNo = MsgBox("¿Estas seguro de agregar el registro?", vbQuestion + vbYesNo, "Confirmar insercion") Then
        Exit Sub
    End If
    With flxResults
        Dim k As Integer
        Dim str As String
        Dim dat As Date
        Dim sql As String
        sql = gs_SQL
        For k = 1 To .Rows - 1
            .Row = k
            .Col = 4
            If .text = "Date" Then
                .Col = 2
                dat = CDate(.text)
                str = Format(Month(dat), "00") & "/" & Format(Day(dat), "00") & "/" & Format(Year(dat), "0000") & " " & Format(Hour(dat), "00") & ":" & Format(Minute(dat), "00") & ":" & Format(Second(dat), "00")
            ElseIf .text = "Text" Then
                .Col = 2
                gfnc_ParseString .text, str
            Else
                .Col = 2
                str = .text
            End If
            sql = Replace(sql, DYNATABLE_FIELD_TAG, str, 1, 1, vbTextCompare)
        Next k
    End With
    
    Dim cd As ADODB.Command
    Set cd = New ADODB.Command
    Set cd.ActiveConnection = cn
    cd.CommandText = sql
    cd.Execute
    
    Unload Me
    
    Exit Sub
    
Handler: MsgBox "No se pudo insertar el registro." & vbCrLf & "Posibles causas de error:" & vbCrLf & _
                "- Algun campo contiene datos no validos." & vbCrLf & _
                "- Error accediendo a la base de datos." & vbCrLf & vbCrLf & _
                "ERROR NUMERO: " & Err.Number & vbCrLf & _
                Err.Description, vbExclamation, "ERROR"
End Sub

Private Sub startEditingFlexgrid(KeyAscii As Integer)
    flxResults.Col = 4
    If flxResults.text = "PK" Then
        MsgBox "No esta permitido modificar este valor" & vbCrLf & "pues es el indice del registro.", vbExclamation, "No permitido"
        flxResults.Col = 2
        flxResults.ColSel = flxResults.Cols - 1
        flxResults.SetFocus
    Else
        flxResults.Col = 2
        gsub_startEditFlexGrid flxResults, txtEdit, KeyAscii
    End If
End Sub

