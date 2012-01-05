VERSION 5.00
Object = "{DAC1C15E-A0D0-11D8-92BC-F3955AEE4860}#3.1#0"; "exHighLightCode.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form testcode 
   Caption         =   "frmTest"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2790
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   2190
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin exHighLightCode.exCodeHighlight exCodeHighlight1 
      Height          =   2700
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5398
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RightMargin     =   40000
      SelRTF          =   $"testcode.frx":0000
      Language        =   1
      OperatorColor   =   255
      CommentColor    =   32768
      LiteralColor    =   16576
      ForeColor       =   0
      FunctionColor   =   12583104
      Author          =   "Esau R.O. [exe_q_tor] ...based in the DevDomainCodeHighlight control."
      LeftMargin      =   180
      LeftMarginColor =   15856113
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu0_0 
         Caption         =   "-"
      End
      Begin VB.Menu txt 
         Caption         =   "test DevPad"
      End
   End
   Begin VB.Menu mnueditar 
      Caption         =   "Editar"
      Begin VB.Menu tabs 
         Caption         =   "tabs"
      End
      Begin VB.Menu mnubold 
         Caption         =   "bold"
      End
      Begin VB.Menu italy 
         Caption         =   "italy"
      End
      Begin VB.Menu sel 
         Caption         =   "sel"
      End
   End
   Begin VB.Menu mnuleng 
      Caption         =   "Script"
      Begin VB.Menu mnuNone 
         Caption         =   "none"
      End
      Begin VB.Menu mnuvbs 
         Caption         =   "vbs"
      End
      Begin VB.Menu mnuc 
         Caption         =   "c++"
      End
      Begin VB.Menu mnuJScript 
         Caption         =   "JScript"
      End
      Begin VB.Menu perl 
         Caption         =   "perl"
      End
      Begin VB.Menu mnuruby 
         Caption         =   "ruby"
      End
      Begin VB.Menu mnupython 
         Caption         =   "python"
      End
      Begin VB.Menu mnusql 
         Caption         =   "SQL"
      End
   End
End
Attribute VB_Name = "testcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const EM_SETTABSTOPS = &HCB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Const m_def_ID_EX = 25647893


Private Sub exCodeHighlight1_SelChange()
    Me.StatusBar1.SimpleText = "line: " & Me.exCodeHighlight1.CursorLine & ", row: " & Me.exCodeHighlight1.CursorRow
End Sub

Private Sub Form_Load()
    Me.exCodeHighlight1.ExID = m_def_ID_EX
    Me.exCodeHighlight1.Language = 2
    
    SendMessage Me.exCodeHighlight1.RichHwnd, EM_SETTABSTOPS, 0&, vbNullString
    SendMessage Me.exCodeHighlight1.RichHwnd, EM_SETTABSTOPS, 1, 16
    exCodeHighlight1.Refresh
    
End Sub

Private Sub Form_Resize()
    Me.exCodeHighlight1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Me.StatusBar1.Height
End Sub

Private Sub italy_Click()
    Me.exCodeHighlight1.SelItalic = True
End Sub

Private Sub mnuabrir_Click()
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'esconde casilla de solo lectura y verifica que el archivo y el path existan
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .DialogTitle = "Indicar el archivo SQL a abrir:"
        .Filter = "Archivo C/C++ (*.cpp)|*.cpp;*.h|Archivo VBS (*.vbs)|*.vbs|Archivo SQL (*.sql)|*.sql|Todos los Archivos(*.*)|*.*"
        .InitDir = App.Path
        'tipo predefinido VBS
        .FilterIndex = 1
        .ShowOpen
        If .FileName <> "" Then
            'cargar el archivo SQL
            Me.exCodeHighlight1.LoadFile .FileName
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "mnuOpen_Click"
    End If
End Sub

Private Sub mnubold_Click()
    Me.exCodeHighlight1.SelBold = True
End Sub

Private Sub mnuJScript_Click()
    Me.exCodeHighlight1.Language = 5
End Sub

Private Sub mnuNone_Click()
    Me.exCodeHighlight1.Language = exNOHighLight
End Sub

Private Sub mnupython_Click()
    Me.exCodeHighlight1.Language = exPython
End Sub

Private Sub mnuruby_Click()
    Me.exCodeHighlight1.Language = exRuby
End Sub

Private Sub mnusalvar_Click()
On Error GoTo Handler
    Me.exCodeHighlight1.SaveFile "c:\test.rtf", 0
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub mnusql_Click()
    Me.exCodeHighlight1.Language = 4
End Sub

Private Sub mnuvbs_Click()
    Me.exCodeHighlight1.Language = 1
End Sub

Private Sub mnuc_Click()
    Me.exCodeHighlight1.Language = 2
End Sub

Private Sub perl_Click()
    Me.exCodeHighlight1.Language = 6
End Sub

Private Sub sel_Click()
    Me.exCodeHighlight1.SelText = "select  " 'iniciar salto de linea al principio sino sera obviado
    Me.exCodeHighlight1.SelText = vbCrLf & "if sxcx elect --asdad"
    Me.exCodeHighlight1.SelText = vbCrLf & "'update"
    Me.exCodeHighlight1.SelText = vbCrLf
End Sub

Private Sub tabs_Click()
    
    SendMessage Me.exCodeHighlight1.RichHwnd, EM_SETTABSTOPS, 0&, vbNullString
    SendMessage Me.exCodeHighlight1.RichHwnd, EM_SETTABSTOPS, 1, 4
    
    exCodeHighlight1.Refresh

End Sub

Private Sub txt_Click()
    Form1.Show vbModal
End Sub
