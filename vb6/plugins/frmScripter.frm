VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScripter 
   Caption         =   "Test scripts"
   ClientHeight    =   2805
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5040
   Icon            =   "frmScripter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4455
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rchtxtResults 
      Height          =   2820
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   4974
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   21000
      TextRTF         =   $"frmScripter.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuActive 
         Caption         =   "&Activado"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Exportar log"
      End
      Begin VB.Menu m01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmScripter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mn_Form_W As Integer
Private mn_Form_H As Integer
Private mn_Rich_W As Integer
Private mn_Rich_H As Integer

Private Sub Form_Load()
    
    mn_Form_W = Me.ScaleWidth
    mn_Form_H = Me.ScaleHeight
    
    mn_Rich_W = rchtxtResults.Width
    mn_Rich_H = rchtxtResults.Height
    
    gb_exSCriptTestActive = True
    Me.mnuActive.Checked = True
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    rchtxtResults.Width = mn_Rich_W + Me.ScaleWidth - mn_Form_W
    rchtxtResults.Height = mn_Rich_H + Me.ScaleHeight - mn_Form_H
    
End Sub

Private Sub mnuActive_Click()
    If gb_exSCriptTestActive Then
        gb_exSCriptTestActive = False
        Me.mnuActive.Checked = False
    Else
        gb_exSCriptTestActive = True
        Me.mnuActive.Checked = True
    End If
End Sub

Private Sub mnuExit_Click()
    gb_exFormVisible = False
    gb_exSCriptTestActive = False
    Unload Me
End Sub

Private Sub mnuLog_Click()
    On Error GoTo ErrorCancel
    
    With CommonDialog
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Salvar el reporte como:"
        .Filter = "Archivos RTF(*.rtf)|*.rtf|Archivos de texto(*.txt)|*.txt|Todos los Archivos(*.*)|*.*"
        'necesario para controlar la extension con que se salvaran los archivos
        'sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .DefaultExt = ""
        .InitDir = App.Path
        'tipo predefinido RTF
        .FilterIndex = 1
        'nombre del reporte inicial
        .FileName = "Reporte_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date)
        .ShowSave
        If .FileName <> "" Then
            '.FilterIndex devuelve la extension seleccionada en el cuadro guardar como
            If .FilterIndex = 1 Then
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea RTF
                If UCase(Right(.FileName, 4)) <> ".RTF" Then
                    .FileName = .FileName & ".rtf"
                End If
                rchtxtResults.SaveFile .FileName, 0
            Else
            'en otro caso guardar como texto
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea TXT
                If UCase(Right(.FileName, 4)) <> ".TXT" Then
                    .FileName = .FileName & ".txt"
                End If
                rchtxtResults.SaveFile .FileName, 1
            End If
        End If
    End With
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical
    End If
End Sub
