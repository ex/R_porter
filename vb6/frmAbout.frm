VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "R_porter 1.2 by ::ex::"
   ClientHeight    =   3570
   ClientLeft      =   2145
   ClientTop       =   1335
   ClientWidth     =   6075
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      Height          =   1080
      Left            =   15
      TabIndex        =   0
      Top             =   2460
      Width           =   6045
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "Siste&ma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4935
         TabIndex        =   3
         Top             =   645
         Width           =   1020
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4935
         TabIndex        =   2
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label lblDisclaimer 
         Caption         =   $"frmAbout.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   4860
      End
   End
   Begin VB.Image imgBanner 
      Height          =   2400
      Left            =   45
      Picture         =   "frmAbout.frx":03FD
      Top             =   60
      Width           =   6000
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF12 Then
        cmdOK.value = True
    End If
End Sub

Private Sub StartSysInfo()

    Dim rc As Long
    Dim SysInfoPath As String
    On Error GoTo SysInfoErr
  
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
    
SysInfoErr: MsgBox "No se pudo acceder a la informacion del sistema", vbOKOnly
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdOK.value = True
End Sub

Private Sub imgBanner_Click()
    cmdOK.value = True
End Sub

Private Sub lblDisclaimer_Click()
    cmdOK.value = True
End Sub

