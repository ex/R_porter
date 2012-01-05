VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   Caption         =   "Espere por favor..."
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3870
      TabIndex        =   3
      Top             =   690
      Width           =   840
   End
   Begin VB.Frame fra 
      Height          =   705
      Left            =   15
      TabIndex        =   1
      Top             =   -30
      Width           =   4710
      Begin MSComctlLib.ImageList ImageList 
         Left            =   3945
         Top             =   780
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProgress.frx":014A
               Key             =   "icoToExcel"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProgress.frx":02AE
               Key             =   "icoToText"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProgress.frx":0412
               Key             =   "icoWait"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMessage 
         Caption         =   "lblMessage"
         Height          =   450
         Left            =   75
         TabIndex        =   2
         Top             =   165
         Width           =   4545
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   705
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mb_Active As Boolean

Private Sub cmdCancel_Click()
    Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mb_Active Then
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Height = 1440
    Me.Width = 4845
End Sub

