VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmReporte 
   Caption         =   "Reporte"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmReporte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rchtxt 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   7091
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReporte.frx":0442
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'
End Sub

Private Sub Form_Resize()

    Me.rchtxt.Width = Me.ScaleWidth - 2 * Me.rchtxt.Left
    Me.rchtxt.Height = Me.ScaleHeight - 75

End Sub
