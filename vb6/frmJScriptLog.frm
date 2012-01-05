VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmJScriptLog 
   Caption         =   "Script log:"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmJScriptLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rchtxtLog 
      Height          =   3930
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6932
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   60000
      TextRTF         =   $"frmJScriptLog.frx":030A
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
End
Attribute VB_Name = "frmJScriptLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next

    If (Me.width < 2655) Then
        Me.width = 2655
    End If
    
    If (Me.height < 2100) Then
        Me.height = 2100
    End If
    
    rchtxtLog.height = Me.ScaleHeight + 15
    rchtxtLog.width = Me.ScaleWidth + 15
    
End Sub

Public Sub echo(line As String)
    rchtxtLog.SelText = line & vbCrLf
End Sub

Public Sub error(line As String)
    rchtxtLog.SelColor = vbRed
    rchtxtLog.SelText = line & vbCrLf
    rchtxtLog.SelColor = vbBlack
End Sub

Public Sub warning(line As String)
    rchtxtLog.SelColor = RGB(126, 126, 126)
    rchtxtLog.SelText = line & vbCrLf
    rchtxtLog.SelColor = vbBlack
End Sub


