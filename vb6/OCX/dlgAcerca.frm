VERSION 5.00
Begin VB.Form dlgAcerca 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dlgAcerca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "exit"
      Height          =   270
      Left            =   3105
      TabIndex        =   1
      Top             =   2070
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"dlgAcerca.frx":000C
      Height          =   1935
      Left            =   2865
      TabIndex        =   0
      Top             =   60
      Width           =   1770
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2250
      Left            =   60
      Picture         =   "dlgAcerca.frx":00ED
      Top             =   75
      Width           =   2730
   End
End
Attribute VB_Name = "dlgAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
