VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   1845
   ClientTop       =   1290
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   705
      Left            =   105
      Picture         =   "frmSplash.frx":000C
      Top             =   3645
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSplash.frx":17DE
      Top             =   0
      Width           =   6225
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
