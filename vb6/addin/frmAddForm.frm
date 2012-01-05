VERSION 5.00
Begin VB.Form frmAddForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de inserción"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmAddForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDSN 
      Height          =   1260
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   3990
      Begin VB.ComboBox cmbProyectos 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   2760
      End
      Begin VB.TextBox txtForm 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1125
         TabIndex        =   3
         Top             =   645
         Width           =   2745
      End
      Begin VB.Label lbl 
         Caption         =   "Formulario:"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lbl 
         Caption         =   "&Proyecto:"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   623
      TabIndex        =   4
      Top             =   1365
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2243
      TabIndex        =   5
      Top             =   1365
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim vbp As VBProject
Dim vbc As VBComponent

    On Error GoTo Handler

    If (Me.cmbProyectos.Text <> "") And (gb_PathFind = True) Then
    
        Set vbp = gVBInstance.VBProjects.Item(Me.cmbProyectos.Text)
        Set vbc = vbp.VBComponents.AddFromTemplate(gs_Path & "\spanish.frm")
        
        vbc.Name = Me.txtForm
        vbc.Properties.Item("Caption") = Me.txtForm
        vbc.Activate
        
    End If
    
    Unload Me
    Exit Sub

Handler:

    If Err.Number = 50135 Then
        MsgBox "Ya existia un formulario con nombre: [" & Me.txtForm & "]" & vbCrLf & "Se insertó un formulario con nombre: [" & vbc.Name & "]", vbExclamation, "Error"
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "cmdAceptar_Click()"
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim vbp As VBProject
Dim k As Integer

    On Error GoTo Handler
    
    If gVBInstance.VBProjects.Count = 0 Then
        MsgBox "Ningún proyecto activo", vbExclamation, "Error"
        Exit Sub
    End If
        
    For k = 1 To gVBInstance.VBProjects.Count
        Set vbp = gVBInstance.VBProjects.Item(k)
        Me.cmbProyectos.AddItem vbp.Name
    Next
   
    Me.txtForm = gs_Form
    Me.cmbProyectos.Text = gVBInstance.VBProjects.StartProject.Name
    
    Exit Sub

Handler:

    If Err.Number = -2147467259 Then
        'No existe proyecto activo
        If Me.cmbProyectos.ListCount > 0 Then
            Me.cmbProyectos.ListIndex = 0
        End If
    Else
        MsgBox Err.Description, vbCritical, "Form_Load()"
    End If

End Sub
