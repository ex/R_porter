VERSION 5.00
Begin VB.Form frmOrdenar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenar los resultados según..."
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   Icon            =   "frmOrdenar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3660
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Ordenar"
      Height          =   315
      Left            =   105
      TabIndex        =   20
      Top             =   2535
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar"
      Height          =   315
      Left            =   2445
      TabIndex        =   22
      Top             =   2550
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A&ceptar"
      Height          =   315
      Left            =   1275
      TabIndex        =   21
      Top             =   2550
      Width           =   1110
   End
   Begin VB.Frame fra 
      Height          =   2475
      Left            =   30
      TabIndex        =   23
      Top             =   -15
      Width           =   3600
      Begin VB.CheckBox chkSortActive 
         Height          =   240
         Index           =   3
         Left            =   3255
         TabIndex        =   28
         Top             =   1575
         Width           =   240
      End
      Begin VB.CheckBox chkSortActive 
         Height          =   240
         Index           =   4
         Left            =   3255
         TabIndex        =   27
         Top             =   2025
         Width           =   240
      End
      Begin VB.CheckBox chkSortActive 
         Height          =   240
         Index           =   2
         Left            =   3255
         TabIndex        =   26
         Top             =   1155
         Width           =   240
      End
      Begin VB.CheckBox chkSortActive 
         Height          =   240
         Index           =   1
         Left            =   3255
         TabIndex        =   25
         Top             =   720
         Width           =   240
      End
      Begin VB.CheckBox chkSortActive 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3255
         TabIndex        =   24
         Top             =   285
         Value           =   1  'Checked
         Width           =   240
      End
      Begin VB.CommandButton cmdDown 
         Height          =   300
         Index           =   0
         Left            =   1500
         Picture         =   "frmOrdenar.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   270
         Width           =   270
      End
      Begin VB.CommandButton cmdUp 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   90
         Picture         =   "frmOrdenar.frx":06B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   270
         Width           =   270
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         Index           =   0
         ItemData        =   "frmOrdenar.frx":07D6
         Left            =   1830
         List            =   "frmOrdenar.frx":07E0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdDown 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1485
         Picture         =   "frmOrdenar.frx":07FD
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2010
         Width           =   270
      End
      Begin VB.CommandButton cmdUp 
         Height          =   300
         Index           =   4
         Left            =   90
         Picture         =   "frmOrdenar.frx":0923
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2010
         Width           =   270
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         Index           =   4
         ItemData        =   "frmOrdenar.frx":0A49
         Left            =   1830
         List            =   "frmOrdenar.frx":0A53
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1980
         Width           =   1365
      End
      Begin VB.CommandButton cmdDown 
         Height          =   300
         Index           =   3
         Left            =   1485
         Picture         =   "frmOrdenar.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1545
         Width           =   270
      End
      Begin VB.CommandButton cmdUp 
         Height          =   300
         Index           =   3
         Left            =   90
         Picture         =   "frmOrdenar.frx":0B96
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1545
         Width           =   270
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         Index           =   3
         ItemData        =   "frmOrdenar.frx":0CBC
         Left            =   1830
         List            =   "frmOrdenar.frx":0CC6
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1530
         Width           =   1365
      End
      Begin VB.CommandButton cmdDown 
         Height          =   300
         Index           =   2
         Left            =   1485
         Picture         =   "frmOrdenar.frx":0CE3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1125
         Width           =   270
      End
      Begin VB.CommandButton cmdUp 
         Height          =   300
         Index           =   2
         Left            =   90
         Picture         =   "frmOrdenar.frx":0E09
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1125
         Width           =   270
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         Index           =   2
         ItemData        =   "frmOrdenar.frx":0F2F
         Left            =   1830
         List            =   "frmOrdenar.frx":0F39
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1110
         Width           =   1365
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         Index           =   1
         ItemData        =   "frmOrdenar.frx":0F56
         Left            =   1830
         List            =   "frmOrdenar.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton cmdUp 
         Height          =   300
         Index           =   1
         Left            =   90
         Picture         =   "frmOrdenar.frx":0F7D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   270
      End
      Begin VB.CommandButton cmdDown 
         Height          =   300
         Index           =   1
         Left            =   1485
         Picture         =   "frmOrdenar.frx":10A3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   690
         Width           =   270
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   345
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   345
         TabIndex        =   17
         Top             =   1995
         Width           =   1155
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   345
         TabIndex        =   13
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   345
         TabIndex        =   9
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblField 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   345
         TabIndex        =   5
         Top             =   675
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmOrdenar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSort_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub cmdAceptar_Click()

    Aplicar_Opciones
    Unload Me
    frmDataControl.cmdSearch.SetFocus

End Sub

Private Sub cmdBuscar_Click()
    
    Aplicar_Opciones
    Unload Me
    frmDataControl.Generar_Lista

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click(Index As Integer)
    '===================================================
    Dim s_temp As String
    '===================================================
    
    s_temp = lblField(Index + 1).Caption
    lblField(Index + 1).Caption = lblField(Index).Caption
    lblField(Index).Caption = s_temp

    s_temp = lblField(Index + 1).Tag
    lblField(Index + 1).Tag = lblField(Index).Tag
    lblField(Index).Tag = s_temp

End Sub

Private Sub cmdUp_Click(Index As Integer)
    '===================================================
    Dim s_temp As String
    '===================================================
    
    s_temp = lblField(Index - 1).Caption
    lblField(Index - 1).Caption = lblField(Index).Caption
    lblField(Index).Caption = s_temp

    s_temp = lblField(Index - 1).Tag
    lblField(Index - 1).Tag = lblField(Index).Tag
    lblField(Index).Tag = s_temp

End Sub

Private Sub Form_Load()
    '===================================================
    Dim k As Integer
    '===================================================
    
    For k = 1 To 5
    
        lblField(k - 1).Tag = gs_DBOrdenarCampo(k)
        
        Select Case gs_DBOrdenarCampo(k)
        
            Case "Parent"
            
                If gb_DBPertenciaPorAlmacenamiento Then
                    lblField(k - 1).Caption = "Medio"
                Else
                    lblField(k - 1).Caption = "Pertenece a"
                End If
            
            Case "Aux"
            
                If gb_DBCampoAuxiliarPorGenero Then
                    lblField(k - 1).Caption = "Género"
                Else
                    lblField(k - 1).Caption = "Tipo"
                End If
            
            Case "Cantidad"
                lblField(k - 1).Caption = gs_DBConCampoDe
            
            Case Else
                lblField(k - 1).Caption = gs_DBOrdenarCampo(k)
                
        End Select
        
        If gb_DBOrdenarAsc(k) Then
            cmbSort(k - 1).ListIndex = 0
        Else
            cmbSort(k - 1).ListIndex = 1
        End If
        
        If gb_DBOrdenarEnabled(k) Then
            chkSortActive(k - 1).value = vbChecked
        Else
            chkSortActive(k - 1).value = vbUnchecked
        End If
        
    Next

End Sub


Private Sub Aplicar_Opciones()
    '===================================================
    Dim k As Integer
    '===================================================
    
    For k = 0 To 4
    
        gs_DBOrdenarCampo(k + 1) = lblField(k).Tag
        
        If (0 = cmbSort(k).ListIndex) Then
            gb_DBOrdenarAsc(k + 1) = True
        Else
            gb_DBOrdenarAsc(k + 1) = False
        End If
        
        If (vbChecked = chkSortActive(k).value) Then
            gb_DBOrdenarEnabled(k + 1) = True
        Else
            gb_DBOrdenarEnabled(k + 1) = False
        End If
        
    Next k
    
End Sub
