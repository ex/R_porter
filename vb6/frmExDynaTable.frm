VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmExDynaTable 
   Caption         =   "frmExDynaTable"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5940
   Icon            =   "frmExDynaTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDownSearch 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4980
      Top             =   1725
   End
   Begin VB.CommandButton cmdDownSearch 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   5145
      Picture         =   "frmExDynaTable.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   75
      Width           =   255
   End
   Begin VB.PictureBox pbxSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   3225
      ScaleHeight     =   1260
      ScaleWidth      =   2145
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton optDBSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.OptionButton optDBSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblOptionSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Todos"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   14
         Top             =   15
         Width           =   2685
      End
      Begin VB.Label lblOptionSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que contenga..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   13
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label lblOptionSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Con palabra completa..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   12
         Top             =   525
         Width           =   2685
      End
      Begin VB.Label lblOptionSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que empiece con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   11
         Top             =   780
         Width           =   2685
      End
      Begin VB.Label lblOptionSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que termine con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   10
         Top             =   1035
         Width           =   2685
      End
   End
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   4755
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   5280
      Top             =   4335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":06B0
            Key             =   "db_new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":080C
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":0950
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":0A94
            Key             =   "db_edit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":0BF0
            Key             =   "db_delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":0D4C
            Key             =   "db_2excel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":0EA8
            Key             =   "db_2txt"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExDynaTable.frx":1004
            Key             =   "db_search"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglst"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   40
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "db_new"
            Object.ToolTipText     =   "Agregar nuevo registro"
            ImageKey        =   "db_new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "db_edit"
            Object.ToolTipText     =   "Editar registro"
            ImageKey        =   "db_edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "db_delete"
            Object.ToolTipText     =   "Eliminar registro"
            ImageKey        =   "db_delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "db_2excel"
            Object.ToolTipText     =   "Exportar a Excel"
            ImageKey        =   "db_2excel"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "db_2txt"
            Object.ToolTipText     =   "Exportar a texto"
            ImageKey        =   "db_2txt"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "db_search"
            Object.ToolTipText     =   "Buscar registros"
            ImageKey        =   "db_search"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cmbFieldSearch 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   1620
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3195
         TabIndex        =   0
         Top             =   30
         Width           =   2190
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   4590
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   8096
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   4
      BackColorFixed  =   16750143
      ForeColorFixed  =   16777215
      BackColorSel    =   16775910
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuExportar 
         Caption         =   "E&xportar resultados..."
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Registros"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOrdenar 
         Caption         =   "&Ordenar columna"
         Shortcut        =   ^Q
         Visible         =   0   'False
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditar 
         Caption         =   "&Editar campo"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "E&liminar registro"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSort 
         Caption         =   "&Ordenar columna"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Editar campo"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&Nuevo registro"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "E&liminar registro"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmExDynaTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const EX_SCROLL_FLX = 1

'***************************************
' VARIABLES PUBLICAS
'***************************************
Public gs_SQL As String
Public gl_IdRow As Long
Public gs_mainTable As String
Public gs_mainField As String
Public gs_indexField As String
Public go_parent As clsExDynaTable

'***************************************
' VARIABLES DE MODULO
'***************************************
Private mz_DeleteConstrains() As exDynaTableDeleteConstrains
Private mn_Constrains As Integer
Private m_topRow As Integer
Private m_WndProcOrg As Long
Private m_HWndSubClassed As Long

Private Enum exdynDB_SearchStyle
   db_Todos = 0
   db_Con = 1
   db_ConPalabra = 2
   db_QueComience = 3
   db_QueTermine = 4
End Enum

Private Const TOP_PBX_TXT = 315

Private m_FormWidth As Integer
Private m_FormHeight As Integer

Private m_FlexWidth As Integer
Private m_FlexHeight As Integer

Private mt_DBSearchStyle As exdynDB_SearchStyle
Private mb_DBSearchStyleActive As Boolean

''*************************************************************
'' TEST SIZE
''*************************************************************
'Private Sub Form_Unload(Cancel As Integer)
'    Dim k As Integer
'    Me.flxResults.ColWidth(k) = Me.flxResults.ColWidth(k)
'    Me.Width = Me.Width
'End Sub
''*************************************************************

'**************************************************************
'* CMBOPTION SEARCH
'**************************************************************
Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscar_Registros
    End If
End Sub

Private Sub txtSearch_GotFocus()
    If mt_DBSearchStyle = db_Todos Then
        txtSearch.text = "[Todos]"
    End If
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.text)
End Sub

Private Sub txtSearch_VerifyAllScan()
    ' por si el usuario quiere buscar algo diferente a todos
    If mt_DBSearchStyle = db_Todos Then
        If (txtSearch.text <> "[Todos]") And (Trim(txtSearch.text) <> "") Then
            lblOptionSearch(db_Todos).ForeColor = &H0&
            lblOptionSearch(db_Todos).BackColor = &HFFFFFF
            mt_DBSearchStyle = db_Con
        End If
    End If
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch_VerifyAllScan
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        If mb_DBSearchStyleActive = False Then
            mb_DBSearchStyleActive = True
            pbxSearch.Top = txtSearch.Top + TOP_PBX_TXT
            pbxSearch.Visible = True
            optDBSearch(mt_DBSearchStyle).value = True
            optDBSearch(mt_DBSearchStyle).SetFocus
            Set cmdDownSearch.Picture = imglst.ListImages.Item("up").Picture
        End If
    End If
End Sub

Private Sub cmdDownSearch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownSearch_Click
    End If
End Sub

Private Sub cmdDownSearch_Click()
    If mb_DBSearchStyleActive = False Then
        mb_DBSearchStyleActive = True
        pbxSearch.Top = txtSearch.Top + TOP_PBX_TXT
        pbxSearch.Visible = True
        optDBSearch(mt_DBSearchStyle).value = True
        optDBSearch(mt_DBSearchStyle).SetFocus
        Set cmdDownSearch.Picture = imglst.ListImages.Item("up").Picture
    Else
        tmrDownSearch.Enabled = False
        mb_DBSearchStyleActive = False
        'primero quitamos enfoque
        txtSearch.SetFocus
        pbxSearch.Visible = False
        Set cmdDownSearch.Picture = imglst.ListImages.Item("down").Picture
    End If
End Sub

Private Sub optDBSearch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    optDBSearch(Index).value = True
    tmrDownSearch.Enabled = True
End Sub

Private Sub optDBSearch_Click(Index As Integer)
Dim Inicio As Single
    If mb_DBSearchStyleActive = True Then
        mt_DBSearchStyle = Index
    End If
End Sub

Private Sub optDBSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownSearch_Click
    End If
End Sub

Private Sub optDBSearch_LostFocus(Index As Integer)
    If (Me.ActiveControl.Name = "cmdDownSearch") Then
        Me.ActiveControl.SetFocus
    Else
        If (Me.ActiveControl.Name = "optDBSearch") Then
            lblOptionSearch(Index).ForeColor = &H0&
            lblOptionSearch(Index).BackColor = &HFFFFFF
            Me.ActiveControl.SetFocus
        Else
            If (Me.ActiveControl.Name = "pbxSearch") Then
                optDBSearch(Index).SetFocus
            Else
                mb_DBSearchStyleActive = False
                pbxSearch.Visible = False
                Set cmdDownSearch.Picture = imglst.ListImages.Item("down").Picture
            End If
        End If
    End If
End Sub

Private Sub optDBSearch_GotFocus(Index As Integer)
    lblOptionSearch(Index).ForeColor = &HFFFFFF
    lblOptionSearch(Index).BackColor = &HFF963F
End Sub

Private Sub lblOptionSearch_Click(Index As Integer)
    optDBSearch(Index).value = True
    optDBSearch(Index).SetFocus
    tmrDownSearch.Enabled = True
End Sub

Private Sub tmrDownSearch_Timer()
    tmrDownSearch.Enabled = False
    mb_DBSearchStyleActive = False
    'primero quitamos enfoque
    txtSearch.SetFocus
    pbxSearch.Visible = False
    Set cmdDownSearch.Picture = imglst.ListImages.Item("down").Picture
End Sub

'***************************************
' MENUS
'***************************************
Private Sub mnuBuscar_Click()
    Buscar_Registros
End Sub

Private Sub mnuExportar_Click()
    gsub_FlxShowSaveAsDialog cmmdlg, flxResults
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuEditar_Click()
    Editar_Registro
End Sub

Private Sub mnuNuevo_Click()
    Agregar_Registro
End Sub

Private Sub mnuBorrar_Click()
    Eliminar_Registro
End Sub

Private Sub mnuDelete_Click()
    mnuBorrar_Click
End Sub

Private Sub mnuEdit_Click()
    mnuEditar_Click
End Sub
Private Sub mnuExit_Click()
    mnuSalir_Click
End Sub

Private Sub mnuNew_Click()
    mnuNuevo_Click
End Sub

Private Sub mnuSort_Click()
    mnuOrdenar_Click
End Sub

Private Sub mnuRefresh_Click()
    Buscar_Registros
End Sub

Private Sub mnuOrdenar_Click()
    
    Dim k As Integer
    Dim oldRow As Long
    Dim oldCol As Long

    With flxResults
        oldRow = .Row
        oldCol = .Col
        If 0 <> .Col Then
            If 0 = .ColData(.Col) Then
                .Sort = flexSortGenericDescending
                .ColData(.Col) = 1
            Else
                .Sort = flexSortGenericAscending
                .ColData(.Col) = 0
            End If
            
            .Redraw = False
            
            .Col = 0
            For k = 1 To .Rows - 1
                .Row = k
                .text = k
            Next
            
            .Redraw = True
            .Row = oldRow
            .Col = oldCol
        End If
    End With
End Sub

'***************************************
' FORMULARIO
'***************************************
Private Sub Form_Load()

    On Error GoTo Handler

#If EX_SCROLL_FLX Then
    If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
        If Not m_WndProcOrg Then
            m_HWndSubClassed = Me.hWnd
            SetWindowLong hWnd, GWL_USERDATA, ObjPtr(Me)
            m_WndProcOrg = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProcExDynaTableForm)
            m_topRow = 1
        End If
    End If
#End If
    
    m_FormWidth = Me.ScaleWidth
    m_FormHeight = Me.ScaleHeight
    
    With flxResults
        m_FlexWidth = .width
        m_FlexHeight = .height
    End With
    
    mt_DBSearchStyle = db_QueComience
    
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.width < 3000 Then
        Me.width = 3000
    End If
    flxResults.width = Me.ScaleWidth - m_FormWidth + m_FlexWidth
    If Me.height < 3000 Then
        Me.height = 3000
    End If
    flxResults.height = Me.ScaleHeight - m_FormHeight + m_FlexHeight
End Sub

'***************************************
' FLEXGRID
'***************************************
Private Sub flxResults_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            Eliminar_Registro
    End Select
End Sub

Private Sub flxResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Me.flxResults.Row <> 0 And Me.flxResults.RowHeight(Me.flxResults.Row) <> 0 Then
            PopupMenu Me.mnuPopUp, 2
        End If
    End If
End Sub

Private Sub flxResults_SelChange()
    If flxResults.RowSel <> flxResults.Row Then
        flxResults.RowSel = flxResults.Row
    End If
End Sub

Private Sub flxResults_DblClick()
    Editar_Registro
End Sub

'***************************************
' MANEJO DE REGISTROS
'***************************************
Private Sub Buscar_Registros()

    Dim rs As ADODB.Recordset
    Dim sMainField As String
    Dim field_variant As String
    Dim s_Field As String
    Dim k As Long
    Dim n_fields As Integer
    Dim b_Redraw As Boolean
    Dim s_query As String
    On Error GoTo Handler

    ' por si el usuario quiere buscar algo diferente a todos
    txtSearch_VerifyAllScan
    
    '==============================================================
    ' Buscamos caracteres de comilla (y los tratamos adecuadamente)
    '
    gfnc_ParseString txtSearch.text, s_Field

    Select Case mt_DBSearchStyle
        Case db_Todos:
            txtSearch.text = "[Todos]"
    
        Case db_Con:
            field_variant = "'%" & Trim(s_Field) & "%'"
            
        Case db_ConPalabra:
            field_variant = "'% " & Trim(s_Field) & " %') OR (" & Me.gs_mainField & " LIKE '" & _
                            Trim(s_Field) & " %') OR (" & Me.gs_mainField & " LIKE '% " & _
                            Trim(s_Field) & "') OR (" & Me.gs_mainField & "='" & Trim(s_Field) & "'"
            
        Case db_QueComience:
            field_variant = "'" & Trim(s_Field) & "%'"
            
        Case db_QueTermine:
            field_variant = "'%" & Trim(s_Field) & "'"
    End Select
    
    If txtSearch.text = "[Todos]" Then
        s_query = gs_SQL
    Else
        s_query = InsertCriteria(gs_SQL, "((" & Me.gs_mainField & " LIKE " & field_variant & "))")
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open s_query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    n_fields = rs.Fields.Count

    If n_fields > 0 Then
    
        With flxResults
    
            .Row = 0
            
            'borramos todas las filas exceptuando la fija y la primera
            .Rows = 2
                    
            If rs.EOF = True Then
                'primera fila invisible no se puede eliminar
                .RowHeight(1) = 0
                '-------------------------------------------------------------------------------
                ' necesaria la siguiente linea para evitar que la primera fila quede seleccionada
                ' de vez en cuando, (parece que se ha colgado el programa) - error del msflexgrid
                .Row = 1
                .Col = 0
                '-------------------------------------------------------------------------------
                .SetFocus
                GoTo SALIR
            End If
            
            Screen.MousePointer = vbHourglass
            .MousePointer = flexHourglass
    
            .Redraw = False
    
            b_Redraw = True
    
            '---------------------------------------
            ' llenar datos
            '
            While rs.EOF = False
    
                .Row = .Rows - 1
                'forzar visible
                .RowHeight(.Row) = -1
    
                .Col = 0
                .CellForeColor = RGB(100, 170, 255)
                .text = .Row
    
                For k = 0 To (n_fields - 1)
                    .Col = k + 1
                    .text = rs.Fields(k).value
                Next k
    
                rs.MoveNext
    
                .Rows = .Rows + 1
    
                If b_Redraw Then
                    If .Row >= CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) + 1 Then
                        .Redraw = True
                        .Refresh
                        .Redraw = False
                        b_Redraw = False
                    End If
                End If
    
            Wend
    
            ' eliminar la ultima fila agregada que esta vacia
            .Rows = .Rows - 1
SALIR:
            .Row = 1
            .Col = 0
            .ColSel = n_fields
    
            .Redraw = True
            .SetFocus
    
        End With
    End If

    Screen.MousePointer = vbDefault
    flxResults.MousePointer = flexDefault

    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Handler:
    Select Case Err.Number
        Case 94, 13
            'uso no valido de NULL
            'cuando el campo esta vacio
            Resume Next
        Case Else
            Screen.MousePointer = vbDefault
            flxResults.MousePointer = flexDefault
            MsgBox Err.Description, vbCritical, "Buscar_Registros"
    End Select
End Sub

Private Sub Eliminar_Registro()

    Dim s_query As String
    Dim id_registro As Long
    Dim fila_eliminada As String
    Dim nombre_fila As String
    Dim fila_eliminada_2 As String
    Dim nombre_fila_2 As String
    Dim rs As ADODB.Recordset
    Dim cd As ADODB.Command
    Dim oldRow As Long
    Dim a As Long
    Dim b As Long
    Dim k As Long
    Dim q As Long
    On Error GoTo Handler

    ' Verify conection and database
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If

    ' Delete selected register
    With flxResults
        If .Row > 0 And .RowHeight(.Row) > 0 Then
            .Redraw = False
            a = .Row
            b = .RowSel
            oldRow = .TopRow
        
            If b = a Then
                .Col = 0
                fila_eliminada = .text
                .Col = 1
                nombre_fila = .text
            Else
                If b < a Then
                    k = b
                    b = a
                    a = k
                End If
                
                .Row = b
                                
                .Col = 0
                fila_eliminada_2 = .text
                .Col = 1
                nombre_fila_2 = .text
                .Row = a
                .Col = 0
                fila_eliminada = .text
                .Col = 1
                nombre_fila = .text
                .RowSel = b
            End If
            
            .Col = 0
            .ColSel = .Cols - 1
            
            .Redraw = True
            
            ' elimina un incomodo movimiento cuando .row esta en las ultimas filas
            .TopRow = oldRow
            
            If a = b Then
                If vbYes = MsgBox("¿Estás seguro de eliminar el registro [" & fila_eliminada & "]:" & _
                                    vbCrLf & nombre_fila & "?", vbExclamation + vbYesNo, "Confirmar eliminación") Then
                    .Col = gl_IdRow
                    id_registro = CLng(.text)
            
                    .Col = 0
                    .ColSel = .Cols - 1
                    
                    ' ejecutar consulta de verificacion
                    For k = 1 To UBound(mz_DeleteConstrains)
                        
                        s_query = "SELECT * FROM " & mz_DeleteConstrains(k).sTableLinked & " WHERE " & mz_DeleteConstrains(k).sFieldLinked & " = " & id_registro
                        
                        Set rs = New ADODB.Recordset
                        rs.Open s_query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
                        
                        If rs.EOF Then
                            ' no hay registros vinculados
                            rs.Close
                        Else
                            rs.Close
                            ' existe al menos un registro vinculado
                            If vbYes = MsgBox("La tabla [" & mz_DeleteConstrains(k).sTableLinked & "] hace referencia al" & _
                                               vbCrLf & "registro de ID [" & id_registro & "] que esta intentando eliminar." & _
                                               vbCrLf & "¿Deseas continuar de todos modos?", vbExclamation + vbYesNo, "Reconfirmar eliminación") Then
                                ' si se quiere eliminar de todos modos
                                Exit For
                            Else
                                ' cancelar eliminacion
                                Exit Sub
                            End If
                        End If
                    Next k
JMP_NO_RESTRICTIONS:
                    Set cd = New ADODB.Command
                    Set cd.ActiveConnection = cn
                    
                    cd.CommandText = "DELETE FROM " & Me.gs_mainTable & " WHERE " & Me.gs_indexField & "=" & id_registro
                    cd.Execute
                    
                    If .Rows > 2 Then
                        oldRow = .Row
                        .RemoveItem (.Row)
                    Else
                        'volver invisible la primera fila
                        .RowHeight(1) = 0
                        Exit Sub
                    End If
                    
                Else
                    Exit Sub
                End If
            Else
                ' TODO (multiple delete)
            End If
                
            '===================================================
            ' re-numerar (optimizar...)
            '
            .Redraw = False
            .Col = 0
            
            For k = 1 To .Rows - 1
                .Row = k
                If .RowHeight(k) <> 0 Then
                    .text = k
                End If
            Next
            .Redraw = True
            '---------------------------------------------------
                
            'seleccionar la fila anterior
            If oldRow > .Rows - 1 Then
                .Row = oldRow - 1
            Else
                .Row = oldRow
            End If
            
            .Col = 0
            .ColSel = .Cols - 1
                
            'verificar si la fila esta visible
            If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
                .TopRow = .Row
            End If

            .SetFocus
        End If
    End With
    
    Exit Sub
    
Handler:
    Select Case Err.Number
        Case 9
            ' La matriz de restricciones esta vacia
            Resume JMP_NO_RESTRICTIONS
        Case Else
            MsgBox Err.Description, vbCritical, "Eliminar_Registro"
            flxResults.Redraw = True
    End Select
End Sub

Private Sub Editar_Registro()
    With flxResults
        If .Row > 0 And .RowHeight(.Row) > 0 Then
            .Col = gl_IdRow
            go_parent.initEditAction flxResults.text
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
End Sub

Private Sub Agregar_Registro()
    go_parent.initInsertAction
End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "db_new"
            Agregar_Registro
        Case "db_edit"
            Editar_Registro
        Case "db_delete"
            Eliminar_Registro
        Case "db_2txt"
            mnuExportar_Click
        Case "db_search"
            Buscar_Registros
    End Select
End Sub

Private Function InsertCriteria(ByVal sql, ByVal str_append) As String

    Dim n As Long
    Dim s_sql As String
    On Error GoTo Handler
    
    s_sql = sql
    InsertCriteria = sql
    
    '-------------------------------------
    ' buscar WHERE
    n = InStr(1, s_sql, "WHERE")
        
    If n = 0 Then
        '-------------------------------------
        ' buscar ORDER BY
        n = InStr(1, s_sql, "ORDER BY")
    
        If n = 0 Then
            s_sql = s_sql & " WHERE " & str_append & Mid(s_sql, n + 1)
        Else
            s_sql = Mid(s_sql, 1, n) & " WHERE " & str_append & Mid(s_sql, n + 1)
        End If
    Else
        s_sql = Mid(s_sql, 1, n + 5) & str_append & " AND " & Mid(s_sql, n + 6)
    End If
    
    InsertCriteria = s_sql
    Exit Function
    
Handler:
End Function

Public Sub AddDeleteConstrains(ByRef strTable As String, ByRef strField As String)

    On Error GoTo Handler
    
    mn_Constrains = mn_Constrains + 1
    ReDim Preserve mz_DeleteConstrains(1 To mn_Constrains)
    mz_DeleteConstrains(mn_Constrains).sTableLinked = strTable
    mz_DeleteConstrains(mn_Constrains).sFieldLinked = strField
    
    Exit Sub
    
Handler: MsgBox Err.Description, vbCritical, "AddDeleteConstrains"
End Sub

'*******************************************************************************
' Funciones para el mouse wheel scroll
'*******************************************************************************
Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If uMsg = WM_MOUSEWHEEL Then
        ' scroll del flexgrid tenga o no el enfoque...
        If (HiWord(wParam) / WHEEL_DELTA) < 0 Then
            ScrollDown
        Else
            ScrollUp
        End If
    End If
    WindowProc = CallWindowProc(m_WndProcOrg, Me.hWnd, uMsg, wParam, lParam)
End Function

Public Sub ScrollUp()
    On Error Resume Next
    If m_topRow > 1 Then
        m_topRow = flxResults.TopRow - 4
        If m_topRow <= 1 Then
            m_topRow = 1
        End If
        flxResults.TopRow = m_topRow
    End If
End Sub

Public Sub ScrollDown()
    On Error Resume Next
    If m_topRow < flxResults.Rows - 4 Then
        m_topRow = flxResults.TopRow + 4
        If m_topRow > flxResults.Rows - 1 Then
            m_topRow = flxResults.Rows - 1
        End If
        flxResults.TopRow = m_topRow
    End If
End Sub
'*******************************************************************************

