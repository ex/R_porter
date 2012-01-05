VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelect 
   Caption         =   "Resultados SELECT"
   ClientHeight    =   4605
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3960
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   3375
      Top             =   4035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   4575
      Left            =   -15
      TabIndex        =   0
      Top             =   30
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   16750143
      ForeColorFixed  =   16777215
      BackColorSel    =   16775910
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
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
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuExportar 
         Caption         =   "Exportar..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const EX_SCROLL_FLX = 1

Private m_FormWidth As Integer
Private m_FormHeight As Integer
Private m_FlexWidth As Integer
Private m_FlexHeight As Integer

Private m_topRow As Integer
Private m_WndProcOrg As Long
Private m_HWndSubClassed As Long

Public sql_str As String

Private Sub Form_Load()

    On Error GoTo Handler
    
#If EX_SCROLL_FLX Then
    If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
        If Not m_WndProcOrg Then
            m_HWndSubClassed = Me.hWnd
            SetWindowLong hWnd, GWL_USERDATA, ObjPtr(Me)
            m_WndProcOrg = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProcSelectForm)
            m_topRow = 1
        End If
    End If
#End If
    
    m_FormWidth = Me.width
    m_FormHeight = Me.ScaleHeight
    
    With flxResults
        
        m_FlexWidth = .width
        m_FlexHeight = .height
        
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        
        .SelectionMode = flexSelectionByRow
        
        .RowHeight(0) = 315
        
    End With
    
Handler:
End Sub

Private Sub Form_Resize()

    On Error GoTo Handler

    If Me.width < 3000 Then
        Me.width = 3000
    End If

    flxResults.width = Me.width - m_FormWidth + m_FlexWidth
    
    If Me.height < 3000 Then
        Me.height = 3000
    End If
    
    flxResults.height = Me.ScaleHeight - m_FormHeight + m_FlexHeight
    
Handler:
End Sub

Private Sub mnuExportar_Click()
    gsub_FlxShowSaveAsDialog cmmdlg, flxResults, sql_str
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

'*******************************************************************************
' Funciones para el mouse wheel scroll
'
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

