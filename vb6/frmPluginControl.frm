VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPluginControl 
   Caption         =   "Administrador de plugins"
   ClientHeight    =   4140
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6270
   Icon            =   "frmPluginControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   7250
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
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPluginRegistrar 
         Caption         =   "&Registrar nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPluginsDescargar 
         Caption         =   "&Descargar"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuPluginsQuitar 
         Caption         =   "&Quitar del registro"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmPluginControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FormWidth As Integer
Private m_FormHeight As Integer
Private m_FlexWidth As Integer
Private m_FlexHeight As Integer

Private Sub flxResults_DblClick()
    EjecutarPlugin
End Sub

Private Sub flxResults_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        EjecutarPlugin
    End If
End Sub

Private Sub flxResults_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            EliminarPlugin
    End Select
End Sub

Private Sub Form_Load()
    '===================================================
    Dim k As Integer
    '===================================================
    On Error GoTo Handler
    
    Me.width = 8745
    m_FormWidth = Me.width
    m_FormHeight = Me.ScaleHeight
    
    With flxResults
    
        m_FlexWidth = .width
        m_FlexHeight = .height
        
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        
        'extension
        .Rows = 2
        .FixedRows = 1
        .Cols = 5
        .FixedCols = 1
        
        'ancho celdas
        .ColWidth(0) = 0
        .ColWidth(1) = 2250
        .ColWidth(2) = 900
        .ColWidth(3) = 1200
        .ColWidth(4) = 4155
        
        .RowHeight(0) = 315
    
        'titulos columna
        .Row = 0
        .Col = 1
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Plugin"
        .Col = 2
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Activo"
        .Col = 3
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Registrado"
        .Col = 4
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Descripción"
        
        .Row = 1
        
        'primera fila invisible
        .RowHeight(1) = 0
    
    End With
    
    frmR_Porter.FillPluginsInFlex flxResults
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "Form_Load"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.width < 3000 Then
        Me.width = 3000
    End If
    flxResults.width = Me.width - m_FormWidth + m_FlexWidth
    If Me.height < 3000 Then
        Me.height = 3000
    End If
    flxResults.height = Me.ScaleHeight - m_FormHeight + m_FlexHeight
End Sub

Private Sub mnuPluginRegistrar_Click()
    frmR_Porter.RegisterNewPlugin
End Sub

Private Sub mnuPluginsDescargar_Click()
    '(TODO)
End Sub

Private Sub mnuPluginsQuitar_Click()
    EliminarPlugin
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub EjecutarPlugin()
    '===================================================
    Dim oldRow As Long
    Dim id_plugin As Long
    '===================================================
    On Error GoTo Handler
    
    With flxResults

        If .RowHeight(.Row) = 0 Or .Row = 0 Then
            Exit Sub
        End If

        .Redraw = False
        oldRow = .TopRow
        .Col = 0
        id_plugin = CLng(.text)
        .Col = 1
        .ColSel = .Cols - 1
        .Redraw = True
        
        'elimina un incomodo movimiento cuando .row esta en las ultimas filas
        .TopRow = oldRow
        
        If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
            .TopRow = .Row
        End If

    End With

    frmR_Porter.ExecutePlugin id_plugin
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "EjecutarPlugin"
End Sub

Private Sub EliminarPlugin()

    Dim id_plugin As Long
    Dim plugin_eliminado As String
    Dim ret As Long
    Dim oldRow As Long
    On Error GoTo Handler
    
    With flxResults
            
        If .Row <> 0 And .RowHeight(.Row) <> 0 Then
    
            .Redraw = False
            oldRow = .TopRow
            .Col = 0
            id_plugin = CLng(.text)
            .Col = 1
            plugin_eliminado = Trim(.text)
            .Col = 1
            .ColSel = .Cols - 1
            .Redraw = True
            
            'elimina un incomodo movimiento cuando .row esta en las ultimas filas
            .TopRow = oldRow
            
            If vbYes = MsgBox("¿Estás seguro de quitar del registro el plugin:" & vbCrLf & _
                              "[" & plugin_eliminado & "]?", vbExclamation + vbYesNo, "Confirmar eliminación") Then
        
                ret = WritePrivateProfileString("Add-Ins32", plugin_eliminado, vbNullString, App.Path & "\R_porter.ini")
                If ret <> 0 Then
                    frmR_Porter.UnregisterPlugin id_plugin
                    If .Rows > 2 Then
                        oldRow = .Row
                        .RemoveItem (.Row)
                    Else
                        'volver invisible la primera fila
                        .RowHeight(1) = 0
                        Exit Sub
                    End If
                End If
            Else
                Exit Sub
            End If
                
            'seleccionar la fila anterior
            If oldRow > .Rows - 1 Then
                .Row = oldRow - 1
            Else
                .Row = oldRow
            End If
            
            .Col = 1
            .ColSel = .Cols - 1
                
            'verificar si la fila esta visible
            If (.RowHeight(.Row) > 0) Then
                If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
                    .TopRow = .Row
                End If
            End If
            .SetFocus
        End If
    End With
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "EliminarPlugin"
End Sub
