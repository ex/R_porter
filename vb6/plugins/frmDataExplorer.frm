VERSION 5.00
Object = "{40EF20E1-7EC5-11D8-95A1-9655FE58C763}#2.0#0"; "exSplit.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDataExplorer 
   Caption         =   "Explorador de BD"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frmDataExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   6195
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView TreeView 
      Height          =   6165
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2965
      _ExtentX        =   5239
      _ExtentY        =   10874
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglst"
      Appearance      =   1
   End
   Begin ComctlLib.ListView ListView 
      Height          =   6165
      Left            =   3010
      TabIndex        =   0
      Top             =   0
      Width           =   6949
      _ExtentX        =   12250
      _ExtentY        =   10874
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin exSplit.SplitRegion Split 
      Height          =   6165
      Left            =   0
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   10874
      FirstControl    =   "TreeView"
      SecondControl   =   "ListView"
      SplitPercent    =   30
      SplitterBarVertical=   -1  'True
      SplitterBarThickness=   45
      MouseIcon       =   "frmDataExplorer.frx":058A
      MousePointer    =   99
   End
   Begin ComctlLib.ImageList imglst 
      Left            =   6270
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":06EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":0C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":0F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":12E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":1634
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDataExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mo_Node As Node
Private mo_Panel As Panel
Private mo_ListItem As ListItem
Private m_SplitWidth As Integer
Private m_SplitHeight As Integer
Private m_Width As Integer
Private m_Height As Integer

' indices de headers del listview
Private m_sizeHeader As Integer
Private m_byteSizeHeader As Integer
Private m_typeHeader As Integer
Private m_dateHeader As Integer
Private m_dateSortedHeader As Integer

' Handle of the ListView
Private m_hwndLV As Long

' Handle of the system small icon imagelists
Private m_himlSysSmall As Long

Private Sub Form_Load()
    '===================================================
    Dim mo_ColumnHeader As ColumnHeader
    '===================================================
   
    Dim dwStyle As Long
    
    ' Initialize the imagelist with an undoc shell call. Is only necessary for
    ' NT4 where the app gets an uninitialized system imagelist copy. See:
    ' http://www.geocities.com/SiliconValley/4942/iconcache.html
    ' [WARN] Call is not exported in stock Win95's Shell32.dll
    On Error Resume Next
    Call FileIconInit(True)
    On Error GoTo 0
    
    With ListView
        
        .View = lvwReport

        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Nombre"
        mo_ColumnHeader.Width = .Width / (10 / 5)
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Tamaño"
        mo_ColumnHeader.Width = .Width / (10 / 1.5)
        mo_ColumnHeader.Alignment = lvwColumnRight
        m_sizeHeader = 1

        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Tipo"
        mo_ColumnHeader.Width = .Width / (10 / 1)
        m_typeHeader = 2
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Fecha"
        mo_ColumnHeader.Width = .Width / (10 / 2.3)
        m_dateHeader = 3
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Bytes"
        mo_ColumnHeader.Width = 0
        m_byteSizeHeader = 4
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.Text = "Date"
        mo_ColumnHeader.Width = 0
        m_dateSortedHeader = 5
        
        m_hwndLV = .hWnd
        
    End With
  
    ' Fist tell the ListView that it will share the imagelist assigned to it, so that
    ' the ListView does not destroy the imagelist when it is itself destroyed.
    dwStyle = GetWindowLong(m_hwndLV, GWL_STYLE)
    If ((dwStyle And LVS_SHAREIMAGELISTS) = False) Then
        Call SetWindowLong(m_hwndLV, GWL_STYLE, dwStyle Or LVS_SHAREIMAGELISTS)
    End If
    
    ' Next get the handle of the system's small icon imagelists
    ' in the module level variables
    m_himlSysSmall = GetSystemImagelist(SHGFI_SMALLICON)

    If (m_himlSysSmall <> 0) Then
    
        ' Assign the respective handle of the imagelist to the ListView.
        ' As far as the VB ListView's internal code is concerned, it's not using
        ' any imagelist, the ListItem icon propertie will return empty.
        Call ListView_SetImageList(m_hwndLV, m_himlSysSmall, LVSIL_SMALL)
        
        ' The only reason we need to subclass the ListView is to prevent it from
        ' removing our system imagelis assignment (which it will do if left unchecked...)
        Call SubClass(m_hwndLV, AddressOf WndProc)
    
    End If
   
    m_SplitWidth = Split.Width
    m_SplitHeight = Split.Height
    m_Width = Me.Width
    m_Height = Me.Height

    StatusBar.Panels(1).Text = "Listo"
    StatusBar.Panels(1).Width = 2610
    Set mo_Panel = StatusBar.Panels.Add()
    mo_Panel.Key = "Size"
    mo_Panel.AutoSize = sbrSpring
    mo_Panel.Alignment = sbrRight
    
    CrearArbol
    
End Sub

Private Sub CrearArbol()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================
    On Error GoTo Handler

    TreeView.Nodes.Clear
    Set mo_Node = TreeView.Nodes.Add(, , "R-DNS", gs_DNS, 1, 1)

    Set rs = New ADODB.Recordset

    '*****************************
    'cargar los medios disponibles
    '*****************************
    query = "SELECT id_storage, name FROM storage WHERE (id_storage > 0) ORDER BY name"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        Set mo_Node = TreeView.Nodes.Add("R-DNS", tvwChild, "S-" & rs!id_storage, rs!Name, 2, 2)
        mo_Node.EnsureVisible
        '#### quite nodo auxiliar (testing) 12/2005
        'Set mo_Node = TreeView.Nodes.Add("M-" & rs!id_storage, tvwChild, "S-" & rs!id_storage, rs!Name, 5, 5)
        rs.MoveNext
    Wend
    rs.Close

    Exit Sub

Handler:    MsgBox Err.Description, vbExclamation, "CrearArbol"
End Sub

' It is only necessary to remove the system imagelist associations from
' the ListView if its LVS_SHAREIMAGELISTS style is not set. If the bit
' was not set and we didn't do this, the ListView would destroy both
' imagelists (and that's why processes get system imagelist copies on NT)
Private Sub Form_Unload(Cancel As Integer)
  Call UnSubClass(m_hwndLV)
  Call ListView_SetImageList(m_hwndLV, 0, LVSIL_SMALL)
  Call ListView_SetImageList(m_hwndLV, 0, LVSIL_NORMAL)
End Sub

Private Sub ListView_DblClick()
    '===================================================
    Dim strQuery As String
    Dim lngKey As Long
    '===================================================
    
    On Error GoTo Handler
    
    If Mid(mo_ListItem.Tag, 1, 2) = "D-" Then
    
        lngKey = CLng(Mid(TreeView.Nodes(mo_ListItem.Tag).Key, 3))
        
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_sys_parent=" & lngKey & ") AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        
        TreeView.Nodes(mo_ListItem.Tag).Selected = True
        
        If TreeView.Nodes(mo_ListItem.Tag).Tag = "" Then
            TreeView.Nodes(mo_ListItem.Tag).Tag = "X"
            CrearVistaDetalle TreeView.Nodes(mo_ListItem.Tag).Key, strQuery, True
        Else
            CrearVistaDetalle TreeView.Nodes(mo_ListItem.Tag).Key, strQuery, False
        End If
    End If
    
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "CrearArbol"
End Sub

Private Sub ListView_ItemClick(ByVal Item As ComctlLib.ListItem)
    Set mo_ListItem = Item
End Sub

Private Sub ListView_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set mo_ListItem = ListView.SelectedItem
        ListView_DblClick
    End If
End Sub

Private Sub TreeView_NodeClick(ByVal Node As ComctlLib.Node)
    '===================================================
    Dim strQuery As String
    Dim lngKey As Long
    '===================================================
    On Error GoTo Handler
    
    If (Mid(Node.Key, 1, 2) = "R-") Then
        ListView.ListItems.Clear
        Exit Sub
    End If
    
    lngKey = CLng(Mid(Node.Key, 3))
    
    If (Mid(Node.Key, 1, 2) = "S-") Then
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_storage=" & lngKey & ") AND (file.id_sys_parent=0) AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        GoTo Process
    End If
    If Mid(Node.Key, 1, 2) = "D-" Then
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_sys_parent=" & lngKey & ") AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        GoTo Process
    End If
    
    ListView.ListItems.Clear
    
    Exit Sub
     
Process:
    If Node.Tag = "" Then
        Node.Tag = "X"
        CrearVistaDetalle Node.Key, strQuery, True
    Else
        CrearVistaDetalle Node.Key, strQuery, False
    End If
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "TreeView_NodeClick"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Split.Width = m_SplitWidth + Me.Width - m_Width
    Split.Height = m_SplitHeight + Me.Height - m_Height
End Sub

Private Sub CrearVistaDetalle(ByVal strNodeKey As String, ByVal strQuery As String, ByVal bolAddDirs As Boolean)
    '===================================================
    Dim rs As ADODB.Recordset
    Dim lngKey As Long
    Dim lngFileSize As Long
    Dim strFileSize  As String
    Dim lngNumFiles As Double
    Dim dblTotSize As Double
    Dim strTotSize As String
    Dim strKeyListItem As String
    Dim strFileName As String
    Dim dateFile As Date
    Dim strFileDate As String
    '===================================================
    On Error GoTo Handler
    
    ' necessary because the items don't display correctly if has been sorted...
    ListView.Sorted = False
    
    ListView.ListItems.Clear
    
    lngKey = CLng(Mid(strNodeKey, 3))
    
    Set rs = New ADODB.Recordset

    rs.Open strQuery, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        
        lngNumFiles = lngNumFiles + 1
        
        strFileName = gfnc_GetFileNameWithoutPath(rs!sys_Name)
        dateFile = rs!fecha
        strFileDate = Format(dateFile, "dd/mm/yyyy") & " " & Format(dateFile, "medium time")
        
        If rs!file_type = "<DIR>" Then
            
            strKeyListItem = "D-" & rs!id_file
            
            AddFolderIconToListView m_hwndLV, strFileName, ListView
            
            ListView.ListItems(ListView.ListItems.Count).Text = strFileName
            ListView.ListItems(ListView.ListItems.Count).Tag = strKeyListItem
            
            If bolAddDirs Then
                'agregar directorio al arbol
                strFileName = gfnc_GetFileNameWithoutPath(rs!sys_Name)
                Set mo_Node = TreeView.Nodes.Add(strNodeKey, tvwChild, strKeyListItem, strFileName, 3, 3)
            End If
            
            ListView.ListItems(ListView.ListItems.Count).SubItems(m_sizeHeader) = " "
            ListView.ListItems(ListView.ListItems.Count).SubItems(m_byteSizeHeader) = " "
        Else

            AddFileIconToListView m_hwndLV, strFileName, ListView, m_typeHeader
            
            ' mostrar tamaño adecuadamente
            lngFileSize = CLng(rs!sys_length)
            dblTotSize = dblTotSize + lngFileSize
            strFileSize = Format(lngFileSize / 1024, "#,###,##0") & " KB"
            
            ListView.ListItems(ListView.ListItems.Count).SubItems(m_sizeHeader) = strFileSize
            ListView.ListItems(ListView.ListItems.Count).SubItems(m_byteSizeHeader) = Format(lngFileSize, "0000000000")
            
        End If
        
        ListView.ListItems(ListView.ListItems.Count).SubItems(m_dateHeader) = strFileDate
        ListView.ListItems(ListView.ListItems.Count).SubItems(m_dateSortedHeader) = Format(dateFile, "yyyy/mm/dd Hh:Nn:Ss")
        
        rs.MoveNext
        
    Wend
    
    rs.Close
    
    ' informacion del status bar
    StatusBar.Panels(1).Text = " " & lngNumFiles & " Elementos"
    If (dblTotSize > 1048576) Then
        strTotSize = Format(dblTotSize / 1048576, "0.00") & " MB"
    Else
        strTotSize = Format(dblTotSize / 1024, "0.00") & " KB"
    End If
    StatusBar.Panels(2).Text = "Total: " & strTotSize

    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "CrearVistaDetalle"
End Sub

'------------------------------------------------------------------------------------------------
' Code added to sort listView by header column
Private Sub ListView_ColumnClick(ByVal colHeader As ComctlLib.ColumnHeader)
    SortListView Me.ListView, colHeader
End Sub

Private Sub SortListView(ByRef lvwControl As ComctlLib.ListView, ByRef colHeader As ComctlLib.ColumnHeader)
    If (colHeader.Index = m_sizeHeader + 1) Then
        lvwControl.SortKey = m_byteSizeHeader
    Else
        If (colHeader.Index = m_dateHeader + 1) Then
            lvwControl.SortKey = m_dateSortedHeader
        Else
            lvwControl.SortKey = colHeader.Index - 1
        End If
    End If
    lvwControl.Sorted = True
    lvwControl.SortOrder = 1 Xor lvwControl.SortOrder  ' toogle sort order
End Sub

