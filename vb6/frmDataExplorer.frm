VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{40EF20E1-7EC5-11D8-95A1-9655FE58C763}#2.0#0"; "exSplit.ocx"
Begin VB.Form frmDataExplorer 
   Caption         =   "Explorador de BD"
   ClientHeight    =   6465
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9960
   Icon            =   "frmDataExplorer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":06EC
            Key             =   "'medium'"
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
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDataExplorer.frx":1B86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuViewAll 
         Caption         =   "Ver &todos"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewByCategory 
         Caption         =   "Ver por categoria"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmDataExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mo_Node As Object
Private mo_Panel As Object
Private mo_ListItem As Object

Private m_SplitWidth As Integer
Private m_SplitHeight As Integer
Private m_width As Integer
Private m_height As Integer

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

' Groups the medias by type
Private m_groupMode As Boolean

Private Sub mnuViewAll_Click()
    If m_groupMode Then
        m_groupMode = False
        mnuViewByCategory.Checked = False
        mnuViewAll.Checked = True
        createTreeView
        ListView.ListItems.Clear
    End If
End Sub

Private Sub mnuViewByCategory_Click()
    If Not m_groupMode Then
        m_groupMode = True
        mnuViewByCategory.Checked = True
        mnuViewAll.Checked = False
        createTreeView
        ListView.ListItems.Clear
    End If
End Sub

Private Sub Form_Load()
    
    Dim mo_ColumnHeader As Object
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
        mo_ColumnHeader.text = "Nombre"
        mo_ColumnHeader.width = .width / (10 / 5)
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.text = "Tamaño"
        mo_ColumnHeader.width = .width / (10 / 1.5)
        mo_ColumnHeader.Alignment = lvwColumnRight
        m_sizeHeader = 1

        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.text = "Tipo"
        mo_ColumnHeader.width = .width / (10 / 1)
        m_typeHeader = 2
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.text = "Fecha"
        mo_ColumnHeader.width = .width / (10 / 2.3)
        m_dateHeader = 3
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.text = "Bytes"
        mo_ColumnHeader.width = 0
        m_byteSizeHeader = 4
        
        Set mo_ColumnHeader = .ColumnHeaders.Add()
        mo_ColumnHeader.text = "Date"
        mo_ColumnHeader.width = 0
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
   
    m_SplitWidth = Split.width
    m_SplitHeight = Split.height
    m_width = Me.width
    m_height = Me.height

    StatusBar.Panels(1).text = "Listo"
    StatusBar.Panels(1).width = 2610
    Set mo_Panel = StatusBar.Panels.Add()
    mo_Panel.key = "Size"
    mo_Panel.AutoSize = sbrSpring
    mo_Panel.Alignment = sbrRight
    
    m_groupMode = True
    mnuViewByCategory.Checked = True
    mnuViewAll.Checked = False
    
    createTreeView
    
End Sub

Private Sub createTreeView()
    
    Dim rs As ADODB.Recordset
    Dim rsMedia As ADODB.Recordset
    On Error GoTo Handler

    TreeView.Nodes.Clear
    Set mo_Node = TreeView.Nodes.Add(, , "R-DSN", gs_DSN, 1, 1)

    Set rs = New ADODB.Recordset

    '*****************************
    'cargar los medios disponibles
    '*****************************
    If m_groupMode Then
        query = "SELECT DISTINCT category.category, category.id_category FROM category, storage WHERE ((category.id_category > 0) AND (category.id_category = storage.id_category)) ORDER BY category.category"
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

        While rs.EOF = False
            Set mo_Node = TreeView.Nodes.Add("R-DSN", tvwChild, "C-" & rs!id_category, rs!category, 6, 6)
            mo_Node.EnsureVisible
            
            query = "SELECT id_storage, name, fecha FROM storage WHERE ((id_storage > 0) AND (id_category = " & rs!id_category & ")) ORDER BY fecha DESC"
            Set rsMedia = New ADODB.Recordset
            rsMedia.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            While rsMedia.EOF = False
                Set mo_Node = TreeView.Nodes.Add("C-" & rs!id_category, tvwChild, "S-" & rsMedia!id_storage, rsMedia!Name, 2, 2)
                rsMedia.MoveNext
            Wend
            rsMedia.Close
            
            rs.MoveNext
        Wend
        rs.Close
    Else
        query = "SELECT id_storage, name FROM storage WHERE (id_storage > 0) ORDER BY name"
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        While rs.EOF = False
            Set mo_Node = TreeView.Nodes.Add("R-DSN", tvwChild, "S-" & rs!id_storage, rs!Name, 2, 2)
            rs.MoveNext
        Wend
        rs.Close
        mo_Node.EnsureVisible
    End If

    Exit Sub

Handler:    MsgBox Err.Description, vbExclamation, "createTreeView"
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
    
    Dim strQuery As String
    Dim lngKey As Long
    On Error GoTo Handler
    
    If Mid(mo_ListItem.Tag, 1, 2) = "D-" Then
    
        lngKey = CLng(Mid(TreeView.Nodes(mo_ListItem.Tag).key, 3))
        
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_sys_parent=" & lngKey & ") AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        
        TreeView.Nodes(mo_ListItem.Tag).Selected = True
        
        If TreeView.Nodes(mo_ListItem.Tag).Tag = "" Then
            TreeView.Nodes(mo_ListItem.Tag).Tag = "X"
            CrearVistaDetalle TreeView.Nodes(mo_ListItem.Tag).key, strQuery, True
        Else
            CrearVistaDetalle TreeView.Nodes(mo_ListItem.Tag).key, strQuery, False
        End If
    End If
    
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "createTreeView"
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

Private Sub TreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnupopup, 2
    End If
End Sub

Private Sub TreeView_NodeClick(ByVal Node As ComctlLib.Node)
    
    Dim strQuery As String
    Dim lngKey As Long
    On Error GoTo Handler
    
    If (Mid(Node.key, 1, 2) = "R-") Then
        ListView.ListItems.Clear
        Exit Sub
    End If
    
    lngKey = CLng(Mid(Node.key, 3))
    
    If (Mid(Node.key, 1, 2) = "S-") Then
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_storage=" & lngKey & ") AND (file.id_sys_parent=0) AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        GoTo Process
    ElseIf Mid(Node.key, 1, 2) = "D-" Then
        strQuery = "SELECT file.id_file, file.sys_name, file.sys_length, file_type.file_type, file.fecha FROM file, file_type WHERE ((file.id_sys_parent=" & lngKey & ") AND (file_type.id_file_type=file.id_file_type)) ORDER BY file_type.file_type ASC, file.name ASC"
        GoTo Process
    ElseIf Mid(Node.key, 1, 2) = "C-" Then
        strQuery = "SELECT storage.id_storage, storage.name, storage_type.storage_type, storage.fecha FROM storage, storage_type WHERE ((storage.id_category=" & lngKey & ") AND (storage.id_storage_type=storage_type.id_storage_type)) ORDER BY storage.name ASC"
        ShowCategoryView Node.key, strQuery
        Exit Sub
    End If
    
    ListView.ListItems.Clear
    Exit Sub
     
Process:
    If Node.Tag = "" Then
        Node.Tag = "X"
        CrearVistaDetalle Node.key, strQuery, True
    Else
        CrearVistaDetalle Node.key, strQuery, False
    End If
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "TreeView_NodeClick"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Split.width = m_SplitWidth + Me.width - m_width
    Split.height = m_SplitHeight + Me.height - m_height
End Sub

Private Sub ShowCategoryView(ByVal strNodeKey As String, ByVal strQuery As String)
    Dim rs As ADODB.Recordset
    Dim lngKey As Long
    Dim lngNumMediums As Double
    Dim strKeyListItem As String
    Dim strMediumName As String
    Dim dateMedium As Date
    Dim strMediumDate As String
    On Error GoTo Handler
    
    ' necessary because the items don't display correctly if has been sorted...
    ListView.Sorted = False
    ListView.ListItems.Clear
    lngKey = CLng(Mid(strNodeKey, 3))
    lngNumMediums = 0
    
    Set rs = New ADODB.Recordset
    rs.Open strQuery, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        lngNumMediums = lngNumMediums + 1
        
        strMediumName = rs!Name
        dateMedium = rs!Fecha
        strMediumDate = Format(dateMedium, "dd/mm/yyyy") & " " & Format(dateMedium, "medium time")
        
        strKeyListItem = "S-" & rs!id_storage
        
        AddMediumIconToListView m_hwndLV, strMediumName, ListView
       
        ListView.ListItems(ListView.ListItems.Count).text = strMediumName
        ListView.ListItems(ListView.ListItems.Count).Tag = strKeyListItem
        
        ListView.ListItems(ListView.ListItems.Count).SubItems(m_typeHeader) = rs!storage_type
        
        ListView.ListItems(ListView.ListItems.Count).SubItems(m_dateHeader) = strMediumDate
        ListView.ListItems(ListView.ListItems.Count).SubItems(m_dateSortedHeader) = Format(dateMedium, "yyyy/mm/dd Hh:Nn:Ss")
        
        rs.MoveNext
    Wend
    rs.Close
    
    ' informacion del status bar
    StatusBar.Panels(1).text = " " & lngNumMediums & " Medios"

    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "ShowCategoryView"
End Sub

Private Sub CrearVistaDetalle(ByVal strNodeKey As String, ByVal strQuery As String, ByVal bolAddDirs As Boolean)
    
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
    On Error GoTo Handler
    
    ' necessary because the items don't display correctly if has been sorted...
    ListView.Sorted = False
    ListView.ListItems.Clear
    lngKey = CLng(Mid(strNodeKey, 3))
    lngNumFiles = 0
    
    Set rs = New ADODB.Recordset
    rs.Open strQuery, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        lngNumFiles = lngNumFiles + 1
        
        strFileName = gfnc_GetFileNameWithoutPath(rs!Sys_Name)
        dateFile = rs!Fecha
        strFileDate = Format(dateFile, "dd/mm/yyyy") & " " & Format(dateFile, "medium time")
        
        If rs!file_type = "<DIR>" Then
            strKeyListItem = "D-" & rs!id_file
            
            AddFolderIconToListView m_hwndLV, strFileName, ListView
            
            ListView.ListItems(ListView.ListItems.Count).text = strFileName
            ListView.ListItems(ListView.ListItems.Count).Tag = strKeyListItem
            
            If bolAddDirs Then
                'agregar directorio al arbol
                strFileName = gfnc_GetFileNameWithoutPath(rs!Sys_Name)
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
    StatusBar.Panels(1).text = " " & lngNumFiles & " Elementos"
    If (dblTotSize > 1048576) Then
        strTotSize = Format(dblTotSize / 1048576, "0.00") & " MB"
    Else
        strTotSize = Format(dblTotSize / 1024, "0.00") & " KB"
    End If
    StatusBar.Panels(2).text = "Total: " & strTotSize

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

