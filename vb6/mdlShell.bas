Attribute VB_Name = "mdlShell"
' =============================================================================
' CODIGO ORIGINAL:
' Brad Martinez http://www.mvps.org
'
' Demonstrates how to assign the system imagelists to the VB ListView,
' without having to use a VB Imagelist. Credit goes to Tom Esh (address
' unknown) for the initial idea of using LVM_SETITEM to set the system
' image indices for listview items.
'
' MODIFICADO POR: Esau Rodriguez para su explorador de archivos
Option Explicit

Public Const MAX_PATH = 260

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function FileIconInit Lib "shell32.dll" Alias "#660" (ByVal cmd As Boolean) As Boolean
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "User32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "User32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

' =============================================================================
' imagelist
Public Const LVS_SHAREIMAGELISTS = &H40

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
End Enum

Public Type LVITEM   ' was LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As Long  ' if String, must be pre-allocated before filled
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type

' LVITEM mask value
Public Const LVIF_IMAGE = &H2

' LVM_GETNEXTITEM lParam
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2

' LVITEM state and stateMask values
Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2

' LVM_GET/SETIMAGELIST wParam
Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1

Public Const WM_NOTIFY = &H4E
Public Const WM_DESTROY = &H2

Public Enum SHGFI_flags
    SHGFI_LARGEICON = &H0           ' sfi.hIcon is large icon
    SHGFI_SMALLICON = &H1           ' sfi.hIcon is small icon
    SHGFI_OPENICON = &H2            ' sfi.hIcon is open icon
    SHGFI_SHELLICONSIZE = &H4       ' sfi.hIcon is shell size (not system size), rtns BOOL
    SHGFI_PIDL = &H8                ' pszPath is pidl, rtns BOOL
    SHGFI_USEFILEATTRIBUTES = &H10  ' pretend pszPath exists, rtns BOOL
    SHGFI_ICON = &H100              ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
    SHGFI_DISPLAYNAME = &H200       ' isf.szDisplayName is filled, rtns BOOL
    SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
    SHGFI_ATTRIBUTES = &H800        ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
    SHGFI_ICONLOCATION = &H1000     ' fills sfi.szDisplayName with filename
                                                          ' containing the icon, rtns BOOL
    SHGFI_EXETYPE = &H2000          ' rtns two ASCII chars of exe type
    SHGFI_SYSICONINDEX = &H4000     ' sfi.iIcon is sys il icon index, rtns hImagelist
    SHGFI_LINKOVERLAY = &H8000      ' add shortcut overlay to sfi.hIcon
    SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
End Enum

' =============================================================================
' SHGetFileInfo
Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As SHGFI_flags) As Long
                              
' =============================================================================
' ListView messages & macros
'
Public Const LVM_FIRST = &H1000
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)

' =============================================================================
' variables estaticas
Private m_bolSytemDirLoaded As Boolean  ' para indicar ruta ya cargada
Private m_strSystemDir As String * MAX_PATH  ' para almacenar la ruta del sistema

' Returns the handle of the small or large icon system imagelist.
'   uFlags - either SHGFI_SMALLICON or SHGFI_LARGEICON
Public Function GetSystemImagelist(uFlags As Long) As Long
    Dim sfi As SHFILEINFO
    ' Any valid file system path can be used to retrieve system image list handles.
    GetSystemImagelist = SHGetFileInfo("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uFlags)
End Function

Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
    ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, iImageList, ByVal himl)
End Function

Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
    ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
End Function

Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
    ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags) 'MAKELPARAM(flags, 0))
End Function

Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
    ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal Abs(fPartialOK)) 'MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As Long, mask As Long) As Boolean
    Dim lvi As LVITEM
    lvi.state = state
    lvi.stateMask = mask
    ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

' Returns the index of the item that is selected and has the focus rectangle (user-defined macro)
Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
    ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function

' Selects the specified item and gives it the focus rectangle.
' If the listview is multiselect (not LVS_SINGLESEL), does not
' de-select any currently selected items (user-defined macro)
Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
    ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, LVIS_FOCUSED Or LVIS_SELECTED)
End Function

' Returns the index of a file's small or large, normal or open icon.
'   sFile   - can be either a file's absolute path, or a nonexistent file or extension
'   uFlags  - either SHGFI_SMALLICON or SHGFI_LARGEICON, and SHGFI_OPENICON
Public Function GetFileIconIndex(sFile As String, uFlags As SHGFI_flags) As Long
    Dim sfi As SHFILEINFO
    sfi.iIcon = -1
    If SHGetFileInfo(sFile, 0, sfi, Len(sfi), uFlags Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES) Then
        GetFileIconIndex = sfi.iIcon
    End If
End Function

' Returns a file's typename.
'   sFile      - can be either a file's absolute path, or a nonexistent file or extension
Public Function GetFileTypeName(sFile As String) As String
    Dim sfi As SHFILEINFO
    If SHGetFileInfo(sFile, 0, sfi, Len(sfi), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
        GetFileTypeName = GetStrFromBufferA(sfi.szTypeName)
    End If
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
    If InStr(sz, vbNullChar) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
    Else
        ' If sz had no null char, the Left$ function
        ' above would return a zero length string ("").
        GetStrFromBufferA = sz
    End If
End Function

' =============================================================================
' SUBCLASIFICACION
Public Function SubClass(hWnd As Long, _
                         lpfnNew As Long, _
                         Optional objNotify As Object = Nothing) As Boolean
    '================================
    Dim lpfnOld As Long
    Dim fSuccess As Boolean
    '================================
    
    On Error GoTo Handler
    
    If GetProp(hWnd, OLDWNDPROC) Then
        SubClass = True
        Exit Function
    End If
    
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)
    
    If lpfnOld Then
        fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
        If (objNotify Is Nothing) = False Then
            fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
        End If
    End If
    
Handler:

    If fSuccess Then
        SubClass = True
    Else
        If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
        MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
               "Err# " & Err.Number & ": " & Err.Description, vbExclamation
    End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
    '================================
    Dim lpfnOld As Long
    '================================
    
    lpfnOld = GetProp(hWnd, OLDWNDPROC)
    
    If lpfnOld Then
        If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
            Call RemoveProp(hWnd, OLDWNDPROC)
            Call RemoveProp(hWnd, OBJECTPTR)
            UnSubClass = True
        End If
    End If
End Function

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg
    ' ======================================================
    ' Prevent the ListView from removing our system imagelist assignment,
    ' which it wil do when it sees no VB ImageList associated with it.
    ' (the ListView can't be subclassed when we're assigning imagelists...)
    Case LVM_SETIMAGELIST
      Exit Function
                
    ' ======================================================
    ' Unsubclass the window.
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
  
End Function

' =============================================================================
' Muestra el icono correspondiente al archivo
Public Sub AddFileIconToListView(hwndLV As Long, strFileName As String, _
                                ByRef lstView As Object, typeHeaderIndex As Integer)
   '================================
    Dim lvi As LVITEM
    Dim sTypeName As String
    Dim li As Object
   '================================
    lvi.mask = LVIF_IMAGE
    
    sTypeName = GetFileTypeName(strFileName)
    
    ' Add a ListItem and SubItem for the icon
    Set li = lstView.ListItems.Add(text:=strFileName)
    li.SubItems(typeHeaderIndex) = Trim(sTypeName)
    
    ' Set the ListItem's image to that of its corresponding icon index.
    ' (real listview and imagelist items are zero-based, the indices of
    ' the icons in the system's small and large imagelists are the same).
    lvi.iItem = li.Index - 1
    lvi.iImage = GetFileIconIndex(strFileName, SHGFI_SMALLICON)
    
    Call ListView_SetItem(hwndLV, lvi)
    
    Call ListView_SetSelectedItem(hwndLV, 0)
    Call ListView_EnsureVisible(hwndLV, 0, False)
    
End Sub

' =============================================================================
' Muestra el icono de CD del sistema
Public Sub AddMediumIconToListView(hwndLV As Long, strFileName As String, _
                                   ByRef lstView As Object)
   '================================
    Dim sfi As SHFILEINFO
    Dim lvi As LVITEM
    Dim sTypeName As String
    Dim li As Object
   '================================
    lvi.mask = LVIF_IMAGE
    
    ' Add a ListItem and SubItem for the icon
    Set li = lstView.ListItems.Add(text:=GetStrFromBufferA(strFileName))
    lvi.iItem = li.Index - 1

    ' Temporary hack (WinXP only?), I can't make the listview show my stoirage icon... :(
    lvi.iImage = 11
    
    Call ListView_SetItem(hwndLV, lvi)
    Call ListView_SetSelectedItem(hwndLV, 0)
    Call ListView_EnsureVisible(hwndLV, 0, False)
End Sub

' =============================================================================
' Muestra el icono de folder del sistema
Public Sub AddFolderIconToListView(hwndLV As Long, strFileName As String, _
                                   ByRef lstView As Object)
   '================================
    Dim sfi As SHFILEINFO
    Dim lvi As LVITEM
    Dim sTypeName As String
    Dim li As Object
   '================================
    lvi.mask = LVIF_IMAGE
    
    ' Add a ListItem and SubItem for the icon
    Set li = lstView.ListItems.Add(text:=GetStrFromBufferA(strFileName))
    
    ' Set the ListItem's image to that of its corresponding icon index.
    ' (real listview and imagelist items are zero-based, the indices of
    ' the icons in the system's small and large imagelists are the same).
    lvi.iItem = li.Index - 1

    ' Pedimos el indice del icono perteneciente al directorio del sistema
    ' (que siempre deberia existir)
    GetStaticSystemDirectory
    If SHGetFileInfo(m_strSystemDir, 0, sfi, Len(sfi), SHGFI_SMALLICON Or SHGFI_SYSICONINDEX) Then
        lvi.iImage = sfi.iIcon
    End If
    
    Call ListView_SetItem(hwndLV, lvi)
    Call ListView_SetSelectedItem(hwndLV, 0)
    Call ListView_EnsureVisible(hwndLV, 0, False)
End Sub

' Se usa para guardar en la variable estatica global [m_strSystemDir] la ruta del sistema
Private Sub GetStaticSystemDirectory()
    ' inicialmente en False por ser variable a nivel de modulo
    If Not m_bolSytemDirLoaded Then
        If GetWindowsDirectory(m_strSystemDir, MAX_PATH) Then
            ' para no volver a llamar a esta funcion
            m_bolSytemDirLoaded = True
        End If
    End If
End Sub
