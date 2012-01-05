Attribute VB_Name = "mslScroll"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_MOUSEWHEELPRESENT = 75
''Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Public Const WM_MOUSEWHEEL = &H20A
Public Const WHEEL_DELTA = 120

Private mySelectForm As frmSelect           ' no se puede declarar este objeto como Form (se cae el IDE)
Private myExDynaTable As frmExDynaTable
Private m_ptrSelectForm As Long
Private m_ptrExDynaTable As Long

Public lpPrevWndProcDataControl As Long

Public Function WndProcDataControlForm(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If uMsg = WM_MOUSEWHEEL Then
        ' scroll del flexgrid tenga o no el enfoque...
        If (HiWord(wParam) / WHEEL_DELTA) < 0 Then
            frmDataControl.ScrollDown
        Else
            frmDataControl.ScrollUp
        End If
    End If
    WndProcDataControlForm = CallWindowProc(lpPrevWndProcDataControl, hWnd, uMsg, wParam, lParam)
End Function

Public Function WndProcSelectForm(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    m_ptrSelectForm = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory mySelectForm, m_ptrSelectForm, 4
    WndProcSelectForm = mySelectForm.WindowProc(hWnd, uMsg, wParam, lParam)
    CopyMemory mySelectForm, 0&, 4
    Set mySelectForm = Nothing
End Function

Public Function WndProcExDynaTableForm(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    m_ptrExDynaTable = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory myExDynaTable, m_ptrExDynaTable, 4
    WndProcExDynaTableForm = myExDynaTable.WindowProc(hWnd, uMsg, wParam, lParam)
    CopyMemory myExDynaTable, 0&, 4
    Set myExDynaTable = Nothing
End Function

Public Function HiWord(dw As Long) As Integer
    If dw And &H80000000 Then
        HiWord = (dw \ 65535) - 1
    Else
        HiWord = dw \ 65535
    End If
End Function

