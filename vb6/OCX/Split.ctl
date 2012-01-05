VERSION 5.00
Begin VB.UserControl SplitRegion 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   6  'Inside Solid
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   19
   ToolboxBitmap   =   "Split.ctx":0000
   Begin VB.Label lblSplitterBar 
      Height          =   735
      Left            =   75
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "SplitRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Control Split (Para cambiar el tamaño de dos controles)
'*******************************************************************************
' Revisado:     esau 10 2001
'               Le agrege la propiedad <MousePointer> y <MouseIcon>
'               para poner un cursor a mi medida.
'*******************************************************************************
' Fuente Original:
'*******************************************************************************
'* CONTROL: SplitRegion
'*    The SplitRegion control provides a display region that is split between
'*    two controls.  The user can then drag the splitter bar to adjust the size
'*    of the two controls within the SplitRegion.  Similarly, the program can
'*    adjust the split percent to control the size of these two controls.  For
'*    performance reasons the SplitRegion is not a container control.  Instead,
'*    the programmer places the SplitRegion and then sets the FirstControl and
'*    the SecondControl properties to the names of the controls to be displayed
'*    within the SplitRegion.
'*
'* PUBLIC PROPERTIES:
'*    FirstControl               Name of the control to occupy the top or left
'*                               position of the split region (R/W)
'*    SecondControl              Name of the control to occupy the bottom or
'*                               right of the split region (R/W)
'*    FirstControlMinSize        Minimum size of the first control in scale
'*                               units of the container (R/W)
'*    SecondControlMinSize       Minimum size of the second control in scale
'*                               units of the container (R/W)
'*    SplitPercent               Percent of shared region currently occupied by
'*                               first control (R/W)
'*    SplitterBarVertical        True if the splitter bar is vertical, false if
'*                               horizontal (R/W)
'*    SplitterBarThickness       Thickness of the splitter bar in scale units of
'*                               the container (R/W)
'*    SplitterBarColor           Background color of the splitter bar (R/W)
'*    AllowControlHiding         True if a control becomes hidden when its
'*                               portion of the split region is reduced to less
'*                               than the control's minimum size. (R/W)
'*    KeepSplitPercentOnResize   Determines if the split percent is kept or if
'*                               split position is kept during resize (R/W)
'*    RightToLeft                Determines the position of the first and second
'*                               controls when the splitter bar is vertical (for
'*                               bi-directional language compatibility) (R/W)
'*
'*    MousePointer               Debe ser 99 para poner el cursor que quieras
'*    MouseIcon                  Imagen del Cursor para el Split
'*
'* PUBLIC METHODS:
'*    Refresh  Refreshes the display of the SplitRegion. (Call whenever the
'*             SplitRegion has been moved but not resized.)
'*
'* PUBLIC EVENTS:
'*    RepositionSplit   Occurs after the splitter bar is moved
'*    Resize            Occurs after the split region is resized but before the
'*                      split is re-adjusted
'*
'* VERSIONS:
'*    1.00  8/2/97      Matthew Carroll
'*******************************************************************************

Option Explicit
Option Compare Text

'*******************************************************************************
'Declaraciones
'*******************************************************************************

Private Const m_iErrInvalidPropertyValue = 380

'* Default values
Private Const m_nControlMinSizeDefault = 400#
Private Const m_nSplitPercentDefault = 50#
Private Const m_fSplitterBarVerticalDefault = False
Private Const m_nSplitterBarThicknessDefault = 100#
Private Const m_fAllowControlHidingDefault = False
Private Const m_fKeepSplitPercentOnResizeDefault = False

'* Property values
Private m_sFirstControl As String
Private m_sSecondControl As String
Private m_nFirstControlMinSize As Single
Private m_nSecondControlMinSize As Single
Private m_nSplitPercent As Single
Private m_fSplitterBarVertical As Boolean
Private m_nSplitterBarThickness As Single
Private m_iSplitterBarThicknessPixels As Long
Private m_fAllowControlHiding As Boolean
Private m_fKeepSplitPercentOnResize As Boolean
Private m_fRightToLeft As Boolean

'* State information
Private m_fEnableSplitAdjustmentOnResize As Boolean  '* Used to prevent adjusting split on initial resize
Private m_fVisible As Boolean    '* Indicates if the control is visible
Private m_fDragging As Boolean   '* Indicates that dragging of the splitter bar is in progress
Private m_nLastSplitPercent As Single  '* SplitPercent at which splitter bar was last drawn
Private m_hDC As Long                  '* Device Conext in use for drawing splitter bar during drag
Private m_nDesiredSplitPercent As Single
Private m_nDesiredSplitterBarX As Single
Private m_nDesiredSplitterBarY As Single

'* Drawing API Declarations
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
      
'* Drawing API Constants
Private Const NULL_BRUSH = 5
Private Const DCX_PARENTCLIP = &H20&
Private Const R2_NOT = 6

'* Container scale information
Private m_nSplitRegionTop As Single          '* Top of the split region in units of its container
Private m_nSplitRegionLeft As Single         '* Left of the split region in units of its container
Private m_nSplitRegionWidth As Single        '* Width of the split region in units of its container
Private m_nSplitRegionHeight As Single       '* Height of the split region in units of its container
Private m_iContainerXDirection As Integer    '* Sign indicating the direction Right in the container scale
Private m_iContainerYDirection As Integer    '* Sign indicating the direction Down in the container scale
Private m_nSplitRegionWidthPixels As Single  '* Width of the split region in pixels
Private m_nSplitRegionHeightPixels As Single '* Height of the split region in pixels
Private m_fContainerInfoAvailable As Boolean '* Indicates if container information is available



'*******************************************************************************
'EVentos
'*******************************************************************************

'*******************************************************************************
'* EVENT:  RepositionSplit
'*    Occurs after the splitter bar is moved
'*******************************************************************************
Public Event RepositionSplit()

'*******************************************************************************
'* EVENT:  Resize
'*    Occurs after the split region is resized but before the split is
'*    re-adjusted
'*******************************************************************************
Public Event Resize()


'*******************************************************************************
'Propiedades
'*******************************************************************************

Public Property Get MouseIcon() As Picture
    Set MouseIcon = lblSplitterBar.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lblSplitterBar.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = lblSplitterBar.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    lblSplitterBar.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property


'*******************************************************************************
'* PROPERTY:  FirstControl
'*    Contains the name of the control to occupy the top or left position of
'*    the split region (R/W)
'*******************************************************************************
Public Property Get FirstControl() As String
    '* Ensure control name is still valid
    If Not fIsValidControlName(m_sFirstControl) Then
        m_sFirstControl = vbNullString
    End If
   
    '* Return control name
    FirstControl = m_sFirstControl
End Property

Public Property Let FirstControl(NewFirstControl As String)
    Dim ctl As Control
   
    '* Ensure the first and second controls are not the same
    If m_sFirstControl = m_sSecondControl Then
        m_sSecondControl = vbNullString
        PropertyChanged "SecondControl"
    End If
   
    '* Ensure control name is valid
    If Not fIsValidControlName(NewFirstControl) Then
        Err.Raise m_iErrInvalidPropertyValue
    End If
   
    '* Update the first control
    m_sFirstControl = NewFirstControl
   
    '* Update display
    UpdateSplitPos
    PropertyChanged "FirstControl"
   
End Property

'*******************************************************************************
'* PROPERTY:  SecondControl
'*    Contains the name of the control to occupy the bottom or right of the
'*    split region (R/W)
'*******************************************************************************
Public Property Get SecondControl() As String
    '* Ensure control name is still valid
    If Not fIsValidControlName(m_sSecondControl) Then
        m_sSecondControl = vbNullString
    End If
   
    '* Return control name
    SecondControl = m_sSecondControl
End Property

Public Property Let SecondControl(NewSecondControl As String)
    Dim ctl As Control
   
    '* Ensure the first and second controls are not the same
    If m_sFirstControl = m_sSecondControl Then
        m_sFirstControl = vbNullString
        PropertyChanged "SecondControl"
    End If
   
    '* Ensure control name is valid
    If Not fIsValidControlName(NewSecondControl) Then
        Err.Raise m_iErrInvalidPropertyValue
    End If
   
    '* Update the first control
    m_sSecondControl = NewSecondControl

    '* Update display
    UpdateSplitPos
    PropertyChanged "SecondControl"

End Property

'*******************************************************************************
'* PROPERTY:  FirstControlMinSize
'*    Contains the minimum size of the first control in scale units of the
'*    container (R/W)
'*******************************************************************************
Public Property Get FirstControlMinSize() As Single
    FirstControlMinSize = m_nFirstControlMinSize
End Property

Public Property Let FirstControlMinSize(NewFirstControlMinSize As Single)
    If Not m_nFirstControlMinSize = NewFirstControlMinSize Then
        m_nFirstControlMinSize = NewFirstControlMinSize
        CheckControlSizes m_nSplitPercent
        UpdateSplitPos
        PropertyChanged "FirstControlMinSize"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  SecondControlMinSize
'*    Contains the minimum size of the second control in scale units of the
'*    container (R/W)
'*******************************************************************************
Public Property Get SecondControlMinSize() As Single
    SecondControlMinSize = m_nSecondControlMinSize
End Property

Public Property Let SecondControlMinSize(NewSecondControlMinSize As Single)
    If Not m_nSecondControlMinSize = NewSecondControlMinSize Then
        m_nSecondControlMinSize = NewSecondControlMinSize
        CheckControlSizes m_nSplitPercent
        UpdateSplitPos
        PropertyChanged "FirstControlMinSize"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  SplitPercent
'*    Contains the percent of shared region currently occupied by first control
'*    (R/W)
'*******************************************************************************
Public Property Get SplitPercent() As Single
Attribute SplitPercent.VB_UserMemId = 0
    SplitPercent = m_nSplitPercent
End Property

Public Property Let SplitPercent(ByVal NewSplitPercent As Single)
    If Not m_nSplitPercent = NewSplitPercent Then
      
        '* Save desired values for use with resizing
        m_nDesiredSplitPercent = NewSplitPercent
        m_nDesiredSplitterBarX = nSplitPercent2XPos(NewSplitPercent)
        m_nDesiredSplitterBarY = nSplitPercent2YPos(NewSplitPercent)
      
        CheckControlSizes NewSplitPercent

        m_nSplitPercent = NewSplitPercent
        UpdateSplitPos
      
        RaiseEvent RepositionSplit
        PropertyChanged "SplitPercent"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  SplitterBarVertical
'*    Contains the true if the splitter bar is vertical, false if horizontal
'*    (R/W)
'*******************************************************************************
Public Property Get SplitterBarVertical() As Boolean
    SplitterBarVertical = m_fSplitterBarVertical
End Property

Public Property Let SplitterBarVertical(NewSplitterBarVertical As Boolean)
    If Not m_fSplitterBarVertical = NewSplitterBarVertical Then
        m_fSplitterBarVertical = NewSplitterBarVertical
        SplitterBarThickness = SplitterBarThickness
        If m_fSplitterBarVertical Then
            lblSplitterBar.Top = 0
            lblSplitterBar.Height = UserControl.ScaleHeight
        Else
            lblSplitterBar.Left = 0
            lblSplitterBar.Width = UserControl.ScaleWidth
        End If
        CheckControlSizes m_nSplitPercent
        UpdateSplitPos
      
        '* Save desired values for use with resizing
        m_nDesiredSplitterBarX = nSplitPercent2XPos(m_nSplitPercent)
        m_nDesiredSplitterBarY = nSplitPercent2YPos(m_nSplitPercent)
      
        PropertyChanged "SplitterBarVertical"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  SplitterBarThickness
'*    Contains the determines the thickness of the splitter bar (R/W)
'*******************************************************************************
Public Property Get SplitterBarThickness() As Single
    SplitterBarThickness = m_nSplitterBarThickness
End Property

Public Property Let SplitterBarThickness(NewSplitterBarThickness As Single)
    If m_fSplitterBarVertical Then
        With UserControl
            '* (Round twips down to prevent rounding errors during display)
            lblSplitterBar.Width = _
                    Fix(NewSplitterBarThickness * .ScaleWidth / _
                    Abs(.ScaleX(.Width, vbTwips, vbContainerSize)))
         
            m_iSplitterBarThicknessPixels = _
             Abs(.ScaleX(NewSplitterBarThickness, vbContainerSize, vbPixels))
        End With
    Else
        With UserControl
            '* (Round twips down to prevent rounding errors during display)
            lblSplitterBar.Height = _
                    Fix(NewSplitterBarThickness * .ScaleHeight / _
                    Abs(.ScaleY(.Height, vbTwips, vbContainerSize)))
         
            m_iSplitterBarThicknessPixels = _
             Abs(.ScaleY(NewSplitterBarThickness, vbContainerSize, vbPixels))
        End With
    End If
    m_nSplitterBarThickness = NewSplitterBarThickness
    CheckControlSizes m_nSplitPercent
    UpdateSplitPos
    PropertyChanged "SplitterBarThickness"
End Property

'*******************************************************************************
'* PROPERTY:  SplitterBarColor
'*    Contains the background color of the splitter bar (R/W)
'*******************************************************************************
Public Property Get SplitterBarColor() As OLE_COLOR
    SplitterBarColor = lblSplitterBar.BackColor
End Property

Public Property Let SplitterBarColor(NewSplitterBarColor As OLE_COLOR)
    If Not lblSplitterBar.BackColor = NewSplitterBarColor Then
        lblSplitterBar.BackColor = NewSplitterBarColor
        PropertyChanged "SplitterBarColor"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  AllowControlHiding
'*    Contains the true if a control becomes hidden when its portion of the
'*    split region is reduced to less than the control’s minimum size. (R/W)
'*******************************************************************************
Public Property Get AllowControlHiding() As Boolean
    AllowControlHiding = m_fAllowControlHiding
End Property

Public Property Let AllowControlHiding(NewAllowControlHiding As Boolean)
    If Not m_fAllowControlHiding = NewAllowControlHiding Then
        m_fAllowControlHiding = NewAllowControlHiding
        CheckControlSizes m_nSplitPercent
        UpdateSplitPos
        PropertyChanged "AllowControlHiding"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  KeepSplitPercentOnResize
'*    Contains the determines if the split percent is kept or if split position
'*    is kept during resize (R/W)
'*******************************************************************************
Public Property Get KeepSplitPercentOnResize() As Boolean
    KeepSplitPercentOnResize = m_fKeepSplitPercentOnResize
End Property

Public Property Let KeepSplitPercentOnResize( _
        NewKeepSplitPercentOnResize As Boolean)
      
    If Not m_fKeepSplitPercentOnResize = NewKeepSplitPercentOnResize Then
        m_fKeepSplitPercentOnResize = NewKeepSplitPercentOnResize
        PropertyChanged "KeepSplitPercentOnResize"
    End If
End Property

'*******************************************************************************
'* PROPERTY:  RightToLeft
'*    Contains the determines the position of the first and second controls
'*    when the splitter bar is vertical (for bi-directional language
'*    compatibility) (R/W)
'*******************************************************************************
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_UserMemId = -611
    RightToLeft = m_fRightToLeft
End Property

Public Property Let RightToLeft(NewRightToLeft As Boolean)
    If Not m_fRightToLeft = NewRightToLeft Then
        m_fRightToLeft = NewRightToLeft
        UpdateSplitPos
      
        '* Save desired values for use with resizing
        m_nDesiredSplitterBarX = nSplitPercent2XPos(m_nSplitPercent)
        m_nDesiredSplitterBarY = nSplitPercent2YPos(m_nSplitPercent)
      
        PropertyChanged "RightToLeft"
    End If
End Property



'*******************************************************************************
'Metodos
'*******************************************************************************

'*******************************************************************************
'* METHOD:  Refresh
'*    Refreshes the display of the SplitRegion. (Call whenever the SplitRegion
'*    has been moved but not resized.)
'*******************************************************************************
Public Sub Refresh()
    Dim ctl As Control
    UpdateContainerInfo
    UpdateSplitPos
    On Error Resume Next '* We don't know if these controls exist and support Refresh
    Set ctl = ctlGetControl(m_sFirstControl)
    ctl.Refresh
    Set ctl = ctlGetControl(m_sSecondControl)
    ctl.Refresh
End Sub

'* Updates the split position to reflect the current split percent
Private Sub UpdateSplitPos()
    Dim nLeft As Single, nTop As Single, nWidth As Single, nHeight As Single
    Dim nSplitRatio As Single
    Dim ctlFirst As Control, ctlSecond As Control
   
    If Not m_fContainerInfoAvailable Then
        Exit Sub
    ElseIf Not fControlsAvailable() Then
        Exit Sub
    End If
   
    Set ctlFirst = ctlGetControl(m_sFirstControl)
    Set ctlSecond = ctlGetControl(m_sSecondControl)
   
    If Not m_fVisible Then
        If Not ctlFirst Is Nothing Then
            ctlFirst.Visible = False
        End If
        If Not ctlSecond Is Nothing Then
            ctlSecond.Visible = False
        End If
        Exit Sub
    End If
   
    nSplitRatio = m_nSplitPercent / 100#
   
    If m_fSplitterBarVertical Then
        '* Position first control
        nWidth = _
                m_nSplitRegionWidth * nSplitRatio - (0.5 * m_nSplitterBarThickness)
        If m_fRightToLeft Then
            nLeft = _
                    m_nSplitRegionLeft + m_iContainerXDirection * _
                    (m_nSplitRegionWidth - nWidth)
        Else
            nLeft = m_nSplitRegionLeft
        End If
        nTop = m_nSplitRegionTop
        nHeight = m_nSplitRegionHeight
        If Not ctlFirst Is Nothing Then
            If nWidth <= 0 Then
                ctlFirst.Visible = False
            Else
                ctlFirst.Move nLeft, nTop, nWidth, nHeight
                ctlFirst.Visible = True
            End If
        End If
      
        '* position splitter bar
        lblSplitterBar.Left = _
                nSplitPercent2XPos(m_nSplitPercent) - (0.5 * lblSplitterBar.Width)
      
        '* Position second control
        If m_fRightToLeft Then
            nLeft = m_nSplitRegionLeft
        Else
            nLeft = _
                    nLeft + m_iContainerXDirection * _
                    (nWidth + m_nSplitterBarThickness)
        End If
        nWidth = _
                m_nSplitRegionWidth * (1 - nSplitRatio) - _
                (0.5 * m_nSplitterBarThickness)
        If Not ctlSecond Is Nothing Then
            If nWidth <= 0 Then
                ctlSecond.Visible = False
            Else
                ctlSecond.Move nLeft, nTop, nWidth, nHeight
                ctlSecond.Visible = True
            End If
        End If
    Else
   
        '* Position first control
        nLeft = m_nSplitRegionLeft
        nWidth = m_nSplitRegionWidth
        nTop = m_nSplitRegionTop
        nHeight = _
                m_nSplitRegionHeight * nSplitRatio - (0.5 * m_nSplitterBarThickness)
        If Not ctlFirst Is Nothing Then
            If nHeight <= 0 Then
                ctlFirst.Visible = False
            Else
                ctlFirst.Move nLeft, nTop, nWidth, nHeight
                ctlFirst.Visible = True
            End If
        End If
      
        '* position splitter bar
        lblSplitterBar.Top = _
                nSplitPercent2YPos(m_nSplitPercent) - (0.5 * lblSplitterBar.Height)
      
        '* Position second control
        nTop = nTop + m_iContainerYDirection * (nHeight + m_nSplitterBarThickness)
        nHeight = _
                m_nSplitRegionHeight * (1 - nSplitRatio) - _
                (0.5 * m_nSplitterBarThickness)
        If Not ctlSecond Is Nothing Then
            If nHeight <= 0 Then
                ctlSecond.Visible = False
            Else
                ctlSecond.Move nLeft, nTop, nWidth, nHeight
                ctlSecond.Visible = True
            End If
        End If
    End If
End Sub

'* Attempts to determine if a control name is a valid control
Private Function fIsValidControlName(sControl As String) As Boolean
    Dim ctl As Control
   
    On Error GoTo fIsValidControlNameErr
      
    If Len(sControl) = 0 Then
        fIsValidControlName = True
    ElseIf fControlsAvailable Then
        fIsValidControlName = True
   
        '* Ensure the first control is not the SplitRegion itself
        If sControl = UserControl.Ambient.DisplayName Then
            fIsValidControlName = False
        End If
   
        '* Check that control exists
        Set ctl = ctlGetControl(sControl)
        If ctl Is Nothing Then
            fIsValidControlName = False
        End If
      
    Else
        fIsValidControlName = True
    End If
   
    Exit Function
fIsValidControlNameErr:
    fIsValidControlName = False
End Function

'* Returns true if control objects are currently available from the host
Private Function fControlsAvailable() As Boolean
   
    On Error GoTo fControlsAvailableErr
   
    If UserControl.ParentControls.Count < 1 Then
        fControlsAvailable = False
    Else
        fControlsAvailable = True
    End If
   
    Exit Function
fControlsAvailableErr:
    fControlsAvailable = False
End Function

'* Gets a reference to the control of the specified name
Private Function ctlGetControl(sName As String) As Control
    Dim i As Long, sCtlCur As String
   
    On Error Resume Next
   
    UserControl.ParentControls.ParentControlsType = vbExtender
    With UserControl.ParentControls
        For i = 0 To .Count - 1
            sCtlCur = .Item(i).Name
            If Err.Number <> 0 Then
                Err.Clear
            ElseIf sName = sCtlCur Then
                Set ctlGetControl = .Item(i)
                Exit Function
            End If
        Next i
    End With
   
    Set ctlGetControl = Nothing

End Function

Private Sub BeginDrag(x As Single, Y As Single)
    Dim nNewSplitPercent As Single
    Dim hBrush As Long
   
    UpdateContainerInfo
   
    '* Prepare DC for drawing
    m_hDC = GetDCEx(UserControl.hwnd, 0&, DCX_PARENTCLIP)
    SelectClipRgn m_hDC, 0&
    SetROP2 m_hDC, R2_NOT
    hBrush = GetStockObject(NULL_BRUSH)
    SelectObject m_hDC, hBrush
        
    nNewSplitPercent = nSplitPos2SplitPercent(x, Y)
    DrawDragLine nNewSplitPercent
    m_nLastSplitPercent = nNewSplitPercent
      
    m_fDragging = True
End Sub

Private Sub Drag(x As Single, Y As Single)
    Dim nNewSplitPercent As Single
   
    Static nLastSplitPercent As Single
    If m_fDragging Then
      
        nNewSplitPercent = nSplitPos2SplitPercent(x, Y)
      
        CheckControlSizes nNewSplitPercent
      
        If Abs(nNewSplitPercent - nLastSplitPercent) > 0.1 Then
            nLastSplitPercent = nNewSplitPercent
            DrawDragLine m_nLastSplitPercent
            DrawDragLine nNewSplitPercent
            m_nLastSplitPercent = nNewSplitPercent
        End If
    End If
End Sub

Private Sub EndDrag(x As Single, Y As Single)
    Dim nNewSplitPercent As Single
    If m_fDragging Then
      
        nNewSplitPercent = nSplitPos2SplitPercent(x, Y)

        DrawDragLine m_nLastSplitPercent
      
        '* Release device context
        ReleaseDC UserControl.hwnd, m_hDC
      
        m_fDragging = False
   
        SplitPercent = nNewSplitPercent
      
    End If
End Sub

Private Sub DrawDragLine(nSplitPercent As Single)
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
   
    If m_fSplitterBarVertical Then
        If m_fRightToLeft Then
            X1 = CLng(m_nSplitRegionWidthPixels * _
                        (1# - (nSplitPercent / 100#))) - _
                        m_iSplitterBarThicknessPixels / 2
        Else
            X1 = CLng(m_nSplitRegionWidthPixels * (nSplitPercent / 100#)) - _
                        m_iSplitterBarThicknessPixels / 2
        End If
        X2 = X1 + m_iSplitterBarThicknessPixels
        Y1 = 0
        Y2 = m_nSplitRegionHeightPixels
    Else
        X1 = 0
        X2 = m_nSplitRegionWidthPixels
        Y1 = CLng(m_nSplitRegionHeightPixels * (nSplitPercent / 100#)) - _
                    m_iSplitterBarThicknessPixels / 2
        Y2 = Y1 + m_iSplitterBarThicknessPixels
    End If
    Rectangle m_hDC, X1, Y1, X2, Y2
End Sub

Private Function nTwips2ContainerSize(nTwips As Single) As Single
    If m_fSplitterBarVertical Then
        nTwips2ContainerSize = _
                UserControl.ScaleY(nTwips, vbTwips, vbContainerSize)
    Else
        nTwips2ContainerSize = _
                UserControl.ScaleX(nTwips, vbTwips, vbContainerSize)
    End If
End Function

'* Attempts to update container scale information for the user control
Private Sub UpdateContainerInfo()
    On Error GoTo UpdateContainerInfoErr
   
    With UserControl.Extender
        m_nSplitRegionLeft = .Left
        m_nSplitRegionTop = .Top
    End With
   
    With UserControl
        m_nSplitRegionWidth = .ScaleX(.Width, vbTwips, vbContainerSize)
        m_nSplitRegionHeight = .ScaleY(.Height, vbTwips, vbContainerSize)
        m_nSplitRegionWidthPixels = .ScaleX(.Width, vbTwips, vbPixels)
        m_nSplitRegionHeightPixels = .ScaleY(.Height, vbTwips, vbPixels)
    End With
   
    If m_nSplitRegionWidth < 0 Then
        m_iContainerXDirection = -1
        m_nSplitRegionWidth = -1 * m_nSplitRegionWidth
    Else
        m_iContainerXDirection = 1
    End If
   
    If m_nSplitRegionHeight < 0 Then
        m_iContainerYDirection = -1
        m_nSplitRegionHeight = -1 * m_nSplitRegionHeight
    Else
        m_iContainerYDirection = 1
    End If
   
    m_fContainerInfoAvailable = True
   
    Exit Sub
UpdateContainerInfoErr:
    Err.Clear
    m_fContainerInfoAvailable = False
End Sub

'* Ensures that the associated controls are not reduced beyond their minimum
'* size.  First, if the SplitRegion is resized to the minimum size of
'* (FirstControlMinSize + SecondControlMinSize + SplitterBarThickness) if it is
'* smaller.  Then the new split percent is adjusted if necessary.  If
'* AllowControlHiding is true and a control would be below its minimum size,
'* the split percent is adjusted to completely hide that control.  Otherwise
'* if a control would be below its minimum size the split percent is adjusted
'* to display that control at its minimum size.
Private Sub CheckControlSizes(nNewSplitPercent As Single)
    Dim nMinSplitRegionSize As Single
    Dim nMinSplitPercent As Single, nMaxSplitPercent As Single
    Dim nSplitRegionSize As Single

    If Not m_fContainerInfoAvailable Then
        Exit Sub
    End If
   
    nMinSplitRegionSize = m_nFirstControlMinSize + _
                                m_nSecondControlMinSize + _
                                m_nSplitterBarThickness
                        
    '* Ensure that SplitRegion can hold both controls
    If m_fSplitterBarVertical Then
        nMinSplitRegionSize = _
                UserControl.ScaleX(nMinSplitRegionSize, vbContainerSize, vbTwips)
        If UserControl.Width < nMinSplitRegionSize Then
            UserControl.Width = nMinSplitRegionSize
        End If
    Else
        nMinSplitRegionSize = _
                UserControl.ScaleY(nMinSplitRegionSize, vbContainerSize, vbTwips)
        If UserControl.Height < nMinSplitRegionSize Then
            UserControl.Height = nMinSplitRegionSize
        End If
    End If
   
    '* Ensure that new split percent is okay
    If m_fSplitterBarVertical Then
        nSplitRegionSize = m_nSplitRegionWidth - m_nSplitterBarThickness
    Else
        nSplitRegionSize = m_nSplitRegionHeight - m_nSplitterBarThickness
    End If
   
    nMinSplitPercent = (m_nFirstControlMinSize / nSplitRegionSize) * 100
    nMaxSplitPercent = (1 - m_nSecondControlMinSize / nSplitRegionSize) * 100
   
    If nNewSplitPercent < nMinSplitPercent Then
        If m_fAllowControlHiding Then
            nNewSplitPercent = 0#
        Else
            nNewSplitPercent = nMinSplitPercent
        End If
        PropertyChanged "SplitPercent"
    ElseIf nNewSplitPercent > nMaxSplitPercent Then
        If m_fAllowControlHiding Then
            nNewSplitPercent = 100#
        Else
            nNewSplitPercent = nMaxSplitPercent
        End If
        PropertyChanged "SplitPercent"
    End If
      
End Sub

'* Takes a position in usercontrol scale and returns the splitpercent
'* that would place the splitter bar across that position
Private Function nSplitPos2SplitPercent(x As Single, Y As Single) As Single
    If m_fSplitterBarVertical Then
        If m_fRightToLeft Then
            nSplitPos2SplitPercent = (1# - x / UserControl.ScaleWidth) * 100#
        Else
            nSplitPos2SplitPercent = (x / UserControl.ScaleWidth) * 100#
        End If
    Else
        nSplitPos2SplitPercent = (Y / UserControl.ScaleHeight) * 100#
    End If
End Function

'* Takes a split percent and returns a corresponding horizontal position in
'* usercontrol scale
Private Function nSplitPercent2XPos(nSplitPercent As Single) As Single
    If m_fSplitterBarVertical Then
        If m_fRightToLeft Then
            nSplitPercent2XPos = (1# - nSplitPercent / 100#) * UserControl.ScaleWidth
        Else
            nSplitPercent2XPos = (nSplitPercent / 100#) * UserControl.ScaleWidth
        End If
    Else
        nSplitPercent2XPos = 0
    End If
End Function

'* Takes a split percent and returns a corresponding vertical position in
'* usercontrol scale
Private Function nSplitPercent2YPos(nSplitPercent As Single) As Single
    If m_fSplitterBarVertical Then
        nSplitPercent2YPos = 0
    Else
        nSplitPercent2YPos = (nSplitPercent / 100#) * UserControl.ScaleHeight
    End If
End Function



'*******************************************************************************
'Manejadores de los Eventos
'*******************************************************************************

Private Sub UserControl_Initialize()
On Error Resume Next
         
    m_fVisible = True
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

         
    m_nFirstControlMinSize = nTwips2ContainerSize(m_nControlMinSizeDefault)
    m_nSecondControlMinSize = nTwips2ContainerSize(m_nControlMinSizeDefault)
    m_nDesiredSplitPercent = m_nSplitPercentDefault
    m_fSplitterBarVertical = m_fSplitterBarVerticalDefault
    m_nSplitterBarThickness = nTwips2ContainerSize(m_nSplitterBarThicknessDefault)
    SplitterBarColor = Ambient.BackColor
    m_fAllowControlHiding = m_fAllowControlHidingDefault
    m_fKeepSplitPercentOnResize = m_fKeepSplitPercentOnResizeDefault
    m_fRightToLeft = Ambient.RightToLeft
    lblSplitterBar.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Paint()
On Error GoTo ErrorHandler

    '* Only enbable adjusting of split percent on resize if in user mode and
    '* control has been shown (i.e. after initial resize)
    If Ambient.UserMode Then
        m_fEnableSplitAdjustmentOnResize = True
    End If

    Exit Sub
ErrorHandler:
    #If DEBUG_MODE Then
        DebugAssert False, "Unexpected error:  " & Err.Description, App.EXEName & ".SplitRegion.UserControl_Paint" ' Do not localize
    #End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
         
    FirstControlMinSize = PropBag.ReadProperty("FirstControlMinSize", nTwips2ContainerSize(m_nControlMinSizeDefault))
    SecondControlMinSize = PropBag.ReadProperty("SecondControlMinSize", nTwips2ContainerSize(m_nControlMinSizeDefault))
    SplitPercent = PropBag.ReadProperty("SplitPercent", m_nSplitPercentDefault)
    SplitterBarVertical = PropBag.ReadProperty("SplitterBarVertical", m_fSplitterBarVerticalDefault)
    SplitterBarThickness = PropBag.ReadProperty("SplitterBarThickness", nTwips2ContainerSize(m_nSplitterBarThicknessDefault))
    SplitterBarColor = PropBag.ReadProperty("SplitterBarColor", Ambient.BackColor)
    AllowControlHiding = PropBag.ReadProperty("AllowControlHiding", m_fAllowControlHidingDefault)
    KeepSplitPercentOnResize = PropBag.ReadProperty("KeepSplitPercentOnResize", m_fKeepSplitPercentOnResizeDefault)
    RightToLeft = PropBag.ReadProperty("RightToLeft", Ambient.RightToLeft)
    m_sFirstControl = PropBag.ReadProperty("FirstControl", vbNullString)
    m_sSecondControl = PropBag.ReadProperty("SecondControl", vbNullString)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lblSplitterBar.MousePointer = PropBag.ReadProperty("MousePointer", 7)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
         
    PropBag.WriteProperty "FirstControl", m_sFirstControl, vbNullString
    PropBag.WriteProperty "SecondControl", m_sSecondControl, vbNullString
    PropBag.WriteProperty "FirstControlMinSize", FirstControlMinSize, nTwips2ContainerSize(m_nControlMinSizeDefault)
    PropBag.WriteProperty "SecondControlMinSize", SecondControlMinSize, nTwips2ContainerSize(m_nControlMinSizeDefault)
    PropBag.WriteProperty "SplitPercent", SplitPercent, m_nSplitPercentDefault
    PropBag.WriteProperty "SplitterBarVertical", SplitterBarVertical, m_fSplitterBarVerticalDefault
    PropBag.WriteProperty "SplitterBarThickness", SplitterBarThickness, nTwips2ContainerSize(m_nSplitterBarThicknessDefault)
    PropBag.WriteProperty "SplitterBarColor", SplitterBarColor, Ambient.BackColor
    PropBag.WriteProperty "AllowControlHiding", AllowControlHiding, m_fAllowControlHidingDefault
    PropBag.WriteProperty "KeepSplitPercentOnResize", KeepSplitPercentOnResize, m_fKeepSplitPercentOnResizeDefault
    PropBag.WriteProperty "RightToLeft", RightToLeft, Ambient.RightToLeft
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lblSplitterBar.MousePointer, 7)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
         
    UpdateContainerInfo
   
    '* Resize splitter bar to fit
    With UserControl
        If m_fSplitterBarVertical Then
            lblSplitterBar.Height = .ScaleHeight
        Else
            lblSplitterBar.Width = .ScaleWidth
        End If
    End With
         
    ' Adjust spit if appropriate to keep split bar stationary
    If Not m_fEnableSplitAdjustmentOnResize Then
        '* Save desired splitter bar position
        m_nDesiredSplitterBarX = nSplitPercent2XPos(m_nDesiredSplitPercent)
        m_nDesiredSplitterBarY = nSplitPercent2YPos(m_nDesiredSplitPercent)
    ElseIf Not m_fKeepSplitPercentOnResize Then
        With lblSplitterBar
            m_nDesiredSplitPercent = nSplitPos2SplitPercent(m_nDesiredSplitterBarX, _
                                            m_nDesiredSplitterBarY)
        End With
      
        If m_nDesiredSplitPercent > 100 Then
            m_nDesiredSplitPercent = 100
        End If
    End If
    m_nSplitPercent = m_nDesiredSplitPercent
   
    '* Prevent resizing below the minimum size required to display both controls
    CheckControlSizes m_nSplitPercent

    '* Raise resize event
    RaiseEvent Resize
   
    '* update split
    UpdateSplitPos
End Sub

Private Sub UserControl_Show()
On Error Resume Next
         
    m_fVisible = True
      
    UpdateContainerInfo
    UpdateSplitPos
End Sub

Private Sub UserControl_Hide()
On Error Resume Next
    
    m_fVisible = False
    UpdateSplitPos
   
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
Dim nScaleAdjustment As Double, nNewSplitterBarThickness As Double
On Error Resume Next
         
    If PropertyName = "ScaleUnits" Then
        If m_fSplitterBarVertical Then
            With UserControl
                nNewSplitterBarThickness = _
                        lblSplitterBar.Width / .ScaleWidth * _
                        Abs(.ScaleX(.Width, vbTwips, vbContainerSize))
            End With
        Else
            With UserControl
                nNewSplitterBarThickness = _
                        lblSplitterBar.Height / .ScaleHeight * _
                        Abs(.ScaleY(.Height, vbTwips, vbContainerSize))
            End With
        End If
        nScaleAdjustment = nNewSplitterBarThickness / m_nSplitterBarThickness
        m_nSplitterBarThickness = nNewSplitterBarThickness
        m_nFirstControlMinSize = m_nFirstControlMinSize * nScaleAdjustment
        m_nSecondControlMinSize = m_nSecondControlMinSize * nScaleAdjustment
        PropertyChanged "SplitterBarThickness"
        PropertyChanged "FirstControlMinSize"
        PropertyChanged "SecondControlMinSize"
    End If
End Sub

Private Sub lblSplitterBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
         
    If Button = vbLeftButton Then
        With UserControl
            BeginDrag lblSplitterBar.Left + .ScaleX(x, vbTwips, .ScaleMode), lblSplitterBar.Top + .ScaleY(Y, vbTwips, .ScaleMode)
        End With
    End If
End Sub

Private Sub lblSplitterBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
         
    With UserControl
        Drag lblSplitterBar.Left + .ScaleX(x, vbTwips, .ScaleMode), lblSplitterBar.Top + .ScaleY(Y, vbTwips, .ScaleMode)
    End With
End Sub

Private Sub lblSplitterBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
         
    With UserControl
        EndDrag lblSplitterBar.Left + .ScaleX(x, vbTwips, .ScaleMode), lblSplitterBar.Top + .ScaleY(Y, vbTwips, .ScaleMode)
    End With
End Sub


