VERSION 5.00
Begin VB.UserControl ocxLightButton 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   ScaleHeight     =   480
   ScaleWidth      =   1035
   ToolboxBitmap   =   "ocxLightButton.ctx":0000
   Begin VB.Timer timer 
      Enabled         =   0   'False
      Left            =   45
      Top             =   570
   End
   Begin VB.Image imgOFF 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1005
   End
   Begin VB.Image imgOK 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1005
   End
   Begin VB.Image imgON 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "ocxLightButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' ocxLightButton
'               Este control es un boton de tres estados, que puede
'               servir cuando los botones de comando estandares
'               comienzan a parecerte muy aburridos.
'*******************************************************************************
' rev:     esau mar-2002
'*******************************************************************************
Option Explicit

'*******************************************************************************
' llamadas al API
'*******************************************************************************
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function GetCursorPos Lib "User32" (ByRef Pnt As POINT) As Long
Private Declare Function ScreenToClient Lib "User32" (ByVal hwnd As Long, ByRef Pnt As POINT) As Long

'*******************************************************************************
' declaracion de tipos
'*******************************************************************************
Private Type POINT
    X As Long
    Y As Long
End Type

'*******************************************************************************
' eventos del control
'*******************************************************************************
Public Event Click()
Public Event OnActivate()
Public Event OnDeactivate()

'*******************************************************************************
' constantes por defecto del control
'*******************************************************************************
Const m_def_Stretch = 0
Const m_def_Activate = False
Const m_def_KTimeOK = 2
Const m_def_TimeWait = 100

'*******************************************************************************
' variables para almacenar propiedades del control
'*******************************************************************************
Private m_Stretch As Boolean
Private m_Activate As Boolean
Private m_KTimeOK As Integer
Private m_TimeWait As Integer

Private bSobre_imgON As Boolean
Private m_MousePointer As Integer

'*******************************************************************************
' propiedades del control
'*******************************************************************************
Public Property Get Stretch() As Boolean
    Stretch = m_Stretch
End Property

Public Property Let Stretch(ByVal NewStretch As Boolean)
    m_Stretch = NewStretch
    PropertyChanged "Stretch"
End Property

Public Property Get Activate() As Boolean
    Activate = m_Activate
End Property

Public Property Let Activate(ByVal NewActivate As Boolean)
    m_Activate = NewActivate
    If Activate Then
        UserControl.MousePointer = m_MousePointer
    Else
        UserControl.MousePointer = vbDefault
    End If
    PropertyChanged "Activate"
End Property

Public Property Get PictureON() As Picture
    Set PictureON = imgON.Picture
End Property

Public Property Set PictureON(ByVal NewPictureON As Picture)
    Set imgON.Picture = NewPictureON
    PropertyChanged "PictureON"
End Property

Public Property Get PictureOFF() As Picture
    Set PictureOFF = imgOFF.Picture
End Property

Public Property Set PictureOFF(ByVal NewPictureOFF As Picture)
    Set imgOFF.Picture = NewPictureOFF
    PropertyChanged "PictureOFF"
End Property

Public Property Get PictureOK() As Picture
    Set PictureOK = imgOK.Picture
End Property

Public Property Set PictureOK(ByVal New_PictureOK As Picture)
    Set imgOK.Picture = New_PictureOK
    PropertyChanged "PictureOK"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get TTT() As String
    TTT = imgON.ToolTipText
End Property

Public Property Let TTT(ByVal New_ToolTipText As String)
    imgON.ToolTipText() = New_ToolTipText
    PropertyChanged "TTT"
End Property

Public Property Get KTimeOK() As Integer
    KTimeOK = m_KTimeOK
End Property

Public Property Let KTimeOK(ByVal New_KTimeOK As Integer)
    m_KTimeOK = New_KTimeOK
    PropertyChanged "KTimeOK"
End Property

Public Property Get Timewait() As Integer
    Timewait = m_TimeWait
End Property

Public Property Let Timewait(ByVal New_Timewait As Integer)
    m_TimeWait = New_Timewait
    PropertyChanged "TimeWait"
End Property

Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
    dlgAcerca.Show vbModal
    Set dlgAcerca = Nothing
End Sub

'*******************************************************************************
' comportamiento del control
'*******************************************************************************

' imgON no puede nunca capturar el control (el cursor por defecto se muestra
' despues de MouseDown) (imgOFF y imgOK tienen el cursor por defecto)
Private Sub imgON_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim L As Long
On Error Resume Next
    If Button = vbLeftButton Then
        'Desactivamos el "corazoncito"
        timer.Enabled = False
        bSobre_imgON = False
        timer.Interval = m_KTimeOK * m_TimeWait
        L = ReleaseCapture()
        imgOK.ZOrder 0
        imgOK.Visible = True
        timer.Enabled = True
    End If
End Sub

Private Sub imgOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Activate Then
        bSobre_imgON = True
        timer.Interval = m_TimeWait
        imgON.ZOrder 0
        timer.Enabled = True
        RaiseEvent OnActivate
    End If
End Sub

'*******************************************************************************
' UserControl events:
'*******************************************************************************
Private Sub UserControl_Resize()
    If m_Stretch = False Then
        ' imgON es la que determina el tamaño
        imgON.Stretch = False
        imgOFF.Stretch = False
        imgOK.Stretch = False
        UserControl.Height = imgON.Height
        UserControl.Width = imgON.Width
    Else
        imgON.Stretch = True
        imgOFF.Stretch = True
        imgOK.Stretch = True
        imgON.Move 0, 0, ScaleWidth, ScaleHeight
        imgOFF.Move 0, 0, ScaleWidth, ScaleHeight
        imgOK.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub

' inicializacion de propiedades por defecto
Private Sub UserControl_InitProperties()
    m_Stretch = m_def_Stretch
    m_Activate = m_def_Activate
    m_KTimeOK = m_def_KTimeOK
    m_TimeWait = m_def_TimeWait
End Sub

' lectura de propiedades
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    m_Activate = PropBag.ReadProperty("Activate", m_def_Activate)
    m_Stretch = PropBag.ReadProperty("Stretch", m_def_Stretch)
    m_KTimeOK = PropBag.ReadProperty("KTimeOK", m_def_KTimeOK)
    m_TimeWait = PropBag.ReadProperty("TimeWait", m_def_TimeWait)
    imgON.Picture = PropBag.ReadProperty("PictureON", Nothing)
    imgOFF.Picture = PropBag.ReadProperty("PictureOFF", Nothing)
    imgOK.Picture = PropBag.ReadProperty("PictureOK", Nothing)
    imgON.ToolTipText = PropBag.ReadProperty("TTT", "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

' escritura de propiedades
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Activate", m_Activate, m_def_Activate)
    Call PropBag.WriteProperty("PictureON", imgON.Picture, Nothing)
    Call PropBag.WriteProperty("PictureOFF", imgOFF.Picture, Nothing)
    Call PropBag.WriteProperty("PictureOK", imgOK.Picture, Nothing)
    Call PropBag.WriteProperty("Stretch", m_Stretch, m_def_Stretch)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("TTT", imgON.ToolTipText, "")
    Call PropBag.WriteProperty("KTimeOK", m_KTimeOK, m_def_KTimeOK)
    Call PropBag.WriteProperty("TimeWait", m_TimeWait, m_def_TimeWait)
End Sub

'*******************************************************************************
' el corazon del control
'*******************************************************************************
Private Sub timer_Timer()
Dim L As Long
Dim mPt As POINT
On Error Resume Next
    If Activate Then
        L = GetCursorPos(mPt)
        L = ScreenToClient(UserControl.hwnd, mPt)
        If bSobre_imgON Then
            ' tenemos que convertir de twips a pixels (dividir entre 15)
            If mPt.X < 0 Or mPt.X >= (UserControl.ScaleWidth / 15) Or mPt.Y < 0 Or mPt.Y >= (UserControl.ScaleHeight / 15) Then
                imgOFF.ZOrder 0
                timer.Enabled = False
                RaiseEvent OnDeactivate
            End If
        Else
            If mPt.X < 0 Or mPt.X >= (UserControl.ScaleWidth / 15) Or mPt.Y < 0 Or mPt.Y >= (UserControl.ScaleHeight / 15) Then
                imgOFF.ZOrder 0
                timer.Enabled = False
                RaiseEvent OnDeactivate
            Else
                imgON.ZOrder 0
                bSobre_imgON = True
                timer.Interval = m_TimeWait
            End If
            imgOK.Visible = False
            RaiseEvent Click
        End If
    Else
        timer.Enabled = False
        imgOFF.ZOrder
        RaiseEvent OnDeactivate
    End If
End Sub


