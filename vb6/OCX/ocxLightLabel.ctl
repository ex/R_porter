VERSION 5.00
Begin VB.UserControl ocxLightLabel 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   ScaleHeight     =   240
   ScaleWidth      =   840
   ToolboxBitmap   =   "ocxLightLabel.ctx":0000
   Begin VB.Timer timer 
      Left            =   0
      Top             =   225
   End
   Begin VB.Label label 
      AutoSize        =   -1  'True
      Caption         =   "LightLabel"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "ocxLightLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' ocxLightLabel - MouseDown
'               Este control es una etiqueta de cuatro estados, que puede
'               servir cuando quieres una etiqueta que emule una
'               direccion de internet (un vinculo)
'               (este control no captura el cursor y funciona
'               con el evento MouseDown, es decir no admite equivocaciones)
'
' Autosize:     Debe ser usado en tiempo de diseño, en tiempo de ejecucion
'               cambiara el tamaño del control dinamicamente, pero a costa de
'               un feo titileo cuando el cursor se pone en puntos criticos
'               del borde, preferible, si vas a hacer algun efecto (como aumentar
'               el tamaño de la fuente cuando se active el control) especifiques
'               el tamaño maximo en tiempo de diseño, con [Autosize = True] luego
'               pones [Autosize = False], el control no se activara sino hasta
'               que el cursor se encuentre sobre el label, no sobre el control
'               no te preocupes de que se active cuando no estes señalando el texto.
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
Const m_def_ColorOFF = 0
Const m_def_ColorON = 0
Const m_def_ColorDOWN = 0
Const m_def_ColorOK = 0
Const m_def_AutoSize = False
Const m_def_KTimeOK = 2
Const m_def_TimeWait = 100
Const m_def_Mousepointer = 0
Const m_def_Activate = False

'*******************************************************************************
' variables para almacenar propiedades del control
'*******************************************************************************
Private m_ColorOFF As OLE_COLOR
Private m_ColorON As OLE_COLOR
Private m_ColorDOWN As OLE_COLOR
Private m_ColorOK As OLE_COLOR
Private m_Activate As Boolean
Private m_AutoSize As Boolean
Private m_KTimeOK As Integer
Private m_TimeWait As Integer
Private m_MousePointer As Integer

Private bActivado As Boolean
Private bPresionado As Boolean
Private bFueActivado As Boolean
Private bMouseMove As Boolean

'*******************************************************************************
' propiedades del control
'*******************************************************************************
Public Property Get Activate() As Boolean
    Activate = m_Activate
End Property

Public Property Let Activate(ByVal New_Activate As Boolean)
    m_Activate = New_Activate
    PropertyChanged "Activate"
    If Activate Then
        bMouseMove = False
    Else
        UserControl.MousePointer = vbDefault
    End If
End Property

Public Property Get Font() As Font
    Set Font = label.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set label.Font = New_Font
    If AutoSize Then
        UserControl_Resize
    End If
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
    Caption = label.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    label.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = label.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    label.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If m_AutoSize = True Then
        UserControl_Resize
    End If
End Property

Public Property Get MousePointer() As Integer
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
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

Public Property Get ColorOFF() As OLE_COLOR
    ColorOFF = m_ColorOFF
End Property

Public Property Let ColorOFF(ByVal New_ColorOFF As OLE_COLOR)
    m_ColorOFF = New_ColorOFF
    label.ForeColor = New_ColorOFF
    PropertyChanged "ColorOFF"
End Property

Public Property Get ColorOK() As OLE_COLOR
    ColorOK = m_ColorOK
End Property

Public Property Let ColorOK(ByVal New_ColorOK As OLE_COLOR)
    m_ColorOK = New_ColorOK
    PropertyChanged "ColorOK"
End Property

Public Property Get ColorON() As OLE_COLOR
    ColorON = m_ColorON
End Property

Public Property Let ColorON(ByVal New_ColorON As OLE_COLOR)
    m_ColorON = New_ColorON
    PropertyChanged "ColorON"
End Property

Public Property Get ColorDOWN() As OLE_COLOR
    ColorDOWN = m_ColorDOWN
End Property

Public Property Let ColorDOWN(ByVal New_ColorDOWN As OLE_COLOR)
    m_ColorDOWN = New_ColorDOWN
    PropertyChanged "ColorDOWN"
End Property

Public Property Get TTT() As String
    TTT = label.ToolTipText
End Property

Public Property Let TTT(ByVal New_ToolTipText As String)
    label.ToolTipText() = New_ToolTipText
    PropertyChanged "TTT"
End Property

Public Property Get Timewait() As Integer
    Timewait = m_TimeWait
End Property

Public Property Let Timewait(ByVal New_Timewait As Integer)
    m_TimeWait = New_Timewait
    PropertyChanged "TimeWait"
End Property

Public Property Get KTimeOK() As Integer
    KTimeOK = m_KTimeOK
End Property

Public Property Let KTimeOK(ByVal New_KTimeOK As Integer)
    m_KTimeOK = New_KTimeOK
    PropertyChanged "KTimeOK"
End Property

Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
    dlgAcerca.Show vbModal
    Set dlgAcerca = Nothing
End Sub

'*******************************************************************************
' comportamiento del control
'*******************************************************************************

Private Sub label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim L As Long
On Error Resume Next
    If Button = vbLeftButton And Activate Then
        timer.Enabled = False
        bActivado = False
        timer.Interval = m_KTimeOK * m_TimeWait
        L = ReleaseCapture()
        label.ForeColor = m_ColorOK
        bPresionado = True
        timer.Enabled = True
    End If
End Sub

Private Sub label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Activate And Not bMouseMove And Not bPresionado Then
        UserControl.MousePointer = m_MousePointer
        bActivado = True
        timer.Interval = m_TimeWait
        label.ForeColor = m_ColorON
        timer.Enabled = True
        bMouseMove = True
        UserControl.MousePointer = m_MousePointer
        RaiseEvent OnActivate
        If AutoSize Then
            UserControl_Resize
        End If
    End If
End Sub

'*******************************************************************************
' UserControl events:
'*******************************************************************************
Private Sub UserControl_Initialize()
    bFueActivado = False
    bPresionado = False
    bActivado = False
    bMouseMove = False
End Sub

Private Sub UserControl_Resize()
    If m_AutoSize Then
        UserControl.Height = label.Height
        UserControl.Width = label.Width
    End If
End Sub

Private Sub UserControl_InitProperties()
    m_Activate = m_def_Activate
    m_ColorOFF = m_def_ColorOFF
    m_ColorON = m_def_ColorON
    m_ColorDOWN = m_def_ColorDOWN
    m_ColorOK = m_def_ColorOK
    m_KTimeOK = m_def_KTimeOK
    m_AutoSize = m_def_AutoSize
    m_TimeWait = m_def_TimeWait
    m_MousePointer = m_def_Mousepointer
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ColorOFF = PropBag.ReadProperty("ColorOFF", m_def_ColorOFF)
    m_ColorON = PropBag.ReadProperty("ColorON", m_def_ColorON)
    m_ColorOK = PropBag.ReadProperty("ColorOK", m_def_ColorOK)
    m_ColorDOWN = PropBag.ReadProperty("ColorDOWN", m_def_ColorDOWN)
    m_Activate = PropBag.ReadProperty("Activate", m_def_Activate)
    m_KTimeOK = PropBag.ReadProperty("KTimeOK", m_def_KTimeOK)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_TimeWait = PropBag.ReadProperty("Timewait", m_def_TimeWait)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_Mousepointer)
    label.Caption = PropBag.ReadProperty("Caption", "LightLbl")
    label.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    label.ForeColor = ColorOFF
    label.ToolTipText = PropBag.ReadProperty("TTT", "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_Mousepointer)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", label.Caption, "LightLbl")
    Call PropBag.WriteProperty("BackColor", label.BackColor, &H8000000F)
    Call PropBag.WriteProperty("TTT", label.ToolTipText, "")
    Call PropBag.WriteProperty("ColorOFF", m_ColorOFF, m_def_ColorOFF)
    Call PropBag.WriteProperty("ColorON", m_ColorON, m_def_ColorON)
    Call PropBag.WriteProperty("ColorOK", m_ColorOK, m_def_ColorOK)
    Call PropBag.WriteProperty("ColorDOWN", m_ColorDOWN, m_def_ColorDOWN)
    Call PropBag.WriteProperty("Activate", m_Activate, m_def_Activate)
    Call PropBag.WriteProperty("Timewait", m_TimeWait, m_def_TimeWait)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("KTimeOK", m_KTimeOK, m_def_KTimeOK)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
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
        If bActivado Then
            ' tenemos que convertir de twips a pixels (dividir entre 15)
            If mPt.X < 0 Or mPt.X >= (UserControl.ScaleWidth / Screen.TwipsPerPixelX) Or mPt.Y < 0 Or mPt.Y >= (UserControl.ScaleHeight / Screen.TwipsPerPixelY) Then
                If bFueActivado Then
                    label.ForeColor = m_ColorDOWN
                Else
                    label.ForeColor = m_ColorOFF
                End If
                bActivado = False
                timer.Enabled = False
                bMouseMove = False
                UserControl.MousePointer = vbDefault
                RaiseEvent OnDeactivate
                If AutoSize Then
                    UserControl_Resize
                End If
            End If
        End If
        If bPresionado Then
            label.ForeColor = m_ColorON
            RaiseEvent Click
            L = GetCursorPos(mPt)
            L = ScreenToClient(UserControl.hwnd, mPt)
            If mPt.X < 0 Or mPt.X >= (UserControl.ScaleWidth / Screen.TwipsPerPixelX) Or mPt.Y < 0 Or mPt.Y >= (UserControl.ScaleHeight / Screen.TwipsPerPixelY) Then
                label.ForeColor = m_ColorDOWN
                timer.Enabled = False
                bMouseMove = False
                UserControl.MousePointer = vbDefault
                RaiseEvent OnDeactivate
            Else
                label.ForeColor = m_ColorON
                bActivado = True
                timer.Interval = m_TimeWait
            End If
            bFueActivado = True
            bPresionado = False
        End If
    Else
        timer.Enabled = False
        label.ForeColor = m_ColorOFF
        UserControl.MousePointer = vbDefault
        RaiseEvent OnDeactivate
        If AutoSize Then
            UserControl_Resize
        End If
    End If
End Sub



