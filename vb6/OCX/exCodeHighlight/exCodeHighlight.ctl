VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl exCodeHighlight 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ClipBehavior    =   0  'None
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   2865
   ToolboxBitmap   =   "exCodeHighlight.ctx":0000
   Begin RichTextLib.RichTextBox rich 
      Height          =   2295
      Left            =   180
      TabIndex        =   0
      Top             =   -15
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4048
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"exCodeHighlight.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "exCodeHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------------
' exCodeHighLight.ocx
'---------------------------------------------------------------------------------------------
' SOURCE:       DevDomainCodeHighlight.ctl
' MODIFICADO:   Esau Mayo 2004
'               Este control lo encontré en el PlanetSourceCode, lo necesitaba para
'               el editor de comandos SQL de R_porter.
'               - Agregada la propiedad RightMargin, no se puede establecer ScrollBars en
'               tiempo de ejecucion por eso esta predefinido a rtfBoth y MultiLine es True
'               - Agregada propiedad BoldKeyword y ItalicComment para poner palabras claves
'               en negrita y comentarios en italica, como yo quiero
'               - Se quito el evento VALIDATE pues genera un error critico en sistemas W98
'               (un problema con la pila al parecer) en XP se colgó una vez pero despues ya no
'               pensaba que era un error interno del control, y se habia arreglado con el ServicePack 5
'               no es un evento que necesite y hasta que tenga tiempo de mejorar el codigo...
'               De paso comente todos los eventos que empiezan con OLE...       (._.)
' 29/08/2004:
'               - Agregada propiedad para poner el tamaño de los tabs
'               - Cambiado metodo SelText (para soportar nuevos tabs)
' 10/11/2006:
'               - Agregado Perl y completado VBScript para que reconozca VB tambien
'---------------------------------------------------------------------------------------------
Option Explicit

#Const EX_NO_COPY_VERS = 1

'---------------------------------------------------------------------------------------------
' eventos
'
Public Event SelChange()
Public Event Change()
Public Event Click()
Public Event DblClick()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'---------------------------------------------------------------------------------------------
' enumeraciones publicas
'
Public Enum ItemCodeType
    enumKeyword = 1
    enumOperator = 2
    enumFunction = 3
    enumComment = 4
    enumLiteral = 5
End Enum

Public Enum ProgrammingLanguage
    exNOHighLight = 0
    exVBScript = 1  ' mixed with VB instructions
    exCPP = 2
    exHtml = 3
    exSQL = 4
    exJscript = 5
    exPerl = 6
    exRuby = 7
    exPython = 8
End Enum

Public Enum enumHighlightCode
    exOnNewLine = 0
    exAsType = 1
End Enum

'---------------------------------------------------------------------------------------------
' api
'
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

'---------------------------------------------------------------------------------------------
' propiedades publicas
'
Public CompareCase As VbCompareMethod
Public GiveCorrectCase As Boolean
Attribute GiveCorrectCase.VB_VarMemberFlags = "400"

'---------------------------------------------------------------------------------------------
' constantes privadas
'
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_FMTLINES = &HC8
Private Const EM_LINELENGTH = &HC1
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF
Private Const EM_SETMARGINS = &HD3
Private Const EM_GETMARGINS = &HD4
Private Const EM_CANUNDO = &HC6
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETHANDLE = &HBD
Private Const EM_GETMODIFY = &HB8
Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_GETRECT = &HB2
Private Const EM_GETSEL = &HB0
Private Const EM_GETTHUMB = &HBE
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_LIMITTEXT = &HC5
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB

Private Const EM_LINESCROLL = &HB6
Private Const EM_REPLACESEL = &HC2
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7
Private Const EM_SETHANDLE = &HBC
Private Const EM_SETMODIFY = &HB9
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SETSEL = &HB1
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_SETWORDBREAKPROC = &HD0
Private Const EM_UNDO = &HC7

Private Const EM_EXLINEFROMCHAR = &H436 '<---RICHEDIT.H:    #define EM_EXLINEFROMCHAR  (WM_USER + 54)
                                        '                   #define WM_USER       0x0400
Private Const WM_SETREDRAW = &HB

Private Const CLR_INVALID = -1

Private Const m_def_Author = "Esau R.O."
Private Const m_def_BoldKeyword = False
Private Const m_def_ItalicComment = False
Private Const m_def_LeftMargin = 150

#If EX_NO_COPY_VERS Then

    Private Const m_def_ID_EX = 25647893
    Public ExID As Long
Attribute ExID.VB_VarMemberFlags = "40"
    
#End If

'---------------------------------------------------------------------------------------------
' tipos de datos privados
'
Private Type HightlightedWord
    Word As String
    WordType As ItemCodeType
End Type

Private Type CommentTag
    CommentStart As String
    CommentEnd As String
End Type

Private Type LiteralTag
    LiteralStart As String
    LiteralEnd As String
End Type

'---------------------------------------------------------------------------------------------
' variables privadas
'
Private bFireSelectionChange As Boolean
Private bListenForChange As Boolean
Private strSeparator(27) As String
Private iSeparatorCount As Integer

Private m_Language As ProgrammingLanguage

Private HighLightWords() As HightlightedWord
Private mHighlightCode As enumHighlightCode

Private m_Comment() As CommentTag
Private m_CommentCount As Integer

Private m_Literal() As LiteralTag
Private m_LiteralCount As Integer

Private WordCount As Integer

Private mKeywordColor As OLE_COLOR
Private mOperatorColor As OLE_COLOR
Private mCommentColor As OLE_COLOR
Private mLiteralColor As OLE_COLOR
Private mForeColor As OLE_COLOR
Private mFunctionColor As OLE_COLOR

Private strKeywordColor As String
Private strOperatorColor As String
Private strCommentColor As String
Private strLiteralColor As String
Private strForeColor As String
Private strFunctionColor As String

Private m_Author As String
Private m_BoldKeyword As Boolean
Private m_ItalicComment As Boolean

Private m_Authorized As Boolean
Private m_LeftMargin As Single
Private m_Line As Long
Private m_Row As Long

'---------------------------------------------------------------------------------------------
' Eventos lanzados por el control
Private Sub rich_Change()
    RaiseEvent Change
End Sub

Private Sub rich_Click()
    RaiseEvent Click
End Sub

Private Sub rich_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rich_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rich_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rich_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rich_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub rich_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If m_Language = exNOHighLight Then
        Exit Sub
    End If
    If KeyCode = vbKeyTab Then ' Indent
        Dim SelStart As Long
        If rich.SelLength > 0 Then
            Dim strLines() As String
            Dim LineCount As Long, k As Long
            Dim strResult As String
            strLines = Split(rich.SelText, vbCrLf)
            LineCount = UBound(strLines)
            If LineCount > 0 Then
                SelStart = rich.SelStart
                For k = 0 To LineCount - 1
                    strResult = strResult & vbTab & strLines(k) & vbCrLf
                Next
                strResult = strResult & vbTab & strLines(k)
                rich.SelText = strResult
                rich.SelStart = SelStart
                rich.SelLength = Len(strResult)
                KeyCode = 0
            End If
        End If
    End If

End Sub

Private Sub rich_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_Language = exNOHighLight Then
        Exit Sub
    End If
    Dim k As Byte
    If mHighlightCode = exAsType Then
        For k = 0 To iSeparatorCount
            If KeyAscii = Asc(strSeparator(k)) Then
                    SendMessage rich.hwnd, WM_SETREDRAW, 0, 0&
                    bFireSelectionChange = False
                    Dim TheStart As Long
                    TheStart = rich.SelStart
                    rich.SelStart = Me.LineStartPos(Me.LineIndex)
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(Me.LineIndex))
                    rich.SelStart = TheStart
                    '-------------------------------------------------------------
                    ' [EX] agregado para que el control escriba en modo simple
                    rich.SelBold = False
                    rich.SelItalic = False
                    rich.SelColor = vbBlack
                    '-------------------------------------------------------------
                    SendMessage rich.hwnd, WM_SETREDRAW, 1, 0&
                    rich.Refresh
                    bFireSelectionChange = True
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub rich_SelChange()
    Static lngLastLine As Long
    Dim lngNewLine As Long
    Dim TheStart As Long
    If m_Language = exNOHighLight Then
        RaiseEvent SelChange
        Exit Sub    ' fix bug when exNOHighLight style active
    End If
    If bFireSelectionChange Then
        If rich.SelLength = 0 Then
                bFireSelectionChange = False
                lngNewLine = Me.LineIndex
                If lngNewLine <> lngLastLine Then
                    On Error GoTo Handler
                    SendMessage rich.hwnd, WM_SETREDRAW, 0, 0&
                    TheStart = rich.SelStart
                    rich.SelStart = Me.LineStartPos(lngLastLine)
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(lngLastLine))
Handler:
                    rich.SelStart = TheStart
                    rich.SelLength = SelLength
                    SendMessage rich.hwnd, WM_SETREDRAW, 1, 0&
                    rich.Refresh
                End If
                lngLastLine = lngNewLine
                bFireSelectionChange = True
        End If
        
        m_Line = SendMessage(rich.hwnd, EM_LINEFROMCHAR, rich.SelStart + rich.SelLength, 0&) + 1
        m_Row = rich.SelStart - SendMessage(rich.hwnd, EM_LINEINDEX, m_Line - 1, 0&) + 1
        RaiseEvent SelChange
    End If
End Sub

'---------------------------------------------------------------------------------------------
' Propiedades publicas del control
Public Property Get Locked() As Boolean
    Locked = rich.Locked
End Property

Public Property Let Locked(newLocked As Boolean)
    rich.Locked = newLocked
    PropertyChanged "Locked"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = rich.BackColor
End Property

Public Property Let BackColor(newColor As OLE_COLOR)
    rich.BackColor = newColor
    PropertyChanged "BackColor"
End Property

Public Property Get Font() As StdFont
    Set Font = rich.Font
End Property

Public Property Set Font(newFont As StdFont)
    Set rich.Font = newFont
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(newForeColor As OLE_COLOR)
    mForeColor = newForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get FunctionColor() As OLE_COLOR
    FunctionColor = mFunctionColor
End Property

Public Property Let FunctionColor(newFunctionColor As OLE_COLOR)
    mFunctionColor = newFunctionColor
    strFunctionColor = GetRTFColor(mFunctionColor)
    PropertyChanged "FunctionColor"
End Property

Public Property Get KeywordColor() As OLE_COLOR
    KeywordColor = mKeywordColor
End Property

Public Property Let KeywordColor(newKeywordColor As OLE_COLOR)
    mKeywordColor = newKeywordColor
    strKeywordColor = GetRTFColor(mKeywordColor)
    PropertyChanged "KeywordColor"
End Property

Public Property Get CommentColor() As OLE_COLOR
    CommentColor = mCommentColor
End Property

Public Property Let CommentColor(newCommentColor As OLE_COLOR)
    mCommentColor = newCommentColor
    strCommentColor = GetRTFColor(mCommentColor)
    PropertyChanged "CommentColor"
End Property

Public Property Get LiteralColor() As OLE_COLOR
    LiteralColor = mLiteralColor
End Property

Public Property Let LiteralColor(newLiteralColor As OLE_COLOR)
    mLiteralColor = newLiteralColor
    strLiteralColor = GetRTFColor(mLiteralColor)
    PropertyChanged "LiteralColor"
End Property

Public Property Get OperatorColor() As OLE_COLOR
    OperatorColor = mOperatorColor
End Property

Public Property Let OperatorColor(newOperatorColor As OLE_COLOR)
    mOperatorColor = newOperatorColor
    strOperatorColor = GetRTFColor(mOperatorColor)
    PropertyChanged "OperatorColor"
End Property

Public Property Get LeftMarginColor() As OLE_COLOR
    LeftMarginColor = UserControl.BackColor
End Property

Public Property Let LeftMarginColor(newLeftMarginColor As OLE_COLOR)
    UserControl.BackColor = newLeftMarginColor
    PropertyChanged "LeftMarginColor"
End Property

Public Property Get HighlightCode() As enumHighlightCode
    HighlightCode = mHighlightCode
End Property

Public Property Let HighlightCode(newHighlightCode As enumHighlightCode)
    mHighlightCode = newHighlightCode
    PropertyChanged "HighlightCode"
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = rich.SelLength
End Property

Public Property Let SelLength(lngNewSelLength As Long)
    rich.SelLength = lngNewSelLength
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = rich.SelStart
End Property

Public Property Let SelStart(lngNewSelStart As Long)
    rich.SelStart = lngNewSelStart
End Property

Public Property Get LineIndex() As Long
    LineIndex = SendMessage(rich.hwnd, EM_LINEFROMCHAR, ByVal -1, 0&)
End Property

Public Property Let LineIndex(lngNewLineIndex As Long)
    rich.SelStart = Abs(LineStartPos(lngNewLineIndex))
End Property

Public Property Get Text() As String
    Text = rich.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    rich.TextRTF = HighlightBlock(vNewValue)
    PropertyChanged "Text"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    SelText = rich.SelText
End Property

Public Property Let SelText(newSelText As String)
    bFireSelectionChange = True
    rich.SelText = newSelText
End Property

Public Property Get Language() As ProgrammingLanguage
    Language = m_Language
End Property

Public Property Let Language(ByVal vNewValue As ProgrammingLanguage)
    
    Dim sData As String
    
    If m_Language <> vNewValue Then
        Select Case vNewValue
            Case exVBScript
                SetVBScript
            Case exCPP
                SetCPP
            Case exSQL
                SetSQL
            Case exHtml
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
                AddCommentTag "<", ">"
            Case exJscript
                SetJScript
            Case exPerl
                SetPerl
            Case exRuby
                SetRuby
            Case exPython
                SetPython
            Case exNOHighLight
                resetLanguage
        End Select
        m_Language = vNewValue

        sData = rich.Text
        rich.TextRTF = ""
        rich.SelRTF = HighlightBlock(sData)
        PropertyChanged "Language"
    End If
End Property

Public Property Get RightMargin() As Single
    RightMargin = rich.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
    rich.RightMargin() = New_RightMargin
    PropertyChanged "RightMargin"
End Property

Public Property Get LeftMargin() As Single
    LeftMargin = m_LeftMargin
End Property

Public Property Let LeftMargin(ByVal New_LeftMargin As Single)
    m_LeftMargin = New_LeftMargin
    PropertyChanged "LeftMargin"
End Property

Public Property Get LineLength(CharacterIndex As Long) As Long
    LineLength = SendMessage(rich.hwnd, EM_LINELENGTH, CharacterIndex, 0&)
End Property

Public Property Get LineStartPos(ByVal LineIndex As Long) As Long
    LineStartPos = SendMessage(rich.hwnd, EM_LINEINDEX, LineIndex, 0&)
End Property

Public Property Get Line(lngLine As Long) As String
    '===================================================
    Dim bReturnedLineBuffer() As Byte
    Dim LengthOfLine As Long
    Dim LineStart As Long
    '===================================================
    LineStart = LineStartPos(LineIndex)
    
    If LineStart = -1 Then Exit Function
    
    LengthOfLine = LineLength(LineStart)
    If LengthOfLine < 1 Then Exit Function
    
    ReDim bReturnedLineBuffer(LengthOfLine)

    bReturnedLineBuffer(0) = LengthOfLine And 255
    bReturnedLineBuffer(1) = LengthOfLine \ 256

    SendMessage rich.hwnd, EM_GETLINE, LineIndex, bReturnedLineBuffer(0)

    Line = Left$(StrConv(bReturnedLineBuffer, vbUnicode), LengthOfLine)
    
End Property

Public Property Get SelBold() As Variant
Attribute SelBold.VB_MemberFlags = "400"
    SelBold = rich.SelBold
End Property

Public Property Let SelBold(ByVal New_SelBold As Variant)
    rich.SelBold = New_SelBold
    PropertyChanged "SelBold"
End Property

Public Property Get SelItalic() As Variant
Attribute SelItalic.VB_MemberFlags = "400"
    SelItalic = rich.SelItalic
End Property

Public Property Let SelItalic(ByVal New_SelItalic As Variant)
    rich.SelItalic = New_SelItalic
    PropertyChanged "SelItalic"
End Property

Public Property Get Author() As String
    Author = m_Author
End Property

Public Property Let Author(ByVal New_Author As String)
    If Ambient.UserMode = False Then
        Err.Raise 387
    Else
        Err.Raise 382
    End If
End Property

Public Property Get BoldKeyword() As Boolean
    BoldKeyword = m_BoldKeyword
End Property

Public Property Let BoldKeyword(ByVal New_BoldKeyword As Boolean)
    m_BoldKeyword = New_BoldKeyword
    PropertyChanged "BoldKeyword"
End Property

Public Property Get ItalicComment() As Boolean
    ItalicComment = m_ItalicComment
End Property

Public Property Let ItalicComment(ByVal New_ItalicComment As Boolean)
    m_ItalicComment = New_ItalicComment
    PropertyChanged "ItalicComment"
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_MemberFlags = "400"
    SelRTF = rich.SelRTF
End Property

Public Property Let SelRTF(ByVal New_SelRTF As String)
    rich.SelRTF = New_SelRTF
    PropertyChanged "SelRTF"
End Property

Public Property Get RichHwnd() As Long
    RichHwnd = rich.hwnd
End Property

Public Property Get CursorLine() As Long
    CursorLine = m_Line
End Property

Public Property Get CursorRow() As Long
    CursorRow = m_Row
End Property

Public Sub SaveFile(ByVal strFilename As String, ByVal vFlags As LoadSaveConstants)
    
    On Error GoTo Handler
    rich.SaveFile strFilename, vFlags
    Exit Sub

Handler:
    MsgBox "exCodeHighLight" & vbCrLf & "[" & Err.Number & "]: " & Err.Description, vbCritical, "SaveFile"
End Sub

Public Sub Refresh()
   rich.Refresh
End Sub

Public Sub AddCommentTag(ByVal CommentTagStart As String, ByVal CommentTagEnd As String)
    ReDim Preserve m_Comment(m_CommentCount)
    With m_Comment(m_CommentCount)
        .CommentStart = CommentTagStart
        .CommentEnd = CommentTagEnd
    End With
    m_CommentCount = m_CommentCount + 1
End Sub

Public Sub AddLiteralTag(ByVal LiteralTagStart As String, ByVal LiteralTagEnd As String)
    ReDim Preserve m_Literal(m_LiteralCount)
    With m_Literal(m_LiteralCount)
        .LiteralStart = LiteralTagStart
        .LiteralEnd = LiteralTagEnd
    End With
    m_LiteralCount = m_LiteralCount + 1
End Sub

Public Sub LoadFile(strFilename)
    '===================================================
    Dim FileNum As Integer
    Dim sData As String
    Dim bInit As Boolean
    '===================================================
    On Error GoTo Handler
   
    FileNum = FreeFile
    bInit = True
    Open strFilename For Input As FileNum
    rich.Text = ""
    rich.TextRTF = ""
    
    'LockWindowUpdate rich.hwnd
    SendMessage rich.hwnd, WM_SETREDRAW, 0, 0&
    bFireSelectionChange = False
    
    Do
        Line Input #FileNum, sData  '<---- Throws error 62 when end of file
        If bInit Then
            bInit = False
        Else
            sData = vbCrLf & sData
        End If

        ' El problema con esta instruccion es que la carga de un archivo es muy lenta e ineficiente
        ' pero no tiene el problema con los tabs
        ' rich.SelText = sData
        
        ' Esta instruccion carga los archivos mucho mas rapido pero no respeta los tabs de 4
        ' se corrigio modificando la funcion: HighlightBlock(...)
        rich.SelRTF = HighlightBlock(sData)
    Loop
    SendMessage rich.hwnd, WM_SETREDRAW, 1, 0& ''
    rich.Refresh
    Exit Sub
    
Handler:
    If Err.Number = 62 Then
        'LockWindowUpdate 0
        SendMessage rich.hwnd, WM_SETREDRAW, 1, 0&
        rich.Refresh
        rich.SelText = vbCrLf
        Close FileNum
        bFireSelectionChange = True
        Exit Sub
    Else
        MsgBox "exCodeHighLight" & vbCrLf & "[" & Err.Number & "]: " & Err.Description, vbCritical, "LoadFile"
        LockWindowUpdate 0
        SendMessage rich.hwnd, WM_SETREDRAW, True, 0&
        'rich.Refresh
    End If
End Sub

Public Sub AddWord(ByVal Word As String, Optional WordType As ItemCodeType = enumKeyword)
    ReDim Preserve HighLightWords(WordCount)
    If WordType = enumComment Then
        AddCommentTag Word, Word
    Else
        If WordType = enumLiteral Then
            AddLiteralTag Word, Word
        Else
            With HighLightWords(WordCount)
                .Word = Word
                .WordType = WordType
            End With
            WordCount = WordCount + 1
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------------
' user control
'---------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    '===================================================
    Dim clx As clsCrypto
    '===================================================
    Set clx = New clsCrypto
    
    clx.SetCod 657865
    '-----------------------------------------------------------------------------------------------
    'm_Author = clx.Encrypt("Esau R.O. [exe_q_tor] ...based in the DevDomainCodeHighlight control.")
    m_Author = clx.Decrypt("LT)""*hnKn*WlVlD8D;1~0*nnn<)Tlx*so*;Ol*vlcv1#)soe1xl9smO]smO;*P1o;~1]n")
    
    Set clx = Nothing
    
    m_BoldKeyword = m_def_BoldKeyword
    m_ItalicComment = m_def_ItalicComment
    
End Sub

Private Sub UserControl_Initialize()
    
    strSeparator(0) = " "
    strSeparator(1) = vbCrLf
    strSeparator(2) = vbTab
    strSeparator(3) = "("
    strSeparator(4) = ")"
    strSeparator(5) = "="
    strSeparator(6) = "+"
    strSeparator(7) = "-"
    strSeparator(8) = "*"
    strSeparator(9) = ">"
    strSeparator(10) = "<"
    strSeparator(11) = "\"
    strSeparator(12) = "/"
    strSeparator(13) = "{"
    strSeparator(14) = "}"
    strSeparator(15) = "["
    strSeparator(16) = "]"
    strSeparator(17) = "|"
    strSeparator(18) = "&"
    strSeparator(19) = ":"
    strSeparator(20) = ";"
    strSeparator(21) = ","
    strSeparator(22) = "?"
    strSeparator(23) = "."
    strSeparator(24) = "$"  ' Perl
    strSeparator(25) = "@"
    strSeparator(26) = "%"
    iSeparatorCount = 26
    bFireSelectionChange = True
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    rich.Text = PropBag.ReadProperty("Text", "")
    rich.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Set rich.Font = PropBag.ReadProperty("Font", rich.Font)
    rich.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    rich.SelBold = PropBag.ReadProperty("SelBold", False)
    rich.SelItalic = PropBag.ReadProperty("SelItalic", False)
    rich.SelRTF = PropBag.ReadProperty("SelRTF", "")
    rich.Locked = PropBag.ReadProperty("Locked", False)
    
    UserControl.BackColor = PropBag.ReadProperty("LeftMarginColor", vbWhite)
    
    Language = PropBag.ReadProperty("Language", exNOHighLight)
    KeywordColor = PropBag.ReadProperty("KeywordColor", vbBlue)
    OperatorColor = PropBag.ReadProperty("OperatorColor", vbBlue)
    CommentColor = PropBag.ReadProperty("CommentColor", vbCyan)
    LiteralColor = PropBag.ReadProperty("LiteralColor", vbRed)
    mForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    FunctionColor = PropBag.ReadProperty("FunctionColor", vbMagenta)
    HighlightCode = PropBag.ReadProperty("HighlightCode", 1)
    m_Author = PropBag.ReadProperty("Author", m_def_Author)
    m_BoldKeyword = PropBag.ReadProperty("BoldKeyword", m_def_BoldKeyword)
    m_ItalicComment = PropBag.ReadProperty("ItalicComment", m_def_ItalicComment)
    m_LeftMargin = PropBag.ReadProperty("LeftMargin", m_def_LeftMargin)
    
End Sub

Private Sub UserControl_Resize()
    rich.Move m_LeftMargin, 0, UserControl.ScaleWidth - m_LeftMargin, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()

    #If EX_NO_COPY_VERS Then
        If Ambient.UserMode Then
            If ExID <> m_def_ID_EX Then
                m_Language = exNOHighLight
                m_BoldKeyword = False
                m_ItalicComment = False
            Else
                m_Authorized = True
            End If
        End If
    #Else
        m_Authorized = True
    #End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", rich.Text, ""
    PropBag.WriteProperty "Font", rich.Font
    PropBag.WriteProperty "BackColor", rich.BackColor, vbWindowBackground
    PropBag.WriteProperty "RightMargin", rich.RightMargin, 0
    PropBag.WriteProperty "SelBold", rich.SelBold, False
    PropBag.WriteProperty "SelItalic", rich.SelItalic, False
    PropBag.WriteProperty "SelRTF", rich.SelRTF, ""
    PropBag.WriteProperty "Locked", rich.Locked, False
    
    PropBag.WriteProperty "Language", m_Language, exNOHighLight
    PropBag.WriteProperty "KeywordColor", mKeywordColor, vbBlue
    PropBag.WriteProperty "OperatorColor", mOperatorColor, vbBlue
    PropBag.WriteProperty "CommentColor", mCommentColor, vbCyan
    PropBag.WriteProperty "LiteralColor", mLiteralColor, vbRed
    PropBag.WriteProperty "ForeColor", mForeColor, vbWindowText
    PropBag.WriteProperty "FunctionColor", mFunctionColor, vbMagenta
    PropBag.WriteProperty "HighlightCode", mHighlightCode, 1
    PropBag.WriteProperty "Author", m_Author, m_def_Author
    PropBag.WriteProperty "BoldKeyword", m_BoldKeyword, m_def_BoldKeyword
    PropBag.WriteProperty "ItalicComment", m_ItalicComment, m_def_ItalicComment
    PropBag.WriteProperty "LeftMargin", m_LeftMargin, m_def_LeftMargin
    PropBag.WriteProperty "LeftMarginColor", UserControl.BackColor, vbWhite
End Sub

'---------------------------------------------------------------------------------------------
' funciones privadas
'
Private Function ColorWord(ByVal sWord As String) As String
    Dim iWord As Integer
    For iWord = 0 To WordCount - 1
    
        If StrComp(sWord, HighLightWords(iWord).Word, CompareCase) = 0 Then
            If GiveCorrectCase Then sWord = HighLightWords(iWord).Word
                '---------------------------------------------------------------------------
                ' se agrego la posibilidad de poner palabras clave en negrita
                If m_BoldKeyword Then
                    ColorWord = "{\cf" & HighLightWords(iWord).WordType & "\b\i0 " & sWord & "}"
                Else
                    ColorWord = "{\cf" & HighLightWords(iWord).WordType & "\b0\i0 " & sWord & "}"
                End If
            Exit Function
        End If
    Next
    ColorWord = "{\cf0\b0\i0 " & sWord & "}"    ' correccion XP
End Function

Private Function GetRTFColor(Color As OLE_COLOR) As String
    Dim lrgb As Long
    lrgb = TranslateColor(Color)
    GetRTFColor = "\red" & (lrgb And &HFF&) & "\green" & (lrgb And &HFF00&) \ &H100 & "\blue" & (lrgb And &HFF0000) \ &H10000 & ";"
End Function

Private Function GetWord(sBlock As String, lngWordStart As Long, lngCharPos As Long, sSep As String) As String
    Dim sWord As String
    On Error GoTo JMP_EXIT
    sWord = Mid$(sBlock, lngWordStart, lngCharPos - lngWordStart)
        If sSep = vbCrLf Then
            sSep = "\par " & vbCrLf
        ElseIf sSep = vbTab Then
                sSep = "\tab "
        ElseIf sSep = "\" Then
                sSep = "\cf2 \\\cf0 "
        ElseIf sSep = "{" Then
                sSep = "\cf2\b0\i0 \{\cf0 "
        ElseIf sSep = "}" Then
                sSep = "\cf2\b0\i0 \}\cf0 "
        ElseIf sSep <> " " And Len(sSep) Then
            sSep = "\cf2\b0\i0 " & sSep & "\cf0 "   'XP FIX (correct bold and italic operators)
        End If
        If lngCharPos - lngWordStart > 0 Then
            GetWord = ColorWord(sWord) & sSep
        Else
            GetWord = sSep
        End If
JMP_EXIT:
End Function

Private Function HighlightComment(sComment As String, sEndofComment As String) As String
    sComment = Replace(sComment, "\", "\\")
    sComment = Replace(sComment, "{", "\{")
    sComment = Replace(sComment, "}", "\}")
    sComment = Replace(sComment, vbCrLf, "\par ")
    If sEndofComment = vbCrLf Then
        sComment = sComment & "\par" & vbCrLf
    Else
        If sEndofComment = vbTab Then
            sComment = sComment & "\tab "
        Else
            sComment = sComment & sEndofComment
        End If
    End If
    
    If ((StartOfLiteral(sComment, 1)) <> -1) Then
        ' si el bloque es de un literal
        HighlightComment = "{\cf5\i0\b0 " & sComment & "}"
    Else
        '---------------------------------------------------------------------------
        ' se agrego la posibilidad de poner comentarios en italica
        If m_ItalicComment Then
            HighlightComment = "{\cf4\i\b0 " & sComment & "}"
        Else
            HighlightComment = "{\cf4\i0\b0 " & sComment & "}" ' correccion XP
        End If
    End If
End Function

Private Function StartOfComment(sBlock As String, lngCharPos As Long) As Integer
    '===================================================
    Dim sChar As String
    Dim k As Byte
    '===================================================
    For k = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(k).CommentStart))
        If sChar = m_Comment(k).CommentStart Then
            StartOfComment = k
            Exit Function
        End If
    Next
    StartOfComment = -1
End Function

Private Function StartOfLiteral(sBlock As String, lngCharPos As Long) As Integer
    '===================================================
    Dim sChar As String
    Dim k As Byte
    '===================================================
    For k = 0 To m_LiteralCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Literal(k).LiteralStart))
        If sChar = m_Literal(k).LiteralStart Then
            StartOfLiteral = k
            Exit Function
        End If
    Next
    StartOfLiteral = -1
End Function

Private Function isSeparator(sBlock As String, lngCharPos As Long) As String
    '===================================================
    Dim sChar As String
    Dim k As Byte
    '===================================================
    For k = 0 To iSeparatorCount
        sChar = Mid$(sBlock, lngCharPos, Len(strSeparator(k)))
        If sChar = strSeparator(k) Then
            isSeparator = sChar
            Exit Function
        End If
    Next
End Function

Private Function EndOfComment(sBlock As String, lngCharPos As Long) As Integer
    '===================================================
    Dim sChar As String
    Dim k As Byte
    '===================================================
    For k = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(k).CommentEnd))
        If sChar = m_Comment(k).CommentEnd Then
            EndOfComment = k
            Exit Function
        End If
    Next
    EndOfComment = -1
End Function

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Function HighlightBlock(sBlock As String) As String
    '===================================================
    Dim lngCharPos As Long
    Dim lngBlockLength As Long
    Dim sWord As String
    Dim lngCommentStartPos As Long
    Dim byteStartOfComment As Integer
    Dim byteEndOfComment As Integer
    Dim sSep As String
    Dim lngWordStart As Long
    Dim sHighlighted As String
    Dim T As Integer
    Dim bWordFound As Boolean
    Dim bLastStepWasComment As Boolean
    '===================================================

    If m_Language = exNOHighLight Or Not m_Authorized Then
        HighlightBlock = sBlock
        Exit Function
    End If
    
    lngBlockLength = Len(sBlock)
    lngWordStart = 1
    byteStartOfComment = -1
    
    For lngCharPos = 1 To lngBlockLength
        
        T = StartOfComment(sBlock, lngCharPos)
        
        If T > -1 And byteStartOfComment = -1 Then
            lngCommentStartPos = lngCharPos
            byteStartOfComment = T
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
           If byteStartOfComment > -1 Then
                byteEndOfComment = EndOfComment(sBlock, lngCharPos)
                If byteEndOfComment > -1 And byteEndOfComment = byteStartOfComment Then
                    
                    sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), m_Comment(byteEndOfComment).CommentEnd)

                    byteStartOfComment = -1
                    bLastStepWasComment = True
                    lngWordStart = lngCharPos + Len(m_Comment(byteEndOfComment).CommentEnd)
                End If
            Else
                If byteStartOfComment = -1 Then
                    
                    sSep = isSeparator(sBlock, lngCharPos)
                    Dim SepLength As ItemCodeType
                    SepLength = Len(sSep)
                    If SepLength > 0 Then
                        If lngCharPos <= lngBlockLength Then
                            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, sSep)
                        End If
                        lngWordStart = lngCharPos + SepLength
                        bLastStepWasComment = False
                    End If
                End If
            End If
        End If
    Next
    
    If byteStartOfComment > -1 Then
        
        Dim lngCommentEndPos As Long
        lngCommentEndPos = InStr(lngCharPos, rich.Text, m_Comment(byteStartOfComment).CommentEnd)
        If lngCommentEndPos = 0 Then lngCommentEndPos = Len(rich.Text)
        sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), "")
    Else
        If bLastStepWasComment Then
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
            If lngBlockLength - lngWordStart >= 0 Then
                sWord = Mid$(sBlock, lngWordStart, (lngBlockLength - lngWordStart) + 1)
                sHighlighted = sHighlighted & ColorWord(sWord)
            End If
        End If
    End If
    
    If Len(sHighlighted) = 0 Then Exit Function
    
    ' se agrego codigo para poner tabs de 4 espacios (para la carga de archivos mas rapida)
    HighlightBlock = "{{\colortbl ;" & strKeywordColor & strOperatorColor & strFunctionColor & strCommentColor & strLiteralColor & vbCrLf & _
                     " \tx420\tx840\tx1260\tx1680\tx2100\tx2520\tx2940\tx3360\tx3780\tx4200\tx4620\tx5040\tx5460\tx5880\tx6300\tx6720\tx7140\tx7560\tx7980\tx8400\tx8820\tx9240\tx9660\tx10080\tx10500\tx10920\tx11340\tx11760\tx12180\tx12600\tx13020\tx13440}" & vbCrLf & _
                     sHighlighted & "}"
    
End Function

Private Sub resetLanguage()
    Erase HighLightWords
    WordCount = 0
    Erase m_Comment
    m_CommentCount = 0
    Erase m_Literal
    m_LiteralCount = 0
End Sub

Private Sub SetJScript()
    
    '------------------------------
    ' Case
    CompareCase = vbBinaryCompare
    GiveCorrectCase = False
    resetLanguage
    
    AddWord "@cc_on"
    AddWord "@if"
    AddWord "@set"
    AddWord "break"
    AddWord "continue"
    AddWord "delete"
    AddWord "do"
    AddWord "else"
    AddWord "for"
    AddWord "function"
    AddWord "if"
    AddWord "in"
    AddWord "new"
    AddWord "return"
    AddWord "switch"
    AddWord "this"
    AddWord "typeof"
    AddWord "var"
    AddWord "void"
    AddWord "while"
    AddWord "with"
    
    '------------------------------
    ' JScript
    AddWord "ScriptEngine", enumFunction
    AddWord "ScriptEngineBuildVersion", enumFunction
    AddWord "ScriptEngineMajorVersion", enumFunction
    AddWord "ScriptEngineMinorVersion", enumFunction
    
    AddWord "abs", enumFunction
    AddWord "acos", enumFunction
    AddWord "Add", enumFunction
    AddWord "anchor", enumFunction
    AddWord "asin", enumFunction
    AddWord "atan", enumFunction
    AddWord "atan2", enumFunction
    AddWord "atEnd", enumFunction
    AddWord "big", enumFunction
    AddWord "blink", enumFunction
    AddWord "bold", enumFunction
    AddWord "BuildPath", enumFunction
    AddWord "ceil", enumFunction
    AddWord "charAt", enumFunction
    AddWord "charCodeAt", enumFunction
    AddWord "Close", enumFunction
    AddWord "compile", enumFunction
    AddWord "concat", enumFunction
    AddWord "Copy", enumFunction
    AddWord "CopyFile", enumFunction
    AddWord "CopyFolder", enumFunction
    AddWord "cos", enumFunction
    AddWord "CreateFolder", enumFunction
    AddWord "CreateTextFile", enumFunction
    AddWord "Delete", enumFunction
    AddWord "DeleteFile", enumFunction
    AddWord "DeleteFolder", enumFunction
    AddWord "dimensions", enumFunction
    AddWord "DriveExists", enumFunction
    AddWord "escape", enumFunction
    AddWord "eval", enumFunction
    AddWord "exec", enumFunction
    AddWord "Exists", enumFunction
    AddWord "exp", enumFunction
    AddWord "FileExists", enumFunction
    AddWord "fixed", enumFunction
    AddWord "floor", enumFunction
    AddWord "FolderExists", enumFunction
    AddWord "fontcolor", enumFunction
    AddWord "fontsize", enumFunction
    AddWord "fromCharCode", enumFunction
    AddWord "GetAbsolutePathName", enumFunction
    AddWord "GetBaseName", enumFunction
    AddWord "getDate", enumFunction
    AddWord "getDay", enumFunction
    AddWord "GetDrive", enumFunction
    AddWord "GetDriveName", enumFunction
    AddWord "GetExtensionName", enumFunction
    AddWord "GetFile", enumFunction
    AddWord "GetFileName", enumFunction
    AddWord "GetFolder", enumFunction
    AddWord "getFullYear", enumFunction
    AddWord "getHours", enumFunction
    AddWord "getItem", enumFunction
    AddWord "getMilliseconds", enumFunction
    AddWord "getMinutes", enumFunction
    AddWord "getMonth", enumFunction
    AddWord "GetParentFolderName", enumFunction
    AddWord "getSeconds", enumFunction
    AddWord "GetSpecialFolder", enumFunction
    AddWord "GetTempName", enumFunction
    AddWord "getTime", enumFunction
    AddWord "getTimezoneOffset", enumFunction
    AddWord "getUTCDate", enumFunction
    AddWord "getUTCDay", enumFunction
    AddWord "getUTCFullYear", enumFunction
    AddWord "getUTCHours", enumFunction
    AddWord "getUTCMilliseconds", enumFunction
    AddWord "getUTCMinutes", enumFunction
    AddWord "getUTCMonth", enumFunction
    AddWord "getUTCSeconds", enumFunction
    AddWord "getVarDate", enumFunction
    AddWord "getYear", enumFunction
    AddWord "indexOf", enumFunction
    AddWord "isFinite", enumFunction
    AddWord "isNaN", enumFunction
    AddWord "italics", enumFunction
    AddWord "item", enumFunction
    AddWord "Items", enumFunction
    AddWord "join", enumFunction
    AddWord "Keys", enumFunction
    AddWord "lastIndexOf", enumFunction
    AddWord "lbound", enumFunction
    AddWord "link", enumFunction
    AddWord "log", enumFunction
    AddWord "match", enumFunction
    AddWord "max", enumFunction
    AddWord "min", enumFunction
    AddWord "Move", enumFunction
    AddWord "MoveFile", enumFunction
    AddWord "moveFirst", enumFunction
    AddWord "MoveFolder", enumFunction
    AddWord "moveNext", enumFunction
    AddWord "OpenAsTextStream", enumFunction
    AddWord "OpenTextFile", enumFunction
    AddWord "parse", enumFunction
    AddWord "parseFloat", enumFunction
    AddWord "parseInt", enumFunction
    AddWord "pow", enumFunction
    AddWord "random", enumFunction
    AddWord "Read", enumFunction
    AddWord "ReadAll", enumFunction
    AddWord "ReadLine", enumFunction
    AddWord "Remove", enumFunction
    AddWord "RemoveAll", enumFunction
    AddWord "replace", enumFunction
    AddWord "reverse", enumFunction
    AddWord "round", enumFunction
    AddWord "search", enumFunction
    AddWord "setDate", enumFunction
    AddWord "setFullYear", enumFunction
    AddWord "setHours", enumFunction
    AddWord "setMilliseconds", enumFunction
    AddWord "setMinutes", enumFunction
    AddWord "setMonth", enumFunction
    AddWord "setSeconds", enumFunction
    AddWord "setTime", enumFunction
    AddWord "setUTCDate", enumFunction
    AddWord "setUTCFullYear", enumFunction
    AddWord "setUTCHours", enumFunction
    AddWord "setUTCMilliseconds", enumFunction
    AddWord "setUTCMinutes", enumFunction
    AddWord "setUTCMonth", enumFunction
    AddWord "setUTCSeconds", enumFunction
    AddWord "setYear", enumFunction
    AddWord "sin", enumFunction
    AddWord "Skip", enumFunction
    AddWord "SkipLine", enumFunction
    AddWord "slice", enumFunction
    AddWord "small", enumFunction
    AddWord "sort", enumFunction
    AddWord "split", enumFunction
    AddWord "sqrt", enumFunction
    AddWord "strike", enumFunction
    AddWord "sub", enumFunction
    AddWord "substr", enumFunction
    AddWord "substring", enumFunction
    AddWord "sup", enumFunction
    AddWord "tan", enumFunction
    AddWord "test", enumFunction
    AddWord "toArray", enumFunction
    AddWord "toGMTString", enumFunction
    AddWord "toLocaleString", enumFunction
    AddWord "toLowerCase", enumFunction
    AddWord "toString", enumFunction
    AddWord "toUpperCase", enumFunction
    AddWord "toUTCString", enumFunction
    AddWord "ubound", enumFunction
    AddWord "unescape", enumFunction
    AddWord "UTC", enumFunction
    AddWord "valueOf", enumFunction
    AddWord "Write", enumFunction
    AddWord "WriteBlankLines", enumFunction
    AddWord "WriteLine", enumFunction
    
    AddWord "ActiveXObject", enumFunction
    AddWord "Array", enumFunction
    AddWord "Boolean", enumFunction
    AddWord "Date", enumFunction
    AddWord "Enumerator", enumFunction
    AddWord "Function", enumFunction
    AddWord "Math", enumFunction
    AddWord "Number", enumFunction
    AddWord "Object", enumFunction
    AddWord "RegExp", enumFunction
    AddWord "TextStream", enumFunction
    AddWord "VBArray", enumFunction
    
    AddWord "arguments", enumFunction
    AddWord "AtEndOfLine", enumFunction
    AddWord "AtEndOfStream", enumFunction
    AddWord "Attributes", enumFunction
    AddWord "AvailableSpace", enumFunction
    AddWord "caller", enumFunction
    AddWord "Column", enumFunction
    AddWord "CompareMode", enumFunction
    AddWord "constructor", enumFunction
    AddWord "Count", enumFunction
    AddWord "DateCreated", enumFunction
    AddWord "DateLastAccessed", enumFunction
    AddWord "DateLastModified", enumFunction
    AddWord "Drive", enumFunction
    AddWord "DriveLetter", enumFunction
    AddWord "Drives", enumFunction
    AddWord "DriveType", enumFunction
    AddWord "E", enumFunction
    AddWord "Files", enumFunction
    AddWord "FileSystem", enumFunction
    AddWord "FreeSpace", enumFunction
    AddWord "global", enumFunction
    AddWord "ignoreCase", enumFunction
    AddWord "index", enumFunction
    AddWord "Infinity", enumFunction
    AddWord "input", enumFunction
    AddWord "IsReady", enumFunction
    AddWord "IsRootFolder", enumFunction
    AddWord "Item", enumFunction
    AddWord "Key", enumFunction
    AddWord "lastIndex", enumFunction
    AddWord "lastMatch", enumFunction
    AddWord "lastParen", enumFunction
    AddWord "leftContext", enumFunction
    AddWord "length", enumFunction
    AddWord "Line", enumFunction
    AddWord "LN10", enumFunction
    AddWord "LN2", enumFunction
    AddWord "LOG10E", enumFunction
    AddWord "LOG2E", enumFunction
    AddWord "MAX_VALUE", enumFunction
    AddWord "MIN_VALUE", enumFunction
    AddWord "multiline", enumFunction
    AddWord "Name", enumFunction
    AddWord "NaN", enumFunction
    AddWord "NEGATIVE_INFINITY", enumFunction
    AddWord "ParentFolder", enumFunction
    AddWord "Path", enumFunction
    AddWord "PI", enumFunction
    AddWord "POSITIVE_INFINITY", enumFunction
    AddWord "prototype", enumFunction
    AddWord "rightContext", enumFunction
    AddWord "RootFolder", enumFunction
    AddWord "SerialNumber", enumFunction
    AddWord "ShareName", enumFunction
    AddWord "ShortName", enumFunction
    AddWord "ShortPath", enumFunction
    AddWord "Size", enumFunction
    AddWord "source", enumFunction
    AddWord "SQRT1_2", enumFunction
    AddWord "SQRT2", enumFunction
    AddWord "SubFolders", enumFunction
    AddWord "TotalSize", enumFunction
    AddWord "Type", enumFunction
    AddWord "VolumeName", enumFunction
    
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "*", enumOperator
    AddWord "/", enumOperator
    AddWord "%", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "^", enumOperator
    AddWord "!", enumOperator
    AddWord "~", enumOperator
    AddWord "=", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    
    AddCommentTag "//", vbCrLf
    'AddCommentTag "/*", "*/"   '[TODO] comentarios multilinea
    AddCommentTag """", """"
    AddCommentTag "'", "'"
    
    '------------------------------
    ' Literales
    AddLiteralTag """", """"
    AddLiteralTag "'", "'"
    
End Sub

Private Sub SetCPP()
    
    '------------------------------
    ' Case
    CompareCase = vbBinaryCompare
    GiveCorrectCase = False
    resetLanguage
    
    '------------------------------
    ' Reserved words
    AddWord "asm"
    AddWord "auto"
    AddWord "bool"
    AddWord "break"
    AddWord "case"
    AddWord "catch"
    AddWord "char"
    AddWord "class"
    AddWord "const"
    AddWord "const_cast"
    AddWord "continue"
    AddWord "default"
    AddWord "delete"
    AddWord "do"
    AddWord "double"
    AddWord "dynamic_cast"
    AddWord "else"
    AddWord "enum"
    AddWord "explicit"
    AddWord "export"
    AddWord "extern"
    AddWord "false"
    AddWord "float"
    AddWord "for"
    AddWord "friend"
    AddWord "goto"
    AddWord "if"
    AddWord "inline"
    AddWord "int"
    AddWord "long"
    AddWord "mutable"
    AddWord "namespace"
    AddWord "new"
    AddWord "operator"
    AddWord "private"
    AddWord "protected"
    AddWord "public"
    AddWord "register"
    AddWord "reinterpret_cast"
    AddWord "return"
    AddWord "short"
    AddWord "signed"
    AddWord "sizeof"
    AddWord "static"
    AddWord "static_cast"
    AddWord "struct"
    AddWord "switch"
    AddWord "template"
    AddWord "this"
    AddWord "throw"
    AddWord "true"
    AddWord "try"
    AddWord "typedef"
    AddWord "typeid"
    AddWord "union"
    AddWord "unsigned"
    AddWord "using"
    AddWord "virtual"
    AddWord "void"
    AddWord "volatile"
    AddWord "wchar_t"
    AddWord "while"
    
    '------------------------------
    ' Preprocesador
    AddWord "#define"
    AddWord "#elif"
    AddWord "#else"
    AddWord "#endif"
    AddWord "#error"
    AddWord "#if"
    AddWord "#ifdef"
    AddWord "#ifndef"
    AddWord "#include"
    AddWord "#line"
    AddWord "#pragma"
    AddWord "#undef"
    
    '------------------------------
    ' STL
    AddWord "algorithm", enumFunction
    AddWord "assign", enumFunction
    AddWord "begin", enumFunction
    AddWord "cctype", enumFunction
    AddWord "cin", enumFunction
    AddWord "count", enumFunction
    AddWord "cout", enumFunction
    AddWord "end", enumFunction
    AddWord "endl", enumFunction
    AddWord "fill", enumFunction
    AddWord "fill_n", enumFunction
    AddWord "find", enumFunction
    AddWord "first", enumFunction
    AddWord "functional", enumFunction
    AddWord "greater", enumFunction
    AddWord "insert", enumFunction
    AddWord "iostream", enumFunction
    AddWord "isalnum", enumFunction
    AddWord "isalpha", enumFunction
    AddWord "iscntrl", enumFunction
    AddWord "isdigit", enumFunction
    AddWord "isgraph", enumFunction
    AddWord "islower", enumFunction
    AddWord "isprint", enumFunction
    AddWord "ispunct", enumFunction
    AddWord "isspace", enumFunction
    AddWord "isupper", enumFunction
    AddWord "isxdigit", enumFunction
    AddWord "iterator", enumFunction
    AddWord "less", enumFunction
    AddWord "map", enumFunction
    AddWord "max", enumFunction
    AddWord "min", enumFunction
    AddWord "npos", enumFunction
    AddWord "ostringstream", enumFunction
    AddWord "push_back", enumFunction
    AddWord "second", enumFunction
    AddWord "size", enumFunction
    AddWord "size_type", enumFunction
    AddWord "sort", enumFunction
    AddWord "sstream", enumFunction
    AddWord "std", enumFunction
    AddWord "string", enumFunction
    AddWord "toupper", enumFunction
    AddWord "tolower", enumFunction
    AddWord "substr", enumFunction
    AddWord "vector", enumFunction
    
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "*", enumOperator
    AddWord "/", enumOperator
    AddWord "%", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "^", enumOperator
    AddWord "!", enumOperator
    AddWord "~", enumOperator
    AddWord "=", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    
    AddCommentTag "//", vbCrLf
    'AddCommentTag "/*", "*/"   '[TODO] comentarios multilinea
    AddCommentTag """", """"
    AddCommentTag "'", "'"
    
    AddLiteralTag """", """"
    AddLiteralTag "'", "'"
    
End Sub

Private Sub SetPerl()
    
    '------------------------------
    ' Case
    CompareCase = vbBinaryCompare
    GiveCorrectCase = False
    resetLanguage
    
    '------------------------------
    ' Reserved words
    AddWord "continue"
    AddWord "do"
    AddWord "else"
    AddWord "elsif"
    AddWord "for"
    AddWord "foreach"
    AddWord "goto"
    AddWord "if"
    AddWord "last"
    AddWord "local"
    AddWord "lock"
    AddWord "map"
    AddWord "my"
    AddWord "next"
    AddWord "package"
    AddWord "redo"
    AddWord "require"
    AddWord "return"
    AddWord "sub"
    AddWord "unless"
    AddWord "until"
    AddWord "use"
    AddWord "while"
    AddWord "STDIN"
    AddWord "STDOUT"
    AddWord "STDERR"
    AddWord "ARGV"
    AddWord "ARGVOUT"
    AddWord "ENV"
    AddWord "INC"
    AddWord "SIG"
    AddWord "TRUE"
    AddWord "FALSE"
    AddWord "__FILE__"
    AddWord "__LINE__"
    AddWord "__PACKAGE__"
    AddWord "__END__"
    AddWord "__DATA__"
    AddWord "lt"
    AddWord "gt"
    AddWord "le"
    AddWord "ge"
    AddWord "eq"
    AddWord "ne"
    AddWord "cmp"
    AddWord "x"
    AddWord "not"
    AddWord "and"
    AddWord "or"
    AddWord "xor"
    AddWord "q"
    AddWord "qq"
    AddWord "qx"
    AddWord "qw"
    
    '------------------------------
    ' Functions
    AddWord "abs", enumFunction
    AddWord "accept", enumFunction
    AddWord "alarm", enumFunction
    AddWord "atan2", enumFunction
    AddWord "bind", enumFunction
    AddWord "binmode", enumFunction
    AddWord "bless", enumFunction
    AddWord "caller", enumFunction
    AddWord "chdir", enumFunction
    AddWord "chmod", enumFunction
    AddWord "chomp", enumFunction
    AddWord "chop", enumFunction
    AddWord "chown", enumFunction
    AddWord "chr", enumFunction
    AddWord "chroot", enumFunction
    AddWord "close", enumFunction
    AddWord "closedir", enumFunction
    AddWord "connect", enumFunction
    AddWord "cos", enumFunction
    AddWord "crypt", enumFunction
    AddWord "dbmclose", enumFunction
    AddWord "dbmopen", enumFunction
    AddWord "defined", enumFunction
    AddWord "delete", enumFunction
    AddWord "die", enumFunction
    AddWord "dump", enumFunction
    AddWord "each", enumFunction
    AddWord "eof", enumFunction
    AddWord "eval", enumFunction
    AddWord "exec", enumFunction
    AddWord "exists", enumFunction
    AddWord "exit", enumFunction
    AddWord "exp", enumFunction
    AddWord "fcntl", enumFunction
    AddWord "fileno", enumFunction
    AddWord "flock", enumFunction
    AddWord "fork", enumFunction
    AddWord "format", enumFunction
    AddWord "formline", enumFunction
    AddWord "getc", enumFunction
    AddWord "getlogin", enumFunction
    AddWord "getpeername", enumFunction
    AddWord "getpgrp", enumFunction
    AddWord "getppid", enumFunction
    AddWord "getpriority", enumFunction
    AddWord "getpwnam", enumFunction
    AddWord "getgrnam", enumFunction
    AddWord "gethostbyname", enumFunction
    AddWord "getnetbyname", enumFunction
    AddWord "getprotobyname", enumFunction
    AddWord "getpwuid", enumFunction
    AddWord "getgrgid", enumFunction
    AddWord "getservbyname", enumFunction
    AddWord "gethostbyaddr", enumFunction
    AddWord "getnetbyaddr", enumFunction
    AddWord "getprotobynumber", enumFunction
    AddWord "getservbyport", enumFunction
    AddWord "getpwent", enumFunction
    AddWord "getgrent", enumFunction
    AddWord "gethostent", enumFunction
    AddWord "getnetent", enumFunction
    AddWord "getprotoent", enumFunction
    AddWord "getservent", enumFunction
    AddWord "setpwent", enumFunction
    AddWord "setgrent", enumFunction
    AddWord "sethostent", enumFunction
    AddWord "setnetent", enumFunction
    AddWord "setprotoent", enumFunction
    AddWord "setservent", enumFunction
    AddWord "endpwent", enumFunction
    AddWord "endgrent", enumFunction
    AddWord "endhostent", enumFunction
    AddWord "endnetent", enumFunction
    AddWord "endprotoent", enumFunction
    AddWord "endservent", enumFunction
    AddWord "getsockname", enumFunction
    AddWord "getsockopt", enumFunction
    AddWord "glob", enumFunction
    AddWord "gmtime", enumFunction
    AddWord "grep", enumFunction
    AddWord "hex", enumFunction
    AddWord "import", enumFunction
    AddWord "index", enumFunction
    AddWord "int", enumFunction
    AddWord "ioctl", enumFunction
    AddWord "join", enumFunction
    AddWord "keys", enumFunction
    AddWord "kill", enumFunction
    AddWord "lc", enumFunction
    AddWord "lcfirst", enumFunction
    AddWord "length", enumFunction
    AddWord "link", enumFunction
    AddWord "listen", enumFunction
    AddWord "localtime", enumFunction
    AddWord "log", enumFunction
    AddWord "lstat", enumFunction
    AddWord "mkdir", enumFunction
    AddWord "msgctl", enumFunction
    AddWord "msgget", enumFunction
    AddWord "msgsnd", enumFunction
    AddWord "msgrcv", enumFunction
    AddWord "no", enumFunction
    AddWord "oct", enumFunction
    AddWord "open", enumFunction
    AddWord "opendir", enumFunction
    AddWord "ord", enumFunction
    AddWord "pack", enumFunction
    AddWord "pipe", enumFunction
    AddWord "pop", enumFunction
    AddWord "pos", enumFunction
    AddWord "print", enumFunction
    AddWord "printf", enumFunction
    AddWord "prototype", enumFunction
    AddWord "push", enumFunction
    AddWord "quotemeta", enumFunction
    AddWord "rand", enumFunction
    AddWord "read", enumFunction
    AddWord "readdir", enumFunction
    AddWord "readlink", enumFunction
    AddWord "recv", enumFunction
    AddWord "ref", enumFunction
    AddWord "rename", enumFunction
    AddWord "reset", enumFunction
    AddWord "reverse", enumFunction
    AddWord "rewinddir", enumFunction
    AddWord "rindex", enumFunction
    AddWord "rmdir", enumFunction
    AddWord "scalar", enumFunction
    AddWord "seek", enumFunction
    AddWord "seekdir", enumFunction
    AddWord "select", enumFunction
    AddWord "semctl", enumFunction
    AddWord "semget", enumFunction
    AddWord "semop", enumFunction
    AddWord "send", enumFunction
    AddWord "setpgrp", enumFunction
    AddWord "setpriority", enumFunction
    AddWord "setsockopt", enumFunction
    AddWord "shift", enumFunction
    AddWord "shmctl", enumFunction
    AddWord "shmget", enumFunction
    AddWord "shmread", enumFunction
    AddWord "shmwrite", enumFunction
    AddWord "shutdown", enumFunction
    AddWord "sin", enumFunction
    AddWord "sleep", enumFunction
    AddWord "socket", enumFunction
    AddWord "socketpair", enumFunction
    AddWord "sort", enumFunction
    AddWord "splice", enumFunction
    AddWord "split", enumFunction
    AddWord "sprintf", enumFunction
    AddWord "sqrt", enumFunction
    AddWord "srand", enumFunction
    AddWord "stat", enumFunction
    AddWord "study", enumFunction
    AddWord "substr", enumFunction
    AddWord "symlink", enumFunction
    AddWord "syscall", enumFunction
    AddWord "sysopen", enumFunction
    AddWord "sysread", enumFunction
    AddWord "sysseek", enumFunction
    AddWord "system", enumFunction
    AddWord "syswrite", enumFunction
    AddWord "tell", enumFunction
    AddWord "telldir", enumFunction
    AddWord "tie", enumFunction
    AddWord "tied", enumFunction
    AddWord "time", enumFunction
    AddWord "times", enumFunction
    AddWord "truncate", enumFunction
    AddWord "uc", enumFunction
    AddWord "ucfirst", enumFunction
    AddWord "umask", enumFunction
    AddWord "undef", enumFunction
    AddWord "unlink", enumFunction
    AddWord "unpack", enumFunction
    AddWord "untie", enumFunction
    AddWord "unshift", enumFunction
    AddWord "utime", enumFunction
    AddWord "values", enumFunction
    AddWord "vec", enumFunction
    AddWord "wait", enumFunction
    AddWord "waitpid", enumFunction
    AddWord "wantarray", enumFunction
    AddWord "warn", enumFunction
    AddWord "write", enumFunction
    
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    AddWord "[", enumOperator
    AddWord "]", enumOperator
    AddWord "-", enumOperator
    AddWord "+", enumOperator
    AddWord "*", enumOperator
    AddWord "%", enumOperator
    AddWord "/", enumOperator
    AddWord "=", enumOperator
    AddWord "~", enumOperator
    AddWord "!", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "<", enumOperator
    AddWord ">", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord ".", enumOperator
    
    AddCommentTag "#", vbCrLf
    AddCommentTag """", """"
    AddCommentTag "'", "'"
    
    AddLiteralTag """", """"
    AddLiteralTag "'", "'"
    
End Sub

Private Sub SetVBScript()
    
    '------------------------------
    ' Case
    CompareCase = vbTextCompare
    GiveCorrectCase = True
    resetLanguage
    
    '------------------------------
    ' (Mixed with VB)
    AddWord "#Const"
    AddWord "#If"
    AddWord "#Else"
    AddWord "#End"
    
    AddWord "AppActivate"
    AddWord "As"
    AddWord "Beep"
    AddWord "Base"
    AddWord "Binary"
    AddWord "Boolean"
    AddWord "ByRef"
    AddWord "Byte"
    AddWord "ByVal"
    AddWord "Call"
    AddWord "Case"
    AddWord "ChDir"
    AddWord "ChDrive"
    AddWord "Const"
    AddWord "Control"
    AddWord "Declare"
    AddWord "Dim"
    AddWord "Do"
    AddWord "DoEvents"
    AddWord "Double"
    AddWord "Loop"
    AddWord "Each"
    AddWord "Else"
    AddWord "Elseif"
    AddWord "Empty"
    AddWord "End"
    AddWord "Enum"
    AddWord "Erase"
    AddWord "Err"
    AddWord "Error"
    AddWord "Event"
    AddWord "Exit"
    AddWord "Explicit"
    AddWord "False"
    AddWord "Friend"
    AddWord "For"
    AddWord "Function"
    AddWord "Get"
    AddWord "GoSub"
    AddWord "GoTo"
    AddWord "If"
    AddWord "Implements"
    AddWord "Input"
    AddWord "Integer"
    AddWord "Is"
    AddWord "Len"
    AddWord "Let"
    AddWord "Load"
    AddWord "Long"
    AddWord "Me"
    AddWord "New"
    AddWord "Next"
    AddWord "Nothing"
    AddWord "Null"
    AddWord "On"
    AddWord "Option"
    AddWord "Optional"
    AddWord "Print"
    AddWord "Private"
    AddWord "Property"
    AddWord "Public"
    AddWord "Randomize"
    AddWord "ReDim"
    AddWord "Rem"
    AddWord "Resume"
    AddWord "Select"
    AddWord "Set"
    AddWord "Single"
    AddWord "Static"
    AddWord "Step"
    AddWord "Stop"
    AddWord "String"
    AddWord "Sub"
    AddWord "Then"
    AddWord "To"
    AddWord "True"
    AddWord "Type"
    AddWord "Unload"
    AddWord "Variant"
    AddWord "Wend"
    AddWord "While"
    AddWord "With"
    AddWord "WithEvents"
   
    '------------------------------
    ' Constantes
    AddWord "vbCr"
    AddWord "vbCrLf"
    AddWord "vbFormFeed"
    AddWord "vbLf"
    AddWord "vbNewLine"
    AddWord "vbNullChar"
    AddWord "vbNullString"
    AddWord "vbTab"
    AddWord "vbVerticalTab"

    AddWord "vbBlack"
    AddWord "vbRed"
    AddWord "vbGreen"
    AddWord "vbYellow"
    AddWord "vbBlue"
    AddWord "vbMagenta"
    AddWord "vbCyan"
    AddWord "vbWhite"

    AddWord "vbBinaryCompare"
    AddWord "vbTextCompare "

    AddWord "vbSunday"
    AddWord "vbMonday"
    AddWord "vbTuesday"
    AddWord "vbWednesday"
    AddWord "vbThursday"
    AddWord "vbFriday"
    AddWord "vbSaturday"
    AddWord "vbFirstJan1"
    AddWord "vbFirstFourDays"
    AddWord "vbFirstFullWeek"
    AddWord "vbUseSystem"
    AddWord "vbUseSystemDayOfWeek"

    AddWord "vbGeneralDate"
    AddWord "vbLongDate"
    AddWord "vbShortDate"
    AddWord "vbLongTime"
    AddWord "vbShortTime "

    AddWord "vbOKOnly"
    AddWord "vbOKCancel"
    AddWord "vbAbortRetryIgnore"
    AddWord "vbYesNoCancel"
    AddWord "vbYesNo"
    AddWord "vbRetryCancel"
    AddWord "vbCritical"
    AddWord "vbQuestion"
    AddWord "vbExclamation"
    AddWord "vbInformation"
    AddWord "vbDefaultButton1"
    AddWord "vbDefaultButton2"
    AddWord "vbDefaultButton3"
    AddWord "vbDefaultButton4"
    AddWord "vbApplicationModal"
    AddWord "vbSystemModal"

    AddWord "vbOK"
    AddWord "vbCancel"
    AddWord "vbAbort"
    AddWord "vbRetry"
    AddWord "vbIgnore"
    AddWord "vbYes"
    AddWord "vbNo"

    '------------------------------
    ' Funciones
    AddWord "Abs", enumFunction
    AddWord "Array", enumFunction
    AddWord "Asc", enumFunction
    AddWord "AscB", enumFunction
    AddWord "AscW", enumFunction
    AddWord "Atn", enumFunction
    AddWord "CBool", enumFunction
    AddWord "CByte", enumFunction
    AddWord "CCur", enumFunction
    AddWord "CDate", enumFunction
    AddWord "CDbl", enumFunction
    AddWord "Chr", enumFunction
    AddWord "ChrB", enumFunction
    AddWord "ChrW", enumFunction
    AddWord "CInt", enumFunction
    AddWord "CLng", enumFunction
    AddWord "Cos", enumFunction
    AddWord "CreateObject", enumFunction
    AddWord "CSng", enumFunction
    AddWord "CStr", enumFunction
    AddWord "Date", enumFunction
    AddWord "DateAdd", enumFunction
    AddWord "DateDiff", enumFunction
    AddWord "DatePart", enumFunction
    AddWord "DateSerial", enumFunction
    AddWord "DateValue", enumFunction
    AddWord "Day", enumFunction
    AddWord "Exp", enumFunction
    AddWord "Filter", enumFunction
    AddWord "Fix", enumFunction
    AddWord "FormatCurrency", enumFunction
    AddWord "FormatDateTime", enumFunction
    AddWord "FormatNumber", enumFunction
    AddWord "FormatPercent", enumFunction
    AddWord "GetObject", enumFunction
    AddWord "Hex", enumFunction
    AddWord "Hour", enumFunction
    AddWord "InputBox", enumFunction
    AddWord "InStr", enumFunction
    AddWord "InStrB", enumFunction
    AddWord "InStrRev", enumFunction
    AddWord "Int", enumFunction
    AddWord "IsArray", enumFunction
    AddWord "IsDate", enumFunction
    AddWord "IsEmpty", enumFunction
    AddWord "IsNull", enumFunction
    AddWord "IsNumeric", enumFunction
    AddWord "IsObject", enumFunction
    AddWord "Join", enumFunction
    AddWord "LBound", enumFunction
    AddWord "LCase", enumFunction
    AddWord "Left", enumFunction
    AddWord "LeftB", enumFunction
    AddWord "Len", enumFunction
    AddWord "LenB", enumFunction
    AddWord "LoadPicture", enumFunction
    AddWord "Log", enumFunction
    AddWord "LTrim", enumFunction
    AddWord "Mid", enumFunction
    AddWord "MidB", enumFunction
    AddWord "Minute", enumFunction
    AddWord "Month", enumFunction
    AddWord "MonthName", enumFunction
    AddWord "MsgBox", enumFunction
    AddWord "Now", enumFunction
    AddWord "Oct", enumFunction
    AddWord "Replace", enumFunction
    AddWord "RGB", enumFunction
    AddWord "Right", enumFunction
    AddWord "RightB", enumFunction
    AddWord "Rnd", enumFunction
    AddWord "Round", enumFunction
    AddWord "RTrim", enumFunction
    AddWord "ScriptEngine", enumFunction
    AddWord "ScriptEngineBuildVersion", enumFunction
    AddWord "ScriptEngineMajorVersion", enumFunction
    AddWord "ScriptEngineMinorVersion", enumFunction
    AddWord "Second", enumFunction
    AddWord "Sgn", enumFunction
    AddWord "Sin", enumFunction
    AddWord "Space", enumFunction
    AddWord "Split", enumFunction
    AddWord "Sqr", enumFunction
    AddWord "StrComp", enumFunction
    AddWord "StrReverse", enumFunction
    AddWord "Tan", enumFunction
    AddWord "Time", enumFunction
    AddWord "TimeSerial", enumFunction
    AddWord "TimeValue", enumFunction
    AddWord "Trim", enumFunction
    AddWord "TypeName", enumFunction
    AddWord "UBound", enumFunction
    AddWord "UCase", enumFunction
    AddWord "VarType", enumFunction
    AddWord "Weekday", enumFunction
    AddWord "WeekdayName", enumFunction
    AddWord "Year", enumFunction

    AddWord "Mod", enumOperator
    AddWord "=", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "/", enumOperator
    AddWord "\", enumOperator
    AddWord "*", enumOperator
    AddWord "&", enumOperator
    AddWord "^", enumOperator
    AddWord "Not", enumOperator
    AddWord "Is", enumOperator
    AddWord "And", enumOperator
    AddWord "Or", enumOperator
    AddWord "xor", enumOperator
    AddWord "Eqv", enumOperator
    AddWord "Imp", enumOperator
    
    AddCommentTag "'", vbCrLf
    AddCommentTag "Rem", vbCrLf
    AddCommentTag """", """"
    
    AddLiteralTag """", """"
    
End Sub

Private Sub SetSQL()
    
    '------------------------------
    ' Case
    CompareCase = vbTextCompare
    GiveCorrectCase = True
    resetLanguage
    
    '------------------------------
    ' modificado para SQL-DAO
    AddWord "ADD"
    AddWord "ALL"
    AddWord "ALTER"
    AddWord "AND"
    AddWord "ANY"
    AddWord "AS"
    AddWord "ASC"
    AddWord "AUTOINCREMENT" 'agregado
    AddWord "AVG"
    AddWord "BETWEEN"
    AddWord "BY"
    AddWord "COLUMN"
    AddWord "CONSTRAINT"
    AddWord "COUNT"
    AddWord "CREATE"
    AddWord "DATABASE"
    AddWord "DELETE"
    AddWord "DESC"
    AddWord "DISALLOW" 'agregado
    AddWord "DISTINCT"
    AddWord "DISTINCTROW" 'agregado
    AddWord "DROP"
    AddWord "EXISTS"
    AddWord "FOREIGN"
    AddWord "FROM"
    AddWord "GROUP"
    AddWord "HAVING"
    AddWord "IGNORE" 'agregado
    AddWord "IN"
    AddWord "INDEX"
    AddWord "INNER"
    AddWord "INSERT"
    AddWord "INTO"
    AddWord "IS"
    AddWord "JOIN"
    AddWord "KEY"
    AddWord "LEFT"
    AddWord "LEVEL"
    AddWord "LIKE"
    AddWord "MAX"
    AddWord "MIN"
    AddWord "NOT"
    AddWord "NULL"
    AddWord "ON"
    AddWord "OPTION"
    AddWord "OR"
    AddWord "ORDER"
    AddWord "OUTER"
    AddWord "OWNERACCESS" 'agregado
    AddWord "PERCENT"
    AddWord "PRIMARY"
    AddWord "PROCEDURE"
    AddWord "REFERENCES"
    AddWord "RESTRICT"
    AddWord "RIGHT"
    AddWord "SELECT"
    AddWord "SET"
    AddWord "SOME"
    AddWord "SUM"
    AddWord "TABLE"
    AddWord "TOP"
    AddWord "UNION"
    AddWord "UNIQUE"
    AddWord "UPDATE"
    AddWord "VALUES"
    AddWord "WHERE"
    AddWord "WITH"
    AddWord "BINARY"
    AddWord "BIT"
    AddWord "BYTE" 'agregado
    AddWord "CHAR"
    AddWord "COUNTER" 'agregado
    AddWord "CURRENCY" 'agregado
    AddWord "DATETIME"
    AddWord "GUID" 'agregado
    AddWord "SINGLE" 'agregado
    AddWord "DOUBLE" 'agregado
    AddWord "SHORT" 'agregado
    AddWord "LONG" 'agregado
    AddWord "LONGTEXT" 'agregado
    AddWord "LONGBINARY" 'agregado
    AddWord "TEXT"
    AddWord "SMALLINT"
    AddWord "INT"
    AddWord "REAL"
    AddWord "FLOAT"
    AddWord "MONEY"
    AddWord "TINYINT"
    AddWord "NUMERIC"
    AddWord "SMALLDATETIME"
    AddWord "VARCHAR"
    AddWord "VARBINARY"
    AddWord "IMAGE"
    AddWord "DECIMAL"
    AddWord "NCHAR"
    AddWord "NTEXT"
    AddWord "NVARCHAR"
    
    AddWord "COUNT", enumFunction
    
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "*", enumOperator
    AddWord "/", enumOperator
    AddWord "%", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "-", enumOperator
    AddWord "=", enumOperator
    AddWord ":", enumOperator
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "[", enumOperator
    AddWord "]", enumOperator
   
    AddWord "TRUE", enumOperator
    AddWord "FALSE", enumOperator
    
    AddCommentTag "--", vbCrLf
    AddCommentTag "'", "'"
    
    AddLiteralTag "'", "'"

End Sub

    
Private Sub SetRuby()
    
    CompareCase = vbTextCompare
    GiveCorrectCase = False
    resetLanguage
    
    AddWord "alias"
    AddWord "and"
    AddWord "begin"
    AddWord "BEGIN"
    AddWord "break"
    AddWord "case"
    AddWord "class"
    AddWord "def"
    AddWord "defined"
    AddWord "do"
    AddWord "else"
    AddWord "elsif"
    AddWord "end"
    AddWord "END"
    AddWord "ensure"
    AddWord "false"
    AddWord "for"
    AddWord "if"
    AddWord "in"
    AddWord "module"
    AddWord "next"
    AddWord "nil"
    AddWord "not"
    AddWord "or"
    AddWord "redo"
    AddWord "rescue"
    AddWord "retry"
    AddWord "return"
    AddWord "self"
    AddWord "super"
    AddWord "then"
    AddWord "true"
    AddWord "undef"
    AddWord "unless"
    AddWord "until"
    AddWord "when"
    AddWord "while"
    AddWord "yield"

    AddWord "ARGF"
    AddWord "ARGV"
    AddWord "DATA"
    AddWord "ENV"
    AddWord "FALSE"
    AddWord "NIL"
    AddWord "RUBY_PLATFORM"
    AddWord "RUBY_RELEASE_DATE"
    AddWord "RUBY_VERSION"
    AddWord "STDERR"
    AddWord "STDIN"
    AddWord "STDOUT"
    AddWord "TRUE"
    
    AddWord "Array", enumFunction
    AddWord "at_exit", enumFunction
    AddWord "autoload", enumFunction
    AddWord "binding", enumFunction
    AddWord "caller", enumFunction
    AddWord "catch", enumFunction
    AddWord "chomp", enumFunction
    AddWord "chomp!", enumFunction
    AddWord "chop", enumFunction
    AddWord "chop!", enumFunction
    AddWord "eval", enumFunction
    AddWord "exec", enumFunction
    AddWord "exit", enumFunction
    AddWord "exit!", enumFunction
    AddWord "fail", enumFunction
    AddWord "Float", enumFunction
    AddWord "fork", enumFunction
    AddWord "format", enumFunction
    AddWord "gets", enumFunction
    AddWord "global_variables", enumFunction
    AddWord "gsub", enumFunction
    AddWord "gsub!", enumFunction
    AddWord "Integer", enumFunction
    AddWord "iterator?", enumFunction
    AddWord "lambda", enumFunction
    AddWord "length", enumFunction
    AddWord "load", enumFunction
    AddWord "local_variables", enumFunction
    AddWord "loop", enumFunction
    AddWord "new", enumFunction
    AddWord "open", enumFunction
    AddWord "p", enumFunction
    AddWord "print", enumFunction
    AddWord "printf", enumFunction
    AddWord "proc", enumFunction
    AddWord "putc", enumFunction
    AddWord "puts", enumFunction
    AddWord "raise", enumFunction
    AddWord "rand", enumFunction
    AddWord "readline", enumFunction
    AddWord "readlines", enumFunction
    AddWord "require", enumFunction
    AddWord "select", enumFunction
    AddWord "sleep", enumFunction
    AddWord "split", enumFunction
    AddWord "sprintf", enumFunction
    AddWord "srand", enumFunction
    AddWord "String", enumFunction
    AddWord "sub", enumFunction
    AddWord "sub!", enumFunction
    AddWord "syscall", enumFunction
    AddWord "system", enumFunction
    AddWord "test", enumFunction
    AddWord "trace_var", enumFunction
    AddWord "trap", enumFunction
    AddWord "untrace_var", enumFunction

    AddWord "ArgumentError", enumFunction
    AddWord "Array", enumFunction
    AddWord "Bignum", enumFunction
    AddWord "Class", enumFunction
    AddWord "Data", enumFunction
    AddWord "Dir", enumFunction
    AddWord "EOFError", enumFunction
    AddWord "Exception", enumFunction
    AddWord "fatal", enumFunction
    AddWord "File", enumFunction
    AddWord "Fixnum", enumFunction
    AddWord "Float", enumFunction
    AddWord "FloatDomainError", enumFunction
    AddWord "Hash", enumFunction
    AddWord "IndexError", enumFunction
    AddWord "Integer", enumFunction
    AddWord "Interrupt", enumFunction
    AddWord "IO", enumFunction
    AddWord "IOError", enumFunction
    AddWord "LoadError", enumFunction
    AddWord "LocalJumpError", enumFunction
    AddWord "MatchingData", enumFunction
    AddWord "Module", enumFunction
    AddWord "NameError", enumFunction
    AddWord "NilClass", enumFunction
    AddWord "NotImplementError", enumFunction
    AddWord "Numeric", enumFunction
    AddWord "Object", enumFunction
    AddWord "Proc", enumFunction
    AddWord "Range", enumFunction
    AddWord "Regexp", enumFunction
    AddWord "RuntimeError", enumFunction
    AddWord "SecurityError", enumFunction
    AddWord "SignalException", enumFunction
    AddWord "StandardError", enumFunction
    AddWord "String", enumFunction
    AddWord "Struct", enumFunction
    AddWord "SyntaxError", enumFunction
    AddWord "SystemCallError", enumFunction
    AddWord "SystemExit", enumFunction
    AddWord "SystemStackError", enumFunction
    AddWord "ThreadError", enumFunction
    AddWord "Time", enumFunction
    AddWord "TypeError", enumFunction
    AddWord "ZeroDivisionError", enumFunction

    AddWord "Comparable", enumFunction
    AddWord "Enumerable", enumFunction
    AddWord "Errno", enumFunction
    AddWord "FileTest", enumFunction
    AddWord "GC", enumFunction
    AddWord "Kernel", enumFunction
    AddWord "Marshal", enumFunction
    AddWord "Math", enumFunction
    AddWord "ObjectSpace", enumFunction
    AddWord "Precision", enumFunction
    AddWord "Process", enumFunction
    
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    AddWord "[", enumOperator
    AddWord "]", enumOperator
    AddWord "-", enumOperator
    AddWord "+", enumOperator
    AddWord "*", enumOperator
    AddWord "%", enumOperator
    AddWord "/", enumOperator
    AddWord "=", enumOperator
    AddWord "~", enumOperator
    AddWord "!", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "<", enumOperator
    AddWord ">", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord ".", enumOperator
    
    AddCommentTag "#", vbCrLf
    AddCommentTag "'", "'"
    AddCommentTag """", """"
    
    AddLiteralTag "'", "'"
    AddLiteralTag """", """"

End Sub

    
Private Sub SetPython()
    
    CompareCase = vbTextCompare
    GiveCorrectCase = False
    resetLanguage
    
    AddWord "and"
    AddWord "assert"
    AddWord "break"
    AddWord "class"
    AddWord "continue"
    AddWord "def"
    AddWord "del"
    AddWord "elif"
    AddWord "else"
    AddWord "except"
    AddWord "exec"
    AddWord "finally"
    AddWord "for"
    AddWord "from"
    AddWord "global"
    AddWord "if"
    AddWord "import"
    AddWord "in"
    AddWord "is"
    AddWord "lambda"
    AddWord "len"
    AddWord "not"
    AddWord "or"
    AddWord "pass"
    AddWord "print"
    AddWord "raise"
    AddWord "return"
    AddWord "self"
    AddWord "try"
    AddWord "while"
    AddWord "None"
    AddWord "True"
    AddWord "False"

    AddWord "__new__"
    AddWord "__init__"
    AddWord "__del__"
    AddWord "__repr__"
    AddWord "__str__"
    AddWord "__lt__"
    AddWord "__ge__"
    AddWord "__eq__"
    AddWord "__ne__"
    AddWord "__gt__"
    AddWord "__ge__"
    AddWord "__cmp__"
    AddWord "__hash__"
    AddWord "__nonzero__"
    AddWord "__unicode__"
    AddWord "__getattr__"
    AddWord "__setattr__"
    AddWord "__delattr__"
    AddWord "__getattribute__"
    AddWord "__get__"
    AddWord "__set__"
    AddWord "__delete__"
    AddWord "__slots__"
    AddWord "__weakref__"
    AddWord "__dict__"
    AddWord "__metaclass__"
    AddWord "__call__"
    AddWord "__len__"
    AddWord "__getitem__"
    AddWord "__setitem__"
    AddWord "__delitem__"
    AddWord "__iter__"
    AddWord "__contains__"
    AddWord "__getslice__"
    AddWord "__setslice__"
    AddWord "__delslice__"
    AddWord "__coerce__"
    AddWord "__class__"
    AddWord "__bases__"
    AddWord "__name__"
    
    AddWord "__import__", enumFunction
    AddWord "abs", enumFunction
    AddWord "basestring", enumFunction
    AddWord "bool", enumFunction
    AddWord "callable", enumFunction
    AddWord "chr", enumFunction
    AddWord "classmethod", enumFunction
    AddWord "cmp", enumFunction
    AddWord "compile", enumFunction
    AddWord "complex", enumFunction
    AddWord "delattr", enumFunction
    AddWord "dict", enumFunction
    AddWord "dir", enumFunction
    AddWord "divmod", enumFunction
    AddWord "enumerate", enumFunction
    AddWord "eval", enumFunction
    AddWord "execfile", enumFunction
    AddWord "file", enumFunction
    AddWord "filter", enumFunction
    AddWord "float", enumFunction
    AddWord "frozenset", enumFunction
    AddWord "getattr", enumFunction
    AddWord "globals", enumFunction
    AddWord "hasattr", enumFunction
    AddWord "hash", enumFunction
    AddWord "help", enumFunction
    AddWord "hex", enumFunction
    AddWord "input", enumFunction
    AddWord "int", enumFunction
    AddWord "isinstance", enumFunction
    AddWord "issubclass", enumFunction
    AddWord "iter", enumFunction
    AddWord "len", enumFunction
    AddWord "list", enumFunction
    AddWord "locals", enumFunction
    AddWord "long", enumFunction
    AddWord "map", enumFunction
    AddWord "max", enumFunction
    AddWord "min", enumFunction
    AddWord "object", enumFunction
    AddWord "oct", enumFunction
    AddWord "open", enumFunction
    AddWord "ord", enumFunction
    AddWord "pow", enumFunction
    AddWord "property", enumFunction
    AddWord "range", enumFunction
    AddWord "raw_input", enumFunction
    AddWord "reduce", enumFunction
    AddWord "reload", enumFunction
    AddWord "repr", enumFunction
    AddWord "reversed", enumFunction
    AddWord "round", enumFunction
    AddWord "set", enumFunction
    AddWord "setattr", enumFunction
    AddWord "slice", enumFunction
    AddWord "sorted", enumFunction
    AddWord "staticmethod", enumFunction
    AddWord "str", enumFunction
    AddWord "sum", enumFunction
    AddWord "super", enumFunction
    AddWord "tuple", enumFunction
    AddWord "type", enumFunction
    AddWord "unichr", enumFunction
    AddWord "unicode", enumFunction
    AddWord "vars", enumFunction
    AddWord "xrange", enumFunction
    AddWord "zip", enumFunction
    
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    AddWord "[", enumOperator
    AddWord "]", enumOperator
    AddWord "-", enumOperator
    AddWord "+", enumOperator
    AddWord "*", enumOperator
    AddWord "%", enumOperator
    AddWord "/", enumOperator
    AddWord "=", enumOperator
    AddWord "~", enumOperator
    AddWord "!", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "<", enumOperator
    AddWord ">", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord ".", enumOperator
    
    AddCommentTag "#", vbCrLf
    AddCommentTag "'", "'"
    AddCommentTag """", """"
    
    AddLiteralTag "'", "'"
    AddLiteralTag """", """"

End Sub

    

