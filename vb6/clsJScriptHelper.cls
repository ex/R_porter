VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJScriptHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MB_ABORTRETRYIGNORE As Long
Public MB_OK As Long
Public MB_OKCANCEL As Long
Public MB_RETRYCANCEL As Long
Public MB_YESNO As Long
Public MB_YESNOCANCEL As Long

Public MB_ICONEXCLAMATION As Long
Public MB_ICONINFORMATION As Long
Public MB_ICONQUESTION As Long
Public MB_ICONERROR As Long

Public IDOK As Long
Public IDCANCEL As Long
Public IDABORT As Long
Public IDRETRY As Long
Public IDIGNORE As Long
Public IDYES As Long
Public IDNO As Long

Private m_Log As frmJScriptLog

Private Sub Class_Initialize()

    MB_OK = 0
    MB_OKCANCEL = 1
    MB_ABORTRETRYIGNORE = 2
    MB_YESNOCANCEL = 3
    MB_YESNO = 4
    MB_RETRYCANCEL = 5

    MB_ICONERROR = 16
    MB_ICONQUESTION = 32
    MB_ICONEXCLAMATION = 48
    MB_ICONINFORMATION = 64
   
    IDOK = 1
    IDCANCEL = 2
    IDABORT = 3
    IDRETRY = 4
    IDIGNORE = 5
    IDYES = 6
    IDNO = 7

    Set m_Log = New frmJScriptLog
    Load m_Log
    m_Log.warning ("*********** Salida de Consola **************")
    
End Sub

'---------------------------------------------------------------------------------
' [WARNING] DEFINIR Y DECLARAR MANUALMENTE TODAS LAS PROPIEDADES Y METODOS AQUI
Public Sub dump()
    m_Log.Visible = True
    m_Log.warning ("MB_OK = 0")
    m_Log.warning ("MB_OKCANCEL = 1")
    m_Log.warning ("MB_ABORTRETRYIGNORE = 2")
    m_Log.warning ("MB_YESNOCANCEL = 3")
    m_Log.warning ("MB_YESNO = 4")
    m_Log.warning ("MB_RETRYCANCEL = 5")
    
    m_Log.warning ("MB_ICONERROR = 16")
    m_Log.warning ("MB_ICONQUESTION = 32")
    m_Log.warning ("MB_ICONEXCLAMATION = 48")
    m_Log.warning ("MB_ICONINFORMATION = 64")
    
    m_Log.warning ("IDOK = 1")
    m_Log.warning ("IDCANCEL = 2")
    m_Log.warning ("IDABORT = 3")
    m_Log.warning ("IDRETRY = 4")
    m_Log.warning ("IDIGNORE = 5")
    m_Log.warning ("IDYES = 6")
    m_Log.warning ("IDNO = 7")
    
    m_Log.warning ("Sub dump()")
    m_Log.warning ("Sub alert(ByVal text As String)")
    m_Log.warning ("Function confirm(ByVal text As String)")
    m_Log.warning ("Function prompt(ByVal text As String)")
    m_Log.warning ("Function InputConsole(ByVal text As String, Optional ByVal title As String = ""JScript"", Optional ByVal default As String = """")")
    m_Log.warning ("Public Function MessageBox(ByVal text As String, Optional ByVal buttons As VbMsgBoxStyle = vbExclamation, Optional ByVal title As String = ""JScript"")")
    m_Log.warning ("Public Sub echo(ByVal text As String)")
    m_Log.warning ("Public Sub error(ByVal text As String)")
    m_Log.warning ("Public Sub warning(ByVal text As String)")
    m_Log.warning ("Public Sub clrscr()")

End Sub

Public Sub alert(ByVal text As String)
    MessageBox text
End Sub

Public Function confirm(ByVal text As String)
    If vbOK = MsgBox(text, vbQuestion + vbOKCancel, "JScript") Then
        confirm = 1
    Else
        confirm = 0
    End If
End Function

Public Function prompt(ByVal text As String)
    prompt = InputConsole(text)
End Function

Public Function InputConsole(ByVal text As String, _
                 Optional ByVal title As String = "JScript", _
                 Optional ByVal default As String = "")
    InputConsole = InputBox(text, title, default)
End Function

Public Function MessageBox(ByVal text As String, _
                 Optional ByVal buttons As VbMsgBoxStyle = vbExclamation, _
                 Optional ByVal title As String = "JScript")
    MessageBox = MsgBox(text, buttons, title)
End Function

Public Sub echo(ByVal text As String)
    m_Log.Visible = True
    m_Log.echo (text)
End Sub

Public Sub error(ByVal text As String)
    m_Log.Visible = True
    m_Log.error (text)
End Sub

Public Sub warning(ByVal text As String)
    m_Log.Visible = True
    m_Log.warning (text)
End Sub

Private Sub Class_Terminate()
    Unload m_Log
End Sub

Public Sub clrscr()
    m_Log.Visible = True
    m_Log.rchtxtLog.text = ""
End Sub

