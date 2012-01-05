VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DAC1C15E-A0D0-11D8-92BC-F3955AEE4860}#5.0#0"; "exHighLightCode.ocx"
Begin VB.Form frmScript 
   Caption         =   "Scripts"
   ClientHeight    =   3690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6390
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin exHighLightCode.exCodeHighlight exCodeHighlight 
      Height          =   3360
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5927
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RightMargin     =   40000
      SelRTF          =   $"frmScript.frx":030A
      Language        =   1
      OperatorColor   =   255
      CommentColor    =   32768
      LiteralColor    =   288688
      ForeColor       =   0
      FunctionColor   =   8388736
      Author          =   "Esau R.O. [exe_q_tor] ...based in the DevDomainCodeHighlight control."
      LeftMargin      =   180
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3390
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2699
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   5745
      Top             =   3105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuActivarScript 
         Caption         =   "Activar &script"
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mnuA1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewScript 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenScript 
         Caption         =   "&Abrir script"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Guardar como..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuA2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuVBScript 
         Caption         =   "VBScript"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuJScript 
         Caption         =   "JScript"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlantillaMain 
         Caption         =   "Plantilla &Main (VBScript)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPlantillaR_porter 
         Caption         =   "Plantilla &R_porter"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuB0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlantillaJScript 
         Caption         =   "Plantilla &Main (JScript)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEjecutar 
         Caption         =   "&Ejecutar"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const m_def_ID_EX = 25647893

Private ScriptControl As clsScript
Private JScriptControl As clsJScript
Private JScriptHelper As clsJScriptHelper
Private m_bolSaveSet As Boolean
Private m_strSaveFile As String
Private m_VBScript As Boolean

Private Sub exCodeHighlight_SelChange()
    status.Panels(1).text = "ln " & exCodeHighlight.CursorLine
    status.Panels(2).text = "chr " & exCodeHighlight.CursorRow
End Sub

'**************************************************************
' FORM
'**************************************************************
Private Sub Form_Load()
    
    On Error GoTo Handler
    
    '---------------------------------------------------
    ' crear objetos script
    Set ScriptControl = New clsScript
    
    Set JScriptControl = New clsJScript
    Set JScriptHelper = New clsJScriptHelper
    JScriptControl.objScript.AddObject "Console", JScriptHelper, True
    '---------------------------------------------------
    
    m_bolSaveSet = False
    
    exCodeHighlight.ExID = m_def_ID_EX
    m_VBScript = False
    exCodeHighlight.Language = exJscript
    mnuVBScript.Checked = False
    mnuJScript.Checked = True
    
    ' necesario para poner los tabs a 4 espacios
    gsub_SetRichTabs exCodeHighlight.RichHwnd, 4
    Exit Sub
    
Handler:    MsgBox Err.Description, vbError, "Form_Load"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If (Me.width < 2655) Then
        Me.width = 2655
    End If
    If (Me.height < 2100) Then
        Me.height = 2100
    End If
    exCodeHighlight.height = Me.ScaleHeight - status.height
    exCodeHighlight.width = Me.ScaleWidth + 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ScriptControl = Nothing
    Set JScriptControl = Nothing
    Set JScriptHelper = Nothing
End Sub

'**************************************************************
' MENUS
'**************************************************************
Private Sub mnuEjecutar_Click()
    If m_VBScript Then
        ejecutarVBScript
    Else
        ejecutarJScript
    End If
End Sub

Private Sub ejecutarJScript()
    
    On Error Resume Next
    Dim modulo As MSScriptControl.Module
   
    Set modulo = JScriptControl.objScript.Modules("Global")
    
    '--------------------------------------
    ' Inicializamos script
    JScriptControl.objScript.Reset
   
    '--------------------------------------
    ' Agregamos objeto de consola en JScript
    JScriptControl.objScript.AddObject "Console", JScriptHelper, True
   
    '--------------------------------------
    ' Agregamos el codigo script
    modulo.AddCode exCodeHighlight.text
    
    If Err.Number <> 0 Then
        If Err.Source <> "R_porter" Then
            '--------------------------------
            ' Procesamos errores de sintaxis
            With JScriptControl.objScript.error
                MsgBox "E" & .Number & ": " & .Description & vbCrLf & "En linea: " & .line & vbCrLf & _
                       "Columna: " & .Column & vbCrLf & .text, vbExclamation, "Error de sintaxis"
            End With
        Else
            GoTo Handler
        End If
    Else
        '--------------------------------
        ' Ejecutamos el codigo de Main()
        modulo.Run "main"
        
        If Err.Number <> 0 Then
            If Err.Source <> "R_porter" Then
                '--------------------------------
                ' Procesamos errores de ejecucion
                With JScriptControl.objScript.error
                    MsgBox "E" & .Number & ": " & .Description & vbCrLf & "En linea: " & .line, vbExclamation, "Runtime error"
                End With
            Else
                GoTo Handler
            End If
        Else
            '--------------------------------
            ' El codigo fue ejecutado
            On Error GoTo Handler
            Exit Sub
        End If
    
    End If
    Exit Sub

Handler:    MsgBox Err.Description, vbExclamation, "Error en el script"
End Sub

Private Sub ejecutarVBScript()
    
    Dim modulo As MSScriptControl.Module
    
    On Error Resume Next
   
    Set modulo = ScriptControl.objScript.Modules("Global")
    
    '--------------------------------------
    ' Inicializamos script
    ScriptControl.objScript.Reset
   
    '--------------------------------------
    ' Agregamos el codigo script
    modulo.AddCode exCodeHighlight.text
    
    If Err.Number <> 0 Then
        If Err.Source <> "R_porter" Then
            '--------------------------------
            ' Procesamos errores de sintaxis
            With ScriptControl.objScript.error
                MsgBox "E" & .Number & ": " & .Description & vbCrLf & "En linea: " & .line & vbCrLf & _
                       "Columna: " & .Column & vbCrLf & .text, vbExclamation, "Error de sintaxis"
            End With
        Else
            GoTo Handler
        End If
    Else
        '--------------------------------
        ' Ejecutamos el codigo de Main()
        modulo.Run "main"
        
        If Err.Number <> 0 Then
            If Err.Source <> "R_porter" Then
                '--------------------------------
                ' Procesamos errores de ejecucion
                With ScriptControl.objScript.error
                    MsgBox "E" & .Number & ": " & .Description & vbCrLf & "En linea: " & .line, vbExclamation, "Runtime error"
                End With
            Else
                GoTo Handler
            End If
        Else
            '--------------------------------
            ' El codigo fue ejecutado
            On Error GoTo Handler
            Exit Sub
        End If
    
    End If
    Exit Sub

Handler:    MsgBox Err.Description, vbExclamation, "Error en el script"
End Sub

Private Sub mnuOpenScript_Click()
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'esconde casilla de solo lectura y verifica que el archivo y el path existan
        .flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .DialogTitle = "Indicar el script a abrir:"
        .Filter = "Scripts JS (*.js)|*.js|Scripts VBS (*.vbs)|*.vbs|Todos los Archivos(*.*)|*.*"
        .InitDir = App.Path & "\scripts"    ' de no existir el directorio usara el directorio activo
        'tipo predefinido JS
        .FilterIndex = 1
        .ShowOpen
        If .filename <> "" Then
            
            If ".js" = Right(.filename, 3) Then
                mnuJScript_Click
            Else
                mnuVBScript_Click
            End If
            
            exCodeHighlight.HighlightCode = exOnNewLine
            'cargar el script
            exCodeHighlight.LoadFile .filename
            exCodeHighlight.HighlightCode = exAsType
            m_bolSaveSet = True
            m_strSaveFile = .filename
            Me.Caption = .filename
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "mnuOpenScript_Click"
    End If
End Sub

Private Sub mnuSave_Click()
    exCodeHighlight.SaveFile m_strSaveFile, rtfText
End Sub

Private Sub mnuSaveAs_Click()
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
        ' Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        ' avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Exportar los resultados como:"
        .Filter = "Archivo JS (*.js)|*.js|Archivo VBS (*.vbs)|*.vbs|Archivo RTF (*.rtf)|*.rtf|Todos los Archivos(*.*)|*.*"
        ' necesario para controlar la extension con que se salvaran los archivos
        ' sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .DefaultExt = ""
        .InitDir = App.Path & "\scripts"    ' de no existir el directorio usara el directorio activo
        ' tipo predefinido JS
        .FilterIndex = 1
        ' nombre del reporte inicial
        .filename = "Nuevo"
        .ShowSave
        If .filename <> "" Then
            ' .FilterIndex devuelve la extension seleccionada en el cuadro guardar como
            Select Case .FilterIndex
                Case 1
                    'forzamos que el archivo sea de tipo JS y se guarde como texto
                    If UCase(Right(.filename, 4)) <> ".JS" Then
                        .filename = .filename & ".js"
                    End If
                    '---------------------------------------------
                    ' salvar script
                    exCodeHighlight.SaveFile .filename, rtfText
                Case 2
                    'forzamos que el archivo sea de tipo VBS y se guarde como texto
                    If UCase(Right(.filename, 4)) <> ".VBS" Then
                        .filename = .filename & ".vbs"
                    End If
                    '---------------------------------------------
                    ' salvar script
                    exCodeHighlight.SaveFile .filename, rtfText
                Case 3
                    'forzamos que el archivo sea RTF
                    If UCase(Right(.filename, 4)) <> ".RTF" Then
                        .filename = .filename & ".rtf"
                    End If
                    '---------------------------------------------
                    ' salvar script
                    exCodeHighlight.SaveFile .filename, rtfRTF
                Case Else
                    '---------------------------------------------
                    ' salvar script
                    exCodeHighlight.SaveFile .filename, rtfText
            End Select
            ' actualizar nombre mostrado en la barra de titulo
            Me.Caption = .filename
            ' actualizar archivo a guardar
            m_strSaveFile = .filename
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "mnuSaveAs_Click()"
    End If
End Sub

Private Sub mnuPlantillaMain_Click()
    With exCodeHighlight
        .SelText = "'-----------------------------------------------------"
        .SelText = vbCrLf & "Option Explicit"
        .SelText = vbCrLf
        .SelText = vbCrLf & "Public Sub Main()"
        .SelText = vbCrLf & vbTab & "MsgBox ""Ejemplo"", vbExclamation"
        .SelText = vbCrLf & "End Sub"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuPlantillaJScript_Click()
    With exCodeHighlight
        .SelText = "//-----------------------------------------------------"
        .SelText = vbCrLf & "function main()"
        .SelText = vbCrLf & "{"
        .SelText = vbCrLf & vbTab & "Console.alert (""Ejemplo"");"
        .SelText = vbCrLf & "}"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuPlantillaR_porter_Click()
    With exCodeHighlight
        .SelText = "Option Explicit"
        .SelText = vbCrLf
        .SelText = vbCrLf & "'--------------------------------------------------------"
        .SelText = vbCrLf & "' Esta funcion es llamada antes de iniciar la busqueda"
        .SelText = vbCrLf & "Public Sub gsub_exPreSearch()"
        .SelText = vbCrLf & vbTab & "' TODO"
        .SelText = vbCrLf & "End Sub"
        .SelText = vbCrLf
        .SelText = vbCrLf & "'--------------------------------------------------------"
        .SelText = vbCrLf & "' Esta funcion es llamada durante la busqueda"
        .SelText = vbCrLf & "Public Sub gsub_exInSearch()"
        .SelText = vbCrLf & vbTab & "' TODO"
        .SelText = vbCrLf & "End Sub"
        .SelText = vbCrLf
        .SelText = vbCrLf & "'--------------------------------------------------------"
        .SelText = vbCrLf & "' Esta funcion es llamada al finalizar la busqueda"
        .SelText = vbCrLf & "Public Sub gsub_exEndSearch()"
        .SelText = vbCrLf & vbTab & "' TODO"
        .SelText = vbCrLf & "End Sub"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuVBScript_Click()
    m_VBScript = True
    mnuVBScript.Checked = True
    mnuJScript.Checked = False
    Me.exCodeHighlight.Language = exVBScript
End Sub

Private Sub mnuJScript_Click()
    m_VBScript = False
    mnuVBScript.Checked = False
    mnuJScript.Checked = True
    Me.exCodeHighlight.Language = exJscript
End Sub

