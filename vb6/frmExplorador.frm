VERSION 5.00
Object = "{B85EE4CE-0C3F-423B-A0E8-96C755EEFE24}#1.0#0"; "exAtlTag.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{40EF20B1-7EC5-11D8-95A1-9655FE58C763}#2.0#0"; "exLightButton.ocx"
Begin VB.Form frmExplorar 
   Caption         =   "R_porter 1.2"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7125
   Icon            =   "frmExplorador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3360
      Top             =   8000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rchtxt 
      Height          =   5700
      Left            =   1935
      TabIndex        =   2
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   10054
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   60000
      TextRTF         =   $"frmExplorador.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraTareas 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5700
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox pctDetener 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   75
         MouseIcon       =   "frmExplorador.frx":0381
         MousePointer    =   99  'Custom
         Picture         =   "frmExplorador.frx":04D3
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   114
         TabIndex        =   5
         Top             =   1830
         Width           =   1710
      End
      Begin VB.PictureBox pctFondoBorde 
         BackColor       =   &H00FA6124&
         BorderStyle     =   0  'None
         FillColor       =   &H00FA6124&
         Height          =   5685
         Left            =   0
         Picture         =   "frmExplorador.frx":397D
         ScaleHeight     =   5685
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   0
         Width           =   1875
         Begin exLightButton.ocxLightButton ELBImprimir 
            Height          =   585
            Left            =   75
            TabIndex        =   6
            Top             =   2580
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmExplorador.frx":8226
            PictureOFF      =   "frmExplorador.frx":B6E0
            PictureOK       =   "frmExplorador.frx":EB9A
            MouseIcon       =   "frmExplorador.frx":12054
            MousePointer    =   99
         End
         Begin exLightButton.ocxLightButton ELBSalvar 
            Height          =   585
            Left            =   75
            TabIndex        =   7
            Top             =   3330
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmExplorador.frx":121B6
            PictureOFF      =   "frmExplorador.frx":15670
            PictureOK       =   "frmExplorador.frx":18B2A
            MouseIcon       =   "frmExplorador.frx":1BFE4
            MousePointer    =   99
         End
         Begin exLightButton.ocxLightButton ELBSalir 
            Height          =   585
            Left            =   75
            TabIndex        =   8
            Top             =   4080
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmExplorador.frx":1C146
            PictureOFF      =   "frmExplorador.frx":1F600
            PictureOK       =   "frmExplorador.frx":22ABA
            MouseIcon       =   "frmExplorador.frx":25F74
            MousePointer    =   99
         End
         Begin ATLTAGLibCtl.exAtlTag exTag 
            Left            =   195
            OleObjectBlob   =   "frmExplorador.frx":260D6
            Top             =   5040
         End
         Begin VB.Label lblAviso 
            BackStyle       =   0  'Transparent
            Caption         =   "Presione ESCAPE para detener........"
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   44751310
            MinWidth        =   25400002
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rchtxtScript 
      Height          =   1050
      Left            =   2010
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1852
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmExplorador.frx":260FA
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar Reporte"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuimprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmExplorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' Formulario de reporte de R_Porter
'               Este programa realiza un reporte de los archivos
'               que encuentra en una unidad o directorio especificado
'               permitiendo imprimirlo y buscar por tipos de archivo
'               ademas de dar algunas opciones de reporte
'               Ademas usa controles ActiveX: ELB, Split y Flash
'*******************************************************************************
' Revisado:     Esau (Agosto 2003) (mejoras para la version 1.1)
'               Anexion a la base de datos
'               Colores configurables
'               Proyecto original no cargaba en XP (se colgaba el vb6, why?)
'               Al parecer es un problema con los controles activeX incrustados
'               este proyecto usa los controles compilados
'               ----------------------------------------------------------------
'               Esau (Marzo 2004) Corregido errores de ScanDir y SetAtrib
'               ----------------------------------------------------------------
'               Esau (Abril 2004) Agregada la tabla de categoria de medio
'*******************************************************************************
' Creado:       Esau (Marzo 2000)
'               Puede detenerse con ESC y/o click en pctDetener
'               Imprime el reporte.
'*******************************************************************************

Option Explicit

Private numDir As Long
Private numFile As Long
Private nTab As Long
Private cad As String
Private isSubDir As Boolean
Private isReadOnly As Boolean
Private isHiden As Boolean
Private isSystem As Boolean
Private Header As String
Private HeaderDat As String
Private strLinea As String
Private numDw As Long
Private InProcess As Boolean
Private lpMsg As MSG
Private crash As Boolean
Private numDirT As Long
Private numFileT As Long
Private m_lngFiles As Long

Private ColorDirNormal As Variant
Private ColorDirHidden As Variant
Private ColorDirReadOnly As Variant
Private ColorDirOther As Variant
Private ColorFileNormal As Variant
Private ColorFileHidden As Variant
Private ColorFileReadOnly As Variant
Private ColorFileOther As Variant
Private ColorFile(1 To 8) As Variant
Private FileNum(1 To 8) As Long

Private rs As ADODB.Recordset

Private ms_DirRaiz As String

Private ml_idStorage As Long

Private ms_TypeOld As String
Private ml_idTypeOld As Long

Private ms_GenreOld As String
Private ml_idGenreOld As Long

Private ms_AuthorOld As String
Private ml_idAuthorOld As Long

Private mb_NoAdvertirRecorteTipo As Boolean
Private mb_NoAdvertirRecorteGenero As Boolean
Private mb_NoAdvertirRecorteAutor As Boolean
Private mb_NoAdvertirRecorteNombre As Boolean
Private mb_NoAdvertirRecorteSysName As Boolean

Private ml_idSysParent As Long

Private mo_IScript As clsIScript
Private ScriptControl As clsScript
Private mo_Modulo As MSScriptControl.Module

Private m_height As Integer
Private m_width As Integer
Private m_rich_height As Integer
Private m_rich_width As Integer
Private m_pic_height As Integer
Private m_fra_height As Integer

'*******************************************************************************
' INICIALIZACION FORMULARIO
'*******************************************************************************
Private Sub Form_Load()
    
    stbar.Style = sbrSimple
    numDirT = 0
    numFileT = 0
    
    m_height = Me.height
    m_width = Me.width
    
    m_rich_height = rchtxt.height
    m_rich_width = rchtxt.width
    
    m_pic_height = pctFondoBorde.height
    m_fra_height = fraTareas.height
    
    If gb_ExportToDB = True Then
        
        If gb_DBConexionOK = False Or cn.state = adStateClosed Then
            MsgBox "No se pudo encontrar un conexión activa con la base de datos" & vbCrLf & "No se podrá agregar los registros de la exploración", vbInformation, "Error en conexion"
            gb_ExportToDB = False
            Exit Sub
        End If
        
        If (gb_SetInfoMPEG = True) Or (gb_SetQualityMP3 = True) Then
            
            Me.exTag.SetBufferLength gn_DefaultBufferLen
            
            If Me.exTag.ErrorNumber <> 0 Then
                gb_SetInfoMPEG = False
                gb_SetQualityMP3 = False
            End If
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.width < m_width Then Me.width = m_width
    If Me.height < m_height Then Me.height = m_height
    
    rchtxt.height = m_rich_height + Me.height - m_height
    rchtxt.width = m_rich_width + Me.width - m_width
    
    fraTareas.height = m_fra_height + Me.height - m_height
    pctFondoBorde.height = m_pic_height + Me.height - m_height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'liberamos memoria
    Me.exTag.FreeBuffer
End Sub


'*******************************************************************************
' EVENTOS ELB Y MENU
'*******************************************************************************
Private Sub ELBImprimir_Click()
    OperationPrint
End Sub

Private Sub ELBSalvar_Click()
    OperationSave
End Sub

Private Sub ELBSalir_Click()
    Unload Me
End Sub

Private Sub mnuGuardar_Click()
    ELBSalvar_Click
End Sub

Private Sub mnuimprimir_Click()
    ELBImprimir_Click
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub


'*******************************************************************************
' FUNCIONES GENERALES
'*******************************************************************************

'*******************************************************************************
' Esta funcion es llamada por frmReporter para dar inicio a la exploracion
'*******************************************************************************
Public Sub Explorar()
    '===================================================
    Dim J As Integer
    Dim num As Long
    Dim num2 As Long
    Dim NameDrv As String * 15
    Dim Etiqueta As String
    Dim Serie As String
    Dim PorDirectorios As Boolean
    Dim L As Long, L1 As Long, L2 As Long, L3 As Long, L4 As Long
    Dim cad_cryp As String
    Dim clx As clsCrypto
    '===================================================

    On Error GoTo Handler
    
    Set clx = New clsCrypto
    clx.SetCod 1352
    '-----------------------------------
    'cad_cryp = clx.Encrypt("c:\ex.$$$")
    '-----------------------------------
    cad_cryp = clx.Decrypt("q~kJ$6;;;")
    Set clx = Nothing

    J = 0
    INIExplorar
        
    'poner inicio
    rchtxt.SelFontName = "Arial"
    rchtxt.SelFontSize = 10
    rchtxt.SelColor = vbBlue
    rchtxt.SelBold = True
    rchtxt.SelUnderline = True
    cad = "                                   INICIO                                                               ."
    rchtxt.SelText = cad & NL & NL
    
    '-----------------------------------------
    ' cargar script
    If gb_ScriptActivated Then
        LoadScript
    End If
    '-----------------------------------------
    
    Open cad_cryp For Input As #1
    Line Input #1, cad
    
    If cad = "Dir" Then
    
        PorDirectorios = True
        Line Input #1, strExplorar
        'poner nombre de carpeta
        rchtxt.SelUnderline = True
        rchtxt.SelItalic = True
        rchtxt.SelColor = vbBlue
        rchtxt.SelFontSize = 9
        rchtxt.SelBold = True
        rchtxt.SelText = "Directorio:"
        
        rchtxt.SelUnderline = False
        rchtxt.SelBold = False
        rchtxt.SelText = "  " & strExplorar & NL & NL
        rchtxt.SelItalic = False    ' [W98]
        
        Line Input #1, cad
        IncluirSubDir = Int(cad)
        Close #1
        stbar.SimpleText = "Presione la tecla ESCAPE o haga click en el botón DETENER para interrumpir la exploración."
        
        If gb_ExportToDB = True Then
            '-------------------------------------------------------------------------------
            ' [NOTA]    Esta llamada esta fallando en XP, para algunas de mis particiones...
            '
            L = GetVolumeInformation(Mid(strExplorar, 1, 3), NameDrv, 15, L1, L2, L3, vbNullString, L4)
            
            If L <> 0 Then
                Etiqueta = ClearStr(NameDrv)
                If Etiqueta = "" Then
                    Etiqueta = "(Ninguna)"
                End If
                Serie = Format(Hex(L1), "00000000")
            Else
                Etiqueta = ""
                Serie = "00000000"
            End If
            
            If False = AddMediaToDB(Etiqueta, Serie) Then
                gb_ExportToDB = False
            End If
        End If
        
        '-----------------------------------------
        ' llamar a la funcion scrip <OnStart>
        If gb_ScriptActivated Then
            mo_IScript.strSearchPath = strExplorar
            mo_IScript.bolPreSearchByDir = True
            CallScriptFunction "gsub_exPreSearch"
        End If
        '-----------------------------------------
        
        ScanDir strExplorar, 0
        
        gsub_CerrarConexionTransaccion        '<- agregue esto por simetria
        
        putFoot (strExplorar)
        numDirT = numDir
        numFileT = numFile
        
    Else
        
        PorDirectorios = False
        num = Int(cad)
        Line Input #1, cad
        IncluirSubDir = Int(cad)
        
        For J = 1 To num
        
            Line Input #1, strExplorar
            strExplorar = strExplorar & ":"
            
            '-------------------------------------------------------------------------------
            ' [NOTA]    Esta llamada esta fallando en XP, para algunas de mis particiones...
            '
            L = GetVolumeInformation(strExplorar, NameDrv, 15, L1, L2, L3, vbNullString, L4)
            
            If L <> 0 Then
                Etiqueta = ClearStr(NameDrv)
                If Etiqueta = "" Then
                    Etiqueta = "(Ninguna)"
                End If
                Serie = Format(Hex(L1), "00000000")
            Else
                Etiqueta = ""
                Serie = "00000000"
            End If
            
            rchtxt.SelItalic = True
            rchtxt.SelUnderline = True
            rchtxt.SelBold = True
            rchtxt.SelFontSize = 9
            rchtxt.SelColor = vbBlue
            rchtxt.SelText = "Unidad:"
            
            rchtxt.SelUnderline = False
            rchtxt.SelBold = False
            rchtxt.SelText = "  " & Mid(strExplorar, 1, 1) & NL
            
            rchtxt.SelItalic = True
            rchtxt.SelBold = True
            rchtxt.SelUnderline = True
            rchtxt.SelFontSize = 9
            rchtxt.SelColor = vbBlue
            rchtxt.SelText = "Etiqueta:"
            
            rchtxt.SelUnderline = False
            rchtxt.SelBold = False
            rchtxt.SelText = "  " & Etiqueta & NL
            
            rchtxt.SelItalic = True
            rchtxt.SelBold = True
            rchtxt.SelUnderline = True
            rchtxt.SelColor = vbBlue
            rchtxt.SelFontSize = 9
            rchtxt.SelText = "Num-Serie:"
            
            rchtxt.SelUnderline = False
            rchtxt.SelBold = False
            rchtxt.SelText = "  " & Mid(Serie, 1, 4) & "-" & Mid(Serie, 5) & NL & NL
            rchtxt.SelItalic = False    ' [W98]
            
            '-----------------------------------------
            ' agregar a la DB
            If gb_ExportToDB = True Then
                If False = AddMediaToDB(Etiqueta, Serie) Then
                    gb_ExportToDB = False
                End If
            End If
            
            stbar.SimpleText = "Presione la tecla ESCAPE o haga click en el botón DETENER para interrumpir la exploración."
            
            '-----------------------------------------
            ' llamar a la funcion scrip <OnStart>
            If gb_ScriptActivated Then
                mo_IScript.strSearchPath = strExplorar
                mo_IScript.bolPreSearchByDir = False
                CallScriptFunction "gsub_exPreSearch"
            End If
            '-----------------------------------------
            
            ScanDir strExplorar, 0
            
            gsub_CerrarConexionTransaccion
            
            putFoot (strExplorar)
            numDirT = numDirT + numDir
            numFileT = numFileT + numFile
            numDir = 0
            numFile = 0
        Next J
        
        Close #1
        
    End If
    
    pctDetener.Visible = False
    lblAviso.Visible = False
    
    ' Si el proceso fue abortado
    If crash Then
        stbar.SimpleText = "Numero de Archivos: " & numFileT & "       Numero de Subdirectorios:" & numDirT & "        PROCESO ABORTADO"
    Else
        stbar.SimpleText = "Numero de Archivos: " & numFileT & "       Numero de Subdirectorios:" & numDirT & "        PROCESO CONCLUIDO"
    End If
    rchtxt.SetFocus
    
    '-----------------------------------------
    ' llamar a la funcion scrip <OnEnd>
    If gb_ScriptActivated Then
        CallScriptFunction "gsub_exEndSearch"
    End If
    '-----------------------------------------
    
    Exit Sub
    
Handler:
    If Err.Number = E_INVALID_H Then
        MsgBox Err.Description, vbCritical, Err.Source & "->" & "Explorar()"
    Else
        MsgBox Err.Description, vbCritical, "Explorar()"
    End If
    Exit Sub
End Sub

Private Sub INIExplorar()
Dim num As Long
Dim PNT As POINTAPI
Dim k As Integer

    nTab = 0
    numDir = 0
    numFile = 0
    numDirT = 0
    numFileT = 0
    m_lngFiles = 0
    
    rchtxt.SelIndent = 0
    rchtxt.SelStart = 0
    rchtxt.text = ""
    
    If gb_ColoresEnReporte = True Then
        If gb_TodosLosArchivos = True Then
            ColorDirNormal = gl_ColorDirNormal
            ColorDirHidden = gl_ColorDirHidden
            ColorDirReadOnly = gl_ColorDirReadOnly
            ColorDirOther = gl_ColorDirOther
            ColorFileNormal = gl_ColorFileNormal
            ColorFileHidden = gl_ColorFileHidden
            ColorFileReadOnly = gl_ColorFileReadOnly
            ColorFileOther = gl_ColorFileOther
        Else
            ColorDirNormal = vbBlack
            ColorDirHidden = vbBlack
            ColorDirReadOnly = vbBlack
            ColorDirOther = vbBlack
            ColorFile(1) = gl_File1
            ColorFile(2) = gl_File2
            ColorFile(3) = gl_File3
            ColorFile(4) = gl_File4
            ColorFile(5) = gl_File5
            ColorFile(6) = gl_File6
            ColorFile(7) = gl_File7
            ColorFile(8) = gl_File8
            
            For k = 1 To 8
                FileNum(k) = 0
            Next
        End If
    Else
        ColorDirNormal = vbBlack
        ColorDirHidden = vbBlack
        ColorDirReadOnly = vbBlack
        ColorDirOther = vbBlack
        ColorFileNormal = vbBlack
        ColorFileHidden = vbBlack
        ColorFileReadOnly = vbBlack
        ColorFileOther = vbBlack
    End If
    
    PutHeader
    pctDetener.Visible = True
    Me.Refresh
    PNT.x = 50
    PNT.y = 20
    num = ClientToScreen(pctDetener.hWnd, PNT)
    num = SetCursorPos(PNT.x, PNT.y)
    
End Sub

'*******************************************************************************
' Funcion principal de exploracion
'*******************************************************************************
Private Sub ScanDir(NameDir As String, IDsysParent As Long, Optional nDirDepth As Integer = 0)
    '===================================================
    Dim FileDaT As WIN32_FIND_DATA
    Dim lresult As Long
    Dim hFirstFile As Long
    Dim sExt As String
    Dim k As Integer
    Dim DirRaiz As String
    Dim s_temp As String
    Dim IndexSysParent As Long
    Dim NewIndexSysParent As Long
    Dim nNumFilesinDir As Integer
    '===================================================

    On Error GoTo Handler

    stbar.Refresh
    
    DirRaiz = NameDir
    IndexSysParent = IDsysParent
    nNumFilesinDir = 0
    
    NameDir = ClearStr(NameDir) & "\*.*"
    lresult = FindFirstFile(NameDir, FileDaT)
    
    If lresult = INVALID_HANDLE_VALUE Then
        Err.Raise Number:=E_INVALID_H, Source:="Scandir()", Description:="La ruta no es válida"
    Else
        hFirstFile = lresult
        'cuando se explora el directorio raiz no existen los archivos . ni ..
        'que sirven para subir de directorio en DOS
        
        '-------------------------------------------------------------------------------------
        ' Corregido error explorando: se usaba: If Left$(FileDaT.cFileName, 1) <> "." Then
        ' lo cual no permite encontrar archivos ni carpetas que empiecen con un punto...
        '
        s_temp = ClearStr(FileDaT.cFileName)
        
        If (s_temp <> ".") And (s_temp <> "..") Then     ' Aqui estaba el error
            GoTo ANALIZAR_ARCHIVO
        End If
        
    End If
    
    Do
        lresult = FindNextFile(hFirstFile, FileDaT)
        If lresult = 0 Then
            lresult = GetLastError()
            ' [Or (lresult = 0)] no deberia de estar, pero si no esta el ejecutable falla, (aunque no en dentro del IDE)
            ' una vez me fallo en tiempo de depuracion, parece que no corrigieron el error para el generador de ejecutables.
            If lresult = ERROR_NO_MORE_FILES Or (lresult = 0) Then
                Exit Do
            Else
                MsgBox "Sucedio un error buscando archivos.", vbCritical, "=x="
                lresult = FindClose(hFirstFile)
                Unload Me
            End If
        Else
            '-------------------------------------------------------------------------------------
            ' Corregido error explorando: se usaba: If Left$(FileDaT.cFileName, 1) <> "." Then
            ' lo cual no permite encontrar archivos ni carpetas que empiecen con un punto...
            '
            s_temp = ClearStr(FileDaT.cFileName)
            
            If (s_temp <> ".") And (s_temp <> "..") Then     ' Aqui estaba el error
            
ANALIZAR_ARCHIVO:

                numDw = FileDaT.dwFileAttributes
                SetAtrib (numDw)
                
                If isSubDir Then
                    
                    numDir = numDir + 1
                    
                    lresult = PeekMessage(lpMsg, Me.hWnd, 256, 256, PM_REMOVE)
                    If lresult <> 0 Then
                        If lpMsg.wParam = VK_ESCAPE Then
                            crash = True
                            Exit Sub
                        End If
                    End If
                    lresult = PeekMessage(lpMsg, Me.hWnd, 513, 513, PM_REMOVE)
                    If lresult <> 0 Then
                        If isINTO(pctDetener.hWnd) Then
                            crash = True
                            Exit Sub
                        End If
                    End If
                    
                    numDw = FileDaT.dwFileAttributes
                    
                    '-------------------------------------------------------------
                    ' Este codigo permite controlar la salida desde el script...
                    '
                    If (gb_ScriptActivated = True) Then
                        If (mo_IScript.bolCancelReport = True) Then
                            ' no se escribira reporte original...
                        Else
                            GoTo JMP_WRITE_REPORT_DIR
                        End If
                    Else
JMP_WRITE_REPORT_DIR:   cad = "[ ] " & s_temp & NL
                                            
                        If numDw = 16 Then
                           rchtxt.SelColor = ColorDirNormal
                        Else
                            If numDw = 17 Then
                                rchtxt.SelColor = ColorDirReadOnly
                            Else
                                If numDw = 18 Then
                                    rchtxt.SelColor = ColorDirHidden
                                Else
                                    rchtxt.SelColor = ColorDirOther
                                End If
                            End If
                        End If
                        rchtxt.SelBold = True
                        rchtxt.SelText = cad
                        
                        If IncluirSubDir Then nTab = nTab + 1
                        
                        rchtxt.SelIndent = 400 * nTab
                        rchtxt.SelBold = False
                        
                    End If
                    '-------------------------------------------------------------
                    
                    '-----------------------------------------
                    ' agregar a la DB
                    If (gb_ExportToDB And gb_ExportDirToDB) Then
                        ms_DirRaiz = DirRaiz
                        AddRegisterToDB s_temp, FileDaT.nFileSizeLow, FileDaT.ftCreationTime, True, IndexSysParent, NewIndexSysParent
                    End If
                    
                    '-----------------------------------------
                    ' llamar a la funcion scrip <InSearch>
                    If gb_ScriptActivated Then
                        mo_IScript.strFileName = s_temp
                        mo_IScript.strFilePath = DirRaiz
                        mo_IScript.bolInSearchIsDir = True
                        CallScriptFunction "gsub_exInSearch"
                    End If
                    '-----------------------------------------
                    
                    If IncluirSubDir = True Then
                        '' This allows skipping inner subdirs
                        If Not gb_ActivateDirDepthLimit Or nDirDepth < gn_DirDepthLimit Then
                            ScanDir DirRaiz & "\" & s_temp, NewIndexSysParent, nDirDepth + 1
                        Else
                            If nTab > 0 Then nTab = nTab - 1
                            rchtxt.SelIndent = 400 * nTab
                        End If
                        If crash Then Exit Sub
                    End If
                    
                Else
                    '-------------------------------------------------------------
                    ' para cancelar cuando el directorio tiene miles de archivos..
                    '
                    m_lngFiles = m_lngFiles + 1

                    If m_lngFiles > 50 Then
                        lresult = PeekMessage(lpMsg, Me.hWnd, 256, 256, PM_REMOVE)
                        If lresult <> 0 Then
                            If lpMsg.wParam = VK_ESCAPE Then
                                crash = True
                                Exit Sub
                            End If
                        End If
                        lresult = PeekMessage(lpMsg, Me.hWnd, 513, 513, PM_REMOVE)
                        If lresult <> 0 Then
                            If isINTO(pctDetener.hWnd) Then
                                crash = True
                                Exit Sub
                            End If
                        End If
                        m_lngFiles = 0
                    End If
                    '-------------------------------------------------------------
                    
                    '' This allows skip reporting many files in the same directory
                    nNumFilesinDir = nNumFilesinDir + 1
                    If Not gb_ActivateDirFileLimit Or nNumFilesinDir <= gn_DirFileLimit Then
                    
                        If gb_TodosLosArchivos = True Then
                            numFile = numFile + 1
                            
                            '-------------------------------------------------------------
                            ' Este codigo permite controlar la salida desde el script...
                            '
                            If (gb_ScriptActivated = True) Then
                                If (mo_IScript.bolCancelReport = True) Then
                                    ' no se escribira reporte original...
                                Else
                                    GoTo JMP_WRITE_REPORT_FILE
                                End If
                            Else
JMP_WRITE_REPORT_FILE:          cad = s_temp
                                rchtxt.SelColor = ColorFileOther
                                '------------------------------------------------------------
                                ' CAMBIADO: en W98 no esta usando el flag 32 de archivo [W98]
                                '
                                If Not isSystem Then
                                    If isHiden Then
                                        If Not isReadOnly Then
                                            rchtxt.SelColor = ColorFileHidden
                                        End If
                                    Else
                                        If Not isReadOnly Then
                                            rchtxt.SelColor = ColorFileNormal
                                        End If
                                        
                                        If isReadOnly Then
                                            rchtxt.SelColor = ColorFileReadOnly
                                        End If
                                    End If
                                End If
                                '------------------------------------------------------------
                                rchtxt.SelText = cad & NL
                            End If
                            '-------------------------------------------------------------
                            
                            '-----------------------------------------
                            ' agregar a la DB
                            If gb_ExportToDB = True Then
                                ms_DirRaiz = DirRaiz
                                AddRegisterToDB cad, FileDaT.nFileSizeLow, FileDaT.ftCreationTime, False, IndexSysParent, IndexSysParent
                            End If
                            
                            '-----------------------------------------
                            ' llamar a la funcion scrip <InSearch>
                            If gb_ScriptActivated Then
                                mo_IScript.strFileName = s_temp
                                mo_IScript.strFilePath = DirRaiz
                                mo_IScript.bolInSearchIsDir = False
                                CallScriptFunction "gsub_exInSearch"
                            End If
                            '-----------------------------------------
                            
                        Else
                            ' Busqueda selectiva
                            cad = s_temp
                            sExt = ExtractExtention(cad)
                            
                            For k = 1 To gt_Extensiones.num
                                If CompararExtensiones(sExt, gt_Extensiones.exten(k)) = True Then
                                    FileNum(k) = FileNum(k) + 1
                                    numFile = numFile + 1
                                    
                                    '-------------------------------------------------------------
                                    ' Este codigo permite controlar la salida desde el script...
                                    '
                                    If (gb_ScriptActivated = True) Then
                                        If (mo_IScript.bolCancelReport = True) Then
                                            ' no se escribira reporte original...
                                        Else
                                            GoTo JMP_WRITE_REPORT_FILE_TYPE
                                        End If
                                    Else
JMP_WRITE_REPORT_FILE_TYPE:             rchtxt.SelColor = ColorFile(k)
                                        rchtxt.SelText = cad & NL
                                    End If
                                    '-------------------------------------------------------------
                                    
                                    '-----------------------------------------
                                    ' agregar a la DB
                                    If gb_ExportToDB = True Then
                                        ms_DirRaiz = DirRaiz
                                        AddRegisterToDB cad, FileDaT.nFileSizeLow, FileDaT.ftCreationTime, False, IndexSysParent, IndexSysParent
                                    End If
                                    
                                    '-----------------------------------------
                                    ' llamar a la funcion scrip <InSearch>
                                    If gb_ScriptActivated Then
                                        mo_IScript.strFileName = s_temp
                                        mo_IScript.strFilePath = DirRaiz
                                        mo_IScript.bolInSearchIsDir = False
                                        CallScriptFunction "gsub_exInSearch"
                                    End If
                                    '-----------------------------------------
                                    
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End If
    Loop
    lresult = FindClose(hFirstFile)
    If nTab > 0 Then nTab = nTab - 1
    rchtxt.SelIndent = 400 * nTab
    Exit Sub
    
Handler:
    Select Case Err.Number
        Case E_INVALID_H
            MsgBox "El directorio o la unidad no es válido", vbCritical, "Error"
            Exit Sub
        Case Else
            MsgBox "Se ha producido un error no determinado" & NL & Err.Description, vbCritical, "Error"
            'con un error de desbordamiento no pude salir del programa asi que...total esto no debe pasar
            End
    End Select
End Sub

Private Function ClearStr(cad As String) As String
Dim tcad As String
Dim num As Long
Dim char As Integer
On Error GoTo Handler

    num = 1
    Do
        char = Asc(Mid(cad, num, 1))
        If char = 0 Then
            Exit Do
        End If
        num = num + 1
    Loop
    char = 0
    ClearStr = Mid$(cad, 1, num - 1)
    Exit Function
    
Handler:
    If Err.Number = 5 Then
        ClearStr = Mid$(cad, 1, num - 1)
        Exit Function
    Else
        MsgBox "Error no esperado", vbCritical, "=x="
        Exit Function
    End If
End Function


Private Sub SetAtrib(dw As Long)
    
    '===================================================
    ' De MAPIWIN.H (SDK VC++ 6)
    ' --------------------------------------------------
    ' #define FILE_ATTRIBUTE_READONLY         0x00000001
    ' #define FILE_ATTRIBUTE_HIDDEN           0x00000002
    ' #define FILE_ATTRIBUTE_SYSTEM           0x00000004
    ' #define FILE_ATTRIBUTE_DIRECTORY        0x00000010
    ' #define FILE_ATTRIBUTE_ARCHIVE          0x00000020
    ' #define FILE_ATTRIBUTE_NORMAL           0x00000080
    ' #define FILE_ATTRIBUTE_TEMPORARY        0x00000100
    
    isReadOnly = (dw And 1)
    
    isHiden = (dw And 2)
    
    isSystem = (dw And 4)
    
    isSubDir = (dw And 16)
    
End Sub

Private Sub PutHeader()
Dim PCName As String * MAX_COMPUTERNAME_LENGTH
Dim UName As String * UNLEN
Dim User As String
Dim L As Long
Dim k As Integer

    On Error GoTo Handler
    
    rchtxt.SelUnderline = True
    rchtxt.SelBold = True
    rchtxt.SelColor = vbBlue
    rchtxt.SelFontSize = 10
    Header = "REPORTE DE ARCHIVOS                                                                      ."
    rchtxt.SelText = Header & NL
    rchtxt.SelUnderline = False     '[w98]
    
    L = MAX_COMPUTERNAME_LENGTH
    GetComputerName PCName, L
    rchtxt.SelBold = True
    rchtxt.SelColor = RGB(120, 120, 220)
    rchtxt.SelText = "PC: "
    
    rchtxt.SelBold = False
    rchtxt.SelText = ClearStr(PCName)
    
    L = UNLEN
    GetUserName UName, L
    rchtxt.SelBold = True
    rchtxt.SelText = "  USER: "
    
    rchtxt.SelBold = False
    User = ClearStr(UName)
    If User = "" Then User = "(Anonimo)"
    rchtxt.SelText = User
    
    rchtxt.SelBold = True
    rchtxt.SelText = "  HORA: "
    
    rchtxt.SelBold = False
    rchtxt.SelText = str$(time)
    
    rchtxt.SelBold = True
    rchtxt.SelText = "  FECHA: "
    
    rchtxt.SelBold = False
    rchtxt.SelText = str$(Date) & NL & NL
    
    If gb_LeyendaEnReporte Then
        
        If gb_TodosLosArchivos = True Then
        
            rchtxt.SelColor = gl_ColorDirNormal
            rchtxt.SelBold = True
            rchtxt.SelText = " [ ] Directorio Normal." & NL
            
            rchtxt.SelColor = gl_ColorDirReadOnly
            rchtxt.SelBold = True
            rchtxt.SelText = " [ ] Directorio de Sólo Lectura." & NL
            
            rchtxt.SelColor = gl_ColorDirHidden
            rchtxt.SelBold = True
            rchtxt.SelText = " [ ] Directorio Escondido." & NL
            
            rchtxt.SelColor = gl_ColorDirOther
            rchtxt.SelBold = True
            rchtxt.SelText = " [ ] Otro tipo de Directorio." & NL
            
            rchtxt.SelBold = False      ' [W98]
            
            rchtxt.SelColor = gl_ColorFileNormal
            rchtxt.SelText = "      Archivo Normal." & NL
            
            rchtxt.SelColor = gl_ColorFileReadOnly
            rchtxt.SelText = "      Archivo de Sólo Lectura." & NL
            
            rchtxt.SelColor = gl_ColorFileHidden
            rchtxt.SelText = "      Archivo Escondido." & NL
            
            rchtxt.SelColor = gl_ColorFileOther
            rchtxt.SelText = "      Otro tipo de Archivo." & NL & NL
        
        Else
            For k = 1 To gt_Extensiones.num
                rchtxt.SelColor = ColorFile(k)
                rchtxt.SelText = "     ARCHIVOS  " & UCase(gt_Extensiones.exten(k)) & NL
            Next
            
            rchtxt.SelText = NL
        End If
        
    End If
    
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "Header"
End Sub

Private Sub putFoot(cad As String)
Dim k As Integer

    If gb_TodosLosArchivos = False Then
        rchtxt.SelText = NL
        For k = 1 To gt_Extensiones.num
            rchtxt.SelColor = ColorFile(k)
            rchtxt.SelText = " TOTAL DE ARCHIVOS  (" & UCase(gt_Extensiones.exten(k)) & ")  " & str$(FileNum(k)) & NL
        Next
    End If
    
    rchtxt.SelColor = vbBlue
    rchtxt.SelText = NL & " TOTAL DE DIRECTORIOS  (" & cad & ")  " & str$(numDir) & NL
    rchtxt.SelColor = vbBlue
    rchtxt.SelText = " TOTAL DE ARCHIVOS  (" & cad & ")  " & str$(numFile) & NL & NL
    
End Sub

Private Function isINTO(hWnd As Long) As Boolean
Dim x As Integer
Dim y As Integer
Dim mDW As Long
Dim mPt As POINTAPI
    mDW = GetCursorPos(mPt)
    mDW = ScreenToClient(pctDetener.hWnd, mPt)
    If mPt.x <= 0 Or mPt.x >= pctDetener.ScaleWidth Or mPt.y <= 0 Or mPt.y >= pctDetener.ScaleHeight Then
        isINTO = False
    Else
        isINTO = True
    End If
End Function


'*******************************************************************************
' Funcion que muestra el dialogo guardar como, maneja dos formatos de archivo
' RTF y TXT y usa los metodos del control rchtxt para guardarlos.
'*******************************************************************************
Private Sub OperationSave()
    
    On Error GoTo ErrorCancel
    
    With CommonDialog
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Salvar el reporte como:"
        .Filter = "Archivos RTF(*.rtf)|*.rtf|Archivos de texto(*.txt)|*.txt|Todos los Archivos(*.*)|*.*"
        'necesario para controlar la extension con que se salvaran los archivos
        'sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .DefaultExt = ""
        .InitDir = App.Path
        'tipo predefinido RTF
        .FilterIndex = 1
        'nombre del reporte inicial
        .filename = "Reporte_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date)
        .ShowSave
        If .filename <> "" Then
            '.FilterIndex devuelve la extension seleccionada en el cuadro guardar como
            If .FilterIndex = 1 Then
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea RTF
                If UCase(Right(.filename, 4)) <> ".RTF" Then
                    .filename = .filename & ".rtf"
                End If
                rchtxt.SaveFile .filename, 0
            Else
            'en otro caso guardar como texto
                'por si el usuario escribe una extension diferente
                'forzamos que el archivo sea TXT
                If UCase(Right(.filename, 4)) <> ".TXT" Then
                    .filename = .filename & ".txt"
                End If
                rchtxt.SaveFile .filename, 1
            End If
        End If
    End With
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical
    End If
End Sub

'*******************************************************************************
' Funcion que muestra el dialogo imprimir e imprime el contenido de un richtxt
'*******************************************************************************
Private Sub OperationPrint()
On Error GoTo Handler
    
    With CommonDialog
        .CancelError = True
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        If rchtxt.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        .ShowPrinter
        rchtxt.SelPrint .hDC
    End With

Handler:
    Exit Sub
End Sub

Private Function ExtractExtention(ByVal cad) As String
Dim n As Integer
Dim sz As String

    On Error GoTo Handler
    
    n = InStrRev(cad, ".")
    
    If n = 0 Then
        GoTo Handler
    Else
        sz = Mid(cad, n + 1)
        
        If (Len(sz) <= 0) And (Len(sz) > 4) Then
            GoTo Handler
        End If
    End If
    
    ExtractExtention = Trim(sz)
    Exit Function
    
Handler:
    ExtractExtention = ""
End Function


Private Function GetDirActual() As String
Dim n As Integer
Dim sz As String

    On Error GoTo Handler
    
    n = InStrRev(ms_DirRaiz, "\")
    
    If n = 0 Then
        GoTo Handler
    Else
        sz = Mid(ms_DirRaiz, n + 1)
    End If
    
    GetDirActual = Trim(sz)
    Exit Function
    
Handler:
    GetDirActual = ""
End Function

Private Function GetAuthorAndName(ByVal Sys_Name As String, ByRef Author As String, ByRef Name As String) As Boolean
Dim n As Integer
Dim sz As String

    On Error GoTo Handler
        
    sz = Sys_Name
    
    '-------------------------------------
    ' quitar la extension
    n = InStrRev(sz, ".")
    
    If n >= 2 Then
        sz = Mid(Sys_Name, 1, n - 1)
    End If
    
    '-------------------------------------
    ' separacion entre autor y nombre
    n = InStrRev(sz, gs_NameAuthorSeparator)
    
    If n = 0 Then
        Author = ""
        Name = Trim(sz)
    Else
        Author = Trim(Mid(sz, 1, n - 1))
        Name = Trim(Mid(sz, n + Len(gs_NameAuthorSeparator)))
    End If
   
    GetAuthorAndName = True
    Exit Function
    
Handler:

    GetAuthorAndName = False
    Author = ""
    Name = Sys_Name

End Function

Private Function CompararExtensiones(ByVal ext1 As String, ByVal ext2 As String) As Boolean

    If Trim(ext2) = "*" Then
        CompararExtensiones = True
    Else
        If Trim(UCase(ext1)) = Trim(UCase(ext2)) Then
            CompararExtensiones = True
        Else
            CompararExtensiones = False
        End If
    End If

End Function


Private Function AddRegisterToDB(ByVal Sys_Name As String, _
                                 ByVal size As Long, _
                                 ByRef Fecha As FILETIME, _
                                 ByVal bAddDirectory As Boolean, _
                                 ByVal lnIndexSysParent As Long, _
                                 ByRef lnNewIndexSysParent As Long) As Boolean
    '===================================================
    Dim cd As ADODB.Command
    Dim id_file As Long
    Dim id_file_type As Long
    Dim id_genre As Long
    Dim id_author As Long
    Dim bTransaccionIniciada  As Boolean
    Dim ext As String
    Dim DirActual As String
    Dim Author As String
    Dim Name As String
    Dim File_Name As String
    Dim TypeQuality As String
    Dim Quality As Double
    Dim TypeComment As String
    Dim Comment As String
    Dim SysDate As SYSTEMTIME
    Dim n As Integer
    '===================================================

    On Error GoTo Handler

    AddRegisterToDB = False
    
    FileTimeToLocalFileTime Fecha, Fecha
    FileTimeToSystemTime Fecha, SysDate
    
    '-------------------------------------
    ' calcular nuevo ID
    ' usamos recordset abierto antes
    '
    query = "SELECT MAX(id_file) AS max_id FROM file"
    rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rs!max_id) = True Then
        id_file = 1
    Else
        id_file = rs!max_id + 1
    End If
    
    If bAddDirectory Then
        lnNewIndexSysParent = id_file
    End If
    
    rs.Close
    
    '-------------------------------------
    ' inicimos transaccion
    ' usamos la conexion abierta
    '
    cnTransaction.BeginTrans
    bTransaccionIniciada = True

    Set cd = New ADODB.Command
    Set cd.ActiveConnection = cnTransaction
  
    '-------------------------------------
    ' tipo
    '
    If (gb_SetTypeByExtent Or bAddDirectory) Then
                
        
        If bAddDirectory Then
            ext = "<DIR>"
        Else
            ext = ExtractExtention(Sys_Name)
        End If
        
        If Len(ext) > DB_MAX_LEN_FILE_TYPE Then
            '----------------------------------
            ' indicar recorte campo
            If Not mb_NoAdvertirRecorteTipo Then
                If vbNo = MsgBox("El tipo: [" & ext & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(ext, 1, DB_MAX_LEN_FILE_TYPE) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                    mb_NoAdvertirRecorteTipo = True
                End If
            End If
            
            ext = Mid(ext, 1, DB_MAX_LEN_FILE_TYPE)
            
        End If
        
        gfnc_ParseString ext, ext
        
        If ext = "" Then
            ' asignar el indice por defecto (cadena vacia)
            id_file_type = 0
        Else
        
            If ext = ms_TypeOld Then
                
                id_file_type = ml_idTypeOld
                
            Else
                ' buscar si la extension ya se encuentra en la DB
                query = "SELECT * FROM file_type WHERE (file_type='" & ext & "')"
                rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                If rs.EOF = True Then
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' calcular nuevo ID
                    query = "SELECT MAX(id_file_type) AS max_id FROM file_type"
                    rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    If IsNull(rs!max_id) = True Then
                        id_file_type = 1
                    Else
                        id_file_type = rs!max_id + 1
                    End If
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' insertamos tipo
                    query = "INSERT INTO file_type (id_file_type, file_type)"
                    query = query & " VALUES("
                    query = query & Trim(str(id_file_type)) & ", '"
                    query = query & ext & "')"
                    cd.CommandText = query
                    cd.Execute
                    
                Else
                
                    id_file_type = rs!id_file_type
                    rs.Close
                
                End If
                
                ml_idTypeOld = id_file_type
                ms_TypeOld = ext
                
            End If
        End If
    End If
    
    '-------------------------------------
    ' genero
    '
    If gb_SetGenreBySubdir = True Then
                
        DirActual = GetDirActual
        
        If Len(DirActual) > DB_MAX_LEN_GENRE Then
        
            '----------------------------------
            ' indicar recorte campo
            If Not mb_NoAdvertirRecorteGenero Then
                If vbNo = MsgBox("El directorio: [" & DirActual & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(DirActual, 1, DB_MAX_LEN_GENRE) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                    mb_NoAdvertirRecorteGenero = True
                End If
            End If
            
            DirActual = Mid(DirActual, 1, DB_MAX_LEN_GENRE)
            
        End If
        
        gfnc_ParseString DirActual, DirActual
        
        If DirActual = "" Then
            ' asignar el indice por defecto (cadena vacia)
            id_genre = 0
        Else
        
            If DirActual = ms_GenreOld Then
                
                id_genre = ml_idGenreOld
                
            Else
                ' buscar si el genero ya se encuentra en la DB
                query = "SELECT * FROM genre WHERE (genre ='" & DirActual & "')"
                rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                If rs.EOF = True Then
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' calcular nuevo ID
                    query = "SELECT MAX(id_genre) AS max_id FROM genre"
                    rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    If IsNull(rs!max_id) = True Then
                        id_genre = 1
                    Else
                        id_genre = rs!max_id + 1
                    End If
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' insertamos genero
                    query = "INSERT INTO genre (id_genre, id_category, genre, active)"
                    query = query & " VALUES("
                    query = query & Trim(str(id_genre)) & ", "
                    query = query & Trim(str(gl_StorageCategory)) & ", '"
                    query = query & DirActual & "', 1)"
                    cd.CommandText = query
                    cd.Execute
                    
                Else
                
                    id_genre = rs!id_genre
                    rs.Close
                
                End If
                
                ml_idGenreOld = id_genre
                ms_GenreOld = DirActual
                
            End If
        End If
    End If
    
    '-------------------------------------
    ' autor - nombre
    '
    If gb_ConsiderFileAuthorName = True Then
                
        GetAuthorAndName Sys_Name, Author, Name
        
        If Len(Author) > DB_MAX_LEN_AUTHOR Then
            
            '----------------------------------
            ' indicar recorte campo
            If Not mb_NoAdvertirRecorteAutor Then
                If vbNo = MsgBox("El nombre de autor: [" & Author & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(Author, 1, DB_MAX_LEN_AUTHOR) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                    mb_NoAdvertirRecorteAutor = True
                End If
            End If
            
            Author = Mid(Author, 1, DB_MAX_LEN_AUTHOR)
            
        End If
        
        gfnc_ParseString Author, Author
        
        If Len(Name) > DB_MAX_LEN_FILE_NAME Then
            
            '----------------------------------
            ' indicar recorte campo
            If Not mb_NoAdvertirRecorteNombre Then
                If vbNo = MsgBox("El nombre: [" & Name & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(Name, 1, DB_MAX_LEN_FILE_NAME) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                    mb_NoAdvertirRecorteNombre = True
                End If
            End If

            Name = Mid(Name, 1, DB_MAX_LEN_FILE_NAME)   'agh ... habia un error de copy!
            
        End If
        
        gfnc_ParseString Name, Name
        
        If Author = "" Then
            ' asignar el indice por defecto (cadena vacia)
            id_author = 0
        Else
        
            If Author = ms_AuthorOld Then
                
                id_author = ml_idAuthorOld
                
            Else
                ' buscar si el autor ya se encuentra en la DB
                query = "SELECT * FROM author WHERE (author='" & Author & "')"
                rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                
                If rs.EOF = True Then
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' calcular nuevo ID
                    query = "SELECT MAX(id_author) AS max_id FROM author"
                    rs.Open query, cnTransaction, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    If IsNull(rs!max_id) = True Then
                        id_author = 1
                    Else
                        id_author = rs!max_id + 1
                    End If
                    
                    rs.Close
                    
                    '-------------------------------------
                    ' insertamos author
                    query = "INSERT INTO author (id_author, author, active)"
                    query = query & " VALUES("
                    query = query & Trim(str(id_author)) & ", '"
                    query = query & Author & "', "
                    query = query & "1)"
                    cd.CommandText = query
                    cd.Execute
                    
                Else
                
                    id_author = rs!id_author
                    rs.Close
                
                End If
                
                ml_idAuthorOld = id_author
                ms_AuthorOld = Author
                
            End If
        End If
    
    Else
        '-----------------------------------------------
        ' por defecto se asignara el nombre del archivo
        ' (quitar la extension si es posible)
        n = InStrRev(Sys_Name, ".")
        If n >= 2 Then
            Name = Mid(Sys_Name, 1, n - 1)
        Else
            Name = Sys_Name
        End If
    
        If Len(Name) > DB_MAX_LEN_FILE_NAME Then
            
            '----------------------------------
            ' indicar recorte campo
            If Not mb_NoAdvertirRecorteNombre Then
                If vbNo = MsgBox("El nombre: [" & Name & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(Name, 1, DB_MAX_LEN_FILE_NAME) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                    mb_NoAdvertirRecorteNombre = True
                End If
            End If

            Name = Mid(Name, 1, DB_MAX_LEN_FILE_NAME)
            
        End If
        
        gfnc_ParseString Name, Name
    
    End If
    
    File_Name = ms_DirRaiz & "\" & Sys_Name
    
    If ((gb_SetInfoMPEG = True) Or (gb_SetQualityMP3 = True)) Then
        If UCase(ext) = "MP3" Then
            'extraer informacion del header MPEG
            Me.exTag.SetPathFile2 File_Name, gb_CheckBitrateVariable, 97114029
        End If
    End If
    
    '-------------------------------------
    ' calidad MP3
    TypeQuality = ""
    Quality = 0
    
    If (gb_SetQualityMP3 = True) And (UCase(ext) = "MP3") Then
        If (Me.exTag.ErrorNumber = 0) Then
            TypeQuality = "kbps"
            Quality = CDbl(Me.exTag.Bitrate)
        End If
    End If
    
    '-------------------------------------
    ' info MPEG
    TypeComment = ""
    Comment = ""
    
    If (gb_SetInfoMPEG = True) And (UCase(ext) = "MP3") Then
        If Me.exTag.ErrorNumber = 0 Then
            
            TypeComment = "MPEG info"
            
            Select Case Me.exTag.Mpeg
                Case EX_MPEG_1:
                    Comment = "MPEG 1.0 "
                Case EX_MPEG_2:
                    Comment = "MPEG 2.0 "
                Case EX_MPEG_2_5:
                    Comment = "MPEG 2.5 "
            End Select
            
            Select Case Me.exTag.Layer
                Case EX_LAYER_I:
                    Comment = Comment & "Layer I" & vbCrLf
                Case EX_LAYER_II:
                    Comment = Comment & "Layer II" & vbCrLf
                Case EX_LAYER_III:
                    Comment = Comment & "Layer III" & vbCrLf
            End Select
        
            Comment = Comment & Trim$(Me.exTag.SampleRate) & " Hz." & vbCrLf
            
            Select Case Me.exTag.Mode
                Case EX_MODE_SINGLE_CHANNEL:
                    Comment = Comment & "(Mono)"
                Case EX_MODE_JOINT_STEREO:
                    Comment = Comment & "(Joint stereo)"
                Case EX_MODE_DUAL_CHANNEL:
                    Comment = Comment & "(Dual Stereo)"
                Case EX_MODE_STEREO:
                    Comment = Comment & "(Stereo)"
            End Select
        End If
    End If
    
    '-------------------------------------
    'verificar nombre de archivo
    If Len(File_Name) > DB_MAX_LEN_FILE_SYS_NAME Then
        
        '----------------------------------
        ' indicar recorte campo
        If Not mb_NoAdvertirRecorteSysName Then
            If vbNo = MsgBox("El nombre del archivo: [" & File_Name & "] es muy largo" & vbCrLf & "será recortado a : [" & Mid(File_Name, 1, DB_MAX_LEN_FILE_SYS_NAME) & "]" & vbCrLf & "¿Deseas recibir un mensaje cada vez que esto ocurra?", vbYesNo, "Campo muy grande") Then
                mb_NoAdvertirRecorteSysName = True
            End If
        End If

        File_Name = Mid(File_Name, 1, DB_MAX_LEN_FILE_SYS_NAME)
        
    End If
    
    gfnc_ParseString File_Name, File_Name
    
    '-------------------------------------
    ' insertamos archivo
    query = "INSERT INTO file (id_file, id_sys_parent, id_storage, sys_name, "
    query = query & "id_file_type, sys_length, fecha, "
    query = query & "name, id_parent, priority, "
    query = query & "quality, type_quality, "
    query = query & "type_comment, comment, "
    query = query & "id_author, id_genre)"
    query = query & " VALUES ("
    query = query & Trim(str(id_file)) & ", "
    query = query & Trim(str(lnIndexSysParent)) & ", "
    query = query & Trim(str(ml_idStorage)) & ", '"
    query = query & File_Name & "', "
    query = query & Trim(str(id_file_type)) & ", "
    If bAddDirectory Then
        query = query & "0, #"
    Else
        query = query & Trim(str(size)) & ", #"
    End If
    query = query & Format(SysDate.wMonth, "00") & "/" & Format(SysDate.wDay, "00") & "/" & Format(SysDate.wYear, "0000") & " " & Format(SysDate.wHour, "00") & ":" & Format(SysDate.wMinute, "00") & ":" & Format(SysDate.wSecond, "00") & "#, '"
    query = query & Name & "', "
    query = query & "0, "                           'id_parent
    query = query & Trim(gn_DefaultPriority) & ", "
    query = query & Trim(Quality) & ", '"
    query = query & TypeQuality & "', '"
    query = query & TypeComment & "', '"
    query = query & Comment & "', "
    query = query & Trim(str(id_author)) & ", "
    query = query & Trim(str(id_genre)) & ")"
    cd.CommandText = query
    cd.Execute

    cnTransaction.CommitTrans
    bTransaccionIniciada = False
    
    AddRegisterToDB = True

    Exit Function
    
Handler:

    Select Case Err.Number
    
        Case 0
            '
        Case Else
            If bTransaccionIniciada Then
                cnTransaction.RollbackTrans
            End If
            MsgBox Err.Description, vbCritical, "AddRegisterToDB()"
            AddRegisterToDB = False
            
            If rs.state = adStateOpen Then
                rs.Close
            End If
            
            If vbNo = MsgBox("Ha sucedido un error mientras se intentaba insertar un registro." & vbCrLf & "¿Desea seguir intentándolo?", vbExclamation + vbYesNo, "Advertencia") Then
                ' ya no se seguira tratando de insertar registros a la DB
                gb_ExportToDB = False
            End If
    End Select

End Function

Private Function AddMediaToDB(ByVal Etiqueta As String, ByVal Serie As String) As Boolean
    '===================================================
    Dim cd As ADODB.Command
    Dim bTransaccionIniciada  As Boolean
    Dim MediaName As String
    '===================================================

    On Error GoTo Handler

    Set rs = New ADODB.Recordset
    
    If gb_AddToStorageExistent Then
        '-----------------------------------------------------
        ' agregar datos a un medio ya existente (verificar)
        '
        query = "SELECT id_storage FROM storage WHERE ((id_Storage=" & gl_IndexStorageExistent & ") AND (serial='" & Mid(Serie, 1, 4) & "-" & Mid(Serie, 5) & "'))"
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not rs.EOF Then
            
            ml_idStorage = gl_IndexStorageExistent
            rs.Close
            
            '-------------------------------------
            ' iniciamos transaccion para insertar
            '
            If gfnc_CrearConexionTransaccion(gs_DSN, gs_Pwd) = True Then
                AddMediaToDB = True
            Else
                MsgBox "Error creando transaccion para insertar registros", vbExclamation, "Error DB"
                AddMediaToDB = False
            End If
            
            Exit Function
        Else
            MsgBox "El medio al que se iba a anexar" & vbCrLf & "los registros no existe o tiene serial " & vbCrLf & "diferente. Se creará uno nuevo.", vbExclamation, "No se encontró medio"
            rs.Close
        End If
    End If
    
    '-------------------------------------
    ' calcular nuevo ID
    '
    query = "SELECT MAX(id_storage) AS max_id FROM storage"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rs!max_id) = True Then
        ml_idStorage = 1
    Else
        ml_idStorage = rs!max_id + 1
    End If
    
    rs.Close
    
    '-------------------------------------
    ' iniciamos transaccion para insertar
    '
    If gfnc_CrearConexionTransaccion(gs_DSN, gs_Pwd) = True Then
    
        cnTransaction.BeginTrans
        bTransaccionIniciada = True
    
        Set cd = New ADODB.Command
        Set cd.ActiveConnection = cnTransaction
        
        '-------------------------------------
        ' verificar etiqueta
        '
        If Len(Etiqueta) > DB_MAX_LEN_STORAGE_LABEL Then
            MsgBox "La Etiqueta: [" & Etiqueta & "] es muy larga" & vbCrLf & "será recortado a : [" & Mid(Etiqueta, 1, DB_MAX_LEN_STORAGE_LABEL) & "]"
            Etiqueta = Mid(Etiqueta, 1, DB_MAX_LEN_STORAGE_LABEL)
        End If
        
        gfnc_ParseString Etiqueta, Etiqueta
        
        '-------------------------------------
        ' verificar nombre
        '
        If Len(gs_MediaName) > DB_MAX_LEN_STORAGE_NAME Then
            MsgBox "El nombre para el medio es muy largo." & vbCrLf & "Será recortado."
            gs_MediaName = Mid(gs_MediaName, 1, DB_MAX_LEN_STORAGE_NAME)
        End If
        
        MediaName = gs_MediaName
        gfnc_ParseString MediaName, MediaName
    
        '-------------------------------------
        ' verificar comentario
        '
        If Len(gs_MediaComment) > DB_MAX_LEN_STORAGE_COMMENT Then
            MsgBox "El comentario para el medio es muy largo. " & vbCrLf & "Será recortado."
            gs_MediaComment = Mid(gs_MediaComment, 1, DB_MAX_LEN_STORAGE_COMMENT)
        End If
        
        gfnc_ParseString gs_MediaComment, gs_MediaComment
    
        '-------------------------------------
        ' insertamos registro
        '
        query = "INSERT INTO storage (id_storage, id_category, id_storage_type, name, "
        query = query & "label, serial, comment, active)"
        query = query & " VALUES("
        query = query & Trim(str(ml_idStorage)) & ", "
        query = query & Trim(str(gl_StorageCategory)) & ", "
        query = query & Trim(str(gl_StorageType)) & ", '"
        query = query & MediaName & "', '"
        query = query & Etiqueta & "', '"
        query = query & Mid(Serie, 1, 4) & "-" & Mid(Serie, 5) & "', '"
        query = query & gs_MediaComment & "', 1)"
        cd.CommandText = query
        cd.Execute
    
        cnTransaction.CommitTrans   ' dejamos abierta la conexion de transaccion
        bTransaccionIniciada = False
    
        AddMediaToDB = True
        
    Else
        MsgBox "Error insertando medio de almacenamiento", vbExclamation, "Error DB"
        AddMediaToDB = False
    End If

    Exit Function
    
Handler:

    Select Case Err.Number
    
        Case 0
            '
        Case Else
            If bTransaccionIniciada Then
                cnTransaction.RollbackTrans
                gsub_CerrarConexionTransaccion
            End If
            MsgBox Err.Description, vbCritical, "AddMediaToDB()"

            If rs.state = adStateOpen Then
                rs.Close
            End If
        
            AddMediaToDB = False
    End Select

End Function

Private Sub rchtxt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'el control richtext no refresca bien
    Me.ZOrder 0
    Me.Refresh
End Sub

Private Sub LoadScript()
    
    On Error GoTo Handler
    
    gb_ScriptActivated = False
    
    '---------------------------------------------------
    ' crear objeto script e Inicializar script
    Set ScriptControl = New clsScript
    Set mo_Modulo = ScriptControl.objScript.Modules("Global")
    ScriptControl.objScript.Reset
   
    '---------------------------------------------------
    ' registrar objeto de interfaz
    Set mo_IScript = New clsIScript
    Set mo_IScript.rchtxtResults = rchtxt
    ScriptControl.objScript.AddObject "IScript", mo_IScript, True
    
    '--------------------------------------
    ' Agregamos el codigo script
    rchtxtScript.LoadFile gs_ScriptFile, rtfText    '<- throws error when the file does't exist
    
    On Error Resume Next
    mo_Modulo.AddCode rchtxtScript.text
    
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
        gb_ScriptActivated = True
    End If
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbExclamation, "Error cargando script"
End Sub

Private Sub CallScriptFunction(ByVal strFunctionName As String)
    
    gb_ScriptActivated = False
    
    On Error Resume Next
    
    '-----------------------------------
    ' Ejecutamos el codigo de la funcion
    mo_Modulo.Run strFunctionName
    
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
        '-------------------------------------
        ' El codigo fue ejecutado exitosamente
        gb_ScriptActivated = True
        Exit Sub
    End If
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbExclamation, "Error cargando script"
End Sub

