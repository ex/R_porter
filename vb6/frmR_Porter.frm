VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{40EF20B1-7EC5-11D8-95A1-9655FE58C763}#3.0#0"; "exLightButton.ocx"
Object = "{40EF20CB-7EC5-11D8-95A1-9655FE58C763}#3.0#0"; "exLightLabel.ocx"
Object = "{40EF20E1-7EC5-11D8-95A1-9655FE58C763}#3.0#0"; "exSplit.ocx"
Begin VB.Form frmR_Porter 
   Caption         =   "R_porter 1.2"
   ClientHeight    =   5985
   ClientLeft      =   300
   ClientTop       =   795
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmR_Porter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgPlugins 
      Left            =   3045
      Top             =   4590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmR_Porter.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmR_Porter.frx":085C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmR_Porter.frx":0DAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   6420
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraExplorar 
      BorderStyle     =   0  'None
      Height          =   5700
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox pctFondoBorde 
         AutoSize        =   -1  'True
         BackColor       =   &H00FA6124&
         BorderStyle     =   0  'None
         FillColor       =   &H00FA6124&
         Height          =   5685
         Left            =   0
         Picture         =   "frmR_Porter.frx":1300
         ScaleHeight     =   5685
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   0
         Width           =   1875
         Begin exLightButton.ocxLightButton ELBAyuda 
            Height          =   585
            Left            =   75
            TabIndex        =   17
            Top             =   3330
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmR_Porter.frx":8E8A
            PictureOFF      =   "frmR_Porter.frx":C344
            PictureOK       =   "frmR_Porter.frx":F7FE
            MouseIcon       =   "frmR_Porter.frx":12CB8
            MousePointer    =   99
         End
         Begin exLightButton.ocxLightButton ELBSalirE 
            Height          =   585
            Left            =   75
            TabIndex        =   18
            Top             =   4080
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmR_Porter.frx":12E1A
            PictureOFF      =   "frmR_Porter.frx":162D4
            PictureOK       =   "frmR_Porter.frx":1978E
            MouseIcon       =   "frmR_Porter.frx":1CC48
            MousePointer    =   99
         End
         Begin exLightButton.ocxLightButton ELBOpciones 
            Height          =   585
            Left            =   75
            TabIndex        =   19
            Top             =   2580
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmR_Porter.frx":1CDAA
            PictureOFF      =   "frmR_Porter.frx":20264
            PictureOK       =   "frmR_Porter.frx":2371E
            MouseIcon       =   "frmR_Porter.frx":26BD8
            MousePointer    =   99
         End
         Begin exLightButton.ocxLightButton ELBIniciar 
            Height          =   585
            Left            =   75
            TabIndex        =   20
            Top             =   1815
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1032
            Activate        =   -1  'True
            PictureON       =   "frmR_Porter.frx":26D3A
            PictureOFF      =   "frmR_Porter.frx":2A1F4
            PictureOK       =   "frmR_Porter.frx":2D6AE
            MouseIcon       =   "frmR_Porter.frx":30B68
            MousePointer    =   99
         End
      End
   End
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5730
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      SimpleText      =   "Listo"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
      EndProperty
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
   Begin VB.Frame fraPExplorar 
      Height          =   4500
      Left            =   1950
      TabIndex        =   0
      Top             =   1200
      Width           =   5175
      Begin exSplit.SplitRegion SplitRegionVertical 
         Height          =   3525
         Left            =   120
         Top             =   885
         Width           =   4920
         _ExtentX        =   5715
         _ExtentY        =   5556
         FirstControl    =   "fraUnidades"
         SecondControl   =   "fraBuscarDir"
         FirstControlMinSize=   1500
         SecondControlMinSize=   1650
         SplitPercent    =   52
         SplitterBarVertical=   -1  'True
         SplitterBarThickness=   90
         MouseIcon       =   "frmR_Porter.frx":30CCA
         MousePointer    =   99
      End
      Begin VB.Frame fraBuscarDir 
         Caption         =   "Directorios:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   2723
         TabIndex        =   11
         Top             =   885
         Width           =   2317
         Begin VB.CheckBox chkIncluirSubDir 
            Caption         =   "Incluir Subdirectorios"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   3225
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.TextBox txtDirActual 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2880
            Width           =   2055
         End
         Begin VB.DirListBox Dir1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2130
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   2070
         End
      End
      Begin VB.Frame fraUnidades 
         Caption         =   "Unidades:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   120
         TabIndex        =   8
         Top             =   885
         Width           =   2513
         Begin MSComctlLib.ImageList imgDrives 
            Left            =   285
            Top             =   2475
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   16777215
            ImageWidth      =   18
            ImageHeight     =   18
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmR_Porter.frx":30E2C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmR_Porter.frx":313E8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmR_Porter.frx":319A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmR_Porter.frx":31F60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmR_Porter.frx":3251C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvwDrives 
            Height          =   2805
            Left            =   135
            TabIndex        =   9
            Top             =   315
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   4948
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            PictureAlignment=   3
            _Version        =   393217
            SmallIcons      =   "imgDrives"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "UNIDADES DISPONIBLES:"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label lblDrives 
            Caption         =   "Ninguna unidad seleccionada"
            Height          =   210
            Left            =   150
            TabIndex        =   10
            Top             =   3120
            Width           =   2115
         End
      End
      Begin VB.Frame fraSeleccion 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   135
         Width           =   1485
         Begin VB.OptionButton optUnidades 
            Caption         =   "Por &Unidades"
            Height          =   210
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   1320
         End
         Begin VB.OptionButton optDirectorio 
            Caption         =   "Por &Directorio"
            Height          =   210
            Left            =   90
            TabIndex        =   6
            Top             =   435
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.Frame fraCorreo 
         Height          =   735
         Left            =   1695
         TabIndex        =   4
         Top             =   135
         Width           =   3345
         Begin exLightLabel.ocxLightLabel ellCorreo 
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   150
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   476
            MousePointer    =   99
            MouseIcon       =   "frmR_Porter.frx":32AD8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "exe_q_tor@hotmail.com"
            ColorOFF        =   8388608
            ColorON         =   16711680
            ColorOK         =   16777215
            ColorDOWN       =   12672845
            Activate        =   -1  'True
            Timewait        =   150
            KTimeOK         =   1
         End
         Begin exLightLabel.ocxLightLabel ellWeb 
            Height          =   255
            Left            =   105
            TabIndex        =   23
            Top             =   420
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   450
            MousePointer    =   99
            MouseIcon       =   "frmR_Porter.frx":32C3A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "http://www.geocities.com/planeta_dev"
            ColorOFF        =   8388608
            ColorON         =   16711680
            ColorOK         =   16777215
            ColorDOWN       =   12672845
            Activate        =   -1  'True
            Timewait        =   150
            KTimeOK         =   1
         End
      End
   End
   Begin VB.Frame fraCabecera 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   1965
      TabIndex        =   16
      Top             =   0
      Width           =   5145
      Begin exLightButton.ocxLightButton ELBAcercaDe 
         Height          =   1065
         Left            =   3915
         TabIndex        =   21
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1879
         Activate        =   -1  'True
         PictureON       =   "frmR_Porter.frx":32D9C
         PictureOFF      =   "frmR_Porter.frx":34DE7
         PictureOK       =   "frmR_Porter.frx":37043
         MouseIcon       =   "frmR_Porter.frx":38DCA
         MousePointer    =   99
      End
      Begin VB.Image imgBanner 
         Height          =   1050
         Left            =   0
         Picture         =   "frmR_Porter.frx":38F2C
         Top             =   45
         Width           =   3915
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuIniciarExploracion 
         Caption         =   "&Iniciar Exploracion"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuB0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnidades 
         Caption         =   "Explorar &Unidades"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuDirectorio 
         Caption         =   "Explorar &Directorio"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScript 
         Caption         =   "S&cript"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuC1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Datos"
      Begin VB.Menu mnuControlRegistros 
         Caption         =   "Ver &registros"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuExploreDB 
         Caption         =   "E&xplorar DB"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuB20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConectarBD 
         Caption         =   "Conectar a la BD"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuAvanzado 
         Caption         =   "Avanzado"
         Begin VB.Menu mnuSQL 
            Caption         =   "Ejecutar comando SQL"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuEstructura 
            Caption         =   "&Ver estructura de la BD"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnuB21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminDSN 
         Caption         =   "Administrar &DSN"
      End
      Begin VB.Menu mnuCrearDSN 
         Caption         =   "&Crear DSN"
      End
      Begin VB.Menu mnuB22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirBD 
         Caption         =   "&Editar BD externamente"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOpcionesGen 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Opciones generales"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBD 
         Caption         =   "O&pciones de anexión a la BD"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMP3 
         Caption         =   "Opciones &MP3"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuColores 
         Caption         =   "Opciones del &reporte"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuControlPlugins 
         Caption         =   "Control de plugins"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuRegistrarPlugin 
         Caption         =   "&Registrar plugin"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuPluginList 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "&Contenido"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuMasAyuda 
         Caption         =   "&Mas ayuda..."
      End
      Begin VB.Menu mnuB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnulvwdrives 
      Caption         =   "mnuPopUpDrives"
      Visible         =   0   'False
      Begin VB.Menu mnuNingunDrive 
         Caption         =   "&Ninguno"
      End
      Begin VB.Menu mnuTodosDrives 
         Caption         =   "Seleccionar &Todos"
      End
   End
End
Attribute VB_Name = "frmR_Porter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' Formulario principal de R_Porter
'               Este programa realiza un reporte de los archivos
'               que encuentra en una unidad o directorio especificado
'               permitiendo imprimirlo y buscar por tipos de archivo
'               ademas de dar algunas opciones de reporte
'               Ademas usa controles ActiveX: ELB, Split y Flash
'*******************************************************************************
' Modificado:   Esau R.O. Julio 2004
'               - Agregado soporte de plugins
'*******************************************************************************
Option Explicit

Private zo_plugins() As Object  'each plugin is added to this array one by one
                                'we store the puglins index in this array along with the
                                'startup argument it requested to be send in the menu.tag for
                                'each entry it added. This way, we know who to call startup for
                                'and what startup argument they are expecting.
                                
Private zs_plugins_name() As String      'nombre de cada plugin
Private zs_plugins_desc() As String      'descripcion de cada plugin
Private zb_plugins_active() As Boolean   'si el plugin esta activo o no
Private zl_plugins_id_menu() As Long     'id de menu
Private mb_plugin_add_to_flex As Boolean
Public intNumPlugins  As Integer         'numero de plugins cargados

Private m_objAccess As Object

Private mb_OptionalDriveSet As Boolean
                               
Private m_ScaleHeight As Integer
Private m_ScaleWidth As Integer
Private m_Height As Integer
Private m_Width As Integer
Private m_fracab_width As Integer
Private m_frapex_width As Integer
Private m_frapex_height As Integer
Private m_fracorr_width As Integer
Private m_elbacerca_left As Integer
Private m_pic_height As Integer
Private m_fra_height As Integer
Private m_split_width As Integer
Private m_split_height As Integer
Private m_unid_height As Integer
Private m_labunid_top As Integer
Private m_dirs_height As Integer
Private m_txtdir_top As Integer
Private m_chkdir_top As Integer

Private m_FormLoading As Boolean    ' extraño error reposicionando controles ahora que agregue resize
                                    ' aveces el statusbar esta mal colocado

'*******************************************************************************
' INICIALIZACION FORMULARIO
'*******************************************************************************
Private Sub Form_Load()
    
    Dim clx As clsCrypto
    On Error Resume Next
    
    mb_OptionalDriveSet = False
    
    m_FormLoading = True
   
    m_pic_height = pctFondoBorde.height
    m_fra_height = fraExplorar.height
    
    m_fracab_width = fraCabecera.width
    m_elbacerca_left = ELBAcercaDe.Left
    
    m_frapex_height = fraPExplorar.height
    m_frapex_width = fraPExplorar.width
    
    m_fracorr_width = fraCorreo.width
    
    m_split_height = SplitRegionVertical.height
    m_split_width = SplitRegionVertical.width
    
    m_unid_height = lvwDrives.height
    m_labunid_top = lblDrives.Top
    
    m_dirs_height = Dir1.height
    m_txtdir_top = txtDirActual.Top
    m_chkdir_top = chkIncluirSubDir.Top
    
    Set clx = New clsCrypto
    clx.SetCod 159
    '--------------------------------------------------------------------
    'ellCorreo.Caption = clx.Encrypt("exeqtor@gmail.com")
    '--------------------------------------------------------------------
    ellCorreo.Caption = clx.Decrypt("-\-_A{%tI8""pF#e{8")
    
    '--------------------------------------------------------------------
    'ellWeb.Caption = clx.Encrypt("http://www.geocities.com/planeta_dev")
    '--------------------------------------------------------------------
    ellWeb.Caption = clx.Decrypt("WAAmi11jjj#I-{epAp-3#e{81mF""*-A"">~-x")
    Set clx = Nothing

    strExplorarU = ""
    strExplorarD = ""
    
    ExplorarDir = True
    txtDirActual = Dir1.Path
    
    IniciarListDrives
    NumDrivesSel = 0
    
    stbarReady
    
    LoadPlugins
    
    If gb_IncluirSubdirectorios Then
        chkIncluirSubDir.value = vbChecked
    Else
        chkIncluirSubDir.value = vbUnchecked
    End If
    
    gb_FrmPluginsActive = True
    
End Sub

Private Sub Form_Activate()
    If m_FormLoading Then
        m_FormLoading = False
        m_Height = Me.height
        m_Width = Me.width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim frm As Form
    Dim k As Integer
    On Error Resume Next
    
    '(FLAW] si no va esta comprobacion el ejecutable se cuelga al salir, pero no el IDE
    If intNumPlugins > 0 Then
        For k = 0 To UBound(zo_plugins)
            Set zo_plugins(k) = Nothing
        Next k
    End If
    
    For Each frm In Forms
        If frm.Name = "frmPluginControl" Then
            Unload frm
        End If
    Next
    
    If Forms.Count > 1 Then
        If vbYes = MsgBox("¿Deseas cerrar todos los formularios abiertos?", vbQuestion + vbYesNo, "Salir") Then
            For Each frm In Forms
                If frm.Name <> "frmR_Porter" Then
                    Unload frm
                End If
            Next
            '(BUG?) algo no se esta descargando...
            End
        End If
    End If
    
    gb_FrmPluginsActive = False

End Sub

'*******************************************************************************
' EVENTOS
'*******************************************************************************
Private Sub Form_Resize()
    
    Dim height As Integer
    Dim width As Integer
    On Error Resume Next
    
    If m_FormLoading Then Exit Sub
    
    If Me.width < m_Width Then Me.width = m_Width
    If Me.height < m_Height Then Me.height = m_Height
    
    height = Me.height - m_Height
    width = Me.width - m_Width
    
    fraExplorar.height = m_fra_height + height
    pctFondoBorde.height = m_pic_height + height
    
    fraCabecera.width = m_fracab_width + width
    ELBAcercaDe.Left = m_elbacerca_left + width
    
    fraPExplorar.width = m_frapex_width + width
    fraPExplorar.height = m_frapex_height + height
    
    fraCorreo.width = m_fracorr_width + width
    
    SplitRegionVertical.width = m_split_width + width
    SplitRegionVertical.height = m_split_height + height
    
    SplitRegionVertical_RepositionSplit
    
    lvwDrives.height = m_unid_height + height
    lblDrives.Top = m_labunid_top + height
    
    Dir1.height = m_dirs_height + height
    txtDirActual.Top = m_txtdir_top + height
    chkIncluirSubDir.Top = m_chkdir_top + height
    
End Sub

'*******************************************************************************
' De los ELB
'*******************************************************************************
Private Sub ELBAyuda_OnActivate()
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Haga click para ver la ayuda."
End Sub

Private Sub ELBAyuda_Click()
Dim hins As Long
Dim cad As String
Dim clx As clsCrypto
    
    On Error GoTo Handler
    
    hins = ShellExecute(Me.hWnd, "open", App.Path & "\help\ayuda.htm", vbNull, vbNull, 0)
    
    If (hins < 33) Then
        Set clx = New clsCrypto
        clx.SetCod 1352
        '--------------------------------------------------------------------------
        'cad = clx.Encrypt("http://www.geocities.com/planeta_dev/spa/r_porter.htm")
        '--------------------------------------------------------------------------
        cad = clx.Decrypt("?rr'~RRggg6eJEq[r[J26qE{R'P5)Jr5`ZJAR2'5R:`'E:rJ:6?r{")
        Set clx = Nothing
        
        hins = ShellExecute(Me.hWnd, "open", cad, vbNull, vbNull, 0)
    End If
    
Handler:
    Exit Sub
End Sub

Private Sub ELBOpciones_OnActivate()
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Haga click para ver el cuadro de dialogo de opciones avanzadas."
End Sub

Private Sub ELBOpciones_Click()
    frmOptions.Show vbModal
End Sub

Private Sub ELBIniciar_OnActivate()
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Haga click para iniciar la búsqueda por unidades o directorio."
End Sub

Private Sub ELBIniciar_Click()
    Iniciar_Exploracion
End Sub

Private Sub ELBSalirE_OnActivate()
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Haga click para salir del programa."
End Sub

Private Sub ELBSalirE_Click()
    Unload Me
End Sub

Private Sub ELBAcercaDe_OnActivate()
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Haga click para ver información del programa."
End Sub

Private Sub ELBAcercaDe_Click()
    frmAbout.Show vbModal
End Sub

'*******************************************************************************
' Del split
'*******************************************************************************
' Para cambiar el tamaño de los controles dentro del los frames del Split
'*******************************************************************************
Private Sub SplitRegionVertical_RepositionSplit()
    Dim width1 As Integer
    Dim width2 As Integer

    width1 = fraUnidades.width - 250
    width2 = fraBuscarDir.width - 250
    
    If width2 < 1900 Then
        chkIncluirSubDir.Caption = "Subdirectorios"
    Else
        chkIncluirSubDir.Caption = "Incluir Subdirectorios"
    End If
    
    lvwDrives.width = fraUnidades.width - 250
    lblDrives.width = width1 - 50
    
    Dir1.width = width2
    Drive1.width = width2
    txtDirActual.width = width2
    chkIncluirSubDir.width = width2
End Sub

'*******************************************************************************
' OptionButton
'*******************************************************************************
Private Sub optDirectorio_Click()
    ExplorarDir = True
    mnuDirectorio.Checked = True
    mnuUnidades.Checked = False
End Sub

Private Sub optUnidades_Click()
    ExplorarDir = False
    mnuDirectorio.Checked = False
    mnuUnidades.Checked = True
End Sub


'*******************************************************************************
' Menu
'*******************************************************************************
Private Sub mnuAbrirBD_Click()

    On Error GoTo Handler
    
    If gb_DBConexionOK Then
    
        Set m_objAccess = GetObject(cn.DefaultDatabase & ".mdb")
        m_objAccess.Visible = True
        
    Else
        gsub_ShowMessageNoConection
    End If
    
    Exit Sub
    
Handler:
    
    Select Case Err.Number
        Case -2147467259
            'error de automatizacion
            MsgBox "No se pudo abrir la base de datos", vbExclamation, "Error al abrir base de datos"
        Case 3704, 432
            'cuando no esta instalado Access
            MsgBox "Para poder abrir la base de datos" & vbCrLf & "necesitas tener instalado MS Access.", vbExclamation, "Error al abrir base de datos"
        Case 2455
            'cuando ya esta abierto Access
            MsgBox "Parece que la base de datos activa" & vbCrLf & "ya se encuentra abierta.", vbExclamation, "Aviso"
        Case Else
            MsgBox Err.Description, vbExclamation, "Error al abrir base de datos"
    End Select
    
End Sub

Private Sub mnuAdminDSN_Click()
Dim ret As Long
    
    On Error Resume Next
    ret = SQLManageDataSources(Me.hWnd)
    
End Sub

Private Sub mnuCrearDSN_Click()
Dim ret As Long
    
    On Error Resume Next
    ret = SQLCreateDataSource(Me.hWnd, "R_porter")
    
End Sub

Private Sub mnuControlRegistros_Click()
    
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            frmDataControl.Show
        Else
            gsub_ShowMessageWrongDB
        End If
    Else
        gsub_ShowMessageNoConection
    End If
    
End Sub

Private Sub mnuOptions_Click()
    ELBOpciones_Click
End Sub

Private Sub mnuBD_Click()
    
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            frmOptions.ssTab.Tab = 1
            frmOptions.Show vbModal
        Else
            gsub_ShowMessageWrongDB
        End If
    Else
        gsub_ShowMessageNoConection
    End If

End Sub

Private Sub mnuColores_Click()
    frmOptions.ssTab.Tab = 3
    frmOptions.Show vbModal
End Sub

Private Sub mnuMP3_Click()
    
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            frmOptions.ssTab.Tab = 2
            frmOptions.Show vbModal
        Else
            gsub_ShowMessageWrongDB
        End If
    Else
        gsub_ShowMessageNoConection
    End If
    
End Sub

Private Sub mnuEstructura_Click()
    
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
    Else
        frmReporte.Show
    End If

End Sub

Private Sub mnuSQL_Click()
    
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
    Else
        frmSQL.Show
    End If

End Sub

Private Sub mnuAcercaDe_Click()
    ELBAcercaDe_Click
End Sub

Private Sub mnuDirectorio_Click()
    optDirectorio.value = True
    mnuDirectorio.Checked = True
    mnuUnidades.Checked = False
    ExplorarDir = True
End Sub

Private Sub mnuUnidades_Click()
    optUnidades.value = True
    mnuUnidades.Checked = True
    mnuDirectorio.Checked = False
    ExplorarDir = False
End Sub

Private Sub mnuContenido_Click()
    ELBAyuda_Click
End Sub

Private Sub mnuMasAyuda_Click()
    '===================================================
    Dim SysPath As String
    '===================================================

    On Error Resume Next
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "SystemRoot", SysPath) Then
        'Probar a extraer la informacion del registro del sistema...
        Shell SysPath & "\hh.exe -r " & SysPath & "\help\windows.chm", vbNormalFocus
    Else
        'para XP...
        If GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "SystemRoot", SysPath) Then
            'Probar a extraer la informacion del registro del sistema...
            Shell SysPath & "\hh.exe -r " & SysPath & "\help\windows.chm", vbNormalFocus
        Else
            'Metodo antiguo (falla si el S.O. activo no está instalado en C)
            Shell "c:\windows\hh.exe -r c:\windows\help\windows.chm", vbNormalFocus
        End If
    End If
    
End Sub

Private Sub mnuIniciarExploracion_Click()
    ELBIniciar_Click
End Sub

Private Sub mnuTodosDrives_Click()
Dim itemx As ListItem
Dim i As Integer
    strExplorarU = ""
    For i = 1 To lvwDrives.ListItems.Count
        Set itemx = lvwDrives.ListItems.Item(i)
        itemx.Checked = True
        strExplorarU = strExplorarU & Left$(itemx.Tag, 1)
    Next i
    NumDrivesSel = lvwDrives.ListItems.Count
    Actualizar_lblDrives
    mnuUnidades_Click
End Sub

Private Sub mnuNingunDrive_Click()
Dim itemx As ListItem
Dim i As Integer
    For i = 1 To lvwDrives.ListItems.Count
        Set itemx = lvwDrives.ListItems.Item(i)
        itemx.Checked = False
    Next i
    NumDrivesSel = 0
    strExplorarU = ""
    Actualizar_lblDrives
    mnuDirectorio_Click
End Sub

Private Sub mnuSalir_Click()
    ELBSalirE_Click
End Sub

Private Sub mnuConectarBD_Click()
    frmConexion.Show vbModal
End Sub

Private Sub mnuScript_Click()
    frmScript.Show vbModeless
End Sub

Private Sub mnuRegistrarPlugin_Click()
    RegisterNewPlugin
End Sub

Private Sub mnuControlPlugins_Click()
    frmPluginControl.Show vbModeless
End Sub

Private Sub mnuExploreDB_Click()
    Load frmDataExplorer
    frmDataExplorer.Show vbModeless
End Sub

'*******************************************************************************
' Mouse Move
'*******************************************************************************
Private Sub pctFondoBorde_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraBuscarDir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraCorreo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraPExplorar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraCabecera_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraSeleccion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbar.Style = sbrNormal
    stbar.Panels.Item(1) = "Seleccione busqueda por directorios o por unidades"
End Sub

Private Sub fraUnidades_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub fraExplorar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarDrives = False
    If stbarDir = False Then
        stbar.Style = sbrNormal
        stbar.Panels.Item(1) = "Escoja una unidad y haga doble click en una carpeta (o pulse ENTER ) para seleccionarla."
        stbarDir = True
    End If
End Sub

Private Sub lvwDrives_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarDir = False
    If stbarDrives = False Then
        stbar.Style = sbrNormal
        stbar.Panels.Item(1) = "Aqui puede seleccionar las unidades que quiera explorar con INICIAR EXPLORACION"
        stbarDrives = True
    End If
End Sub

Private Sub stbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stbarReady
End Sub

'*******************************************************************************
' Controles LightLabel:
'*******************************************************************************
Private Sub ellCorreo_Click()
Dim cad As String
Dim clx As clsCrypto
    
    On Error Resume Next
    
    Set clx = New clsCrypto
    clx.SetCod 1179
    '---------------------------------------------------
    'cad = clx.Encrypt("mailto:exeqtor@gmail.com")
    '---------------------------------------------------
    cad = clx.Decrypt("K@+Cf[/ F kf[v_oK@+C\0[K")
    Set clx = Nothing
    
    ShellExecute Me.hWnd, "open", cad, vbNull, vbNull, 0

End Sub

Private Sub ellWeb_Click()
Dim cad As String
Dim clx As clsCrypto
    
    On Error Resume Next
    
    Set clx = New clsCrypto
    clx.SetCod 9204
    '--------------------------------------------------------------------
    'cad = clx.Encrypt("http://www.geocities.com/planeta_dev/spa/r_porter.htm")
    '--------------------------------------------------------------------
    cad = clx.Decrypt("%&&e188>>>0|l]\-&-l}0\]M8eELUl&L<dl68}eL8u<e]u&lu0%&M")
    Set clx = Nothing
        
    ShellExecute Me.hWnd, "open", cad, vbNull, vbNull, 0
    
End Sub

Private Sub ellCorreo_OnActivate()
    ellCorreo.Font.Bold = True
    ellCorreo.Font.Underline = True
End Sub

Private Sub ellCorreo_OnDeactivate()
    ellCorreo.Font.Bold = False
    ellCorreo.Font.Underline = False
End Sub

Private Sub ellWeb_OnActivate()
    ellWeb.Font.Bold = True
    ellWeb.Font.Underline = True
End Sub

Private Sub ellWeb_OnDeactivate()
    ellWeb.Font.Bold = False
    ellWeb.Font.Underline = False
End Sub

'*******************************************************************************
' Controles de unidades de disco y directorio
'*******************************************************************************
Private Sub Drive1_Change()
On Error GoTo DriveHandler
    
    Dir1.Path = Mid(Drive1.Drive, 1, 2) & "\"
    txtDirActual.text = Dir1.Path
    Exit Sub
DriveHandler:
    Drive1.Drive = Dir1.Path
    txtDirActual.text = Dir1.Path
    txtDirActual.SetFocus
    txtDirActual.SelStart = 260
    Exit Sub
End Sub

Private Sub Dir1_Change()
On Error GoTo FIN
    strExplorarD = Dir1.Path
    txtDirActual = Dir1.Path
    txtDirActual.SetFocus
    txtDirActual.SelStart = 260
    mnuDirectorio_Click
FIN:
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
    End If
End Sub

'*******************************************************************************
' ListView de unidades
'*******************************************************************************
Private Sub lvwDrives_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim str As String
Dim pos As Long
    If Item.Checked Then
        NumDrivesSel = NumDrivesSel + 1
        strExplorarU = strExplorarU & Left$(Item.Tag, 1)
    Else
        NumDrivesSel = NumDrivesSel - 1
        pos = InStr(strExplorarU, Left$(Item.Tag, 1))
        If pos = 1 Then
            strExplorarU = Mid$(strExplorarU, 2)
        Else
            If pos = Len(strExplorarU) Then
                strExplorarU = Mid$(strExplorarU, 1, pos - 1)
            Else
                strExplorarU = Mid$(strExplorarU, 1, pos - 1) & Mid$(strExplorarU, pos + 1)
            End If
        End If
    End If
    Actualizar_lblDrives
    mnuUnidades_Click
End Sub

Private Sub lvwDrives_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnulvwdrives
    End If
End Sub

'*******************************************************************************
' CHECKBOX
'*******************************************************************************
Private Sub chkIncluirSubDir_Click()
    If chkIncluirSubDir.value = vbChecked Then
        gb_IncluirSubdirectorios = True
    Else
        gb_IncluirSubdirectorios = False
    End If
End Sub


'*******************************************************************************
' FUNCIONES GENERALES
'*******************************************************************************

'*******************************************************************************
' Inicializacion
'*******************************************************************************
Private Sub IniciarListDrives()

    Dim DirBuffer As String * MAX_TAM_DRIVES
    Dim LngBuffer As Long
    Dim DirCad As String
    Dim chars As String
    Dim char As Long
    Dim k As Integer

    LngBuffer = GetLogicalDriveStrings(MAX_TAM_DRIVES, DirBuffer)
    If LngBuffer > MAX_TAM_DRIVES Then
        MsgBox "Demasiadas unidades logicas.", vbCritical, "Error"
    End If
    k = 0
    Do
        DirCad = ""
        Do
            k = k + 1
            chars = Mid$(DirBuffer, k, 1)
            char = Asc(chars)
            If char = 0 Then
                Exit Do
            End If
            DirCad = DirCad & chars
        Loop
        PonerIcono (DirCad)
        chars = Mid(DirBuffer, k + 1, 1)
        char = Asc(chars)
        If char = 0 Then
            Exit Do
        End If
    Loop
End Sub

Private Sub PonerIcono(Name As String)
    '===================================================
    Dim J As Long, L As Long, L1 As Long, L2 As Long, L3 As Long, L4 As Long
    Dim itmX As ListItem
    Dim mName As String
    Dim NameDrv As String * 15
    '===================================================
    
    J = GetDriveType(Name)
    J = J - 1
    
    mName = UCase$(Mid$(Name, 1, 2))
    
    Select Case J
    
        Case 1:
            Set itmX = lvwDrives.ListItems.Add
            itmX.SmallIcon = 1
            itmX.Tag = mName
            If mName = "A:" Or mName = "B:" Then
              itmX.text = "Disco de 3½ (" & mName & ")"
            Else
              itmX.text = "Disco extraible (" & mName & ")"
            End If
            
        Case 2:
            Set itmX = lvwDrives.ListItems.Add
            itmX.SmallIcon = 2
            itmX.Tag = mName
            L = GetVolumeInformation(Name, NameDrv, 15, L1, L2, L3, vbNullString, L4)
            If Asc(NameDrv) <> 0 Then
             itmX.text = NameDrv
             itmX.text = itmX.text & " (" & mName & ")"
            Else
             itmX.text = "(" & mName & ")"
            End If
            
        Case 3:
            Set itmX = lvwDrives.ListItems.Add
            itmX.SmallIcon = 3
            itmX.Tag = mName
            itmX.text = "Disco Remoto (" & mName & ")"
            
        Case 4:
            Set itmX = lvwDrives.ListItems.Add
            itmX.SmallIcon = 4
            itmX.Tag = mName
            itmX.text = "CD-ROM (" & mName & ")"
            
            If Not mb_OptionalDriveSet Then
                gs_OptionalDrive = UCase(Left(mName, 1))
                mb_OptionalDriveSet = True
            End If
            
        Case 5:
            Set itmX = lvwDrives.ListItems.Add
            itmX.SmallIcon = 5
            itmX.Tag = mName
            itmX.text = "Disco en Memoria (" & mName & ")"
            
    End Select
End Sub

'*******************************************************************************
' Otros
'*******************************************************************************
Private Sub Iniciar_Exploracion()
Dim mfrmExplorar As frmExplorar
Dim k As Long
Dim J As Integer

    On Error GoTo Handler
    
    If ExplorarDir And strExplorarD = "" Then
        strExplorarD = Dir1.Path
    End If
    
    If Not ExplorarDir And strExplorarU = "" Then
        MsgBox "Debe seleccionar primero alguna unidad disponible.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Open "C:\ex.$$$" For Output As #1
    
    If ExplorarDir Then
        If Right$(strExplorarD, 1) = "\" Then
            strExplorarD = Mid$(strExplorarD, 1, 2)
        End If
        Print #1, "Dir"
        Print #1, strExplorarD
        Print #1, str$(chkIncluirSubDir.value)
    Else
        k = Len(strExplorarU)
        Print #1, str$(k)
        Print #1, str$(chkIncluirSubDir.value)
        J = 1
        Do
            Print #1, Mid$(strExplorarU, J, 1)
            J = J + 1
        Loop While J <= k
    End If
    
    Close #1
    
    Set mfrmExplorar = New frmExplorar
    mfrmExplorar.Show
    mfrmExplorar.Explorar
    Set mfrmExplorar = Nothing
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "Iniciar_Exploracion()"
End Sub

Private Sub Actualizar_lblDrives()
    Select Case NumDrivesSel
    Case 0
        lblDrives = "Ninguna unidad seleccionada."
    Case 1
        lblDrives = "Una unidad seleccionada."
    Case Else
        lblDrives = "(" & str$(NumDrivesSel) & " ) unidades seleccionadas."
    End Select
End Sub

Private Sub stbarReady()
    If stbar.Style = sbrNormal Then
        stbar.Style = sbrSimple
        stbarDir = False
        stbarDrives = False
    End If
End Sub

Private Sub LoadPlugins()
    '===================================================
    '(TODO):    quitar limite de memoria
    Dim strbuff As String * 256
    Dim ret As Long
    Dim plugs() As String
    Dim k As Integer
    Dim PluginID As String
    '===================================================
    On Error GoTo Handler
    
    ret = GetPrivateProfileString("Add-Ins32", vbNullString, "ERROR", strbuff, Len(strbuff), App.Path & "\R_porter.ini")
    If (ret > 0) Then
        plugs = gfnc_ZGetStrings(strbuff, Len(strbuff))
        If Not gfnc_ZIsEmpty(plugs) Then
            intNumPlugins = 0
            If plugs(0) <> "ERROR" Then
                For k = 0 To UBound(plugs)
                    ret = GetPrivateProfileString("Add-Ins32", plugs(k), "ERROR", strbuff, Len(strbuff), App.Path & "\R_porter.ini")
                    If (ret > 0) And (strbuff <> "ERROR") Then
                        PluginID = gfnc_GetFileNameWithoutExt(plugs(k)) & ".exPlugin"
                        mb_plugin_add_to_flex = False
                        TryLoadPlugin PluginID, plugs(k)
                    End If
                Next k
                Exit Sub 'Se cargo plugins normalmente
            Else
                ' en caso de error crear el archivo ini
                gsub_Create_R_porter_ADDIN
            End If
        End If
    End If
    
    Exit Sub
Handler:
    MsgBox Err.Description, vbExclamation, "LoadPlugins"
End Sub

Private Function TryLoadPlugin(ByVal plugin As String, ByVal file_plugin As String) As Boolean
    '===================================================
    Dim ret As Long
    '===================================================
    On Error GoTo Handler
    
    TryLoadPlugin = False
    ReDim Preserve zo_plugins(intNumPlugins)

    On Error Resume Next
    Set zo_plugins(intNumPlugins) = CreateObject(plugin)
    
    If Err.Number = 0 Then
        
        zo_plugins(intNumPlugins).SetHost Me, file_plugin
        intNumPlugins = intNumPlugins + 1
        TryLoadPlugin = True
    Else
        MsgBox "No se pudo reconocer el plugin:" & vbCrLf & _
               "[" & file_plugin & "]", vbExclamation, "Error cargando plugin"
        ' quitar el plugin del registro
        ret = WritePrivateProfileString("Add-Ins32", file_plugin, vbNullString, App.Path & "\R_porter.ini")
    End If
    
    Exit Function

Handler:    MsgBox Err.Description, vbExclamation, "TryLoadPlugin"
End Function

Public Function RegisterPlugin(ByVal intMenu As Integer, _
                               ByVal strMenuName As String, _
                               ByVal strDescription As String, _
                               ByVal intStartupArgument As Integer, _
                               ByVal strPluginName As String)
    '===================================================
    Dim k As Integer
    Dim id_plugin As Long
    '===================================================
    On Error Resume Next
    
    id_plugin = UBound(zo_plugins)
    
    k = mnuPluginList.Count
    Load mnuPluginList(k)
    mnuPluginList(k).Caption = strMenuName
    mnuPluginList(k).Visible = True
    mnuPluginList(k).Tag = id_plugin & "<>" & intStartupArgument
    
    ' guardar nombre del plugin (actualmente el nombre del archivo)
    ReDim Preserve zs_plugins_name(id_plugin)
    zs_plugins_name(UBound(zs_plugins_name)) = strPluginName
    
    ' agregar descripcion
    ReDim Preserve zs_plugins_desc(id_plugin)
    zs_plugins_desc(UBound(zs_plugins_desc)) = strDescription
    
    ' establecer id de menu
    ReDim Preserve zl_plugins_id_menu(id_plugin)
    zl_plugins_id_menu(UBound(zl_plugins_id_menu)) = k
    
    ' establecer como activo
    ReDim Preserve zb_plugins_active(id_plugin)
    zb_plugins_active(UBound(zb_plugins_active)) = True
    
    If mb_plugin_add_to_flex Then
        AddPluginToFlex id_plugin, zs_plugins_name(id_plugin), strDescription
    End If
    
End Function

Private Sub mnuPluginList_Click(Index As Integer)
    '===================================================
    Dim z_str() As String
    '===================================================
    On Error GoTo Handler
    z_str = Split(mnuPluginList(Index).Tag, "<>")
    If zb_plugins_active(CInt(z_str(0))) Then
        zo_plugins(CInt(z_str(0))).StartUp CInt(z_str(1))
    End If
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "mnuPluginList_Click"
End Sub

Public Sub FillPluginsInFlex(ByVal flx As MSHFlexGrid)
    '===================================================
    Dim k As Integer
    Dim b_Redraw As Byte
    Dim active_plugins As Long
    '==================================================
    On Error GoTo Handler
    
    With flx
        
        .Row = 0
        'borramos todas las filas exceptuando la fija y la primera
        .Rows = 2
                
        If gfnc_ZIsEmpty(zo_plugins) Then
            'primera fila invisible no se puede eliminar
            .RowHeight(1) = 0
            GoTo SALIR
        End If
        
        .Redraw = False
        
        b_Redraw = 1
        active_plugins = 0
        
        For k = 1 To intNumPlugins
        
            If zb_plugins_active(k - 1) Then
            
                active_plugins = active_plugins + 1
                .Rows = .Rows + 1
                .Row = .Rows - 1
                'forzar visible
                .RowHeight(.Row) = -1
                
                '***************************
                ' ID
                '***************************
                .Col = 0
                .text = k - 1
                '***************************
                ' PluginNombre
                '***************************
                .Col = 1
                .CellAlignment = flexAlignLeftCenter
                Set .CellPicture = imgPlugins.ListImages.Item(1).Picture
                .text = "     " & zs_plugins_name(k - 1)
                '***************************
                ' Activo
                '***************************
                .Col = 2
                .CellAlignment = flexAlignLeftCenter
                Set .CellPicture = imgPlugins.ListImages.Item(2).Picture
                .text = "     Activo"
                '***************************
                ' Registrado
                '***************************
                .Col = 3
                .CellAlignment = flexAlignLeftCenter
                Set .CellPicture = imgPlugins.ListImages.Item(2).Picture
                .text = "     Registrado"
                '***************************
                ' Descripcion
                '***************************
                .Col = 4
                .CellAlignment = flexAlignLeftCenter
                .text = zs_plugins_desc(k - 1)
    
                If b_Redraw = 1 Then
                    If k >= CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) + 1 Then
                        .Redraw = True
                        .Refresh
                        .Redraw = False
                        b_Redraw = 0
                    End If
                End If
            End If
        Next k
        
        '************************************************
        'eliminar la primera fila invisible
        '************************************************
        If active_plugins > 0 Then
            .RemoveItem (1)
        End If
        
SALIR:  'seleccionar el primero
        .Row = 1
        .Col = 1
        .ColSel = 4
        
        .Redraw = True
    
    End With
        
    Screen.MousePointer = vbDefault
    flx.MousePointer = flexDefault
    
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "FillPluginsInFlex"
    flx.Redraw = True
    Screen.MousePointer = vbDefault
    flx.MousePointer = flexDefault
End Sub

Public Sub UnregisterPlugin(ByVal id_plugin As Long)
    On Error Resume Next
    mnuPluginList(zl_plugins_id_menu(id_plugin)).Visible = False
    zb_plugins_active(id_plugin) = False
    Set zo_plugins(id_plugin) = Nothing
    intNumPlugins = intNumPlugins - 1
End Sub

Public Sub ExecutePlugin(ByVal id_plugin As Long)
    On Error Resume Next
    mnuPluginList_Click (zl_plugins_id_menu(id_plugin))
End Sub

Public Sub AddPluginToFlex(ByVal id_plugin As Long, ByVal nombre As String, ByVal descripcion As String)
    
    On Error GoTo Handler
    
    With frmPluginControl.flxResults
        
        .Redraw = False
        
        .Rows = .Rows + 1
        .Row = .Rows - 1
        'forzar visible
        .RowHeight(.Row) = -1
        
        '***************************
        ' ID
        '***************************
        .Col = 0
        .text = id_plugin
        '***************************
        ' PluginNombre
        '***************************
        .Col = 1
        .CellAlignment = flexAlignLeftCenter
        Set .CellPicture = imgPlugins.ListImages.Item(1).Picture
        .text = "     " & nombre
        '***************************
        ' Activo
        '***************************
        .Col = 2
        .CellAlignment = flexAlignLeftCenter
        Set .CellPicture = imgPlugins.ListImages.Item(2).Picture
        .text = "     Activo"
        '***************************
        ' Registrado
        '***************************
        .Col = 3
        .CellAlignment = flexAlignLeftCenter
        Set .CellPicture = imgPlugins.ListImages.Item(2).Picture
        .text = "     Registrado"
        '***************************
        ' Descripcion
        '***************************
        .Col = 4
        .CellAlignment = flexAlignLeftCenter
        .text = descripcion

        'seleccionar el agregado
        .Row = .Rows - 1
        .Col = 1
        .ColSel = 4
        
        .Redraw = True
    
    End With
        
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "AddPluginToFlex"
End Sub

'--------------------------------------------------------------------------------
'   Muestra el cuadro [Abrir ...] para escoger el plugin a registrar
'
Public Sub RegisterNewPlugin()
    '===================================================
    Dim ret As Long
    Dim PluginID As String
    '===================================================
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'esconde casilla de solo lectura y verifica que el archivo y el path existan
        .flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .DialogTitle = "Indicar el plugin a registrar:"
        .Filter = "Plugins (*.dll)|*.dll|Todos los Archivos(*.*)|*.*"
        .InitDir = App.Path & "\plugins"    ' de no existir el directorio usara el directorio activo
        'tipo predefinido DLL
        .FilterIndex = 1
        'nombre del reporte inicial
        .ShowOpen
        If .filename <> "" Then
            If UCase(Right(.filename, 4)) = ".DLL" Then
                Shell "regsvr32 """ & .filename & """", vbNormalFocus
                'intentar cargar el plugin dinamicamente
                PluginID = gfnc_GetFileNameWithoutExt(.FileTitle) & ".exPlugin"
                mb_plugin_add_to_flex = True
                If True = TryLoadPlugin(PluginID, .FileTitle) Then
                    ret = WritePrivateProfileString("Add-Ins32", .FileTitle, .filename, App.Path & "\R_porter.ini")
                End If
            End If
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "mnuRegistrarPlugin_Click"
    End If
End Sub

Public Function plg_GetDNS() As String
    ' devolver la cadena DNS
    plg_GetDNS = gs_DSN
End Function

Public Function plg_GetPWD() As String
    ' devolver el pwd
    plg_GetPWD = gs_Pwd
End Function

'-----------------------------------------------------------------------------------------------------
' Ejecuta el plugin cargado que tenga como nombre "strPluginName"
'
Public Function gfnc_ScriptExecutePlugin(ByVal strPluginName As String, ByVal intParam As Integer, Optional ArrayParam As Variant = Nothing) As Boolean
    '======================================
    Dim k As Integer
    Dim find As Boolean
    '======================================
    On Error GoTo Handler
    
    find = False
    gfnc_ScriptExecutePlugin = False
    
    'buscar el plugin en la coleccion
    For k = 1 To intNumPlugins
        If zb_plugins_active(k - 1) Then
            If UCase(Trim(zs_plugins_name(k - 1))) = UCase(Trim(strPluginName)) Then
                find = True
                Exit For
            End If
        End If
    Next k
    
    ' si se encontro
    If find Then
        gfnc_ScriptExecutePlugin = zo_plugins(k - 1).StartUp(intParam, ArrayParam)
    End If
    
    Exit Function

Handler:    MsgBox Err.Number & ":" & Err.Description, vbCritical, "gfnc_ScriptExecutePlugin"
End Function


