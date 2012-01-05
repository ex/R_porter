VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{40EF20B1-7EC5-11D8-95A1-9655FE58C763}#2.0#0"; "exLightButton.ocx"
Object = "{40EF20E1-7EC5-11D8-95A1-9655FE58C763}#2.0#0"; "exSplit.ocx"
Begin VB.Form frmDataControl 
   Caption         =   "Base de datos"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   Icon            =   "frmDataControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9375
   StartUpPosition =   1  'CenterOwner
   Begin exLightButton.ocxLightButton cmdVerDetalle 
      Height          =   750
      Left            =   4650
      TabIndex        =   40
      Top             =   3360
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   1323
      Activate        =   -1  'True
      PictureON       =   "frmDataControl.frx":058A
      PictureOFF      =   "frmDataControl.frx":0A8C
      PictureOK       =   "frmDataControl.frx":0F8E
      MouseIcon       =   "frmDataControl.frx":1490
      MousePointer    =   99
   End
   Begin exSplit.SplitRegion ctrSpliter 
      Height          =   5025
      Left            =   30
      Top             =   1410
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8864
      FirstControl    =   "flxResults"
      SecondControl   =   "flxDetails"
      FirstControlMinSize=   3000
      SecondControlMinSize=   120
      SplitterBarVertical=   -1  'True
      SplitterBarThickness=   90
      MouseIcon       =   "frmDataControl.frx":15F2
      MousePointer    =   99
   End
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   8430
      Top             =   7500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstA 
      Left            =   6000
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   105
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":1754
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":2272
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":2D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":38AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":43CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":4EEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrDownName 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2850
      Top             =   8500
   End
   Begin VB.PictureBox pbxSearchAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   1020
      ScaleHeight     =   1260
      ScaleWidth      =   3000
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8000
      Visible         =   0   'False
      Width           =   3030
      Begin VB.OptionButton optDBAuthor 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.OptionButton optDBAuthor 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBAuthor 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBAuthor 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBAuthor 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblOptionAuthor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Todos"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   39
         Top             =   15
         Width           =   2685
      End
      Begin VB.Label lblOptionAuthor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que contenga..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   38
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label lblOptionAuthor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que contenga... (palabra completa)"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   37
         Top             =   525
         Width           =   2685
      End
      Begin VB.Label lblOptionAuthor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que empiece con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   36
         Top             =   780
         Width           =   2685
      End
      Begin VB.Label lblOptionAuthor 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que termine con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   35
         Top             =   1035
         Width           =   2685
      End
   End
   Begin VB.Timer tmrDownAuthor 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2850
      Top             =   8500
   End
   Begin VB.PictureBox pbxSearchName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   1020
      ScaleHeight     =   1260
      ScaleWidth      =   3000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   3030
      Begin VB.OptionButton optDBName 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.OptionButton optDBName 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBName 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBName 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBName 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.Label lblOptionName 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que termine con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   28
         Top             =   1035
         Width           =   2685
      End
      Begin VB.Label lblOptionName 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que empiece con..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   27
         Top             =   780
         Width           =   2685
      End
      Begin VB.Label lblOptionName 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que contenga... (palabra completa)"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   26
         Top             =   525
         Width           =   2685
      End
      Begin VB.Label lblOptionName 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Que contenga..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   25
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label lblOptionName 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Todos"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   24
         Top             =   15
         Width           =   2685
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin VB.Frame fraBusqueda 
      Height          =   1395
      Left            =   30
      TabIndex        =   15
      Top             =   -45
      Width           =   9330
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         ItemData        =   "frmDataControl.frx":523C
         Left            =   5805
         List            =   "frmDataControl.frx":523E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   3030
      End
      Begin VB.CommandButton cmdSort 
         Height          =   315
         Left            =   3300
         Picture         =   "frmDataControl.frx":5240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   975
         Width           =   330
      End
      Begin VB.CommandButton cmdDownAuthor 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3735
         Picture         =   "frmDataControl.frx":538A
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   615
         Width           =   255
      End
      Begin VB.CommandButton cmdDownName 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3735
         Picture         =   "frmDataControl.frx":54B0
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   990
         TabIndex        =   1
         Text            =   "[Todos]"
         Top             =   210
         Width           =   3030
      End
      Begin VB.ComboBox cmbGenre 
         Height          =   315
         ItemData        =   "frmDataControl.frx":55D6
         Left            =   5805
         List            =   "frmDataControl.frx":55D8
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   3030
      End
      Begin VB.ComboBox cmbParent 
         Height          =   315
         ItemData        =   "frmDataControl.frx":55DA
         Left            =   5805
         List            =   "frmDataControl.frx":55DC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   585
         Width           =   3030
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   315
         Left            =   3660
         Picture         =   "frmDataControl.frx":55DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   975
         Width           =   330
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Opciones de busqueda..."
         Height          =   315
         Left            =   990
         TabIndex        =   10
         Top             =   975
         Width           =   2265
      End
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   990
         TabIndex        =   3
         Text            =   "[Todos]"
         Top             =   585
         Width           =   3030
      End
      Begin VB.Label lblCategory 
         Caption         =   "&Categoria"
         Height          =   225
         Left            =   4575
         TabIndex        =   4
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lblName 
         Caption         =   "&Nombre"
         Height          =   225
         Left            =   255
         TabIndex        =   0
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lblAuthor 
         Caption         =   "&Autor"
         Height          =   225
         Left            =   255
         TabIndex        =   2
         Top             =   645
         Width           =   915
      End
      Begin VB.Label lblBelongTo 
         Caption         =   "&Perteneciente a:"
         Height          =   225
         Left            =   4575
         TabIndex        =   6
         Top             =   645
         Width           =   1200
      End
      Begin VB.Label lblGenre 
         Caption         =   "&Género"
         Height          =   225
         Left            =   4575
         TabIndex        =   8
         Top             =   1020
         Width           =   1200
      End
   End
   Begin MSComctlLib.ImageList imglstB 
      Left            =   2850
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   13027270
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":5728
            Key             =   "down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":585E
            Key             =   "up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":5994
            Key             =   "left_off"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":5E96
            Key             =   "left_on"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":6398
            Key             =   "left_ok"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":689A
            Key             =   "right_off"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":6D9C
            Key             =   "right_on"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataControl.frx":729E
            Key             =   "right_ok"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxResults 
      Height          =   5025
      Left            =   30
      TabIndex        =   13
      Top             =   1410
      Width           =   4613
      _ExtentX        =   8149
      _ExtentY        =   8864
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   19
      BackColorFixed  =   16750143
      ForeColorFixed  =   16777215
      BackColorSel    =   16775910
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetails 
      Height          =   5025
      Left            =   4733
      TabIndex        =   14
      Top             =   1410
      Width           =   4613
      _ExtentX        =   8123
      _ExtentY        =   8864
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   23
      Cols            =   3
      BackColorFixed  =   16777215
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   -2147483643
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Menu munArchivo 
      Caption         =   "&Datos"
      Begin VB.Menu mnuExploreDB 
         Caption         =   "E&xplorar BD"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuText2Search 
         Caption         =   "Texto a buscar..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuActualizar 
         Caption         =   "&Actualizar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeeDetails 
         Caption         =   "Ver &detalles"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "&Opciones..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportarResultados 
         Caption         =   "&Exportar resultados..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "&Avanzado"
      Begin VB.Menu mnuSQL 
         Caption         =   "Ejecutar comando &SQL"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEstructura 
         Caption         =   "&Ver estructura de la BD"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuRegistros 
      Caption         =   "&Registros"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Editar registro"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&Nuevo registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar registro"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuOcultar 
         Caption         =   "Oc&ultar registro"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "E&liminar registro"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuOtherAction 
         Caption         =   "Mas..."
         Begin VB.Menu mnuRenameLowerCase 
            Caption         =   "Cambiar nombres a minuscula"
         End
         Begin VB.Menu mnuEraseEmptyGenres 
            Caption         =   "Borrar generos sin registros asociados"
         End
         Begin VB.Menu mnuEraseAuthorsWithoutData 
            Caption         =   "Borrar autores sin registros asociados"
         End
         Begin VB.Menu mnuAuthors2LowerCase 
            Caption         =   "Renombrar todos los autores a minuscula"
         End
      End
      Begin VB.Menu mnur10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdenar 
         Caption         =   "&Ordenar..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnur1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "&Abrir archivo asociado"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOptionalDrive 
         Caption         =   "Establecer &unidad opcional..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveFileAs 
         Caption         =   "&Guardar archivo como..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuTablas 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu mnuMedios 
         Caption         =   "&Medios"
      End
      Begin VB.Menu mnuAutores 
         Caption         =   "&Autores"
      End
      Begin VB.Menu mnuCategorias 
         Caption         =   "&Categorías"
      End
      Begin VB.Menu mnuMediaTypes 
         Caption         =   "&Tipos de medio"
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Arc&hivos"
      End
      Begin VB.Menu mnuFileType 
         Caption         =   "Ti&pos de archivo"
      End
      Begin VB.Menu mnuGeneros 
         Caption         =   "&Géneros"
      End
      Begin VB.Menu mnuSubGeneros 
         Caption         =   "&Sub géneros"
      End
      Begin VB.Menu mnuGrupos 
         Caption         =   "Grup&os"
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuEditar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuQuitHidden 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Ocultar"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar de la lista"
      End
      Begin VB.Menu mnu51 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOcultarDetalles 
         Caption         =   "&Ocultar detalles"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVerDetalles 
         Caption         =   "&Ver detalles"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Abrir archivo asociado"
      End
      Begin VB.Menu mnuGuardarArchivoComo 
         Caption         =   "&Guardar archivo como..."
      End
      Begin VB.Menu mnu21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerSQL 
         Caption         =   "&Ver consulta SQL generada"
      End
   End
End
Attribute VB_Name = "frmDataControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const EX_SCROLL_FLX = 1

Private ms_temp As String
Private mb_DetalleMostrado As Boolean
Private mn_WidthDetalle As Integer

Private m_strQuery As String
Private m_OldSpliterPercent As Double

Private m_NumberColor As Long
Private m_NormalColor As Long
Private m_HiddenColor As Long

Private Const k_MinSplitPercent = 90
Private Const k_HideSplitPercent = 99

'**************************************************************
'* MENUS
'**************************************************************
Private Sub mnuActualizar_Click()
    Actualizar_CombosDeBusqueda
End Sub

Private Sub mnuAuthors2LowerCase_Click()
    RenameAllAuthors2LowerCase
End Sub

Private Sub mnuBuscar_Click()
    Generar_Lista
End Sub

Private Sub mnuEraseAuthorsWithoutData_Click()
    EraseAllAuthorWithoutData
End Sub

Private Sub mnuEraseEmptyGenres_Click()
    EraseAllEmptyGenres
End Sub

Private Sub mnuHide_Click()
    Ocultar_Registro True
End Sub

Private Sub mnuDelete_Click()
    Eliminar_Registro
End Sub

Private Sub mnuEdit_Click()
    Modificar_Registro
End Sub

Private Sub mnuExecute_Click()
    Ejecutar_Archivo
End Sub

Private Sub mnuExportarResultados_Click()
    gsub_FlxShowSaveAsDialog cmmdlg, flxResults
End Sub

Private Sub mnuMostrar_Click()
    If Not gb_DBShowHiddenFiles Then
        MsgBox FRM_DATA_CONTROL_1, vbExclamation
    Else
        Ocultar_Registro False
    End If
End Sub

Private Sub mnuOcultar_Click()
    Ocultar_Registro True
End Sub

Private Sub mnuOptionalDrive_Click()
    Dim sDrive As String
    Dim ascii As Integer
    On Error Resume Next
    
    sDrive = InputBox(FRM_DATA_CONTROL_2, FRM_DATA_CONTROL_3, gs_OptionalDrive)
    If (Trim(sDrive) <> "") Then
        ascii = Asc(UCase(Left(sDrive, 1)))
        If (ascii >= 65) And (ascii <= 90) Then
            gs_OptionalDrive = chr(ascii)
            MsgBox FRM_DATA_CONTROL_4 & ": [" & gs_OptionalDrive & "]", vbExclamation, FRM_DATA_CONTROL_4
            Exit Sub
        End If
    End If
    MsgBox FRM_DATA_CONTROL_5, vbExclamation, RPORTER_LOC_ERROR
End Sub

Private Sub mnuQuitHidden_Click()
    mnuMostrar_Click
End Sub

Private Sub mnuRenameLowerCase_Click()
    RenameNames2LowerCase
End Sub

Private Sub mnuSaveFileAs_Click()
    Guardar_Archivo_Como
End Sub

Private Sub mnuCategorias_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "category", "category", FRM_DATA_CONTROL_6, 2190
    clxTable.AddFields False, False, True, "category", "id_category", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "category"
    clxTable.SqlAdditionalWHERE = "(id_category > 0)"
    clxTable.SqlAdditionalORDER_BY = "category"
    clxTable.AddDeleteConstrains "author", "id_category"
    clxTable.AddDeleteConstrains "storage", "id_category"
    clxTable.AddDeleteConstrains "parent", "id_category"
    clxTable.AddDeleteConstrains "genre", "id_category"
    clxTable.Caption = FRM_DATA_CONTROL_8
    clxTable.ShowForm
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuCategorias_Click"
End Sub

Private Sub mnuSubGeneros_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "sub_genre", "sub_genre", FRM_DATA_CONTROL_9, 2190
    clxTable.AddFields False, True, False, "genre", "genre", FRM_DATA_CONTROL_10, 1500, "", "sub_genre", "id_genre", "id_genre"
    clxTable.AddFields False, False, True, "sub_genre", "id_sub_genre", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "sub_genre, genre"
    clxTable.SqlAdditionalWHERE = "(sub_genre.id_sub_genre>0)"
    clxTable.SqlAdditionalORDER_BY = "sub_genre.sub_genre"
    clxTable.AddDeleteConstrains "file", "id_sub_genre"
    clxTable.Caption = FRM_DATA_CONTROL_11
    clxTable.ShowForm
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuSubGeneros_Click"
End Sub

Private Sub mnuAutores_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "author", "author", FRM_DATA_CONTROL_12, 3300
    clxTable.AddFields False, True, False, "category", "category", FRM_DATA_CONTROL_6, 1500, "", "author", "id_category", "id_category"
    clxTable.AddFields False, False, False, "author", "active", FRM_DATA_CONTROL_13, 600
    clxTable.AddFields False, False, True, "author", "id_author", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "author, category"
    clxTable.SqlAdditionalWHERE = "(author.id_author > 0)"
    clxTable.SqlAdditionalORDER_BY = "author.author"
    clxTable.AddDeleteConstrains "file", "id_author"
    clxTable.Caption = FRM_DATA_CONTROL_14
    clxTable.ShowForm 6960
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuAutores_Click"
End Sub

Private Sub mnuGrupos_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "parent", "parent", FRM_DATA_CONTROL_15, 2190
    clxTable.AddFields False, True, False, "category", "category", FRM_DATA_CONTROL_6, 1500, "", "parent", "id_category", "id_category"
    clxTable.AddFields False, False, True, "parent", "id_parent", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "parent, category"
    clxTable.SqlAdditionalWHERE = "(parent.id_parent > 0)"
    clxTable.SqlAdditionalORDER_BY = "parent.parent"
    clxTable.AddDeleteConstrains "file", "id_parent"
    clxTable.Caption = FRM_DATA_CONTROL_16
    clxTable.ShowForm
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuGrupos_Click"
End Sub

Private Sub mnuMedios_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "storage", "name", FRM_DATA_CONTROL_17, 1155
    clxTable.AddFields False, True, False, "storage_type", "storage_type", FRM_DATA_CONTROL_18, 585, "", "storage", "id_storage_type", "id_storage_type"
    clxTable.AddFields False, True, False, "category", "category", FRM_DATA_CONTROL_6, 1860, "", "storage", "id_category", "id_category"
    clxTable.AddFields False, False, False, "storage", "label", FRM_DATA_CONTROL_19, 1335
    clxTable.AddFields False, False, False, "storage", "serial", FRM_DATA_CONTROL_20, 1095
    clxTable.AddFields False, False, False, "storage", "fecha", FRM_DATA_CONTROL_21, 1230
    clxTable.AddFields False, False, False, "storage", "comment", FRM_DATA_CONTROL_22, 2460
    clxTable.AddFields False, False, False, "storage", "active", FRM_DATA_CONTROL_13, 600
    clxTable.AddFields False, False, True, "storage", "id_storage", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "storage, storage_type, category"
    clxTable.SqlAdditionalWHERE = "(storage.id_storage > 0)"
    clxTable.SqlAdditionalORDER_BY = "storage.name"
    clxTable.AddDeleteConstrains "file", "id_storage"
    clxTable.Caption = FRM_DATA_CONTROL_23
    clxTable.ShowForm 8220
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuMedios_Click"
End Sub

Private Sub mnuGeneros_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "genre", "genre", "Género", 2190
    clxTable.AddFields False, True, False, "category", "category", "Categoria", 1860, "", "genre", "id_category", "id_category"
    clxTable.AddFields False, False, False, "genre", "active", "Activo", 600
    clxTable.AddFields False, False, True, "genre", "id_genre", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "genre, category"
    clxTable.SqlAdditionalWHERE = "(genre.id_genre > 0)"
    clxTable.SqlAdditionalORDER_BY = "genre.genre"
    clxTable.AddDeleteConstrains "file", "id_genre"
    clxTable.AddDeleteConstrains "sub_genre", "id_genre"
    clxTable.Caption = "Géneros"
    clxTable.ShowForm 6180
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuGeneros_Click"
End Sub

Private Sub mnuMediaTypes_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "storage_type", "storage_type", "Tipo de medio", 2190
    clxTable.AddFields False, False, True, "storage_type", "id_storage_type", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "storage_type"
    clxTable.SqlAdditionalWHERE = "(id_storage_type > 0)"
    clxTable.SqlAdditionalORDER_BY = "storage_type"
    clxTable.AddDeleteConstrains "storage", "id_storage_type"
    clxTable.Caption = "Tipos de medio"
    clxTable.ShowForm 6180
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuMediaTypes_Click"
End Sub

Private Sub mnuFiles_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "file", "name", "Nombre", 1250    ' Main table - main search field
    clxTable.AddFields False, False, False, "file", "sys_name", "Archivo", 1500
    clxTable.AddFields False, True, False, "file_type", "file_type", "Tipo", 585, "", "file", "id_file_type", "id_file_type"
    clxTable.AddFields False, False, False, "file", "sys_length", "Tamaño", 900
    clxTable.AddFields False, True, False, "storage", "name", "Medio", 585, "", "file", "id_storage", "id_storage"
    clxTable.AddFields False, False, False, "file", "hidden", "Escondido", 900
    clxTable.AddFields False, True, False, "parent", "parent", "Pertence a", 585, "", "file", "id_parent", "id_parent"
    clxTable.AddFields False, False, False, "file", "id_sys_parent", "ID contenedor", 585
    clxTable.AddFields False, True, False, "author", "author", "Autor", 585, "", "file", "id_author", "id_author"
    clxTable.AddFields False, True, False, "genre", "genre", "Genero", 585, "", "file", "id_genre", "id_genre"
    clxTable.AddFields False, True, False, "sub_genre", "sub_genre", "Sub genero", 585, "", "file", "id_sub_genre", "id_sub_genre"
    clxTable.AddFields False, False, False, "file", "fecha", "Fecha", 1335
    clxTable.AddFields False, False, False, "file", "priority", "Prioridad", 720
    clxTable.AddFields False, False, False, "file", "quality", "Calidad", 900
    clxTable.AddFields False, False, False, "file", "type_quality", "Tipo calidad", 900
    clxTable.AddFields False, False, False, "file", "comment", "Comentario", 600
    clxTable.AddFields False, False, False, "file", "type_comment", "Observacion", 600
    clxTable.AddFields False, False, True, "file", "id_file", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "file, file_type, storage, parent, author, genre, sub_genre"
    clxTable.SqlAdditionalWHERE = "(file.id_file > 0)"
    clxTable.SqlAdditionalORDER_BY = "file.name"
    clxTable.Caption = "Archivos"
    clxTable.dontSearchAtBeginning = True ' Table is big, so don't search automatically
    clxTable.ShowForm 8220
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuFiles_Click"
End Sub

Private Sub mnuFileType_Click()
    Dim clxTable As New clsExDynaTable
    On Error GoTo Handler
    
    clxTable.AddFields True, False, False, "file_type", "file_type", "Tipo de archivo", 2190
    clxTable.AddFields False, False, True, "file_type", "id_file_type", FRM_DATA_CONTROL_7, 540
    clxTable.SqlFROM = "file_type"
    clxTable.SqlAdditionalWHERE = "(id_file_type > 0)"
    clxTable.SqlAdditionalORDER_BY = "file_type"
    clxTable.AddDeleteConstrains "file", "id_file_type"
    clxTable.Caption = "Tipos de archivo"
    clxTable.ShowForm 6180
    Exit Sub

Handler:    MsgBox Err.Description, vbCritical, "mnuFileType_Click"
End Sub

Private Sub mnuNew_Click()
    Agregar_Registro
End Sub

Private Sub mnuOpciones_Click()
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            frmOpcBusqueda.Show vbModal
            UpdateForm
            cmdSearch.SetFocus
        Else
            gsub_ShowMessageWrongDB
            Unload Me
        End If
    Else
        gsub_ShowMessageNoConection
        Unload Me
    End If
End Sub

Private Sub mnuOrdenar_Click()
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            frmOrdenar.Show vbModal
        Else
            gsub_ShowMessageWrongDB
            Unload Me
        End If
    Else
        gsub_ShowMessageNoConection
        Unload Me
    End If
End Sub

Private Sub mnuQuitar_Click()
    Eliminar_Resultados
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuSeeDetails_Click()
    cmdVerDetalle_Click
End Sub

Private Sub mnuEditar_Click()
    mnuEdit_Click
End Sub

Private Sub mnuNuevo_Click()
    mnuNew_Click
End Sub

Private Sub mnuEliminar_Click()
    mnuDelete_Click
End Sub

Private Sub mnuOcultarDetalles_Click()
    cmdVerDetalle_Click
End Sub

Private Sub mnuText2Search_Click()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.text)
    txtName.SetFocus
End Sub

Private Sub mnuVerDetalles_Click()
    cmdVerDetalle_Click
    Mostrar_Detalles
End Sub

Private Sub mnuAbrir_Click()
    mnuExecute_Click
End Sub

Private Sub mnuGuardarArchivoComo_Click()
    mnuSaveFileAs_Click
End Sub

Private Sub mnuVerSQL_Click()
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
        Unload Me
    Else
        frmSQL.Show
        frmSQL.SetSQL m_strQuery
    End If
End Sub

Private Sub mnuSQL_Click()
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
        Unload Me
    Else
        frmSQL.Show
    End If
End Sub

Private Sub mnuEstructura_Click()
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
        Unload Me
    Else
        frmReporte.Show
    End If
End Sub

Private Sub mnuExploreDB_Click()
    Load frmDataExplorer
    frmDataExplorer.Show vbModeless
End Sub

'**************************************************************
'* CMBOPTION AUTHOR
'**************************************************************
Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch.value = True
    End If
End Sub

Private Sub txtAuthor_GotFocus()
    If gt_DBAuthorSearchStyle = db_Todos Then
        txtAuthor.text = "[Todos]"
    End If
    
    txtAuthor.SelStart = 0
    txtAuthor.SelLength = Len(txtAuthor.text)
End Sub

Private Sub txtAuthor_VerifyAllScan()
    ' por si el usuario quiere buscar algo diferente a todos
    If gt_DBAuthorSearchStyle = db_Todos Then
        If (txtAuthor.text <> "[Todos]") And (Trim(txtAuthor.text) <> "") Then
            lblOptionAuthor(db_Todos).ForeColor = &H0&
            lblOptionAuthor(db_Todos).BackColor = &HFFFFFF
            gt_DBAuthorSearchStyle = db_Con
        End If
    End If
End Sub

Private Sub txtAuthor_LostFocus()
    txtAuthor_VerifyAllScan
End Sub

Private Sub txtAuthor_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        If gb_DBAuthorSearchStyleActive = False Then
            gb_DBAuthorSearchStyleActive = True
            pbxSearchAuthor.Top = txtAuthor.Top + 270
            pbxSearchAuthor.Visible = True
                    
            optDBAuthor(gt_DBAuthorSearchStyle).value = True
            optDBAuthor(gt_DBAuthorSearchStyle).SetFocus
            
            Set cmdDownAuthor.Picture = imglstB.ListImages.Item("up").Picture
        End If
    End If
End Sub

Private Sub cmdDownAuthor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownAuthor_Click
    End If
End Sub

Private Sub cmdDownAuthor_Click()
    If gb_DBAuthorSearchStyleActive = False Then
        gb_DBAuthorSearchStyleActive = True
        pbxSearchAuthor.Top = txtAuthor.Top + 270
        pbxSearchAuthor.Visible = True
                
        optDBAuthor(gt_DBAuthorSearchStyle).value = True
        optDBAuthor(gt_DBAuthorSearchStyle).SetFocus
        
        Set cmdDownAuthor.Picture = imglstB.ListImages.Item("up").Picture
    Else
        tmrDownAuthor.Enabled = False
        gb_DBAuthorSearchStyleActive = False
        'primero quitamos enfoque
        txtAuthor.SetFocus
        pbxSearchAuthor.Visible = False
        Set cmdDownAuthor.Picture = imglstB.ListImages.Item("down").Picture
    End If
End Sub

Private Sub optDBAuthor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    optDBAuthor(Index).value = True
    tmrDownAuthor.Enabled = True
End Sub

Private Sub optDBAuthor_Click(Index As Integer)
    If gb_DBAuthorSearchStyleActive = True Then
        gt_DBAuthorSearchStyle = Index
    End If
End Sub

Private Sub optDBAuthor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownAuthor_Click
    End If
End Sub

Private Sub optDBAuthor_LostFocus(Index As Integer)
    If (Me.ActiveControl.Name = "cmdDownAuthor") Then
        Me.ActiveControl.SetFocus
    Else
        If (Me.ActiveControl.Name = "optDBAuthor") Then
            lblOptionAuthor(Index).ForeColor = &H0&
            lblOptionAuthor(Index).BackColor = &HFFFFFF
            Me.ActiveControl.SetFocus
        Else
            If (Me.ActiveControl.Name = "pbxSearchAuthor") Then
                optDBAuthor(Index).SetFocus
            Else
                gb_DBAuthorSearchStyleActive = False
                pbxSearchAuthor.Visible = False
                Set cmdDownAuthor.Picture = imglstB.ListImages.Item("down").Picture
            End If
        End If
    End If
End Sub

Private Sub optDBAuthor_GotFocus(Index As Integer)
    lblOptionAuthor(Index).ForeColor = &HFFFFFF
    lblOptionAuthor(Index).BackColor = &HFF963F
End Sub

Private Sub lblOptionAuthor_Click(Index As Integer)
    optDBAuthor(Index).value = True
    optDBAuthor(Index).SetFocus
    tmrDownAuthor.Enabled = True
End Sub

Private Sub tmrDownAuthor_Timer()
    tmrDownAuthor.Enabled = False
    gb_DBAuthorSearchStyleActive = False
    'primero quitamos enfoque
    txtAuthor.SetFocus
    pbxSearchAuthor.Visible = False
    Set cmdDownAuthor.Picture = imglstB.ListImages.Item("down").Picture
End Sub

'**************************************************************
'* CMBOPTION NAME
'**************************************************************
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch.value = True
    End If
End Sub

Private Sub txtName_GotFocus()
    If gt_DBNameSearchStyle = db_Todos Then
        txtName.text = "[Todos]"
    End If
    
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.text)
End Sub

Private Sub txtName_VerifyAllScan()
    ' por si el usuario quiere buscar algo diferente a todos
    If gt_DBNameSearchStyle = db_Todos Then
        If (txtName.text <> "[Todos]") And (Trim(txtName.text) <> "") Then
            lblOptionName(db_Todos).ForeColor = &H0&
            lblOptionName(db_Todos).BackColor = &HFFFFFF
            gt_DBNameSearchStyle = db_Con
        End If
    End If
End Sub

Private Sub txtName_LostFocus()
    txtName_VerifyAllScan
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        If gb_DBNameSearchStyleActive = False Then
            gb_DBNameSearchStyleActive = True
            pbxSearchName.Top = txtName.Top + 270
            pbxSearchName.Visible = True
                    
            optDBName(gt_DBNameSearchStyle).value = True
            optDBName(gt_DBNameSearchStyle).SetFocus
            
            Set cmdDownName.Picture = imglstB.ListImages.Item("up").Picture
        End If
    End If
End Sub

Private Sub cmdDownName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownName_Click
    End If
End Sub

Private Sub cmdDownName_Click()
    If gb_DBNameSearchStyleActive = False Then
        gb_DBNameSearchStyleActive = True
        pbxSearchName.Top = txtName.Top + 270
        pbxSearchName.Visible = True
                
        optDBName(gt_DBNameSearchStyle).value = True
        optDBName(gt_DBNameSearchStyle).SetFocus
        
        Set cmdDownName.Picture = imglstB.ListImages.Item("up").Picture
    Else
        tmrDownName.Enabled = False
        gb_DBNameSearchStyleActive = False
        'primero quitamos enfoque
        txtName.SetFocus
        pbxSearchName.Visible = False
        Set cmdDownName.Picture = imglstB.ListImages.Item("down").Picture
    End If
End Sub

Private Sub optDBName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    optDBName(Index).value = True
    tmrDownName.Enabled = True
End Sub

Private Sub optDBName_Click(Index As Integer)
    If gb_DBNameSearchStyleActive = True Then
        gt_DBNameSearchStyle = Index
    End If
End Sub

Private Sub optDBName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownName_Click
    End If
End Sub

Private Sub optDBName_LostFocus(Index As Integer)
    If (Me.ActiveControl.Name = "cmdDownName") Then
        Me.ActiveControl.SetFocus
    Else
        If (Me.ActiveControl.Name = "optDBName") Then
            lblOptionName(Index).ForeColor = &H0&
            lblOptionName(Index).BackColor = &HFFFFFF
            Me.ActiveControl.SetFocus
        Else
            If (Me.ActiveControl.Name = "pbxSearchName") Then
                optDBName(Index).SetFocus
            Else
                gb_DBNameSearchStyleActive = False
                pbxSearchName.Visible = False
                Set cmdDownName.Picture = imglstB.ListImages.Item("down").Picture
            End If
        End If
    End If
End Sub

Private Sub optDBName_GotFocus(Index As Integer)
    lblOptionName(Index).ForeColor = &HFFFFFF
    lblOptionName(Index).BackColor = &HFF963F
End Sub

Private Sub lblOptionName_Click(Index As Integer)
    optDBName(Index).value = True
    optDBName(Index).SetFocus
    
    tmrDownName.Enabled = True
End Sub

Private Sub tmrDownName_Timer()
    tmrDownName.Enabled = False
    gb_DBNameSearchStyleActive = False
    'primero quitamos enfoque
    txtName.SetFocus
    pbxSearchName.Visible = False
    Set cmdDownName.Picture = imglstB.ListImages.Item("down").Picture
End Sub

'**************************************************************
'* COMBOBOX
'**************************************************************
Private Sub cmbCategory_Click()
    
    On Error GoTo Handler
    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------
    
    Actualizar_ComboPariente
    Actualizar_ComboGenero
    Exit Sub
    
Handler: MsgBox Err.Description, vbCritical, "cmbCategory_Click"
End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch.value = True
    End If
End Sub

Private Sub cmbParent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch.value = True
    End If
End Sub

Private Sub cmbGenre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch.value = True
    End If
End Sub

'**************************************************************
'* COMMANDBUTTONS
'**************************************************************
Private Sub cmdSort_Click()
    mnuOrdenar_Click
End Sub

Private Sub cmdOptions_Click()
    mnuOpciones_Click
End Sub

Private Sub cmdSearch_Click()
    mnuBuscar_Click
End Sub

Private Sub cmdVerDetalle_Click()

    On Error Resume Next
   
    If mb_DetalleMostrado = True Then

        mb_DetalleMostrado = False
        
        Set cmdVerDetalle.PictureOFF = imglstB.ListImages.Item("left_off").Picture
        Set cmdVerDetalle.PictureOK = imglstB.ListImages.Item("left_ok").Picture
        Set cmdVerDetalle.PictureON = imglstB.ListImages.Item("left_on").Picture
        cmdVerDetalle.Tag = "left"
        
        m_OldSpliterPercent = ctrSpliter.SplitPercent
        ctrSpliter = k_HideSplitPercent

        flxDetails.ScrollBars = flexScrollBarNone

        flxResults.SetFocus
        
        mnuOcultarDetalles.Visible = False
        mnuVerDetalles.Visible = True
        mnuSeeDetails.Caption = "Ver &detalles"
        
    Else

        mb_DetalleMostrado = True
        
        Set cmdVerDetalle.PictureOFF = imglstB.ListImages.Item("right_off").Picture
        Set cmdVerDetalle.PictureOK = imglstB.ListImages.Item("right_ok").Picture
        Set cmdVerDetalle.PictureON = imglstB.ListImages.Item("right_on").Picture
        cmdVerDetalle.Tag = "right"
                
        If m_OldSpliterPercent < k_MinSplitPercent Then
            ctrSpliter.SplitPercent = m_OldSpliterPercent
        Else
            ' para mostrar detalles cuando el control de detalles se haya reducido mucho...
            ' (trata de mostrar la tabla de detalles completa)
            If Me.height > 8025 Then
                ctrSpliter.SplitPercent = 100 * (Me.width - 4710) / Me.width
            Else
                ctrSpliter.SplitPercent = 100 * (Me.width - 4950) / Me.width
            End If

        End If

        flxDetails.ScrollBars = flexScrollBarBoth

        mnuOcultarDetalles.Visible = True
        mnuVerDetalles.Visible = False
        mnuSeeDetails.Caption = "Ocultar &detalles"

    End If
    
    Form_Resize

End Sub

'**************************************************************
'* SPLIT CONTROL
'**************************************************************
Private Sub ctrSpliter_RepositionSplit()

    On Error Resume Next
    
    cmdVerDetalle.Top = ctrSpliter.Top + (ctrSpliter.height - cmdVerDetalle.height) / 2
    cmdVerDetalle.Left = flxDetails.Left - 90

    
    If mb_DetalleMostrado Then
        
        ' si se hace muy pequeño el detalle, mejor ocultarlo...
        If (ctrSpliter.SplitPercent > k_MinSplitPercent) Then
            cmdVerDetalle_Click
        End If
    
    Else
    
        ' si se hace el detalle suficientemente visible cuando esta ocultado, cambiar a estado a visible
        If (cmdVerDetalle.Tag = "left") And (ctrSpliter.SplitPercent < k_MinSplitPercent) Then
            
            mb_DetalleMostrado = True

            Set cmdVerDetalle.PictureOFF = imglstB.ListImages.Item("right_off").Picture
            Set cmdVerDetalle.PictureOK = imglstB.ListImages.Item("right_ok").Picture
            Set cmdVerDetalle.PictureON = imglstB.ListImages.Item("right_on").Picture
            cmdVerDetalle.Tag = "right"

            flxDetails.ScrollBars = flexScrollBarBoth

            mnuOcultarDetalles.Visible = True
            mnuVerDetalles.Visible = False
            mnuSeeDetails.Caption = "Ocultar &detalles"
            
        End If
        
    End If

End Sub

'**************************************************************
'* FLEXGRID
'**************************************************************
Private Sub flxDetails_DblClick()
    cmdVerDetalle_Click
End Sub

Private Sub flxDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxResults.SetFocus
    End If
End Sub

Private Sub flxDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Shift = vbCtrlMask Then
            Ejecutar_Archivo
        End If
    End If
End Sub


Private Sub flxResults_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        If flxResults.Row <> 0 And flxResults.RowHeight(flxResults.Row) <> 0 Then
            PopupMenu mnupopup, 2
        End If
    End If
    If KeyCode = 13 Then
        If (Shift = vbCtrlMask) Or (Shift = vbAltMask) Then
            Ejecutar_Archivo
        End If
    End If
End Sub

Private Sub flxResults_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If Shift = vbShiftMask Then
                Eliminar_Registro
            End If
            If Shift = 0 Then
                Eliminar_Resultados
            End If
    End Select
End Sub

Private Sub flxResults_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mostrar_Detalles
        With flxDetails
            .Row = 1
            .TopRow = 1
        End With
    End If
End Sub

Private Sub flxResults_DblClick()
    '**************************************************************
    ' TO DO: hacer nueva busqueda ordenada segun la columna
    ' donde se hizo doble click
    '**************************************************************
    Mostrar_Detalles
End Sub

Private Sub flxResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If flxResults.Row <> 0 And flxResults.RowHeight(flxResults.Row) <> 0 Then
            PopupMenu mnupopup, 2
        End If
    End If
End Sub

'**************************************************************
'* FORMULARIO
'**************************************************************
Private Sub Form_Load()
    
    Dim k As Integer
    On Error GoTo Handler
    
#If EX_SCROLL_FLX Then
    If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
        ' verificar si el mouse soporta o no la ruedita...
        lpPrevWndProcDataControl = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndProcDataControlForm)
        flxDetails.Tag = 1
        flxResults.Tag = 1
    End If
#End If

    mb_DetalleMostrado = True
        
    Me.Top = 255
    Me.Left = 60
    Me.width = 9840
    Me.height = 7245
       
    Actualizar_CombosDeBusqueda
    
    m_NumberColor = RGB(100, 170, 255)
    m_NormalColor = RGB(0, 0, 0)
    m_HiddenColor = RGB(128, 128, 128)
    
    '--------------------------------
    ' inicializar flxResults
    With flxResults
    
        .height = 4920
        
        .SelectionMode = flexSelectionByRow
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        
        'extension
        .Rows = 2
        .FixedRows = 1
        .Cols = 7
        .FixedCols = 1
        
        'ancho celdas
        .ColWidth(0) = 0
        .ColWidth(1) = 450
        .ColWidth(2) = 4005
        .ColWidth(3) = 2265
        .ColWidth(4) = 840
        .ColWidth(5) = 675
        .ColWidth(6) = 915
        
        'alto celda titulo
        .RowHeight(0) = 315
        
        'titulos columna
        .Row = 0
        .Col = 1
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Nº"
        .Col = 2
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Nombre"
        .Col = 3
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Autor"
        .Col = 4
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .Col = 5
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .text = "Género"
        .Col = 6
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
   
        .Row = 1
        
        'primera fila invisible
        .RowHeight(1) = 0

    End With
        
    Establecer_TitulosFlex
    
    '--------------------------------
    ' inicializar flxDetails
    With flxDetails

        .height = 4920

        .SelectionMode = flexSelectionByRow
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth

        'extension
        .Rows = 12
        .FixedRows = 1
        .Cols = 3
        .FixedCols = 1

        'ancho columnas
        .ColWidth(0) = 105
        .ColWidth(1) = 720
        .ColWidth(2) = 3645

        .ForeColor = vbBlack

        'cabecera
        .RowHeight(0) = 300
        .Row = 0
        For k = 1 To 2
            .Col = k
            .CellBackColor = RGB(63, 150, 255)
        Next k

        'filas
        .Col = 1

        .RowHeight(1) = 810

        .Row = 2
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Fecha"

        .Row = 3
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Tamaño"

        .Row = 4
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Tipo"
        .RowHeight(4) = 300

        .RowHeight(5) = 300

        .RowHeight(6) = 660
        .Row = 6
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Autor"

        .RowHeight(7) = 480
        .Row = 7
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Género"

        .Row = 8
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Calidad"

        .RowHeight(9) = 660
        .Row = 9
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "De"

        .RowHeight(10) = 690
        .Row = 10
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Medio"

        .RowHeight(11) = 840
        .Row = 11
        .CellTextStyle = flexTextRaisedLight
        .CellForeColor = &HFF963F
        .CellAlignment = flexAlignLeftTop
        .text = "Observ."


        For k = 1 To 2
          .Col = k
          .Row = 2
          .CellBackColor = &HFFFAE6
          .Row = 4
          .CellBackColor = &HFFFAE6
          .Row = 6
          .CellBackColor = &HFFFAE6
          .Row = 8
          .CellBackColor = &HFFFAE6
          .Row = 10
          .CellBackColor = &HFFFAE6
        Next k

        .Row = 1
        .Col = 1

    End With
    
    cmdVerDetalle_Click
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, Me.Name & "::Form_Load()"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "frmRegistros" Then
            Unload frm
        End If
    Next
End Sub

Private Sub Form_Resize()
    
    Dim nTop As Integer
    On Error Resume Next

    If Me.width < 9150 Then
        ' nothing
    Else
        
        fraBusqueda.width = Me.ScaleWidth - 60
        ctrSpliter.width = Me.ScaleWidth - 60
        
        If mb_DetalleMostrado = True Then
            ' nothing
        Else
            ctrSpliter.SplitPercent = k_HideSplitPercent
        End If
    End If
    
    
    If Me.height < 4500 Then
        ' nothing
    Else
        ctrSpliter.height = Me.ScaleHeight - 1485
    End If
    
    ctrSpliter_RepositionSplit
    
End Sub

Private Sub UpdateForm()
    If gb_DBNameFromFile Then
        lblName.Caption = "Archivo"
    Else
        lblName.Caption = "Nombre"
    End If
End Sub

'**************************************************************
'* FUNCIONES GENERALES
'**************************************************************

'--------------------------------------------------------------
' Genera los resultados, construye el query SQL basandose en
' las opciones globales establecidas en este formulario como
' en [frmOpcBusqueda]
'
Public Sub Generar_Lista()
    '===================================================
    Dim k As Integer
    Dim rs As ADODB.Recordset
    Dim querybase As String
    Dim queryorder As String
    Dim name_variant As String
    Dim author_variant As String
    Dim b_Redraw As Byte
    Dim l_FileZise As Long
    Dim s_Autor As String
    Dim s_Nombre As String
    Dim s_ConFechaDe As String
    Dim s_DiaSiguiente As String
    Dim d_DiaSiguiente As Date
    Dim table_variant As String
    Dim field_variant As String
    Dim table_aux_variant As String
    Dim field_aux_variant As String
    Dim field_variant_num As String
    Dim id_aux_variant As String
    Dim id_variant As String
    Dim lpMsg As MSG
    Dim lresult As Long
    Dim title_variant As String
    Dim queryHidden As String
    '===================================================
    
    On Error GoTo Handler
    
    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------
    
    ' por si el usuario quiere buscar algo diferente a todos
    txtName_VerifyAllScan
    txtAuthor_VerifyAllScan
    
    '==============================================================
    ' Buscamos caracteres de comilla (y los tratamos adecuadamente)
    '
    gfnc_ParseString txtName.text, s_Nombre
    gfnc_ParseString txtAuthor.text, s_Autor
    
    Limpiar_Detalles
    
    Screen.MousePointer = vbHourglass
    flxResults.MousePointer = flexHourglass
    
    '==============================================================
    ' Decidir si se mostraran en los resultados el Medio o el
    ' Pariente al cual pertenece.
    '
    If gb_DBPertenciaPorAlmacenamiento = True Then
        
        table_variant = "storage"
        field_variant = "name"
        id_variant = "id_storage"
        
    Else
    
        table_variant = "parent"
        field_variant = "parent"
        id_variant = "id_parent"
        
    End If
    
    '==============================================================
    ' Decidir si se se buscara por el nombre del
    ' titulo o el nombre del archivo
    '
    If gb_DBNameFromFile = True Then
        title_variant = "sys_name"
    Else
        title_variant = "name"
    End If
    
    
    '==============================================================
    ' Decidir si se mostraran en los resultados el Genero o el
    ' Tipo al cual pertenece el archivo.
    '
    If gb_DBCampoAuxiliarPorGenero = True Then
        
        table_aux_variant = "genre"
        field_aux_variant = "genre"
        id_aux_variant = "id_genre"
        
    Else
    
        table_aux_variant = "file_type"
        field_aux_variant = "file_type"
        id_aux_variant = "id_file_type"
        
    End If
    '--------------------------------------------------------------
    
    '==============================================================
    ' Escoger tabla a mostrar
    '
    Select Case gs_DBConCampoDe
        Case "Prioridad"
            field_variant_num = "priority"
        Case "Calidad"
            field_variant_num = "quality"
        Case "Tamaño"
            field_variant_num = "sys_length"
        Case "Fecha"
            field_variant_num = "fecha"
    End Select
    '--------------------------------------------------------------
    
    Select Case gt_DBNameSearchStyle
    
        Case db_Todos:
            txtName.text = "[Todos]"
    
        Case db_Con:
            ' verificar si se quiere busqueda compuesta
            Dim zsSqlParts() As String
            Dim zsSqlOperators() As String
            
            If gfnc_GetLogicalParts(Trim(s_Nombre), zsSqlParts, zsSqlOperators) Then
                '-----------------------------------------
                ' generar cadena compuesta con operadores
                For k = 1 To UBound(zsSqlOperators)
                    name_variant = name_variant & "'%" & Trim(zsSqlParts(k)) & "%') " & zsSqlOperators(k) & " (file." & title_variant & " LIKE "
                Next k
                ' el bucle for deja el contador con un valor superior a su limite superior
                name_variant = name_variant & "'%" & Trim(zsSqlParts(k)) & "%'"
            Else
                name_variant = "'%" & Trim(s_Nombre) & "%'"
            End If
            
        Case db_ConPalabra:
            name_variant = "'% " & Trim(s_Nombre) & " %') OR (file." & title_variant & " LIKE '" & Trim(s_Nombre) & " %') OR (file." & title_variant & " LIKE '% " & Trim(s_Nombre) & "') OR (file." & title_variant & "='" & Trim(s_Nombre) & "'"
            
        Case db_QueComience:
            name_variant = "'" & Trim(s_Nombre) & "%'"
            
        Case db_QueTermine:
            name_variant = "'%" & Trim(s_Nombre) & "'"
            
    End Select
        
    Select Case gt_DBAuthorSearchStyle
    
        Case db_Todos:
            txtAuthor.text = "[Todos]"
    
        Case db_Con:
            author_variant = "'%" & Trim(s_Autor) & "%'"
            
        Case db_ConPalabra:
            author_variant = "'% " & Trim(s_Autor) & " %') OR (author.author LIKE '" & Trim(s_Autor) & " %') OR (author.author LIKE '% " & Trim(s_Autor) & "') OR (author.author='" & Trim(s_Autor) & "'"
            
        Case db_QueComience:
            author_variant = "'" & Trim(s_Autor) & "%'"
            
        Case db_QueTermine:
            author_variant = "'%" & Trim(s_Autor) & "'"
            
    End Select
        
    If gb_DBShowHiddenFiles Then
        queryHidden = ", file.hidden "
    Else
        queryHidden = " "
    End If
    
    If ("[Todos]" <> cmbCategory.text) And ("[Todos]" = cmbGenre.text) And ("[Todos]" = cmbParent.text) Then
        '---------------------------------------------------------------------------------------------
        ' caso especial: mostrar todos los registros de la categoria seleccionada
        ' [TODO] no estoy muy seguro del DISTINCT quizas pueda optimizarse con otro tipo de query...
        If gb_DBPertenciaPorAlmacenamiento = True Then
        
            querybase = "SELECT DISTINCT file.id_file, file." & title_variant & " AS title, file." & field_variant_num & " AS quantity, author.author, "
            querybase = querybase & table_variant & "." & field_variant & " AS belong_to, "
            querybase = querybase & table_aux_variant & "." & field_aux_variant & " AS aux_field" & queryHidden
            querybase = querybase & "FROM file, author, " & table_variant & ", " & table_aux_variant & ", category "
            
        Else
        
            querybase = "SELECT DISTINCT file.id_file, file." & title_variant & " AS title, file." & field_variant_num & " AS quantity, author.author, "
            querybase = querybase & table_variant & "." & field_variant & " AS belong_to, "
            querybase = querybase & table_aux_variant & "." & field_aux_variant & " AS aux_field" & queryHidden
            querybase = querybase & "FROM file, author, " & table_variant & ", " & table_aux_variant & ", storage, category "
        
        End If
        
    Else
    
        querybase = "SELECT file.id_file, file." & title_variant & " AS title, file." & field_variant_num & " AS quantity, author.author, "
        querybase = querybase & table_variant & "." & field_variant & " AS belong_to, "
        querybase = querybase & table_aux_variant & "." & field_aux_variant & " AS aux_field" & queryHidden
        querybase = querybase & "FROM file, author, " & table_variant & ", " & table_aux_variant & " "
        
    End If
    
    If txtAuthor.text = "[Todos]" Then
        
        If txtName.text = "[Todos]" Then
        
            If cmbParent.text = "[Todos]" Then
            
                If cmbGenre.text = "[Todos]" Then
                
                    If cmbCategory.text = "[Todos]" Then
                    
                        query = querybase & "WHERE ((author.id_author=file.id_author) AND "
                        query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                        query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                    Else
                        '----------------------------------------------------------------------------
                        ' caso especial: mostrar todos los registros de la categoria seleccionada
                        If gb_DBPertenciaPorAlmacenamiento = True Then
                        
                            query = querybase & "WHERE ((storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(author.id_author=file.id_author) AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        Else
                        
                            query = querybase & "WHERE ((storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(file.id_storage=storage.id_storage) AND "
                            query = query & "(author.id_author=file.id_author) AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        End If
                    
                    End If
                
                Else
                    
                    query = querybase & "WHERE ((file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                End If
            
            Else
                
                If cmbGenre.text = "[Todos]" Then
                
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                Else
                    
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "(file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                End If
                
            End If
        
        Else
            '*******************
            'se busca por nombre
            '*******************
            If cmbParent.text = "[Todos]" Then
            
                If cmbGenre.text = "[Todos]" Then
                
                    If cmbCategory.text = "[Todos]" Then
                    
                        query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND " 'NOTA: no quitar parentesis adicionales
                        query = query & "(author.id_author=file.id_author) AND "
                        query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                        query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                    Else
                        '----------------------------------------------------------------------------
                        ' caso especial: mostrar solo los registros de la categoria seleccionada
                        If gb_DBPertenciaPorAlmacenamiento = True Then
                        
                            query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND "
                            query = query & "(author.id_author=file.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        Else
                        
                            query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND "
                            query = query & "(author.id_author=file.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(file.id_storage=storage.id_storage) AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        End If
                    
                    End If
                
                Else
                    
                    query = querybase & "WHERE ((file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                End If
            
            Else
                
                If cmbGenre.text = "[Todos]" Then
                
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                Else
                    
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "(file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                End If
                
            End If
        
        End If
                
    Else
        '***********************
        'se busca por autor
        '***********************
        If txtName.text = "[Todos]" Then
        
            If cmbParent.text = "[Todos]" Then
            
                If cmbGenre.text = "[Todos]" Then
                
                    If cmbCategory.text = "[Todos]" Then
                    
                        query = querybase & "WHERE (((author.author LIKE " & author_variant & ")) AND " 'NOTA: no quitar parentesis adicionales
                        query = query & "(file.id_author=author.id_author) AND "
                        query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                        query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                    Else
                        '----------------------------------------------------------------------------
                        ' caso especial: mostrar solo los registros de la categoria seleccionada
                        If gb_DBPertenciaPorAlmacenamiento = True Then
                        
                            query = querybase & "WHERE (((author.author LIKE " & author_variant & ")) AND "
                            query = query & "(file.id_author=author.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        Else
                        
                            query = querybase & "WHERE (((author.author LIKE " & author_variant & ")) AND "
                            query = query & "(file.id_author=author.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(file.id_storage=storage.id_storage) AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        End If
                    
                    End If
                
                Else
                    
                    query = querybase & "WHERE ((file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                End If
            
            Else
                
                If cmbGenre.text = "[Todos]" Then
                
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                Else
                    
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "(file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                End If
                
            End If
        
        Else
            '***************************
            'se busca por nombre y autor
            '***************************
            If cmbParent.text = "[Todos]" Then
            
                If cmbGenre.text = "[Todos]" Then
                
                    If cmbCategory.text = "[Todos]" Then
                    
                        query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND " 'NOTA: no quitar parentesis adicionales
                        query = query & "((author.author LIKE " & author_variant & ")) AND "
                        query = query & "(file.id_author=author.id_author) AND "
                        query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                        query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                        
                    Else
                        '----------------------------------------------------------------------------
                        ' caso especial: mostrar solo los registros de la categoria seleccionada
                        If gb_DBPertenciaPorAlmacenamiento = True Then
                        
                            query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND "
                            query = query & "((author.author LIKE " & author_variant & ")) AND "
                            query = query & "(file.id_author=author.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        Else
                        
                            query = querybase & "WHERE (((file." & title_variant & " LIKE " & name_variant & ")) AND "
                            query = query & "((author.author LIKE " & author_variant & ")) AND "
                            query = query & "(file.id_author=author.id_author) AND "
                            query = query & "(storage.id_category=" & cmbCategory.ItemData(cmbCategory.ListIndex) & ") AND "
                            query = query & "(file.id_storage=storage.id_storage) AND "
                            query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                            query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                            
                        End If
                    
                    End If
                
                Else
                    
                    query = querybase & "WHERE ((file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                End If
            
            Else
                
                If cmbGenre.text = "[Todos]" Then
                
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                
                Else
                    
                    query = querybase & "WHERE ((file." & id_variant & "=" & cmbParent.ItemData(cmbParent.ListIndex) & ") AND "
                    query = query & "(file.id_genre=" & cmbGenre.ItemData(cmbGenre.ListIndex) & ") AND "
                    query = query & "((file." & title_variant & " LIKE " & name_variant & ")) AND "
                    query = query & "((author.author LIKE " & author_variant & ")) AND "
                    query = query & "(author.id_author=file.id_author) AND "
                    query = query & "(" & table_variant & "." & id_variant & "=file." & id_variant & ") AND "
                    query = query & "(" & table_aux_variant & "." & id_aux_variant & "=file." & id_aux_variant & ")"
                    
                End If
                
            End If
        
        End If
        
    End If
                
    '==============================================================
    ' opciones de filtrado adicional (prioridad y calidad)
    '
    If gb_DBConCalidad = True Then
    
        If gb_DBConPrioridad = True Then
            
            Select Case gt_DBPrioritySearchStyle
            
                Case db_Menor:
                    query = query & " AND (file.priority<" & gs_DBConPrioridadDe & ") AND "

                Case db_MenorIgual:
                    query = query & " AND (file.priority<=" & gs_DBConPrioridadDe & ") AND "

                Case db_Igual:
                    query = query & " AND (file.priority=" & gs_DBConPrioridadDe & ") AND "

                Case db_MayorIgual:
                    query = query & " AND (file.priority>=" & gs_DBConPrioridadDe & ") AND "

                Case db_Mayor:
                    query = query & " AND (file.priority>" & gs_DBConPrioridadDe & ") AND "

            End Select
            
            Select Case gt_DBQualitySearchStyle
            
                Case db_Menor:
                    query = query & "(file.quality<" & gs_DBConCalidadDe & ") "
            
                Case db_MenorIgual:
                    query = query & "(file.quality<=" & gs_DBConCalidadDe & ") "
                    
                Case db_Igual:
                    query = query & "(file.quality=" & gs_DBConCalidadDe & ") "
                    
                Case db_MayorIgual:
                    query = query & "(file.quality>=" & gs_DBConCalidadDe & ") "
                    
                Case db_Mayor:
                    query = query & "(file.quality>" & gs_DBConCalidadDe & ") "
                    
            End Select

        Else
            
            Select Case gt_DBQualitySearchStyle
            
                Case db_Menor:
                    query = query & " AND (file.quality<" & gs_DBConCalidadDe & ") "
            
                Case db_MenorIgual:
                    query = query & " AND (file.quality<=" & gs_DBConCalidadDe & ") "
                    
                Case db_Igual:
                    query = query & " AND (file.quality=" & gs_DBConCalidadDe & ") "
                    
                Case db_MayorIgual:
                    query = query & " AND (file.quality>=" & gs_DBConCalidadDe & ") "
                    
                Case db_Mayor:
                    query = query & " AND (file.quality>" & gs_DBConCalidadDe & ") "
                    
            End Select
            
        End If
    Else
        If gb_DBConPrioridad = True Then
            
            Select Case gt_DBPrioritySearchStyle
                
                Case db_Menor:
                    query = query & " AND (file.priority<" & gs_DBConPrioridadDe & ") "
            
                Case db_MenorIgual:
                    query = query & " AND (file.priority<=" & gs_DBConPrioridadDe & ") "
                    
                Case db_Igual:
                    query = query & " AND (file.priority=" & gs_DBConPrioridadDe & ") "
                    
                Case db_MayorIgual:
                    query = query & " AND (file.priority>=" & gs_DBConPrioridadDe & ") "
                    
                Case db_Mayor:
                    query = query & " AND (file.priority>" & gs_DBConPrioridadDe & ") "
                    
            End Select
            
        Else
            ' vacio (para que continue con los filtros adicionales)
        End If
    End If
    '--------------------------------------------------------------
    
    '==============================================================
    ' opciones de filtrado adicional (tamaño, fecha y tipo)
    '
    '--------------------------------------------------------------
    ' generar fechas auxiliares para filtrar por fecha
    '
    If gb_DBConFecha = True Then
        
        s_ConFechaDe = Format(Month(gd_DBConFechaDe), "00") & "/" & Format(Day(gd_DBConFechaDe), "00") & "/" & Trim(str(Year(gd_DBConFechaDe)))
        
        If (gt_DBDateSearchStyle = db_Igual) Or (gt_DBDateSearchStyle = db_Mayor) Or (gt_DBDateSearchStyle = db_MenorIgual) Then
        
            d_DiaSiguiente = DateAdd("d", 1, gd_DBConFechaDe)   '[EX] thnxs for the function...
            s_DiaSiguiente = Format(Month(d_DiaSiguiente), "00") & "/" & Format(Day(d_DiaSiguiente), "00") & "/" & Trim(str(Year(d_DiaSiguiente)))
        
        End If
        
    End If
    
    If gb_DBConTamanyo = True Then
    
        If gb_DBConFecha = True Then
        
            If gb_DBConTipo = True Then
    
                Select Case gt_DBFileSizeSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.sys_length<" & gs_DBConTamanyoDe & ") "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.sys_length<=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Igual:
                        query = query & " AND (file.sys_length=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.sys_length>=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Mayor:
                        query = query & " AND (file.sys_length>" & gs_DBConTamanyoDe & ") "
                        
                End Select
            
                Select Case gt_DBDateSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.fecha<#" & s_ConFechaDe & "#) "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.fecha<#" & s_DiaSiguiente & "#) "
                        
                    Case db_Igual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) AND (file.fecha<#" & s_DiaSiguiente & "#) "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) "
                        
                    Case db_Mayor:
                        query = query & " AND (file.fecha>=#" & s_DiaSiguiente & "#) "
                        
                End Select
                
                query = query & " AND (file.id_file_type=" & gl_DBConTipoDe & ")) "
                
            Else
            
                Select Case gt_DBFileSizeSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.sys_length<" & gs_DBConTamanyoDe & ") "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.sys_length<=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Igual:
                        query = query & " AND (file.sys_length=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.sys_length>=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Mayor:
                        query = query & " AND (file.sys_length>" & gs_DBConTamanyoDe & ") "
                        
                End Select
            
                Select Case gt_DBDateSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.fecha<#" & s_ConFechaDe & "#)) "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.fecha<#" & s_DiaSiguiente & "#)) "
                        
                    Case db_Igual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) AND (file.fecha<#" & s_DiaSiguiente & "#)) "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#)) "
                        
                    Case db_Mayor:
                        query = query & " AND (file.fecha>=#" & s_DiaSiguiente & "#)) "
                        
                End Select
            
            End If
            
        Else    '(gb_DBConFecha = False)
            
            If gb_DBConTipo = True Then
    
                Select Case gt_DBFileSizeSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.sys_length<" & gs_DBConTamanyoDe & ") "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.sys_length<=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Igual:
                        query = query & " AND (file.sys_length=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.sys_length>=" & gs_DBConTamanyoDe & ") "
                        
                    Case db_Mayor:
                        query = query & " AND (file.sys_length>" & gs_DBConTamanyoDe & ") "
                        
                End Select
            
                query = query & " AND (file.id_file_type=" & gl_DBConTipoDe & ")) "
                
            Else
            
                Select Case gt_DBFileSizeSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.sys_length<" & gs_DBConTamanyoDe & ")) "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.sys_length<=" & gs_DBConTamanyoDe & ")) "
                        
                    Case db_Igual:
                        query = query & " AND (file.sys_length=" & gs_DBConTamanyoDe & ")) "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.sys_length>=" & gs_DBConTamanyoDe & ")) "
                        
                    Case db_Mayor:
                        query = query & " AND (file.sys_length>" & gs_DBConTamanyoDe & ")) "
                        
                End Select
            
            End If
            
        End If
            
    Else    ' (gb_DBConTamanyo = False)
    
        If gb_DBConFecha = True Then
        
            If gb_DBConTipo = True Then
    
                Select Case gt_DBDateSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.fecha<#" & s_ConFechaDe & "#) "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.fecha<#" & s_DiaSiguiente & "#) "
                        
                    Case db_Igual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) AND (file.fecha<#" & s_DiaSiguiente & "#) "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) "
                        
                    Case db_Mayor:
                        query = query & " AND (file.fecha>=#" & s_DiaSiguiente & "#) "
                        
                End Select
                
                query = query & " AND (file.id_file_type=" & gl_DBConTipoDe & ")) "
                
            Else
            
                Select Case gt_DBDateSearchStyle
                                
                    Case db_Menor:
                        query = query & " AND (file.fecha<#" & s_ConFechaDe & "#)) "
                
                    Case db_MenorIgual:
                        query = query & " AND (file.fecha<#" & s_DiaSiguiente & "#)) "
                        
                    Case db_Igual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#) AND (file.fecha<#" & s_DiaSiguiente & "#)) "
                        
                    Case db_MayorIgual:
                        query = query & " AND (file.fecha>=#" & s_ConFechaDe & "#)) "
                        
                    Case db_Mayor:
                        query = query & " AND (file.fecha>=#" & s_DiaSiguiente & "#)) "
                        
                End Select
            
            End If
            
        Else    '(gb_DBConFecha = False)
            
            If gb_DBConTipo = True Then
    
                query = query & " AND (file.id_file_type=" & gl_DBConTipoDe & ")) "
                
            End If
            
        End If
    
    End If
    
    If Not gb_DBShowHiddenFiles = True Then
        
        query = query & " AND (file.hidden=0)) "
        
    Else
    
        query = query & ") "
        
    End If
    '--------------------------------------------------------------
    
    '==============================================================
    ' Agregar criterio de ordenacion
    '
    queryorder = "ORDER BY"
    
    For k = 1 To 5
        
        If gb_DBOrdenarEnabled(k) Then
            
            queryorder = queryorder & FieldOrder(gs_DBOrdenarCampo(k))
            
            If gb_DBOrdenarAsc(k) Then
                queryorder = queryorder & ","
            Else
                queryorder = queryorder & " DESC,"
            End If
            
        End If
        
    Next
    
    queryorder = Left(queryorder, Len(queryorder) - 1)
    
    query = query & queryorder
    '--------------------------------------------------------------
    
    m_strQuery = query
                
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    With flxResults
        
        .Row = 0
        
        'borramos todas las filas exceptuando la fija y la primera
        .Rows = 2
                
        If rs.EOF = True Then
            'primera fila invisible no se puede eliminar
            .RowHeight(1) = 0
            GoTo SALIR
        End If
        
        .Redraw = False
        
        k = 0
        b_Redraw = 1
        
        Dim textColor As Long
        
        While rs.EOF = False
            
            k = k + 1
            .Rows = .Rows + 1
            .Row = .Rows - 1
            'forzar visible
            .RowHeight(.Row) = -1
            
            '***************************
            ' ID
            '***************************
            .Col = 0
            .text = rs!id_file
            '***************************
            ' Numeracion
            '***************************
            .Col = 1
            .CellForeColor = m_NumberColor
            .text = k
            
            If gb_DBShowHiddenFiles Then
                If rs!Hidden = 1 Then
                    textColor = m_HiddenColor
                Else
                    textColor = m_NormalColor
                End If
            Else
                textColor = m_NormalColor
            End If
            
            '***************************
            ' Nombre
            '***************************
            .Col = 2
            .CellAlignment = flexAlignLeftCenter
            .CellForeColor = textColor
            If gb_DBNameFromFile And Not gb_DBShowPathInFileName Then
                .text = gfnc_GetFileNameWithoutPath(rs!title)
            Else
                .text = rs!title
            End If
            '***************************
            ' Autor
            '***************************
            .Col = 3
            .CellAlignment = flexAlignLeftCenter
            .CellForeColor = textColor
            .text = rs!Author
            '***************************
            ' Pertenece a
            '***************************
            .Col = 4
            .CellAlignment = flexAlignLeftCenter
            .CellForeColor = textColor
            .text = rs!belong_to
            '***************************
            ' Genero
            '***************************
            .Col = 5
            .CellAlignment = flexAlignLeftCenter
            .CellForeColor = textColor
            .text = rs!aux_field
            '***************************
            ' Prioridad, calidad,
            ' tamaño o fecha
            '***************************
            .Col = 6
            .CellAlignment = flexAlignRightCenter
            .CellForeColor = textColor
            Select Case gs_DBConCampoDe
                Case "Tamaño"
                    l_FileZise = CLng(rs!quantity)
                    If (l_FileZise > 1048576) Then
                        .text = Format(l_FileZise / 1048576, "0.00") & " MB"
                    Else
                        .text = Format(l_FileZise / 1024, "0.00") & " KB"
                    End If
                Case "Fecha"
                    '-------------------------------------------------------------------------------
                    ' [NOTA] Format(rs!quantity, "dd/mm/yyy") me seguia mostrando valores de hora
                    ' por eso tuve que usar el Left()...
                    '
                    .text = Left$(Format(rs!quantity, "dd/mm/yyy"), 8)
                    '-------------------------------------------------------------------------------
                Case Else
                    .text = rs!quantity
            End Select

            rs.MoveNext
            
            If b_Redraw = 1 Then
                If k >= CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) + 1 Then
                    .Redraw = True
                    .Refresh
                    .Redraw = False
                    b_Redraw = 0
                End If
            End If
            
            '---------------------------------------------------------------
            ' cancelar con [ESC]
            If 0 = (k Mod 100) Then
                lresult = PeekMessage(lpMsg, Me.hWnd, 256, 256, PM_REMOVE)
                If lresult <> 0 Then
                    If lpMsg.wParam = VK_ESCAPE Then
                        If vbYes = MsgBox("¿Estas seguro de cancelar?", vbExclamation + vbYesNo, "Confirmar") Then
                            .Redraw = True
                            .SetFocus
                            GoTo SALIR
                        End If
                    End If
                End If
            End If
            '---------------------------------------------------------------
            
        Wend
        
        '************************************************
        'eliminar la primera fila invisible
        '************************************************
        .RemoveItem (1)

SALIR:

        'seleccionar el primero
        .Row = 1
        .Col = 1
        .ColSel = 6
        
        .Redraw = True
        .SetFocus
    
    End With
        
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    flxResults.MousePointer = flexDefault
    
    Exit Sub
    
Handler:
    
    Select Case Err.Number
    
        Case 94, 13
            'uso no valido de NULL
            'cuando el campo esta vacio
            Resume Next
        
        Case 3709
        
            MsgBox "Se ha perdido conexión con la base de datos" & vbCrLf & "Verifique la conexión y vuelva a cargar el formulario.", vbExclamation, "Conexión perdida"
            Screen.MousePointer = vbDefault
            Unload Me
        
        Case Else
            MsgBox Err.Description, vbCritical, "Generar_Lista"
            flxResults.Redraw = True
            Screen.MousePointer = vbDefault
            flxResults.MousePointer = flexDefault
            
    End Select
    
End Sub

'------------------------------------------------------------------------------
' Actualiza los combos de Categoria, Genero y Pariente de este formulario
'
Public Sub Actualizar_CombosDeBusqueda()
    
    On Error GoTo Handler
    
    Actualizar_ComboCategoria
    Actualizar_ComboGenero
    Actualizar_ComboPariente
    
    Exit Sub
    
Handler:

    Select Case Err.Number
    
        Case 3709
            MsgBox "Se ha perdido conexión con la base de datos" & vbCrLf & "Verifique la conexión y vuelva a cargar el formulario.", vbExclamation, "Conexión perdida"
            Unload Me
        
        Case Else
            MsgBox Err.Description, vbCritical, ":Actualizar_CombosDeBusqueda()"
            
    End Select
    
End Sub

Private Sub Actualizar_ComboCategoria()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================

    On Error GoTo Handler

    Set rs = New ADODB.Recordset

    '--------------------------------
    'cargar la tabla de categorias
    '
    cmbCategory.Clear
    cmbCategory.AddItem "[Todos]"
    cmbCategory.ItemData(cmbCategory.NewIndex) = 0
    
    query = "SELECT * FROM category WHERE (id_category>0) ORDER BY category"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbCategory.AddItem rs!category
        cmbCategory.ItemData(cmbCategory.NewIndex) = rs!id_category
        rs.MoveNext
    Wend
    rs.Close
    
    cmbCategory.ListIndex = 0
    '------------------------------------

    Exit Sub

Handler:

    Err.Raise Number:=Err.Number

End Sub

Public Sub Actualizar_ComboPariente()
    Dim rs As ADODB.Recordset
    Dim id_category  As Long
    
    On Error GoTo Handler
    Set rs = New ADODB.Recordset
    
    ' Cargar la tabla de pertenece a
    cmbParent.Clear
    cmbParent.AddItem "[Todos]"
    
    id_category = cmbCategory.ItemData(cmbCategory.ListIndex)
    
    If 0 = id_category Then
        ' [Todos]
        If gb_DBPertenciaPorAlmacenamiento = True Then
            query = "SELECT * FROM storage WHERE active=1 ORDER BY name"
        Else
            query = "SELECT * FROM parent WHERE (id_parent > 0) ORDER BY parent"
        End If
    Else
        ' Categoria seleccionada
        If gb_DBPertenciaPorAlmacenamiento = True Then
            query = "SELECT * FROM storage WHERE ((active=1) AND (id_category=" & Trim(str(id_category)) & ")) ORDER BY id_storage DESC"
        Else
            query = "SELECT * FROM parent WHERE ((id_parent > 0) AND (id_category=" & Trim(str(id_category)) & ")) ORDER BY id_storage DESC"
        End If
        '--------------------------------------------------
    End If
    
    If gb_DBPertenciaPorAlmacenamiento = True Then
    
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        While rs.EOF = False
            cmbParent.AddItem rs!Name
            cmbParent.ItemData(cmbParent.NewIndex) = rs!id_storage
            rs.MoveNext
        Wend
            
    Else
        
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        While rs.EOF = False
            cmbParent.AddItem rs!Parent
            cmbParent.ItemData(cmbParent.NewIndex) = rs!id_parent
            rs.MoveNext
        Wend
        
    End If
    
    rs.Close
    cmbParent.ListIndex = 0
    Exit Sub
    
Handler:
    Err.Raise Number:=Err.Number
End Sub

Public Sub Actualizar_ComboGenero()
    '===================================================
    Dim rs As ADODB.Recordset
    Dim id_category As Long
    '===================================================
    
    On Error GoTo Handler
    
    Set rs = New ADODB.Recordset
    
    '--------------------------------
    ' cargar la tabla de géneros
    '
    cmbGenre.Clear
    cmbGenre.AddItem "[Todos]"

    id_category = cmbCategory.ItemData(cmbCategory.ListIndex)

    If 0 = id_category Then
        '--------------------------------------------------
        '[Todos]
        '
        query = "SELECT * FROM genre WHERE active=1 ORDER BY genre"
        '--------------------------------------------------
    Else
        '--------------------------------------------------
        'Categoria seleccionada
        '
        query = "SELECT * FROM genre WHERE ((active=1) AND (id_category=" & Trim(str(id_category)) & ")) ORDER BY genre"
        '--------------------------------------------------
    End If
    
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While rs.EOF = False
        cmbGenre.AddItem rs!genre
        cmbGenre.ItemData(cmbGenre.NewIndex) = rs!id_genre
        rs.MoveNext
    Wend
    rs.Close
    
    cmbGenre.ListIndex = 0
    
    Exit Sub
    
Handler:

    Err.Raise Number:=Err.Number
    
End Sub

'--------------------------------------------------------------------------
' Elimina las filas seleccionadas del flxgrid de resultados
' Renumera los resultados
'
Private Sub Eliminar_Resultados()

    Dim k As Long
    Dim a As Long
    Dim b As Long
    On Error GoTo Handler
    
    With flxResults
        
        a = .Row
        b = .RowSel
    
        If b < a Then
            k = b
            b = a
            a = k
        End If
            
        For k = b To a Step -1
            .RemoveItem (k)
        Next
        
        '===================================================
        ' re-numerar (optimizar...)
        '
        .Redraw = False
        .Col = 1
        
        For k = 1 To .Rows - 1
            .Row = k
            If .RowHeight(k) <> 0 Then
                .text = k
            End If
        Next
        
        .Redraw = True
        '---------------------------------------------------
        
        .Row = a
        .ColSel = .Cols - 1
    
        'verificar si la fila esta visible
        If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
            .TopRow = .Row
        End If

        .SetFocus
    
    End With
    
    Exit Sub
    
Handler:

    If Err.Number = 30015 Then
        '--------------------------------------------------------
        ' no se pude quitar la ultima fila no fija
        ' (cuando se intenta eliminar la primera fila visible)
        '
        flxResults.RowHeight(k) = 0
        Exit Sub
    Else
        If Err.Number = 30009 Then
            '--------------------------------------------------------
            ' valor de fila no valido
            ' (cuando se borra la ultima fila)
            '
            Resume Next
        Else
            MsgBox Err.Description, vbExclamation, "Eliminar_Resultados()"
        End If
    End If
    
    
End Sub

'--------------------------------------------------------------------------
' Establece el los archivos de las filas seleccionadas como ocultos
' Renumera los resultados
'
Private Sub Ocultar_Registro(ByRef bolHide As Boolean)
    '===================================================
    Dim id_registro As Long
    Dim fila_eliminada As String
    Dim nombre_fila As String
    Dim fila_eliminada_2 As String
    Dim nombre_fila_2 As String
    Dim cd As ADODB.Command
    Dim oldRow As Long
    Dim a As Long
    Dim b As Long
    Dim k As Long
    Dim q As Long
    '===================================================

    On Error GoTo Handler

    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------

    'ocultar (bolHide = True) o mostrar (bolHide = False) registro seleccionado
    With flxResults
            
        If .Row <> 0 And .RowHeight(.Row) <> 0 Then
    
            .Redraw = False
            
            a = .Row
            b = .RowSel
        
            oldRow = .TopRow
        
            If b = a Then
                
                .Col = 0
                id_registro = CLng(.text)
                .Col = 1
                fila_eliminada = .text
                .Col = 2
                nombre_fila = .text
                
                .Col = 1
                .ColSel = .Cols - 1

            Else
                
                If b < a Then
                    k = b
                    b = a
                    a = k
                End If
                
                .Row = b
                                
                .Col = 1
                fila_eliminada_2 = .text
                .Col = 2
                nombre_fila_2 = .text
                
                .Row = a
                                
                .Col = 1
                fila_eliminada = .text
                .Col = 2
                nombre_fila = .text
                
                .Col = 1
                .ColSel = .Cols - 1
                
                .RowSel = b
                
            End If
            
            .Redraw = True
            
            'elimina un incomodo movimiento cuando .row esta en las ultimas filas
            .TopRow = oldRow
            
            Dim strMsg As String
            
            If a = b Then
            
                If bolHide Then
                    strMsg = "ocultar el"
                Else
                    strMsg = "quitar el atributo de oculto del"
                End If
                
                If vbYes = MsgBox("Estás a punto de " & strMsg & " registro [" & fila_eliminada & "]:" & vbCrLf & nombre_fila & vbCrLf & "¿Estás seguro de continuar?", vbInformation + vbYesNo, "Confirmar") Then
            
                    Set cd = New ADODB.Command
                    Set cd.ActiveConnection = cn
                    
                    If bolHide Then
                        cd.CommandText = "UPDATE file SET hidden=1 WHERE id_file=" & id_registro
                    Else
                        cd.CommandText = "UPDATE file SET hidden=0 WHERE id_file=" & id_registro
                    End If
                    cd.Execute
                    
                    If gb_DBShowHiddenFiles Then
                        'solo cambiar de color
                        .Row = a
                        For q = 2 To .Cols - 1
                            .Col = q
                            If bolHide Then
                                .CellForeColor = m_HiddenColor
                            Else
                                .CellForeColor = m_NormalColor
                            End If
                        Next q
                    Else
                        If .Rows > 2 Then
                            oldRow = .Row
                            .RemoveItem (.Row)
                        Else
                            'volver invisible la primera fila
                            .RowHeight(1) = 0
                            Limpiar_Detalles
                            Exit Sub
                        End If
                    End If
                    
                Else
                    Exit Sub
                End If
                
            Else
            
                If bolHide Then
                    strMsg = "ocultar "
                Else
                    strMsg = "quitar el atributo de oculto de "
                End If
                
                If vbYes = MsgBox("Estás a punto de " & strMsg & Trim(str(b - a + 1)) & " registros." & vbCrLf & "Desde el registro [" & fila_eliminada & "] hasta el [" & fila_eliminada_2 & "]" & vbCrLf & "¿Estás seguro de continuar?", vbInformation + vbYesNo, "Confirmar") Then
            
                    Set cd = New ADODB.Command
                    Set cd.ActiveConnection = cn
                    
                    For k = b To a Step -1
                        
                        .Row = k
                        .Col = 0
                        id_registro = CLng(.text)
                        
                        If bolHide Then
                            cd.CommandText = "UPDATE file SET hidden=1 WHERE id_file=" & id_registro
                        Else
                            cd.CommandText = "UPDATE file SET hidden=0 WHERE id_file=" & id_registro
                        End If
                        cd.Execute
    
                        If gb_DBShowHiddenFiles Then
                            'solo cambiar de color
                            .Row = k
                            For q = 2 To .Cols - 1
                                .Col = q
                                If bolHide Then
                                    .CellForeColor = m_HiddenColor
                                Else
                                    .CellForeColor = m_NormalColor
                                End If
                            Next q
                        Else
                            If .Rows > 2 Then
                                .RemoveItem (.Row)
                            Else
                                'volver invisible la primera fila
                                .RowHeight(1) = 0
                                Limpiar_Detalles
                                Exit Sub
                            End If
                        End If
                        
                    Next
                    
                    oldRow = a

                Else
                    Exit Sub
                End If
            
            End If
                
            '===================================================
            ' re-numerar (optimizar...)
            '
            If Not gb_DBShowHiddenFiles Then
                .Redraw = False
                .Col = 1
                
                For k = 1 To .Rows - 1
                    
                    .Row = k
                
                    If .RowHeight(k) <> 0 Then
                        .text = k
                    End If
                    
                Next
            End If
            
            .Redraw = True
            '---------------------------------------------------
                
            'seleccionar la fila anterior
            If oldRow > .Rows - 1 Then
                .Row = oldRow - 1
            Else
                .Row = oldRow
            End If
            
            .Col = 1
            .ColSel = .Cols - 1
                
            'verificar si la fila esta visible
            If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
                .TopRow = .Row
            End If

            .SetFocus
        
        End If
            
    End With
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "Ocultar_Registro()"
    flxResults.Redraw = True

End Sub

Private Sub Eliminar_Registro()
    '===================================================
    Dim id_registro As Long
    Dim fila_eliminada As String
    Dim nombre_fila As String
    Dim fila_eliminada_2 As String
    Dim nombre_fila_2 As String
    Dim cd As ADODB.Command
    Dim oldRow As Long
    Dim a As Long
    Dim b As Long
    Dim k As Long
    '===================================================

    On Error GoTo Handler

    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------

    'eliminar registro seleccionado
    With flxResults
            
        If .Row <> 0 And .RowHeight(.Row) <> 0 Then
    
            .Redraw = False
            
            a = .Row
            b = .RowSel
        
            oldRow = .TopRow
        
            If b = a Then
                
                .Col = 0
                id_registro = CLng(.text)
                .Col = 1
                fila_eliminada = .text
                .Col = 2
                nombre_fila = .text
                
                .Col = 1
                .ColSel = .Cols - 1

            Else
                
                If b < a Then
                    k = b
                    b = a
                    a = k
                End If
                
                .Row = b
                                
                .Col = 1
                fila_eliminada_2 = .text
                .Col = 2
                nombre_fila_2 = .text
                
                .Row = a
                                
                .Col = 1
                fila_eliminada = .text
                .Col = 2
                nombre_fila = .text
                
                .Col = 1
                .ColSel = .Cols - 1
                
                .RowSel = b
                
            End If
            
            .Redraw = True
            
            'elimina un incomodo movimiento cuando .row esta en las ultimas filas
            .TopRow = oldRow
            
            If a = b Then
            
                If vbYes = MsgBox("¿Estás seguro de eliminar el registro [" & fila_eliminada & "]:" & vbCrLf & nombre_fila & "?", vbExclamation + vbYesNo, "Confirmar eliminación") Then
            
                    Set cd = New ADODB.Command
                    Set cd.ActiveConnection = cn
                    
                    cd.CommandText = "DELETE FROM file WHERE id_file=" & id_registro
                    cd.Execute
                    
                    If .Rows > 2 Then
                        oldRow = .Row
                        .RemoveItem (.Row)
                    Else
                        'volver invisible la primera fila
                        .RowHeight(1) = 0
                        Limpiar_Detalles
                        Exit Sub
                    End If
                    
                Else
                    Exit Sub
                End If
                
            Else
            
                If vbYes = MsgBox("¿Estás seguro de eliminar " & Trim(str(b - a + 1)) & " registros?" & vbCrLf & "Desde el registro [" & fila_eliminada & "] hasta el [" & fila_eliminada_2 & "]", vbExclamation + vbYesNo, "Confirmar eliminación") Then
            
                    Set cd = New ADODB.Command
                    Set cd.ActiveConnection = cn
                    
                    For k = b To a Step -1
                        
                        .Row = k
                        .Col = 0
                        id_registro = CLng(.text)
                        
                        cd.CommandText = "DELETE FROM file WHERE id_file=" & id_registro
                        cd.Execute
    
                        If .Rows > 2 Then
                            .RemoveItem (.Row)
                        Else
                            'volver invisible la primera fila
                            .RowHeight(1) = 0
                            Limpiar_Detalles
                            Exit Sub
                        End If
                        
                    Next
                    
                    oldRow = a

                Else
                    Exit Sub
                End If
            
            End If
                
            '===================================================
            ' re-numerar (optimizar...)
            '
            .Redraw = False
            .Col = 1
            
            For k = 1 To .Rows - 1
                
                .Row = k
            
                If .RowHeight(k) <> 0 Then
                    .text = k
                End If
                
            Next
            
            .Redraw = True
            '---------------------------------------------------
                
            'seleccionar la fila anterior
            If oldRow > .Rows - 1 Then
                .Row = oldRow - 1
            Else
                .Row = oldRow
            End If
            
            .Col = 1
            .ColSel = .Cols - 1
                
            'verificar si la fila esta visible
            If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
                .TopRow = .Row
            End If

            .SetFocus
        
        End If
            
    End With
    
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "Eliminar_Registro()"
    flxResults.Redraw = True
End Sub

Private Sub Agregar_Registro()
    '===================================================
    Dim frmAddRegister As Form
    '===================================================

    On Error GoTo Handler

    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------
    
    If Not gb_AddRegisterStarted Then
    
EX_NEW_FORM:

        Set frmAddRegister = New frmRegistros
        
        gb_AddRegisterStarted = True
        gb_DBAddRegister = True
        
        frmAddRegister.Caption = "[Nuevo Registro]"
        frmAddRegister.Show vbModeless
        frmAddRegister.SetFocus
        
    Else
    
        For Each frmAddRegister In Forms
            If frmAddRegister.Caption = "[Nuevo Registro]" Then
                frmAddRegister.Visible = True
                frmAddRegister.SetFocus
                Exit Sub
            End If
        Next
        
        '-------------------------------------------------------------
        ' Es posible que la variable [gb_AddRegisterStarted] se quede
        ' mal si el intento de carga se cancelo por falta de memoria..
        '
        GoTo EX_NEW_FORM
        
    End If
    
    Exit Sub

Handler:
    
    MsgBox Err.Description, vbCritical, "Agregar_Registro"

End Sub

Private Sub Modificar_Registro()
    '===================================================
    Dim frmEditRegistros As Form
    Dim mycad As String
    '===================================================

    On Error GoTo Handler

    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------

    With flxResults

        If .Row <> 0 And .RowHeight(.Row) <> 0 Then

            .Redraw = False
            
            .Col = 0
            gl_DB_IDRegistroModificar = CLng(.text)
            .Col = 2
            mycad = .text

            .Col = 1
            .ColSel = 6
            .Redraw = True
            
            gb_DBAddRegister = False

            Set frmEditRegistros = New frmRegistros
            frmEditRegistros.Caption = "[Editando] :: [" & mycad & "]"
            frmEditRegistros.gml_RowEditRegister = .Row
            frmEditRegistros.Show vbModeless
            frmEditRegistros.SetFocus

        End If

    End With
    
    Exit Sub

Handler:
    
    MsgBox Err.Description, vbCritical, "Modificar_Registro"
    flxResults.Redraw = True
    
End Sub

Private Sub Mostrar_Detalles()
    '===================================================
    Dim rs As ADODB.Recordset
    Dim oldRow As Long
    Dim sApellidoP As String
    Dim sApellidoM As String
    Dim sName As String
    Dim sAuthor As String
    Dim id_file As Long
    Dim sAddress As String
    Dim sDistrict As String
    Dim sCity As String
    Dim sObserv As String
    Dim myPicture As Object
    Dim nPrioridad As Byte
    Dim strFileName As String
    Dim strFileExtention As String
    '===================================================
    
    On Error GoTo Handler
    
    '--------------------------------
    ' verificar conexion y BD
    If gb_DBConexionOK Then
        If gb_DBFormatOK Then
            ' todo OK
        Else
            gsub_ShowMessageWrongDB
            Exit Sub
        End If
    Else
        gsub_ShowMessageNoConection
        Exit Sub
    End If
    '--------------------------------
    
    If mb_DetalleMostrado = False Then
        cmdVerDetalle_Click
    End If

    With flxResults

        If .RowHeight(.Row) = 0 Or .Row = 0 Then
            Exit Sub
        End If

        .Redraw = False

        oldRow = .TopRow
        
        .Col = 0
        id_file = CLng(.text)
        .Col = 2
        sName = .text
        .Col = 3
        sAuthor = .text
        .Col = 1
        .ColSel = 6

        .Redraw = True
        
        'elimina un incomodo movimiento cuando .row esta en las ultimas filas
        .TopRow = oldRow
        
        If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
            .TopRow = .Row
        End If

    End With

    flxDetails.Row = 0
    flxDetails.Col = 2
    flxDetails.text = sName

    'buscar los otros detalles
    query = "SELECT file.sys_name, file.priority, file.sys_length, "
    query = query & "file.fecha, file.quality, file.type_quality, "
    query = query & "file.comment, file.type_comment, "
    query = query & "file_type.file_type, parent.parent, genre.genre, sub_genre.sub_genre, "
    query = query & "storage.name, storage.label, storage.serial, "
    query = query & "storage_type.storage_type "
    query = query & "FROM file, file_type, parent, genre, sub_genre, storage, storage_type "
    query = query & "WHERE ((file.id_file=" & id_file & ") AND "
    query = query & "(file_type.id_file_type=file.id_file_type) AND "
    query = query & "(parent.id_parent=file.id_parent) AND "
    query = query & "(genre.id_genre=file.id_genre) AND "
    query = query & "(sub_genre.id_sub_genre=file.id_sub_genre) AND "
    query = query & "(storage.id_storage=file.id_storage) AND "
    query = query & "(storage_type.id_storage_type=storage.id_storage_type))"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF = True Then
        MsgBox "Probablemente el registro hace referencia a campos" & vbCrLf & "que ya no están en la BD o han sido modificados externamente." & vbCrLf & "Intente corregir el error editando el registro manualmente.", vbExclamation, "[Error] :: [No se pueden mostrar detalles del registro]"
        rs.Close
        Set rs = Nothing
        Limpiar_Detalles
        Exit Sub
    End If

    With flxDetails

        .Col = 2

        .Row = 1
        .WordWrap = True
        .CellAlignment = flexAlignLeftTop
        strFileName = rs!Sys_Name
        .text = strFileName
        
        .Row = 2
        .CellAlignment = flexAlignLeftTop
        .text = rs!Fecha
        
        .Row = 3
        .CellAlignment = flexAlignLeftTop
        .text = Format(rs!sys_length, "###,###,##0 Bytes")
        
        .Row = 4
        .CellAlignment = flexAlignLeftTop
        strFileExtention = UCase(rs!file_type)
       
        If strFileExtention <> "<DIR>" Then
            .text = "[" & strFileExtention & "]" & "  " & GetFileTypeDescription(strFileName)
        Else
            .text = "Directorio"
        End If
        
        .Row = 5
        'insertar imagen
        nPrioridad = rs!Priority
        If nPrioridad <> 0 Then
            If nPrioridad < 5 Then
                Set .CellPicture = imglstA.ListImages.Item(nPrioridad).Picture
            Else
                Set .CellPicture = imglstA.ListImages.Item(5).Picture
            End If
        Else
            Set .CellPicture = Nothing
        End If

        .Row = 6
        .WordWrap = True
        .CellAlignment = flexAlignLeftTop
        .text = sAuthor

        .Row = 7
        .CellAlignment = flexAlignLeftTop
        .text = rs!genre & NL & rs!sub_genre

        .Row = 8
        .CellAlignment = flexAlignLeftTop
        .text = rs!Quality & " " & rs!type_quality

        .Row = 9
        .WordWrap = True
        .CellAlignment = flexAlignLeftTop
        .text = rs!Parent

        .Row = 10
        .WordWrap = True
        .CellAlignment = flexAlignLeftTop
        .text = "[" & rs!storage_type & "]   " & rs!Name & NL & rs!Label & NL & rs!serial
        
        .Row = 11
        .WordWrap = True
        .CellAlignment = flexAlignLeftTop
        .text = rs!type_comment & vbCrLf & rs!Comment

        .Col = 1
        .Row = 2

    End With

    rs.Close
    Set rs = Nothing

    Exit Sub

Handler:
    
    Select Case Err.Number
    
        Case 94
            'uso no valido de NULL
            'cuando el campo esta vacio
            Resume Next
        
        Case Else
            MsgBox Err.Description, vbCritical, "Mostrar_Detalles()"
        
    End Select
    
End Sub

Public Sub Actualizar_Lista(rowEdit As Long)
    '===================================================
    Dim l_FileZise As Long
    '===================================================
    
    On Error Resume Next
    
    With flxResults

        .Row = rowEdit
        
        '***************************
        'nombre
        '***************************
        .Col = 2
        .CellAlignment = flexAlignLeftCenter
        .text = gs_DBRegName
        '***************************
        ' Autor
        '***************************
        .Col = 3
        .CellAlignment = flexAlignLeftCenter
        .text = gs_DBRegAuthor
        '***************************
        ' (Medio / Pertenece a)
        '***************************
        .Col = 4
        .CellAlignment = flexAlignLeftCenter
        If gb_DBPertenciaPorAlmacenamiento Then
            .text = gs_DBRegStorage
        Else
            .text = gs_DBRegParent
        End If
        '***************************
        ' (Genero / Tipo)
        '***************************
        .Col = 5
        .CellAlignment = flexAlignLeftCenter
        If gb_DBCampoAuxiliarPorGenero Then
            .text = gs_DBRegGenre
        Else
            .text = gs_DBRegFileType
        End If
        '***************************
        ' Prioridad, calidad,
        ' tamaño o fecha
        '***************************
        .Col = 6
        .CellAlignment = flexAlignRightCenter
        
        Select Case gs_DBConCampoDe
            Case "Tamaño"
                l_FileZise = CLng(gs_DBRegFileSize)
                If (l_FileZise > 1048576) Then
                    .text = Format(l_FileZise / 1048576, "0.00") & " MB"
                Else
                    .text = Format(l_FileZise / 1024, "0.00") & " KB"
                End If
            Case "Fecha"
                '-------------------------------------------------------------------------------
                ' [NOTA] Format(rs!quantity, "dd/mm/yyy") me seguia mostrando valores de hora
                ' por eso tuve que usar el Left()...
                '
                .text = Left$(Format(gd_DBRegFileDate, "dd/mm/yyy"), 8)
                '-------------------------------------------------------------------------------
            Case "Prioridad"
                .text = gs_DBRegPriority
            Case "Calidad"
                .text = gs_DBRegQuality
        End Select
        
        .Col = 1
        .ColSel = .Cols - 1
        
        '**************************************************************
        ' TO DO: modificarlo para usar variables globales (mas rapido)
        ' La actualizacion del registro no es instantanea y aveces
        ' no se muestra correctamente (pues aun no esta actualizado)
        '
        Sleep DB_EDIT_SLEEP
        '**************************************************************

        Mostrar_Detalles

        .SetFocus
        
    End With

End Sub

Public Sub Agregar_Lista()
    '===================================================
    Dim l_FileZise As Long
    '===================================================
    
    On Error Resume Next
    
    With flxResults

        If .Rows = 2 And .RowHeight(1) = 0 Then
            .RowHeight(1) = -1
            .Row = 1
        Else
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .RowHeight(.Row) = -1
        End If
        
        '***************************
        ' ID
        '***************************
        .Col = 0
        .text = gl_DBRegIDNew
        '***************************
        ' Numeracion
        '***************************
        .Col = 1
        .CellForeColor = RGB(100, 170, 255)
        .CellAlignment = flexAlignRightCenter
        .text = "+" & Trim(.Rows - 1)
        '***************************
        'nombre
        '***************************
        .Col = 2
        .CellAlignment = flexAlignLeftCenter
        .text = gs_DBRegName
        '***************************
        ' Autor
        '***************************
        .Col = 3
        .CellAlignment = flexAlignLeftCenter
        .text = gs_DBRegAuthor
        '***************************
        ' (Medio / Pertenece a)
        '***************************
        .Col = 4
        .CellAlignment = flexAlignLeftCenter
        If gb_DBPertenciaPorAlmacenamiento Then
            .text = gs_DBRegStorage
        Else
            .text = gs_DBRegParent
        End If
        '***************************
        ' (Genero / Tipo)
        '***************************
        .Col = 5
        .CellAlignment = flexAlignLeftCenter
        If gb_DBCampoAuxiliarPorGenero Then
            .text = gs_DBRegGenre
        Else
            .text = gs_DBRegFileType
        End If
        '***************************
        ' Prioridad, calidad,
        ' tamaño o fecha
        '***************************
        .Col = 6
        .CellAlignment = flexAlignRightCenter
        
        Select Case gs_DBConCampoDe
            Case "Tamaño"
                l_FileZise = CLng(gs_DBRegFileSize)
                If (l_FileZise > 1048576) Then
                    .text = Format(l_FileZise / 1048576, "0.00") & " MB"
                Else
                    .text = Format(l_FileZise / 1024, "0.00") & " KB"
                End If
            Case "Fecha"
                '-------------------------------------------------------------------------------
                ' [NOTA] Format(rs!quantity, "dd/mm/yyy") me seguia mostrando valores de hora
                ' por eso tuve que usar el Left()...
                '
                .text = Left$(Format(gd_DBRegFileDate, "dd/mm/yyy"), 8)
                '-------------------------------------------------------------------------------
            Case "Prioridad"
                .text = gs_DBRegPriority
            Case "Calidad"
                .text = gs_DBRegQuality
        End Select
        
        .Col = 1
        .ColSel = .Cols - 1
       
        '**************************************************************
        ' TO DO: modificarlo para usar variables globales (mas rapido)
        ' La actualizacion del registro no es instantanea y aveces
        ' no se muestra correctamente (pues aun no esta actualizado)
        '
        Sleep DB_EDIT_SLEEP
        '**************************************************************
        
        Mostrar_Detalles

        .SetFocus
        
    End With

End Sub


Private Sub Limpiar_Detalles()
    With flxDetails
        .Col = 2
        .Row = 0
        .text = ""
        .Row = 1
        .text = ""
        .Row = 2
        .text = ""
        .Row = 3
        .text = ""
        .Row = 4
        .text = ""
        .Row = 5
        Set .CellPicture = Nothing
        .Row = 6
        .text = ""
        .Row = 7
        .text = ""
        .Row = 8
        .text = ""
        .Row = 9
        .text = ""
        .Row = 10
        .text = ""
        .Row = 11
        .text = ""
        .Col = 1
    End With
End Sub

Private Sub Sleep(ByVal secs As Double)
    
    Dim dblEndTime As Double
   
    dblEndTime = Timer + secs
    
    Do While dblEndTime > Timer
        ' permite a las otras aplicaciones procesar sus eventos.
        DoEvents
    Loop

End Sub

Public Sub Establecer_TitulosFlex()

    With flxResults
        
        .Row = 0
        .Col = 4
        If gb_DBPertenciaPorAlmacenamiento = True Then
            .text = "Medio"
        Else
            .text = "Pertenece a"
        End If
        
        .Col = 5
        If gb_DBCampoAuxiliarPorGenero = True Then
            .text = "Género"
        Else
            .text = "Tipo"
        End If
        
        .Col = 6
        .text = gs_DBConCampoDe
    End With

End Sub


Private Sub Ejecutar_Archivo()

    Dim hins As Long
    Dim file_path As String
    Dim id_file As Long
    Dim oldRow As Long
    Dim rs As ADODB.Recordset
    On Error GoTo Handler
    
    With flxResults

        If .RowHeight(.Row) = 0 Or .Row = 0 Then
            Exit Sub
        End If

        .Redraw = False
        oldRow = .TopRow
        
        .Col = 0
        id_file = CLng(.text)
        .Col = 1
        .ColSel = .Cols - 1

        .Redraw = True
        
        ' Elimina un incomodo movimiento cuando .row esta en las ultimas filas
        .TopRow = oldRow
        
        If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
            .TopRow = .Row
        End If
    End With
    
    query = "SELECT sys_name FROM file WHERE (id_file=" & id_file & ")"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF = False Then
        '--------------------------------------------------------------------------
        '[NOTA] (el SDK de Win98 dice que en vez de SW_NORMAL debe ir 0)
        file_path = Trim(rs!Sys_Name)
        hins = ShellExecute(vbNull, "open", file_path, vbNull, vbNull, 0)
        
        If (hins < 33) Then
            '--------------------------------------------------------------------------
            ' intentar con el drive opcional...
            file_path = gs_OptionalDrive & Mid(file_path, 2)
            hins = ShellExecute(Me.hWnd, "open", file_path, vbNull, vbNull, 0)
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
Handler:
    Exit Sub

End Sub

'--------------------------------------------------------------------------
' [NOTA]    Esta funcion esta ligada al querybase del procedimiento
'           Generar_Lista() de frmDataControl
Private Function FieldOrder(ByVal FieldName As String) As String
    Select Case FieldName
        Case "Cantidad"
            FieldOrder = " 3"
        Case "Aux"
            FieldOrder = " 6"
        Case "Parent"
            FieldOrder = " 5"
        Case "Autor"
            FieldOrder = " 4"
        Case "Nombre"
            FieldOrder = " 2"
    End Select
End Function

'*******************************************************************************
' Funcion que muestra el dialogo guardar como
Private Sub Guardar_Archivo_Como()
    
    Dim path_file As String
    Dim name_file As String
    Dim id_file As Long
    Dim oldRow As Long
    Dim rs As ADODB.Recordset
    On Error GoTo Handler
    
    With flxResults

        If .RowHeight(.Row) = 0 Or .Row = 0 Then
            Exit Sub
        End If

        .Redraw = False
        oldRow = .TopRow
        
        .Col = 0
        id_file = CLng(.text)
        .Col = 1
        .ColSel = .Cols - 1

        .Redraw = True
        
        ' Elimina un incomodo movimiento cuando .row esta en las ultimas filas
        .TopRow = oldRow
        
        If (.Row < .TopRow) Or (.Row >= .TopRow + CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) - 1) Then
            .TopRow = .Row
        End If

    End With
    
    query = "SELECT sys_name FROM file WHERE (id_file=" & id_file & ")"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF = False Then
        '-------------------------------------------------
        ' Extraer el nombre del archivo para guardarlo
        path_file = rs!Sys_Name
        name_file = Right(path_file, Len(path_file) - InStrRev(path_file, "\"))
        ex_SaveFileAs path_file, name_file
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
Handler:    MsgBox Err.Description, vbExclamation, "Error guardando archivo"
End Sub

Private Sub ex_SaveFileAs(ByVal file_path As String, ByVal File_Name As String)
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
    
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Salvar el archivo como:"
        .Filter = "Todos los Archivos(*.*)|*.*"
        'necesario para controlar la extension con que se salvaran los archivos
        'sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .InitDir = gs_CopyPath
        'nombre del reporte inicial
        .filename = File_Name
        .ShowSave
        
        If .filename <> "" Then
            '------------------------------------------------------
            ' salvar archivo y establecer directorio para guardar
            '
            gs_CopyPath = Left(.filename, InStrRev(.filename, "\"))
            FileCopy file_path, .filename
        End If
        
    End With
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, vbExclamation, "Error al guardar archivo"
    End If
End Sub

'*******************************************************************************
' Funciones para el mouse wheel scroll
'*******************************************************************************
Public Sub ScrollUp()

    Dim flexGrid As MSHFlexGrid
    On Error Resume Next
    
    If Me.ActiveControl.Name = "flxResults" Then Set flexGrid = flxResults
    If Me.ActiveControl.Name = "flxDetails" Then Set flexGrid = flxDetails
    
    If flexGrid.Tag > 1 Then
        flexGrid.Tag = flexGrid.TopRow - 4
        If flexGrid.Tag <= 1 Then
            flexGrid.Tag = 1
        End If
        flexGrid.TopRow = flexGrid.Tag
    End If
End Sub

Public Sub ScrollDown()
    
    Dim flexGrid As MSHFlexGrid
    On Error Resume Next
    
    If Me.ActiveControl.Name = "flxResults" Then Set flexGrid = flxResults
    If Me.ActiveControl.Name = "flxDetails" Then Set flexGrid = flxDetails

    If flexGrid.Tag < flexGrid.Rows - 4 Then
        flexGrid.Tag = flexGrid.TopRow + 4
        If flexGrid.Tag > flexGrid.Rows - 1 Then
            flexGrid.Tag = flexGrid.Rows - 1
        End If
        flexGrid.TopRow = flexGrid.Tag
    End If
End Sub

'*******************************************************************************
' Funciones para mostrar icono asociado
'*******************************************************************************
Private Function IconToPicture(hIcon As Long) As IPictureDisp
    '================================
    Dim cls_id As CLSID
    Dim hRes As Long
    Dim new_icon As TypeIcon
    Dim lpUnk As IUnknown
    '================================
    
    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    
    If hRes = 0 Then Set IconToPicture = lpUnk
    
End Function

Private Function GetFileIcon(ByRef strFileName As String, ByRef iconSize As Long, ByRef SHFILEINFO As SHFILEINFO) As IPictureDisp
    '================================
    Dim hIcon As Long
    '================================

    SHGetFileInfo strFileName, 0, SHFILEINFO, Len(SHFILEINFO), SHGFI_ICON + SHGFI_USEFILEATTRIBUTES + iconSize
    
    hIcon = SHFILEINFO.hIcon
    
    Set GetFileIcon = IconToPicture(hIcon)
    
End Function

Private Function GetFileTypeDescription(ByRef strFileName As String) As String
    '================================
    Dim shInfo As SHFILEINFO
    '================================

    SHGetFileInfo strFileName, 0, shInfo, Len(shInfo), SHGFI_TYPENAME + SHGFI_USEFILEATTRIBUTES
    
    GetFileTypeDescription = shInfo.szTypeName
    
End Function

'********************************************************************************************
' FUNCIONES DE EDICION COMPLEMENTARIAS
'********************************************************************************************
' eliminar todos los generos sin archivos asociados
Private Sub EraseAllEmptyGenres()
    '================================
    Dim rs As ADODB.Recordset
    Dim cd As ADODB.Command
    Dim strName As String
    Dim strNameFix As String
    Dim id As Long
    '================================
    On Error GoTo Handler
    
    If vbNo = MsgBox("Esta a punto de eliminar todos los generos" & vbCrLf & _
                      "que no tienen registros asociado." & vbCrLf & _
                      "¿Desea continuar?", vbExclamation + vbYesNo, "Advertencia") Then
        Exit Sub
    End If
    
    query = "SELECT genre.genre, genre.id_genre, COUNT(file.id_genre) " & _
            "FROM file RIGHT JOIN genre ON (file.id_genre=genre.id_genre) " & _
            "GROUP BY genre.genre, genre.id_genre, file.id_genre " & _
            "HAVING COUNT(file.id_genre)=0"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set cd = New ADODB.Command
    Set cd.ActiveConnection = cn

    While rs.EOF = False

        id = rs!id_genre
        
        cd.CommandText = "DELETE FROM genre WHERE id_genre=" & id
        cd.Execute
        
        rs.MoveNext
    Wend
    rs.Close
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbCritical, "EraseAllAuthorWithoutData"
End Sub

' eliminar todos los autores sin archivos asociados
Private Sub EraseAllAuthorWithoutData()
    '================================
    Dim rs As ADODB.Recordset
    Dim cd As ADODB.Command
    Dim strName As String
    Dim strNameFix As String
    Dim id As Long
    '================================
    On Error GoTo Handler
    
    If vbNo = MsgBox("Esta a punto de eliminar todos los autores" & vbCrLf & _
                      "que no tienen registros asociado." & vbCrLf & _
                      "¿Desea continuar?", vbExclamation + vbYesNo, "Advertencia") Then
        Exit Sub
    End If
    
    query = "SELECT  author.author, author.id_author, COUNT(file.id_author) " & _
            "FROM file RIGHT JOIN author ON (file.id_author=author.id_author) " & _
            "GROUP BY author.author, author.id_author, file.id_author " & _
            "HAVING COUNT(file.id_author) = 0"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set cd = New ADODB.Command
    Set cd.ActiveConnection = cn

    While rs.EOF = False

        id = rs!id_author
        
        cd.CommandText = "DELETE FROM author WHERE id_author=" & id
        cd.Execute
        
        rs.MoveNext
    Wend
    rs.Close
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbCritical, "EraseAllAuthorWithoutData"
End Sub

' convertir todos los nombres de archivo seleccionados a minusculas
Private Sub RenameAllAuthors2LowerCase()
    '================================
    Dim rs As ADODB.Recordset
    Dim strName As String
    Dim strNameFix As String
    Dim id As Long
    '================================
    On Error GoTo Handler
    
    If vbNo = MsgBox("Esta a punto de renombrar todos los nombres" & vbCrLf & _
                      "de todos los autores a minúsculas." & vbCrLf & _
                      "¿Desea continuar?", vbExclamation + vbYesNo, "Advertencia") Then
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT id_author, author FROM author", cn, adOpenDynamic, adLockOptimistic, adCmdText
    
    While rs.EOF = False

        strName = rs!Author
        id = rs!id_author
        
        strNameFix = LCase(strName)
        If strName <> strNameFix Then
            
            strName = strNameFix
            gfnc_ParseString strName, strNameFix
            rs!Author = strNameFix
            rs.Update
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbCritical, "RenameAllAuthors2LowerCase"
End Sub

' convertir todos los nombres de archivo seleccionados a minusculas
Private Sub RenameNames2LowerCase()
    '================================
    Dim rs As ADODB.Recordset
    Dim cd As ADODB.Command
    Dim strName As String
    Dim strNameFix As String
    Dim id As Long
    '================================
    On Error GoTo Handler
    
    If Not gb_DBNameFromFile Then
    
        If vbNo = MsgBox("Esta a punto de renombrar todos los nombres" & vbCrLf & _
                          "de la última busqueda a minúsculas." & vbCrLf & _
                          "¿Desea continuar?", vbExclamation + vbYesNo, "Advertencia") Then
            Exit Sub
        End If
        
        Set rs = New ADODB.Recordset
        rs.Open m_strQuery, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        Set cd = New ADODB.Command
        Set cd.ActiveConnection = cn
    
        While rs.EOF = False
    
            strName = rs!title
            id = rs!id_file
            
            strNameFix = LCase(strName)
            If strName <> strNameFix Then
                
                strName = strNameFix
                gfnc_ParseString strName, strNameFix
                
                cd.CommandText = "UPDATE file SET name='" & strNameFix & "' WHERE id_file=" & id
                cd.Execute
                
            End If
            rs.MoveNext
        Wend
        rs.Close
    
        Generar_Lista
        
    Else
        MsgBox "Selecciona busqueda por [Nombre] y no [Archivo]"
    End If
    
    Exit Sub

Handler:
    MsgBox Err.Description, vbCritical, "RenameNames2LowerCase"
End Sub

'''*******************************************************************************
''' hacer alguna modificacion a la BD
''Private Sub MakeSomethingToRegisters()
''
''    Dim rs As ADODB.Recordset
''    Dim str As String
''
''    Set rs = New ADODB.Recordset
''    query = "SELECT name FROM file WHERE name LIKE '-%'"
''
''    rs.Open query, cn, adOpenDynamic, adLockOptimistic, adCmdText
''
''    While rs.EOF = False
''
''        str = rs!Name
''
''        If (Mid(str, 1, 2) = "- ") Then
''            str = Mid(str, 3)
''        Else
''            MsgBox str
''        End If
''
''        rs!Name = str
''        rs.Update
''        rs.MoveNext
''    Wend
''    rs.Close
''
''End Sub
'''*******************************************************************************
