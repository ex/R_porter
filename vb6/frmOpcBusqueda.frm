VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpcBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de búsqueda"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmOpcBusqueda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Otras opciones:"
      Height          =   675
      Left            =   30
      TabIndex        =   80
      Top             =   5385
      Width           =   5010
      Begin VB.CheckBox chkShowHiddenFiles 
         Caption         =   "&Mostrar archivos ocultos"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   2100
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre a usar:"
      Height          =   675
      Left            =   30
      TabIndex        =   79
      Top             =   4695
      Width           =   5010
      Begin VB.CheckBox chkShowPath 
         Caption         =   "Con &ruta..."
         Height          =   195
         Left            =   3795
         TabIndex        =   21
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton optByFileName 
         Caption         =   "Archi&vo"
         Height          =   195
         Left            =   2715
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optByTitleName 
         Caption         =   "&Título"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   555
      TabIndex        =   23
      Top             =   6240
      Width           =   1215
   End
   Begin VB.PictureBox pbxSearchDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   2655
      ScaleHeight     =   1260
      ScaleWidth      =   2235
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7710
      Visible         =   0   'False
      Width           =   2265
      Begin VB.OptionButton optDBDate 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.OptionButton optDBDate 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBDate 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBDate 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBDate 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.Label lblOptionDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   77
         Top             =   1035
         Width           =   1920
      End
      Begin VB.Label lblOptionDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   76
         Top             =   780
         Width           =   1920
      End
      Begin VB.Label lblOptionDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   75
         Top             =   525
         Width           =   1920
      End
      Begin VB.Label lblOptionDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   74
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label lblOptionDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   73
         Top             =   15
         Width           =   1920
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin VB.PictureBox pbxSearchFileSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   120
      ScaleHeight     =   1260
      ScaleWidth      =   2235
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7710
      Visible         =   0   'False
      Width           =   2265
      Begin VB.OptionButton optDBFileSize 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.OptionButton optDBFileSize 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBFileSize 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBFileSize 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBFileSize 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblOptionFileSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   71
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label lblOptionFileSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   70
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label lblOptionFileSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   69
         Top             =   525
         Width           =   1920
      End
      Begin VB.Label lblOptionFileSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   68
         Top             =   780
         Width           =   1920
      End
      Begin VB.Label lblOptionFileSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   67
         Top             =   1035
         Width           =   1920
      End
   End
   Begin VB.PictureBox pbxSearchPriority 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   2655
      ScaleHeight     =   1260
      ScaleWidth      =   2235
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   2265
      Begin VB.OptionButton optDBPriority 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.OptionButton optDBPriority 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBPriority 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBPriority 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBPriority 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblOptionPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   57
         Top             =   15
         Width           =   1920
      End
      Begin VB.Label lblOptionPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   56
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label lblOptionPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   55
         Top             =   525
         Width           =   1920
      End
      Begin VB.Label lblOptionPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   54
         Top             =   780
         Width           =   1920
      End
      Begin VB.Label lblOptionPriority 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   53
         Top             =   1035
         Width           =   1920
      End
   End
   Begin VB.PictureBox pbxSearchQuality 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   120
      ScaleHeight     =   1260
      ScaleWidth      =   2235
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   2265
      Begin VB.OptionButton optDBQuality 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.OptionButton optDBQuality 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   285
         Width           =   210
      End
      Begin VB.OptionButton optDBQuality 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   540
         Width           =   210
      End
      Begin VB.OptionButton optDBQuality 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   795
         Width           =   210
      End
      Begin VB.OptionButton optDBQuality 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1050
         Width           =   210
      End
      Begin VB.Label lblOptionQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   63
         Top             =   1035
         Width           =   1920
      End
      Begin VB.Label lblOptionQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Mayor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   62
         Top             =   780
         Width           =   1920
      End
      Begin VB.Label lblOptionQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   61
         Top             =   525
         Width           =   1920
      End
      Begin VB.Label lblOptionQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor igual que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   60
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label lblOptionQuality 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Menor que..."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   59
         Top             =   0
         Width           =   1920
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FCDBAF&
         X1              =   0
         X2              =   3060
         Y1              =   1005
         Y2              =   1005
      End
   End
   Begin VB.Timer tmrDownQuality 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2040
      Top             =   1245
   End
   Begin VB.Timer tmrDownPriority 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4575
      Top             =   135
   End
   Begin VB.Frame fraBelong 
      Caption         =   "Mostrar en los resultados el:"
      Height          =   675
      Left            =   30
      TabIndex        =   31
      Top             =   3315
      Width           =   5010
      Begin VB.OptionButton optParent 
         Caption         =   "Grupo de &pertenencia"
         Height          =   195
         Left            =   2715
         TabIndex        =   16
         Top             =   300
         Width           =   1845
      End
      Begin VB.OptionButton optStorage 
         Caption         =   "Medio de &almacenamiento"
         Height          =   195
         Left            =   225
         TabIndex        =   15
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha"
      Height          =   1080
      Left            =   2565
      TabIndex        =   30
      Top             =   1125
      Width           =   2475
      Begin VB.Timer tmrDownDate 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   2010
         Top             =   120
      End
      Begin VB.CommandButton cmdDownDate 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         Picture         =   "frmOpcBusqueda.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   630
         Width           =   255
      End
      Begin VB.CheckBox chkDeFecha 
         Caption         =   "Buscar con &fecha:"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   330
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpckFecha 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "              dd/MM/yyy"
         Format          =   59965443
         CurrentDate     =   38039
      End
   End
   Begin VB.Frame fraCalidad 
      Caption         =   "Calidad"
      Height          =   1080
      Left            =   30
      TabIndex        =   29
      Top             =   1125
      Width           =   2475
      Begin VB.CommandButton cmdDownQuality 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2100
         Picture         =   "frmOpcBusqueda.frx":0270
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   255
      End
      Begin VB.CheckBox chkConCalidad 
         Caption         =   "Buscar con cali&dad:"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtQuality 
         Height          =   315
         Left            =   90
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   2280
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      Height          =   1080
      Left            =   30
      TabIndex        =   28
      Top             =   30
      Width           =   2475
      Begin VB.ComboBox cmbFileType 
         Height          =   315
         ItemData        =   "frmOpcBusqueda.frx":0396
         Left            =   90
         List            =   "frmOpcBusqueda.frx":0398
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2280
      End
      Begin VB.CheckBox chkDelTipo 
         Caption         =   "Buscar del &tipo:"
         Height          =   225
         Left            =   90
         TabIndex        =   0
         Top             =   330
         Width           =   2175
      End
   End
   Begin VB.Frame fraPrioridad 
      Caption         =   "Prioridad"
      Height          =   1080
      Left            =   2565
      TabIndex        =   27
      Top             =   30
      Width           =   2475
      Begin VB.CommandButton cmdDownPriority 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2100
         Picture         =   "frmOpcBusqueda.frx":039A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   255
      End
      Begin VB.CheckBox chkConPrioridad 
         Caption         =   "Buscar con p&rioridad:"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtPriority 
         Height          =   315
         Left            =   90
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A&plicar"
      Height          =   315
      Left            =   1920
      TabIndex        =   24
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar"
      Height          =   315
      Left            =   3300
      TabIndex        =   25
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame fraTamanyo 
      Caption         =   "Tamaño"
      Height          =   1080
      Left            =   30
      TabIndex        =   65
      Top             =   2220
      Width           =   2475
      Begin VB.Timer tmrDownFileSize 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   2010
         Top             =   120
      End
      Begin VB.CheckBox chkDeTamanyo 
         Caption         =   "Buscar con &tamaño:"
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   2175
      End
      Begin VB.CommandButton cmdDownFileSize 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2100
         Picture         =   "frmOpcBusqueda.frx":04C0
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   255
      End
      Begin VB.TextBox txtFileSize 
         Height          =   315
         Left            =   90
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "0"
         Top             =   600
         Width           =   2280
      End
   End
   Begin VB.Frame fraCantidad 
      Caption         =   "Mostrar en resultados:"
      Height          =   1080
      Left            =   2565
      TabIndex        =   64
      Top             =   2220
      Width           =   2475
      Begin VB.ComboBox cmbResultados 
         Height          =   315
         ItemData        =   "frmOpcBusqueda.frx":05E6
         Left            =   90
         List            =   "frmOpcBusqueda.frx":05E8
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2280
      End
      Begin VB.Label lblTabla 
         Caption         =   "Campo numérico a &mostrar:"
         Height          =   225
         Left            =   105
         TabIndex        =   13
         Top             =   330
         Width           =   2205
      End
   End
   Begin VB.Frame fraAux 
      Caption         =   "Seleccionar campo auxiliar a mostrar:"
      Height          =   675
      Left            =   30
      TabIndex        =   78
      Top             =   4005
      Width           =   5010
      Begin VB.OptionButton optGenre 
         Caption         =   "&Género al que pertenece"
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   300
         Width           =   2175
      End
      Begin VB.OptionButton optFileType 
         Caption         =   "T&ipo de archivo"
         Height          =   195
         Left            =   2715
         TabIndex        =   18
         Top             =   300
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmOpcBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mt_QualitySearchStyle As DB_CompareStyle
Private mt_PrioritySearchStyle As DB_CompareStyle
Private mt_DateSearchStyle As DB_CompareStyle
Private mt_FileSizeSearchStyle As DB_CompareStyle

Private mb_DBPertenciaPorAlmacenamiento As Boolean

Private Sub optByFileName_Click()
    If optByFileName.value = True Then
        chkShowPath.Visible = True
    End If
End Sub

Private Sub optByTitleName_Click()
    If optByTitleName.value = True Then
        chkShowPath.Visible = False
    End If
End Sub

'**************************************************************
'* CMBOPTION Quality
'**************************************************************
Private Sub txtQuality_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub txtQuality_GotFocus()

    txtQuality.SelStart = 0
    txtQuality.SelLength = Len(txtQuality.Text)

End Sub

Private Sub txtQuality_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        
        If gb_DBQualitySearchStyleActive = False Then
        
            gb_DBQualitySearchStyleActive = True
            pbxSearchQuality.Top = txtQuality.Top + 345 + 1095
            pbxSearchQuality.Visible = True
                    
            optDBQuality(mt_QualitySearchStyle).value = True
            optDBQuality(mt_QualitySearchStyle).SetFocus
            
            Set cmdDownQuality.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
            
        End If
    End If

End Sub

Private Sub cmdDownQuality_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownQuality_Click
    End If
End Sub

Private Sub cmdDownQuality_Click()
    
    If gb_DBQualitySearchStyleActive = False Then
    
        gb_DBQualitySearchStyleActive = True
        pbxSearchQuality.Top = txtQuality.Top + 345 + 1095
        pbxSearchQuality.Visible = True
                
        optDBQuality(mt_QualitySearchStyle).value = True
        optDBQuality(mt_QualitySearchStyle).SetFocus
        
        Set cmdDownQuality.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
        
    Else
    
        tmrDownQuality.Enabled = False
        gb_DBQualitySearchStyleActive = False
        'primero quitamos enfoque
        txtQuality.SetFocus
        pbxSearchQuality.Visible = False
        Set cmdDownQuality.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
    End If

End Sub

Private Sub optDBQuality_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    optDBQuality(Index).value = True
    tmrDownQuality.Enabled = True
    
End Sub

Private Sub optDBQuality_Click(Index As Integer)
    '===================================================
    Dim Inicio As Single
    '===================================================
    
    If gb_DBQualitySearchStyleActive = True Then
        mt_QualitySearchStyle = Index
    End If
    
End Sub

Private Sub optDBQuality_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownQuality_Click
    End If

End Sub

Private Sub optDBQuality_LostFocus(Index As Integer)

    If (ActiveControl.Name = "cmdDownQuality") Then
        ActiveControl.SetFocus
    Else
    
        If (ActiveControl.Name = "optDBQuality") Then
    
            lblOptionQuality(Index).ForeColor = &H0&
            lblOptionQuality(Index).BackColor = &HFFFFFF
            ActiveControl.SetFocus
        
        Else
            
            If (ActiveControl.Name = "pbxSearchQuality") Then
                optDBQuality(Index).SetFocus
            Else
                gb_DBQualitySearchStyleActive = False
                pbxSearchQuality.Visible = False
                Set cmdDownQuality.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
            End If
            
        End If
    End If
    
End Sub

Private Sub optDBQuality_GotFocus(Index As Integer)

    lblOptionQuality(Index).ForeColor = &HFFFFFF
    lblOptionQuality(Index).BackColor = &HFF963F

End Sub

Private Sub lblOptionQuality_Click(Index As Integer)
    
    optDBQuality(Index).value = True
    optDBQuality(Index).SetFocus
    
    tmrDownQuality.Enabled = True
    
End Sub

Private Sub tmrDownQuality_Timer()
    
    tmrDownQuality.Enabled = False
    gb_DBQualitySearchStyleActive = False
    'primero quitamos enfoque
    txtQuality.SetFocus
    pbxSearchQuality.Visible = False
    Set cmdDownQuality.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
End Sub


'**************************************************************
'* CMBOPTION Priority
'**************************************************************
Private Sub txtPriority_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub txtPriority_GotFocus()

    txtPriority.SelStart = 0
    txtPriority.SelLength = Len(txtPriority.Text)

End Sub

Private Sub txtPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        
        If gb_DBPrioritySearchStyleActive = False Then
        
            gb_DBPrioritySearchStyleActive = True
            pbxSearchPriority.Top = txtPriority.Top + 345
            pbxSearchPriority.Visible = True
                    
            optDBPriority(mt_PrioritySearchStyle).value = True
            optDBPriority(mt_PrioritySearchStyle).SetFocus
            
            Set cmdDownPriority.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
            
        End If
    End If

End Sub

Private Sub cmdDownPriority_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownPriority_Click
    End If
End Sub

Private Sub cmdDownPriority_Click()
    
    If gb_DBPrioritySearchStyleActive = False Then
    
        gb_DBPrioritySearchStyleActive = True
        pbxSearchPriority.Top = txtPriority.Top + 345
        pbxSearchPriority.Visible = True
                
        optDBPriority(mt_PrioritySearchStyle).value = True
        optDBPriority(mt_PrioritySearchStyle).SetFocus
        
        Set cmdDownPriority.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
        
    Else
    
        tmrDownPriority.Enabled = False
        gb_DBPrioritySearchStyleActive = False
        'primero quitamos enfoque
        txtPriority.SetFocus
        pbxSearchPriority.Visible = False
        Set cmdDownPriority.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
    End If

End Sub

Private Sub optDBPriority_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    optDBPriority(Index).value = True
    tmrDownPriority.Enabled = True
    
End Sub

Private Sub optDBPriority_Click(Index As Integer)
Dim Inicio As Single
    
    If gb_DBPrioritySearchStyleActive = True Then
        mt_PrioritySearchStyle = Index
    End If
    
End Sub

Private Sub optDBPriority_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownPriority_Click
    End If

End Sub

Private Sub optDBPriority_LostFocus(Index As Integer)

    If (ActiveControl.Name = "cmdDownPriority") Then
        ActiveControl.SetFocus
    Else
    
        If (ActiveControl.Name = "optDBPriority") Then
    
            lblOptionPriority(Index).ForeColor = &H0&
            lblOptionPriority(Index).BackColor = &HFFFFFF
            ActiveControl.SetFocus
        
        Else
            
            If (ActiveControl.Name = "pbxSearchPriority") Then
                optDBPriority(Index).SetFocus
            Else
                gb_DBPrioritySearchStyleActive = False
                pbxSearchPriority.Visible = False
                Set cmdDownPriority.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
            End If
            
        End If
    End If
    
End Sub

Private Sub optDBPriority_GotFocus(Index As Integer)

    lblOptionPriority(Index).ForeColor = &HFFFFFF
    lblOptionPriority(Index).BackColor = &HFF963F

End Sub

Private Sub lblOptionPriority_Click(Index As Integer)
    
    optDBPriority(Index).value = True
    optDBPriority(Index).SetFocus
    
    tmrDownPriority.Enabled = True
    
End Sub

Private Sub tmrDownPriority_Timer()
    
    tmrDownPriority.Enabled = False
    gb_DBPrioritySearchStyleActive = False
    'primero quitamos enfoque
    txtPriority.SetFocus
    pbxSearchPriority.Visible = False
    Set cmdDownPriority.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
End Sub

'**************************************************************
'* CMBOPTION FileSize
'**************************************************************
Private Sub txtFileSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub txtFileSize_GotFocus()

    txtFileSize.SelStart = 0
    txtFileSize.SelLength = Len(txtFileSize.Text)

End Sub

Private Sub txtFileSize_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        
        If gb_DBFileSizeSearchStyleActive = False Then
        
            gb_DBFileSizeSearchStyleActive = True
            pbxSearchFileSize.Top = txtFileSize.Top + 345 + 2190
            pbxSearchFileSize.Visible = True
                    
            optDBFileSize(mt_FileSizeSearchStyle).value = True
            optDBFileSize(mt_FileSizeSearchStyle).SetFocus
            
            Set cmdDownFileSize.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
            
        End If
    End If

End Sub

Private Sub cmdDownFileSize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownFileSize_Click
    End If
End Sub

Private Sub cmdDownFileSize_Click()
    
    If gb_DBFileSizeSearchStyleActive = False Then
    
        gb_DBFileSizeSearchStyleActive = True
        pbxSearchFileSize.Top = txtFileSize.Top + 345 + 2190
        pbxSearchFileSize.Visible = True
                
        optDBFileSize(mt_FileSizeSearchStyle).value = True
        optDBFileSize(mt_FileSizeSearchStyle).SetFocus
        
        Set cmdDownFileSize.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
        
    Else
    
        tmrDownFileSize.Enabled = False
        gb_DBFileSizeSearchStyleActive = False
        'primero quitamos enfoque
        txtFileSize.SetFocus
        pbxSearchFileSize.Visible = False
        Set cmdDownFileSize.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
    End If

End Sub

Private Sub optDBFileSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    optDBFileSize(Index).value = True
    tmrDownFileSize.Enabled = True
    
End Sub

Private Sub optDBFileSize_Click(Index As Integer)
    '===================================================
    Dim Inicio As Single
    '===================================================
    
    If gb_DBFileSizeSearchStyleActive = True Then
        mt_FileSizeSearchStyle = Index
    End If
    
End Sub

Private Sub optDBFileSize_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownFileSize_Click
    End If

End Sub

Private Sub optDBFileSize_LostFocus(Index As Integer)

    If (ActiveControl.Name = "cmdDownFileSize") Then
        ActiveControl.SetFocus
    Else
    
        If (ActiveControl.Name = "optDBFileSize") Then
    
            lblOptionFileSize(Index).ForeColor = &H0&
            lblOptionFileSize(Index).BackColor = &HFFFFFF
            ActiveControl.SetFocus
        
        Else
            
            If (ActiveControl.Name = "pbxSearchFileSize") Then
                optDBFileSize(Index).SetFocus
            Else
                gb_DBFileSizeSearchStyleActive = False
                pbxSearchFileSize.Visible = False
                Set cmdDownFileSize.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
            End If
            
        End If
    End If
    
End Sub

Private Sub optDBFileSize_GotFocus(Index As Integer)

    lblOptionFileSize(Index).ForeColor = &HFFFFFF
    lblOptionFileSize(Index).BackColor = &HFF963F

End Sub

Private Sub lblOptionFileSize_Click(Index As Integer)
    
    optDBFileSize(Index).value = True
    optDBFileSize(Index).SetFocus
    
    tmrDownFileSize.Enabled = True
    
End Sub

Private Sub tmrDownFileSize_Timer()
    
    tmrDownFileSize.Enabled = False
    gb_DBFileSizeSearchStyleActive = False
    'primero quitamos enfoque
    txtFileSize.SetFocus
    pbxSearchFileSize.Visible = False
    Set cmdDownFileSize.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
End Sub

'**************************************************************
'* CMBOPTION Date
'**************************************************************
Private Sub dtpckFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub dtpckFecha_KeyDown(KeyCode As Integer, Shift As Integer)

    If ((KeyCode = 40) And (Shift = 2)) Or ((KeyCode = 40) And (Shift = 4)) Then
        
        If gb_DBDateSearchStyleActive = False Then
        
            gb_DBDateSearchStyleActive = True
            pbxSearchDate.Top = dtpckFecha.Top + 345 + 1095
            pbxSearchDate.Visible = True
                    
            optDBDate(mt_DateSearchStyle).value = True
            optDBDate(mt_DateSearchStyle).SetFocus
            
            Set cmdDownDate.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
            
        End If
    End If

End Sub

Private Sub cmdDownDate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        cmdDownDate_Click
    End If
End Sub

Private Sub cmdDownDate_Click()
    
    If gb_DBDateSearchStyleActive = False Then
    
        gb_DBDateSearchStyleActive = True
        pbxSearchDate.Top = dtpckFecha.Top + 345 + 1095
        pbxSearchDate.Visible = True
                
        optDBDate(mt_DateSearchStyle).value = True
        optDBDate(mt_DateSearchStyle).SetFocus
        
        Set cmdDownDate.Picture = frmDataControl.imglstB.ListImages.Item("up").Picture
        
    Else
    
        tmrDownDate.Enabled = False
        gb_DBDateSearchStyleActive = False
        'primero quitamos enfoque
        dtpckFecha.SetFocus
        pbxSearchDate.Visible = False
        Set cmdDownDate.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
    End If

End Sub

Private Sub optDBDate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    optDBDate(Index).value = True
    tmrDownDate.Enabled = True
    
End Sub

Private Sub optDBDate_Click(Index As Integer)
Dim Inicio As Single
    
    If gb_DBDateSearchStyleActive = True Then
        mt_DateSearchStyle = Index
    End If
    
End Sub

Private Sub optDBDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Or KeyCode = 13 Then
        cmdDownDate_Click
    End If

End Sub

Private Sub optDBDate_LostFocus(Index As Integer)

    If (ActiveControl.Name = "cmdDownDate") Then
        ActiveControl.SetFocus
    Else
    
        If (ActiveControl.Name = "optDBDate") Then
    
            lblOptionDate(Index).ForeColor = &H0&
            lblOptionDate(Index).BackColor = &HFFFFFF
            ActiveControl.SetFocus
        
        Else
            
            If (ActiveControl.Name = "pbxSearchDate") Then
                optDBDate(Index).SetFocus
            Else
                gb_DBDateSearchStyleActive = False
                pbxSearchDate.Visible = False
                Set cmdDownDate.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
            End If
            
        End If
    End If
    
End Sub

Private Sub optDBDate_GotFocus(Index As Integer)

    lblOptionDate(Index).ForeColor = &HFFFFFF
    lblOptionDate(Index).BackColor = &HFF963F

End Sub

Private Sub lblOptionDate_Click(Index As Integer)
    
    optDBDate(Index).value = True
    optDBDate(Index).SetFocus
    
    tmrDownDate.Enabled = True
    
End Sub

Private Sub tmrDownDate_Timer()
    
    tmrDownDate.Enabled = False
    gb_DBDateSearchStyleActive = False
    'primero quitamos enfoque
    dtpckFecha.SetFocus
    pbxSearchDate.Visible = False
    Set cmdDownDate.Picture = frmDataControl.imglstB.ListImages.Item("down").Picture
    
End Sub

'**************************************************************
'* CHECKBOXS
'**************************************************************
Private Sub chkConPrioridad_Click()
    
    If chkConPrioridad.value = vbChecked Then
        txtPriority.Enabled = True
        cmdDownPriority.Enabled = True
    Else
        txtPriority.Enabled = False
        cmdDownPriority.Enabled = False
    End If

End Sub

Private Sub chkConCalidad_Click()
    
    If chkConCalidad.value = vbChecked Then
        txtQuality.Enabled = True
        cmdDownQuality.Enabled = True
    Else
        txtQuality.Enabled = False
        cmdDownQuality.Enabled = False
    End If

End Sub

Private Sub chkDeFecha_Click()
    
    If chkDeFecha.value = vbChecked Then
        dtpckFecha.Enabled = True
        cmdDownDate.Enabled = True
    Else
        dtpckFecha.Enabled = False
        cmdDownDate.Enabled = False
    End If

End Sub

Private Sub chkDeTamanyo_Click()
    
    If chkDeTamanyo.value = vbChecked Then
        txtFileSize.Enabled = True
        cmdDownFileSize.Enabled = True
    Else
        txtFileSize.Enabled = False
        cmdDownFileSize.Enabled = False
    End If

End Sub

Private Sub chkDelTipo_Click()
    
    If chkDelTipo.value = vbChecked Then
        cmbFileType.Enabled = True
    Else
        cmbFileType.Enabled = False
    End If

End Sub

'**************************************************************
'* CMDBUTTONS
'**************************************************************
Private Sub cmdBuscar_Click()
    
    On Error Resume Next
    
    Aplicar_Opciones
    Unload Me
    frmDataControl.Generar_Lista

End Sub

Private Sub cmdAceptar_Click()
    
    On Error Resume Next
    
    Aplicar_Opciones
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'**************************************************************
'* COMBOS
'**************************************************************
Private Sub cmbResultados_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

Private Sub cmbFileType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.value = True
    End If
End Sub

'**************************************************************
'* FORM
'**************************************************************
Private Sub Form_Load()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================
    
    On Error GoTo Handler
    
    Set rs = New ADODB.Recordset

    '************************************
    'cargar combo de campos por mostrar
    '************************************
    cmbResultados.Clear
    cmbResultados.AddItem "Tamaño"                      ' [COD-001]
        cmbResultados.AddItem "Prioridad"
    cmbResultados.AddItem "Calidad"
    cmbResultados.AddItem "Fecha"

    '************************************
    'cargar la tabla de tipos de archivo
    '************************************
    query = "SELECT * FROM file_type WHERE (id_file_type > 0) ORDER BY file_type"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbFileType.AddItem UCase(rs!file_type)
        cmbFileType.ItemData(cmbFileType.NewIndex) = rs!id_file_type
        rs.MoveNext
    Wend
    rs.Close

    '************************************
    'trabajamos con variables locales
    '(para el caso de cancelar)
    '************************************
    mt_PrioritySearchStyle = gt_DBPrioritySearchStyle
    mt_QualitySearchStyle = gt_DBQualitySearchStyle
    mt_DateSearchStyle = gt_DBDateSearchStyle
    mt_FileSizeSearchStyle = gt_DBFileSizeSearchStyle
    
    '************************************
    'inicializamos controles
    '************************************
    txtPriority = gs_DBConPrioridadDe
    txtQuality = gs_DBConCalidadDe
    txtFileSize = gs_DBConTamanyoDe
    dtpckFecha.value = gd_DBConFechaDe
    
    If cmbFileType.ListCount > 0 Then           ' por si no hay tipos en la BD
        cmbFileType.ListIndex = gl_DBFileTypeIndex
    End If
    
    cmbResultados.ListIndex = gn_DBFieldIndex
    
    If gb_DBConPrioridad = True Then
        chkConPrioridad.value = vbChecked
        txtPriority.Enabled = True
        cmdDownPriority.Enabled = True
    Else
        chkConPrioridad.value = vbUnchecked
        txtPriority.Enabled = False
        cmdDownPriority.Enabled = False
    End If
    
    If gb_DBConCalidad = True Then
        chkConCalidad.value = vbChecked
        txtQuality.Enabled = True
        cmdDownQuality.Enabled = True
    Else
        chkConCalidad.value = vbUnchecked
        txtQuality.Enabled = False
        cmdDownQuality.Enabled = False
    End If
    
    If gb_DBConFecha = True Then
        chkDeFecha.value = vbChecked
        dtpckFecha.Enabled = True
        cmdDownDate.Enabled = True
    Else
        chkDeFecha.value = vbUnchecked
        dtpckFecha.Enabled = False
        cmdDownDate.Enabled = False
    End If
    
    If gb_DBConTamanyo = True Then
        chkDeTamanyo.value = vbChecked
        txtFileSize.Enabled = True
        cmdDownFileSize.Enabled = True
    Else
        chkDeTamanyo.value = vbUnchecked
        txtFileSize.Enabled = False
        cmdDownFileSize.Enabled = False
    End If
    
    If gb_DBConTipo = True Then
        chkDelTipo.value = vbChecked
        cmbFileType.Enabled = True
    Else
        chkDelTipo.value = vbUnchecked
        cmbFileType.Enabled = False
    End If
    
    mb_DBPertenciaPorAlmacenamiento = gb_DBPertenciaPorAlmacenamiento
    If gb_DBPertenciaPorAlmacenamiento = True Then
        optStorage.value = True
    Else
        optParent.value = True
    End If
    
    If gb_DBCampoAuxiliarPorGenero = True Then
        optGenre.value = True
    Else
        optFileType.value = True
    End If

    If gb_DBNameFromFile = True Then
        optByFileName.value = True
    Else
        optByTitleName.value = True
    End If
    
    If gb_DBShowPathInFileName = True Then
        chkShowPath.value = vbChecked
    Else
        chkShowPath.value = vbUnchecked
    End If
    
    If gb_DBShowHiddenFiles = True Then
        chkShowHiddenFiles.value = vbChecked
    Else
        chkShowHiddenFiles.value = vbUnchecked
    End If
    
    Exit Sub
    
Handler:

    Select Case Err.Number
    
        Case 3709
            MsgBox "Se ha perdido conexión con la base de datos" & vbCrLf & "Verifique la conexión y vuelva a cargar el formulario.", vbExclamation, "Conexión perdida"
            Me.cmdAceptar.Enabled = False
        
        Case Else
            MsgBox Err.Description, vbCritical, "Form_Load"
            
    End Select
    
End Sub

Private Sub Aplicar_Opciones()
    
    On Error GoTo Handler
    
    If CInt(txtPriority) < 0 Then
        MsgBox "La prioridad debe ser >= 0", vbExclamation, "Error"
        txtPriority.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtQuality) < 0 Then
        MsgBox "La calidad debe ser >= 0", vbExclamation, "Error"
        txtQuality.SetFocus
        Exit Sub
    End If
    
    If CLng(txtFileSize) < 0 Then
        MsgBox "El tamaño debe ser >= 0", vbExclamation, "Error"
        txtQuality.SetFocus
        Exit Sub
    End If
    
    If chkConPrioridad.value = vbUnchecked Then
        gb_DBConPrioridad = False
    Else
        gb_DBConPrioridad = True
        gs_DBConPrioridadDe = Trim(txtPriority)
    End If
        
    If chkConCalidad.value = vbUnchecked Then
        gb_DBConCalidad = False
    Else
        gb_DBConCalidad = True
        gs_DBConCalidadDe = Trim(txtQuality)
    End If
        
    If chkDeTamanyo.value = vbUnchecked Then
        gb_DBConTamanyo = False
    Else
        gb_DBConTamanyo = True
        gs_DBConTamanyoDe = Trim(txtFileSize)
    End If
    
    If chkDeFecha.value = vbUnchecked Then
        gb_DBConFecha = False
    Else
        gb_DBConFecha = True
        gd_DBConFechaDe = dtpckFecha.value
    End If
        
    If chkDelTipo.value = vbUnchecked Then
        gb_DBConTipo = False
    Else
        gb_DBConTipo = True
        
        If cmbFileType.ListCount > 0 Then       ' por si no hay tipos en la BD
            gl_DBConTipoDe = cmbFileType.ItemData(cmbFileType.ListIndex)
        End If
        
    End If
        
    gs_DBConCampoDe = cmbResultados.Text
    
    If optStorage.value = True Then
        gb_DBPertenciaPorAlmacenamiento = True
    Else
        gb_DBPertenciaPorAlmacenamiento = False
    End If
        
    If optGenre.value = True Then
        gb_DBCampoAuxiliarPorGenero = True
    Else
        gb_DBCampoAuxiliarPorGenero = False
    End If
        
    If optByFileName.value = True Then
        gb_DBNameFromFile = True
    Else
        gb_DBNameFromFile = False
    End If
        
    If chkShowPath.value = vbChecked Then
        gb_DBShowPathInFileName = True
    Else
        gb_DBShowPathInFileName = False
    End If
        
    If chkShowHiddenFiles.value = vbChecked Then
        gb_DBShowHiddenFiles = True
    Else
        gb_DBShowHiddenFiles = False
    End If
        
    '-----------------------------------------------------------------
    ' actualizar registros del combo: frmDataControl.cmbParent
    '
    If (gb_DBPertenciaPorAlmacenamiento <> mb_DBPertenciaPorAlmacenamiento) Then
        frmDataControl.Actualizar_ComboPariente
    End If
    
    '-----------------------------------------------------------------
    ' actualizar titulos de frmDataControl.flxResults
    '
    frmDataControl.Establecer_TitulosFlex
    
    '************************************
    'guardar opciones globales
    '************************************
    gt_DBPrioritySearchStyle = mt_PrioritySearchStyle
    gt_DBQualitySearchStyle = mt_QualitySearchStyle
    gl_DBFileTypeIndex = cmbFileType.ListIndex
    gt_DBDateSearchStyle = mt_DateSearchStyle
    gt_DBFileSizeSearchStyle = mt_FileSizeSearchStyle
    gn_DBFieldIndex = cmbResultados.ListIndex
    '------------------------------------
    
    Exit Sub
    
Handler:

    If Err.Number = 13 Then
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "Aceptar_Opciones()"
    End If
    
End Sub

