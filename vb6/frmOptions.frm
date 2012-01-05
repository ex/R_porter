VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Opciones del programa"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   3450
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar"
      Height          =   315
      Left            =   2212
      TabIndex        =   42
      Top             =   5355
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   787
      TabIndex        =   41
      Top             =   5355
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5235
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   9234
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&1) General "
      TabPicture(0)   =   "frmOptions.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOpciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2) Anexion BD "
      TabPicture(1)   =   "frmOptions.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOpcionesDB"
      Tab(1).Control(1)=   "fraOtrosBD"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3) MP3   "
      TabPicture(2)   =   "frmOptions.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraMP3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4) Reporte "
      TabPicture(3)   =   "frmOptions.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pbxFrame"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox pbxFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000007&
         Height          =   4635
         Left            =   -74880
         ScaleHeight     =   4635
         ScaleWidth      =   3930
         TabIndex        =   50
         Top             =   420
         Width           =   3930
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            Height          =   315
            Index           =   0
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1020
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "8"
            Height          =   315
            Index           =   7
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3915
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "7"
            Height          =   315
            Index           =   6
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   3495
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "6"
            Height          =   315
            Index           =   5
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3090
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "5"
            Height          =   315
            Index           =   4
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2670
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "4"
            Height          =   315
            Index           =   3
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2265
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3"
            Height          =   315
            Index           =   2
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1845
            Width           =   360
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2"
            Height          =   315
            Index           =   1
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1425
            Width           =   360
         End
         Begin VB.Frame fraSelect 
            BackColor       =   &H00FFFFFF&
            Height          =   825
            Left            =   150
            TabIndex        =   51
            Top             =   120
            Width           =   2565
            Begin VB.OptionButton optPorTipo 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Por &tipo"
               Height          =   225
               Left            =   105
               TabIndex        =   32
               Top             =   525
               Width           =   840
            End
            Begin VB.OptionButton optPorAtributo 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Por &atributos"
               Height          =   225
               Left            =   105
               TabIndex        =   31
               Top             =   225
               Width           =   1200
            End
         End
         Begin VB.Label lblDirNormal 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Directorio normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   180
            TabIndex        =   59
            Top             =   1095
            Width           =   1470
         End
         Begin VB.Label lblFileOther 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Otro tipo de archivo "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   58
            Top             =   3990
            Width           =   1470
         End
         Begin VB.Label lblFileHidden 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo escondido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   57
            Top             =   3570
            Width           =   1470
         End
         Begin VB.Label lblFileReadOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo de sólo lectura"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   56
            Top             =   3165
            Width           =   1725
         End
         Begin VB.Label lblFileNormal 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivo normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   55
            Top             =   2745
            Width           =   1470
         End
         Begin VB.Label lblDirOther 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Otro tipo de directorio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   54
            Top             =   2340
            Width           =   1845
         End
         Begin VB.Label lblDirHidden 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Directorio escondido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   53
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblDirReadOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Directorio de sólo lectura"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   52
            Top             =   1515
            Width           =   2100
         End
      End
      Begin VB.Frame fraType 
         Height          =   1230
         Left            =   120
         TabIndex        =   49
         Top             =   3840
         Width           =   3930
         Begin VB.Frame fraTipo 
            Enabled         =   0   'False
            Height          =   660
            Left            =   105
            TabIndex        =   60
            Top             =   420
            Width           =   3720
            Begin VB.ComboBox cmbTipoArchivo 
               Height          =   315
               ItemData        =   "frmOptions.frx":037A
               Left            =   570
               List            =   "frmOptions.frx":037C
               TabIndex        =   11
               Text            =   "*"
               ToolTipText     =   "Ingrese extensiones separadas por puntos y comas (cpp;c;h)"
               Top             =   210
               Width           =   3045
            End
            Begin VB.Label lblType 
               Caption         =   "Ti&po"
               Height          =   255
               Left            =   105
               TabIndex        =   10
               Top             =   255
               Width           =   435
            End
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "&Todos los archivos"
            Height          =   300
            Left            =   105
            TabIndex        =   9
            Top             =   180
            Value           =   1  'Checked
            Width           =   1845
         End
      End
      Begin VB.Frame fraOtrosBD 
         Height          =   960
         Left            =   -74880
         TabIndex        =   48
         Top             =   4095
         Width           =   3930
         Begin VB.CheckBox chkAddDirectories 
            Caption         =   "&Añadir información de directorios a la BD"
            Height          =   300
            Left            =   165
            TabIndex        =   24
            Top             =   210
            Width           =   3435
         End
         Begin VB.CheckBox chkAddRegistersToMedia 
            Caption         =   "Agregar &registros al medio existente"
            Height          =   300
            Left            =   165
            TabIndex        =   25
            Top             =   540
            Width           =   3435
         End
      End
      Begin VB.Frame fraMP3 
         Height          =   4650
         Left            =   -74880
         TabIndex        =   47
         Top             =   420
         Width           =   3930
         Begin VB.CheckBox chkMPEG 
            Caption         =   "Añadir información &MPEG - Audio (MP3)"
            Height          =   300
            Left            =   120
            TabIndex        =   30
            Top             =   1125
            Width           =   3435
         End
         Begin VB.Frame fraDatosMP3 
            Height          =   915
            Left            =   105
            TabIndex        =   43
            Top             =   135
            Width           =   3720
            Begin VB.TextBox txtBuffer 
               Height          =   330
               Left            =   2130
               MaxLength       =   4
               TabIndex        =   29
               Top             =   465
               Width           =   1005
            End
            Begin VB.CheckBox chkCalidad 
               Caption         =   "Extraer ca&lidad de archivos MP3 (kbps)"
               Height          =   300
               Left            =   75
               TabIndex        =   26
               Top             =   150
               Width           =   3150
            End
            Begin VB.CheckBox chkVBR 
               Caption         =   "&VBR"
               Height          =   195
               Left            =   75
               TabIndex        =   27
               Top             =   525
               Width           =   1005
            End
            Begin VB.Label label 
               Caption         =   "&Buffer:"
               Height          =   225
               Index           =   3
               Left            =   1500
               TabIndex        =   28
               Top             =   510
               Width           =   600
            End
         End
      End
      Begin VB.Frame fraOpcionesDB 
         Height          =   3660
         Left            =   -74880
         TabIndex        =   45
         Top             =   420
         Width           =   3930
         Begin VB.TextBox txtMediaComment 
            Height          =   780
            Left            =   825
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            ToolTipText     =   "Observaciones opcionales para el medio."
            Top             =   1020
            Width           =   3000
         End
         Begin VB.TextBox txtMediaName 
            Height          =   330
            Left            =   825
            TabIndex        =   13
            ToolTipText     =   "Escribe aquí el nombre que tendrá el medio"
            Top             =   240
            Width           =   1590
         End
         Begin VB.Frame fraOptionsDB 
            Height          =   1740
            Left            =   90
            TabIndex        =   46
            Top             =   1800
            Width           =   3750
            Begin VB.CheckBox chkTipo 
               Caption         =   "Establecer &tipo por extensión"
               Height          =   300
               Left            =   60
               TabIndex        =   20
               Top             =   540
               Width           =   2475
            End
            Begin VB.CheckBox chkGenero 
               Caption         =   "Establecer género por &subdirectorio"
               Height          =   300
               Left            =   60
               TabIndex        =   19
               Top             =   210
               Width           =   2820
            End
            Begin VB.CheckBox chkNombre 
               Caption         =   "Considerar arc&hivo como [autor - nombre]"
               Height          =   300
               Left            =   60
               TabIndex        =   21
               Top             =   885
               Width           =   3225
            End
            Begin VB.TextBox txtDefaultPriority 
               Height          =   330
               Left            =   2190
               MaxLength       =   2
               TabIndex        =   23
               Top             =   1260
               Width           =   1005
            End
            Begin VB.Label label 
               Caption         =   "Prio&ridad predeterminada:"
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   22
               Top             =   1290
               Width           =   1920
            End
         End
         Begin VB.ComboBox cmbStorageType 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Escoge el tipo de medio"
            Top             =   240
            Width           =   1365
         End
         Begin VB.ComboBox cmbCategory 
            Height          =   315
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Escoge la categoría a la que pertenecerá el medio"
            Top             =   630
            Width           =   3000
         End
         Begin VB.Label label 
            Caption         =   "&Notas:"
            Height          =   285
            Index           =   1
            Left            =   75
            TabIndex        =   17
            Top             =   1050
            Width           =   615
         End
         Begin VB.Label label 
            Caption         =   "&Medio:"
            Height          =   285
            Index           =   0
            Left            =   75
            TabIndex        =   12
            Top             =   300
            Width           =   480
         End
         Begin VB.Label label 
            Caption         =   "&Categoria:"
            Height          =   285
            Index           =   4
            Left            =   75
            TabIndex        =   15
            Top             =   690
            Width           =   735
         End
      End
      Begin VB.Frame fraOpciones 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   120
         TabIndex        =   44
         Top             =   420
         Width           =   3930
         Begin VB.TextBox txtFileLimit 
            Height          =   315
            Left            =   2775
            TabIndex        =   65
            Top             =   2970
            Width           =   645
         End
         Begin VB.CheckBox chkFileLimit 
            Caption         =   "Li&mitar archivos por directorio:"
            Height          =   300
            Left            =   120
            TabIndex        =   64
            Top             =   2970
            Width           =   2640
         End
         Begin VB.CheckBox chkDirLimit 
            Caption         =   "&Limitar profundidad de directorio:"
            Height          =   300
            Left            =   120
            TabIndex        =   63
            Top             =   2580
            Width           =   2640
         End
         Begin VB.Frame fraScript 
            Height          =   660
            Left            =   90
            TabIndex        =   6
            Top             =   1845
            Width           =   3720
            Begin VB.CommandButton cmdSearchScript 
               Caption         =   "..."
               Height          =   315
               Left            =   3345
               TabIndex        =   8
               Top             =   210
               Width           =   270
            End
            Begin VB.TextBox txtScriptFile 
               Height          =   315
               Left            =   570
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   210
               Width           =   2760
            End
            Begin VB.Label lblScript 
               Caption         =   "Scr&ipt"
               Height          =   255
               Left            =   105
               TabIndex        =   61
               Top             =   255
               Width           =   435
            End
         End
         Begin VB.CheckBox chkScript 
            Caption         =   "Activar &scripts durante exploración"
            Height          =   300
            Left            =   105
            TabIndex        =   5
            Top             =   1605
            Width           =   2805
         End
         Begin VB.CheckBox chkIncluirSubDir 
            Caption         =   "Incluir &subdirectorios en la exploración"
            Height          =   255
            Left            =   105
            TabIndex        =   1
            Top             =   210
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkAgregarDB 
            Caption         =   "Agregar resultados a la &BD"
            Height          =   300
            Left            =   105
            TabIndex        =   2
            Top             =   525
            Width           =   3150
         End
         Begin VB.CheckBox chkColores 
            Caption         =   "Emplear &colores en el reporte"
            Height          =   300
            Left            =   105
            TabIndex        =   3
            Top             =   885
            Value           =   1  'Checked
            Width           =   2490
         End
         Begin VB.CheckBox chkLeyenda 
            Caption         =   "Poner &leyenda en el reporte"
            Height          =   300
            Left            =   105
            TabIndex        =   4
            Top             =   1245
            Value           =   1  'Checked
            Width           =   2400
         End
         Begin VB.TextBox txtDirLimit 
            Height          =   315
            Left            =   2775
            TabIndex        =   62
            Top             =   2580
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAgregarDB_Click()
    If chkAgregarDB.value = vbUnchecked Then
        fraMP3.Enabled = False
        fraOpcionesDB.Enabled = False
        fraOtrosBD.Enabled = False
    Else
        If gb_DBConexionOK And gb_DBFormatOK Then
            fraMP3.Enabled = True
            fraOpcionesDB.Enabled = True
            fraOtrosBD.Enabled = True
        End If
    End If
End Sub

Private Sub chkDirLimit_Click()
    If chkDirLimit.value = vbUnchecked Then
        txtDirLimit.Enabled = False
    Else
        txtDirLimit.Enabled = True
    End If
End Sub

Private Sub chkFileLimit_Click()
    If chkFileLimit.value = vbUnchecked Then
        txtFileLimit.Enabled = False
    Else
        txtFileLimit.Enabled = True
    End If
End Sub

Private Sub chkScript_Click()
    If chkScript.value = vbUnchecked Then
        txtScriptFile.Enabled = False
        cmdSearchScript.Enabled = False
    Else
        txtScriptFile.Enabled = True
        cmdSearchScript.Enabled = True
    End If
End Sub

Private Sub cmdSearchScript_Click()

    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'esconde casilla de solo lectura y verifica que el archivo y el path existan
        .flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .DialogTitle = "Indicar el script a utilizar:"
        .Filter = "Scripts VBS (*.vbs)|*.vbs|Todos los Archivos(*.*)|*.*"
        .InitDir = App.Path & "\scripts"    ' de no existir el directorio usara el directorio activo
        'tipo predefinido VBS
        .FilterIndex = 1
        .ShowOpen
        If .filename <> "" Then
            'cargar el script
            txtScriptFile.text = .filename
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "cmdSearchScript_Click"
    End If
End Sub

'*******************************************************************************
' INICIALIZACION FORMULARIO
'*******************************************************************************
Private Sub Form_Load()

    LoadGeneralOptions
    LoadReportOptions
    LoadDBOptions

End Sub

'*******************************************************************************
' ACEPTAR CAMBIOS
'*******************************************************************************
Private Sub cmdAceptar_Click()
    On Error GoTo Handler
    SaveGeneralOptions
    SaveReportOptions
    SaveDBOptions
    Unload Me
    Exit Sub
Handler:    MsgBox Err.Description, vbCritical, "cmdAceptar_Click"
End Sub

'*******************************************************************************
' IGNORAR CAMBIOS
'*******************************************************************************
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'*******************************************************************************
' OPCIONES GENERALES
'*******************************************************************************
Private Sub LoadGeneralOptions()

    If gb_IncluirSubdirectorios Then
        chkIncluirSubDir.value = vbChecked
    Else
        chkIncluirSubDir.value = vbUnchecked
    End If
    
    '=============================================
    ' Agregar algunos tipos predefinidos
    '
    cmbTipoArchivo.AddItem "*"
    cmbTipoArchivo.AddItem "mp3;ogg;mpc;wav;mid;mod"
    cmbTipoArchivo.AddItem "pdf;ps;doc;chm;djvu"
    cmbTipoArchivo.AddItem "exe;zip;rar;msi;iso"
    cmbTipoArchivo.AddItem "mpg;mpeg;avi;wmv;mov;rm;mkv"
    cmbTipoArchivo.AddItem "htm;html;js;css;vbs;php;asp"
    cmbTipoArchivo.AddItem "frm;cls;bas;ctl;dob;vbs"
    '---------------------------------------------

    If gb_DBConexionOK And gb_DBFormatOK Then
        chkAgregarDB.Enabled = True
    Else
        chkAgregarDB.Enabled = False
        gb_ExportToDB = False
    End If
    
    If gb_ColoresEnReporte = True Then
        chkColores.value = vbChecked
    Else
        chkColores.value = vbUnchecked
    End If

    If gb_LeyendaEnReporte = True Then
        chkLeyenda.value = vbChecked
    Else
        chkLeyenda.value = vbUnchecked
    End If

    If gb_TodosLosArchivos = True Then
        chkTodos.value = vbChecked
    Else
        chkTodos.value = vbUnchecked
    End If

    If gb_ExportToDB = True Then
        chkAgregarDB.value = vbChecked
    Else
        chkAgregarDB.value = vbUnchecked
    End If
    chkAgregarDB_Click

    txtScriptFile = gs_ScriptFile
    
    If gb_ScriptActivated = True Then
        chkScript.value = vbChecked
    Else
        chkScript.value = vbUnchecked
    End If
    chkScript_Click

    If gb_ActivateDirDepthLimit = True Then
        chkDirLimit.value = vbChecked
    Else
        chkDirLimit.value = vbUnchecked
    End If
    txtDirLimit.text = gn_DirDepthLimit
    chkDirLimit_Click

    If gb_ActivateDirFileLimit = True Then
        chkFileLimit.value = vbChecked
    Else
        chkFileLimit.value = vbUnchecked
    End If
    txtFileLimit.text = gn_DirFileLimit
    chkFileLimit_Click

End Sub

Private Sub SaveGeneralOptions()
    
    If chkIncluirSubDir.value = vbChecked Then
        gb_IncluirSubdirectorios = True
        frmR_Porter.chkIncluirSubDir.value = vbChecked
    Else
        gb_IncluirSubdirectorios = False
        frmR_Porter.chkIncluirSubDir.value = vbUnchecked
    End If

    'setear los parametros
    If chkColores.value = vbChecked Then
        gb_ColoresEnReporte = True
    Else
        gb_ColoresEnReporte = False
    End If
    
    If chkAgregarDB.value = vbChecked Then
        gb_ExportToDB = True
    Else
        gb_ExportToDB = False
    End If
    
    If chkLeyenda.value = vbChecked Then
        gb_LeyendaEnReporte = True
    Else
        gb_LeyendaEnReporte = False
    End If
    
    If chkTodos.value = vbChecked Then
        gb_TodosLosArchivos = True
    Else
        ' Verificar extensiones
        If True = VerificarExtensiones() Then
            gb_TodosLosArchivos = False
        Else
            MsgBox "Campo de extensiones incorrecto." & vbCrLf & "Intente por ejemplo:" & vbCrLf & "mp3;mpeg;gs;*", vbExclamation, "Error"
            gb_TodosLosArchivos = True
            cmbTipoArchivo.SetFocus
            Exit Sub
        End If
    End If

    If chkScript.value = vbChecked Then
        If Trim(txtScriptFile) = "" Then
            MsgBox "Debe seleccionar el archivo con el script." & vbCrLf & "Se desactivará la opción de script.", vbExclamation, "Script no seleccionado"
            chkScript.value = vbUnchecked
        Else
            gb_ScriptActivated = True
            gs_ScriptFile = Trim(txtScriptFile)
        End If
    Else
        gb_ScriptActivated = False
    End If

    If chkDirLimit.value = vbChecked Then
        gb_ActivateDirDepthLimit = True
        gn_DirDepthLimit = CInt(txtDirLimit.text)
    Else
        gb_ActivateDirDepthLimit = False
    End If

    If chkFileLimit.value = vbChecked Then
        gb_ActivateDirFileLimit = True
        gn_DirFileLimit = CInt(txtFileLimit.text)
    Else
        gb_ActivateDirFileLimit = False
    End If

End Sub

Private Sub chkColores_Click()
    If chkColores.value = vbChecked Then
        chkLeyenda.Enabled = True
    Else
        chkLeyenda.Enabled = False
        chkLeyenda.value = vbUnchecked
    End If
End Sub

Private Sub chkTodos_Click()
    '=============================================
    Dim k As Integer
    '=============================================

    On Error GoTo Handler

    If chkTodos.value = vbChecked Then
        fraTipo.Enabled = False
        cmbTipoArchivo.text = "*"
    Else
        fraTipo.Enabled = True
        cmbTipoArchivo.text = ""

        For k = 1 To gt_Extensiones.num
            cmbTipoArchivo.text = cmbTipoArchivo.text & Trim(gt_Extensiones.exten(k))
            If k < gt_Extensiones.num Then
                cmbTipoArchivo.text = cmbTipoArchivo.text & ";"
            End If
        Next

        cmbTipoArchivo.SetFocus ' <- throw error number 5

    End If
    
    Exit Sub
    
Handler:

    If Err.Number = 5 Then
        ' no se puede dar enfoque en evento load
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "chkTodos_Click()"
    End If
    
End Sub

Private Function VerificarExtensiones() As Boolean
    '=============================================
    Dim cad As String
    '=============================================

    On Error GoTo Handler
    
    cad = cmbTipoArchivo.text
    
    If cad = "" Then
        GoTo Handler
    Else
        gt_Extensiones.num = 0
        
        Do
            If Extract(cad) = False Then
                GoTo Handler
            Else
                If cad = "" Then
                    Exit Do
                End If
            End If
        Loop
    End If
    
    VerificarExtensiones = True
    Exit Function
    
Handler:

    VerificarExtensiones = False
    Exit Function

End Function

Private Function Extract(ByRef cad As String) As Boolean
    '=============================================
    Dim sz As String
    Dim n As Integer
    '=============================================

    On Error GoTo Handler
    
    If cad = "" Then
        GoTo Handler
    Else
        n = InStr(1, cad, ";")
        
        If n = 0 Then
            cad = Trim(cad)
            If (Len(cad) <= 4) And (Len(cad) > 0) Then
                gt_Extensiones.num = gt_Extensiones.num + 1
                gt_Extensiones.exten(gt_Extensiones.num) = cad
                cad = ""
            Else
                GoTo Handler
            End If
        Else
            sz = Mid(cad, 1, n - 1)
            sz = Trim(sz)
            If (Len(sz) <= 4) And (Len(sz) > 0) Then
                gt_Extensiones.num = gt_Extensiones.num + 1
                gt_Extensiones.exten(gt_Extensiones.num) = sz
                cad = Mid(cad, n + 1)
            Else
                GoTo Handler
            End If
        End If
    End If
    
    Extract = True
    Exit Function
    
Handler:

    Extract = False
    Exit Function

End Function

'*******************************************************************************
' OPCIONES REPORTE
'*******************************************************************************
Private Sub LoadReportOptions()

    If gb_TodosLosArchivos = True Then
        Me.optPorAtributo.value = True
    Else
        Me.optPorTipo.value = True
    End If

End Sub

Private Sub SaveReportOptions()
    
    If Me.optPorAtributo.value = True Then
        gl_ColorDirNormal = Me.lblDirNormal.ForeColor
        gl_ColorDirReadOnly = Me.lblDirReadOnly.ForeColor
        gl_ColorDirHidden = Me.lblDirHidden.ForeColor
        gl_ColorDirOther = Me.lblDirOther.ForeColor
        gl_ColorFileNormal = Me.lblFileNormal.ForeColor
        gl_ColorFileReadOnly = Me.lblFileReadOnly.ForeColor
        gl_ColorFileHidden = Me.lblFileHidden.ForeColor
        gl_ColorFileOther = Me.lblFileOther.ForeColor
    Else
        gl_File1 = Me.lblDirNormal.ForeColor
        gl_File2 = Me.lblDirReadOnly.ForeColor
        gl_File3 = Me.lblDirHidden.ForeColor
        gl_File4 = Me.lblDirOther.ForeColor
        gl_File5 = Me.lblFileNormal.ForeColor
        gl_File6 = Me.lblFileReadOnly.ForeColor
        gl_File7 = Me.lblFileHidden.ForeColor
        gl_File8 = Me.lblFileOther.ForeColor
    End If

End Sub

Private Sub cmdChange_Click(Index As Integer)

    On Error GoTo Handler
    
    With frmR_Porter.cmmdlg
    
        .CancelError = True
        Select Case Index
            Case 0
                .Color = Me.lblDirNormal.ForeColor
            Case 1
                .Color = Me.lblDirReadOnly.ForeColor
            Case 2
                .Color = Me.lblDirHidden.ForeColor
            Case 3
                .Color = Me.lblDirOther.ForeColor
            Case 4
                .Color = Me.lblFileNormal.ForeColor
            Case 5
                .Color = Me.lblFileReadOnly.ForeColor
            Case 6
                .Color = Me.lblFileHidden.ForeColor
            Case 7
                .Color = Me.lblFileOther.ForeColor
        End Select
        
        .flags = cdlCCFullOpen + cdlCCRGBInit
        .ShowColor
        
        Select Case Index
            Case 0
                Me.lblDirNormal.ForeColor = .Color
            Case 1
                Me.lblDirReadOnly.ForeColor = .Color
            Case 2
                Me.lblDirHidden.ForeColor = .Color
            Case 3
                Me.lblDirOther.ForeColor = .Color
            Case 4
                Me.lblFileNormal.ForeColor = .Color
            Case 5
                Me.lblFileReadOnly.ForeColor = .Color
            Case 6
                Me.lblFileHidden.ForeColor = .Color
            Case 7
                Me.lblFileOther.ForeColor = .Color
        End Select
        
    End With
    
Handler:    Exit Sub
End Sub

Private Sub optPorAtributo_Click()

    If Me.optPorAtributo.value = True Then
        Me.lblDirHidden.Caption = "Directorio escondido"
        Me.lblDirNormal.Caption = "Directorio normal"
        Me.lblDirOther.Caption = "Otro tipo de directorio"
        Me.lblDirReadOnly.Caption = "Directorio de sólo lectura"
        Me.lblFileHidden.FontBold = False
        Me.lblFileHidden.Caption = "Archivo escondido"
        Me.lblFileNormal.FontBold = False
        Me.lblFileNormal.Caption = "Archivo normal"
        Me.lblFileOther.FontBold = False
        Me.lblFileOther.Caption = "Otro tipo de archivo"
        Me.lblFileReadOnly.FontBold = False
        Me.lblFileReadOnly.Caption = "Archivo de sólo lectura"
        
        Me.lblDirNormal.ForeColor = gl_ColorDirNormal
        Me.lblDirReadOnly.ForeColor = gl_ColorDirReadOnly
        Me.lblDirHidden.ForeColor = gl_ColorDirHidden
        Me.lblDirOther.ForeColor = gl_ColorDirOther
        Me.lblFileNormal.ForeColor = gl_ColorFileNormal
        Me.lblFileReadOnly.ForeColor = gl_ColorFileReadOnly
        Me.lblFileHidden.ForeColor = gl_ColorFileHidden
        Me.lblFileOther.ForeColor = gl_ColorFileOther
    End If

End Sub

Private Sub optPorTipo_Click()
    
    If Me.optPorTipo.value = True Then
        Me.lblDirNormal.Caption = "ARCHIVO TIPO 1"
        Me.lblDirReadOnly.Caption = "ARCHIVO TIPO 2"
        Me.lblDirHidden.Caption = "ARCHIVO TIPO 3"
        Me.lblDirOther.Caption = "ARCHIVO TIPO 4"
        Me.lblFileNormal.FontBold = True
        Me.lblFileNormal.Caption = "ARCHIVO TIPO 5"
        Me.lblFileReadOnly.FontBold = True
        Me.lblFileReadOnly.Caption = "ARCHIVO TIPO 6"
        Me.lblFileHidden.FontBold = True
        Me.lblFileHidden.Caption = "ARCHIVO TIPO 7"
        Me.lblFileOther.FontBold = True
        Me.lblFileOther.Caption = "ARCHIVO TIPO 8"
        
        Me.lblDirNormal.ForeColor = gl_File1
        Me.lblDirReadOnly.ForeColor = gl_File2
        Me.lblDirHidden.ForeColor = gl_File3
        Me.lblDirOther.ForeColor = gl_File4
        Me.lblFileNormal.ForeColor = gl_File5
        Me.lblFileReadOnly.ForeColor = gl_File6
        Me.lblFileHidden.ForeColor = gl_File7
        Me.lblFileOther.ForeColor = gl_File8
    End If
    
End Sub

'*******************************************************************************
' OPCIONES ANEXION BD & MP3
'*******************************************************************************
Private Sub LoadDBOptions()

    On Error GoTo Handler
    
    '===================================================
    ' Por si no esta iniciada la conexion...
    '
    If gb_DBConexionOK And gb_DBFormatOK Then
        FillStorageTypes
        FillStorageCategory
        Me.Caption = Me.Caption & "  [Conexión exitosa]"
    Else
        Me.Caption = Me.Caption & "  [Sin Conexión con BD]"
        Exit Sub
    End If
    '---------------------------------------------------

    txtDefaultPriority = gn_DefaultPriority

    txtMediaName = gs_MediaName
    txtMediaName.MaxLength = DB_MAX_LEN_STORAGE_NAME
    
    txtMediaComment = gs_MediaComment
    txtMediaComment.MaxLength = DB_MAX_LEN_STORAGE_COMMENT
    
    txtBuffer = gn_DefaultBufferLen
    If gb_CheckBitrateVariable = True Then
        chkVBR.value = vbChecked
    Else
        chkVBR.value = vbUnchecked
    End If
    
    chkCalidad_Click
    
    If gb_SetGenreBySubdir = True Then
        chkGenero.value = vbChecked
    Else
        chkGenero.value = vbUnchecked
    End If

    If gb_SetQualityMP3 = True Then
        chkCalidad.value = vbChecked
    Else
        chkCalidad.value = vbUnchecked
    End If

    If gb_SetInfoMPEG = True Then
        chkMPEG.value = vbChecked
    Else
        chkMPEG.value = vbUnchecked
    End If

    If gb_SetTypeByExtent = True Then
        chkTipo.value = vbChecked
    Else
        chkTipo.value = vbUnchecked
    End If

    If gb_ConsiderFileAuthorName = True Then
        chkNombre.value = vbChecked
    Else
        chkNombre.value = vbUnchecked
    End If
    
    If gb_ExportDirToDB = True Then
        chkAddDirectories.value = vbChecked
    Else
        chkAddDirectories.value = vbUnchecked
    End If
    
    If gb_AddToStorageExistent = True Then
        chkAddRegistersToMedia.value = vbChecked
    Else
        chkAddRegistersToMedia.value = vbUnchecked
    End If
    
    Exit Sub
    
Handler:    MsgBox Err.Description, vbCritical, "LoadDBOptions()"
End Sub

Private Sub SaveDBOptions()
    
    On Error GoTo Handler
    
    '===================================================
    ' Por si no esta iniciada la conexion...
    '
    If gb_DBConexionOK And gb_DBFormatOK Then
        ' continuar normalmente
    Else
        Exit Sub
    End If
    '---------------------------------------------------
    
    If Trim(txtMediaName) = "" Then
        MsgBox "Debe especificar un nombre para el medio de almacenamiento", vbExclamation
        txtMediaName.SetFocus
        Exit Sub
    Else
        gs_MediaName = Trim(txtMediaName)       'se corregira la cadena despues (en la insercion  a la BD)
    End If
    
    If chkAgregarDB.value = vbChecked Then
        ' Verificar si medio no existe ya
        Verify_Storage gs_MediaName
    End If
    
    If cmbStorageType.ListCount > 0 Then
        gl_IndexStorageType = cmbStorageType.ListIndex
        gl_StorageType = cmbStorageType.ItemData(cmbStorageType.ListIndex)
    End If
    
    If cmbCategory.ListCount > 0 Then
        gs_StorageCategory = cmbCategory.text
        gl_StorageCategory = cmbCategory.ItemData(cmbCategory.ListIndex)
    End If
    
    gs_MediaComment = Trim(txtMediaComment)     'se corregira la cadena despues (en la insercion  a la BD)
    
    If chkCalidad.value = vbChecked Then
        gb_SetQualityMP3 = True
    Else
        gb_SetQualityMP3 = False
    End If
    
    If chkMPEG.value = vbChecked Then
        gb_SetInfoMPEG = True
    Else
        gb_SetInfoMPEG = False
    End If
    
    If chkGenero.value = vbChecked Then
        gb_SetGenreBySubdir = True
    Else
        gb_SetGenreBySubdir = False
    End If
    
    If chkNombre.value = vbChecked Then
        gb_ConsiderFileAuthorName = True
    Else
        gb_ConsiderFileAuthorName = False
    End If
    
    If chkTipo.value = vbChecked Then
        gb_SetTypeByExtent = True
    Else
        gb_SetTypeByExtent = False
    End If
    
    If CByte(txtDefaultPriority) < 0 Then
        MsgBox "La prioridad debe ser un número entero", vbExclamation, "Error"
        txtDefaultPriority.SetFocus
        Exit Sub
    Else
        gn_DefaultPriority = CByte(txtDefaultPriority)
    End If
    
    If CInt(txtBuffer) < 512 Then
        MsgBox "El buffer debe ser un número entero (preferentemente entre 1024 y 8192)", vbExclamation, "Error"
        txtBuffer.SetFocus
        Exit Sub
    Else
        gn_DefaultBufferLen = CInt(txtBuffer)
    End If
    
    If chkVBR.value = vbChecked Then
        gb_CheckBitrateVariable = True
    Else
        gb_CheckBitrateVariable = False
    End If
    
    If chkAddDirectories.value = vbChecked Then
        gb_ExportDirToDB = True
    Else
        gb_ExportDirToDB = False
    End If
    
    If chkAddRegistersToMedia.value = vbChecked Then
        gb_AddToStorageExistent = True
    Else
        gb_AddToStorageExistent = False
    End If
    
    Unload Me
    
    Exit Sub
    
Handler:

    If Err.Number = 13 Then
        'txtDefaultPriority no era un numero
        'txtBuffer no era un numero
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "cmdAceptar"
    End If
    
End Sub

Private Sub chkAddRegistersToMedia_Click()
    '===================================================
    Dim rs As ADODB.Recordset
    Dim strTempMediaName As String
    '===================================================
    
    On Error GoTo Handler
    
    If chkAddRegistersToMedia.value = vbChecked Then
        
        gfnc_ParseString Trim(txtMediaName), strTempMediaName     'esta cadena viene sin corregir
        
        Set rs = New ADODB.Recordset
        query = "SELECT * FROM storage WHERE name='" & strTempMediaName & "'"
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        If rs.EOF = False Then
            gl_IndexStorageExistent = rs!id_storage
            gb_AddToStorageExistent = True
        Else
            MsgBox "El medio: [" & txtMediaName & "] no existe", vbExclamation, "Error"
            gl_IndexStorageExistent = 0
            chkAddRegistersToMedia.value = vbUnchecked
        End If
        rs.Close
        
    Else
        gl_IndexStorageExistent = 0
    End If
    
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "chkAddRegistersToMedia_Click"
End Sub

Private Sub chkCalidad_Click()
    If chkCalidad.value = vbChecked Then
        chkVBR.Enabled = True
        txtBuffer.Enabled = True
    Else
        chkVBR.Enabled = False
        txtBuffer.Enabled = False
    End If
End Sub

Private Sub FillStorageTypes()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================

    On Error GoTo Handler

    Set rs = New ADODB.Recordset

    '************************************
    'cargar la tabla de tipos de medios
    '************************************
    query = "SELECT * FROM storage_type ORDER BY storage_type"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbStorageType.AddItem rs!storage_type
        cmbStorageType.ItemData(cmbStorageType.NewIndex) = rs!id_storage_type
        rs.MoveNext
    Wend
    rs.Close

    cmbStorageType.ListIndex = gl_IndexStorageType
    '------------------------------------

    Exit Sub

Handler:

    If Err.Number = 380 Then
        'indice fuera de rango para el cmbStorageType
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "FillStorageTypes()"
    End If
End Sub

Private Sub FillStorageCategory()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================

    On Error GoTo Handler

    Set rs = New ADODB.Recordset

    '************************************
    'cargar la tabla de categorias
    '************************************
    query = "SELECT * FROM category ORDER BY category"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbCategory.AddItem rs!category
        cmbCategory.ItemData(cmbCategory.NewIndex) = rs!id_category
        rs.MoveNext
    Wend
    rs.Close
    

    cmbCategory.text = gs_StorageCategory
    '------------------------------------

    Exit Sub

Handler:

    If Err.Number = 383 Then
        'el texto no pudo ser encontrado en el cmbCategory
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "FillStorageCategory()"
    End If
End Sub

Private Sub txtDefaultPriority_GotFocus()
    txtDefaultPriority.SelStart = 0
    txtDefaultPriority.SelLength = Len(txtDefaultPriority.text)
End Sub

Private Sub Verify_Storage(ByVal storage_name As String)
    '===================================================
    Dim rs As ADODB.Recordset
    Dim strTempMediaName As String
    '===================================================

    On Error GoTo Handler

    gfnc_ParseString storage_name, strTempMediaName     'esta cadena viene sin corregir
    
    Set rs = New ADODB.Recordset
    query = "SELECT * FROM storage WHERE name='" & strTempMediaName & "'"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF = False Then
        If chkAddRegistersToMedia.value = vbUnchecked Then
            If vbYes = MsgBox("Ya existe un medio registrado con el nombre:" & vbCrLf & _
                              "[" & storage_name & "]. Si deseas agregar nuevos registros" & vbCrLf & _
                              "a este medio debes activar la casilla:" & vbCrLf & _
                              "[Agregar registros al medio existente] de [Otros]" & vbCrLf & _
                              "¿Deseas hacerlo ahora?" & vbCrLf & _
                              "-- [Si] agregará los registros a ese medio --", vbExclamation + vbYesNo, "Advertencia") Then
                chkAddRegistersToMedia.value = vbChecked
            Else
                chkAddRegistersToMedia.value = vbUnchecked
            End If
        End If
    Else
        If chkAddRegistersToMedia.value = vbChecked Then
            MsgBox "Tienes activa la casilla:" & vbCrLf & _
                   "[Agregar registros al medio existente] de [Otros]" & vbCrLf & _
                   "Pero el medio: [" & txtMediaName & "] no existe" & vbCrLf & _
                   "Se creará un nuevo medio", vbExclamation, "No existe medio"
            chkAddRegistersToMedia.value = vbUnchecked
        End If
    End If
    
    rs.Close
    Exit Sub

Handler:
    MsgBox Err.Description, vbCritical, "Verify_Storage"
End Sub

