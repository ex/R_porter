VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRegistros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmRegistros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   6405
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3825
      TabIndex        =   46
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   315
      Left            =   2160
      TabIndex        =   45
      Top             =   4140
      Width           =   1215
   End
   Begin TabDlg.SSTab sstab 
      Height          =   4005
      Left            =   0
      TabIndex        =   47
      Top             =   45
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&1)     General       "
      TabPicture(0)   =   "frmRegistros.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGeneral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2)      Otros        "
      TabPicture(1)   =   "frmRegistros.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraLaboral"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3)       Medio       "
      TabPicture(2)   =   "frmRegistros.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraReferencias"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraLaboral 
         Height          =   3585
         Left            =   -74910
         TabIndex        =   50
         Top             =   315
         Width           =   6855
         Begin VB.TextBox txtTipoCalidad 
            Height          =   315
            Left            =   2820
            TabIndex        =   26
            Top             =   1470
            Width           =   1155
         End
         Begin VB.TextBox txtCalidad 
            Height          =   315
            Left            =   1575
            TabIndex        =   25
            Top             =   1470
            Width           =   1185
         End
         Begin VB.TextBox txtSysLength 
            Height          =   315
            Left            =   1575
            MaxLength       =   15
            TabIndex        =   21
            Top             =   660
            Width           =   2400
         End
         Begin VB.ComboBox cmbFileType 
            Height          =   315
            ItemData        =   "frmRegistros.frx":019E
            Left            =   1575
            List            =   "frmRegistros.frx":01A0
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   255
            Width           =   2400
         End
         Begin VB.TextBox txtFileObservation 
            Height          =   1155
            Left            =   1575
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   2280
            Width           =   2400
         End
         Begin VB.TextBox txtFileTypeObserv 
            Height          =   315
            Left            =   1575
            TabIndex        =   28
            Top             =   1875
            Width           =   2400
         End
         Begin MSComCtl2.DTPicker dtFileDate 
            Height          =   315
            Left            =   1575
            TabIndex        =   23
            Top             =   1065
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyy          HH:mm:ss "
            Format          =   59113475
            CurrentDate     =   27431
         End
         Begin VB.Label lbl 
            Caption         =   "&Comentario"
            Height          =   195
            Index           =   20
            Left            =   135
            TabIndex        =   29
            Top             =   2385
            Width           =   1290
         End
         Begin VB.Label lbl 
            Caption         =   "&Observación"
            Height          =   195
            Index           =   19
            Left            =   135
            TabIndex        =   27
            Top             =   1980
            Width           =   1380
         End
         Begin VB.Label lbl 
            Caption         =   "&Calidad"
            Height          =   195
            Index           =   18
            Left            =   135
            TabIndex        =   24
            Top             =   1560
            Width           =   1320
         End
         Begin VB.Label lbl 
            Caption         =   "Ta&maño   (bytes)"
            Height          =   195
            Index           =   17
            Left            =   135
            TabIndex        =   20
            Top             =   735
            Width           =   1305
         End
         Begin VB.Label lbl 
            Caption         =   "&Fecha archivo"
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   22
            Top             =   1155
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "T&ipo archivo"
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   18
            Top             =   315
            Width           =   1200
         End
      End
      Begin VB.Frame fraReferencias 
         Height          =   3585
         Left            =   -74910
         TabIndex        =   49
         Top             =   315
         Width           =   6855
         Begin VB.TextBox txtStorageCategory 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1065
            Width           =   2400
         End
         Begin VB.TextBox txtStorageObservation 
            ForeColor       =   &H80000012&
            Height          =   750
            Left            =   1575
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   2685
            Width           =   2400
         End
         Begin VB.TextBox txtEtiqueta 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1470
            Width           =   2400
         End
         Begin VB.TextBox txtIDStorage 
            ForeColor       =   &H80000010&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   255
            Width           =   2400
         End
         Begin VB.TextBox txtSerial 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   2280
            Width           =   2400
         End
         Begin VB.ComboBox cmbStorage 
            Height          =   315
            ItemData        =   "frmRegistros.frx":01A2
            Left            =   1575
            List            =   "frmRegistros.frx":01A4
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   660
            Width           =   2400
         End
         Begin VB.TextBox txtStorageType 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   660
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker dtFechaMedio 
            Height          =   315
            Left            =   1575
            TabIndex        =   40
            Top             =   1875
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyy          HH:mm:ss"
            Format          =   59113475
            CurrentDate     =   27431
         End
         Begin VB.Label lbl 
            Caption         =   "&Tipo"
            Height          =   195
            Index           =   14
            Left            =   4065
            TabIndex        =   52
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lbl 
            Caption         =   "&Serial"
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   41
            Top             =   2385
            Width           =   1200
         End
         Begin VB.Label lbl 
            Caption         =   "&Fecha medio"
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   39
            Top             =   1965
            Width           =   1230
         End
         Begin VB.Label lbl 
            Caption         =   "ID medio"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   31
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label lbl 
            Caption         =   "&Medio"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   33
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lbl 
            Caption         =   "Categoria"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   35
            Top             =   1155
            Width           =   990
         End
         Begin VB.Label lbl 
            Caption         =   "&Etiqueta"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   37
            Top             =   1545
            Width           =   810
         End
         Begin VB.Label lbl 
            Caption         =   "Obser&vación"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   43
            Top             =   2805
            Width           =   1260
         End
      End
      Begin VB.Frame fraGeneral 
         Height          =   3585
         Left            =   90
         TabIndex        =   48
         Top             =   315
         Width           =   6855
         Begin VB.CommandButton cmdSearch 
            Height          =   315
            Left            =   6510
            TabIndex        =   4
            Top             =   660
            Width           =   255
         End
         Begin VB.ComboBox cmbFirstLetter 
            Height          =   315
            ItemData        =   "frmRegistros.frx":01A6
            Left            =   1575
            List            =   "frmRegistros.frx":0201
            TabIndex        =   8
            Top             =   1470
            Width           =   660
         End
         Begin VB.TextBox txtIDFile 
            ForeColor       =   &H80000010&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   " "
            Top             =   255
            Width           =   2400
         End
         Begin VB.ComboBox cmbGenre 
            Height          =   315
            ItemData        =   "frmRegistros.frx":025C
            Left            =   1575
            List            =   "frmRegistros.frx":025E
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2685
            Width           =   5205
         End
         Begin VB.ComboBox cmbSubGenre 
            Height          =   315
            ItemData        =   "frmRegistros.frx":0260
            Left            =   1575
            List            =   "frmRegistros.frx":0262
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   3090
            Width           =   5205
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1575
            TabIndex        =   6
            Top             =   1065
            Width           =   5205
         End
         Begin VB.TextBox txtArchivo 
            ForeColor       =   &H80000003&
            Height          =   315
            Left            =   1575
            TabIndex        =   3
            Top             =   660
            Width           =   4905
         End
         Begin VB.TextBox txtPrioridad 
            Height          =   315
            Left            =   1575
            MaxLength       =   2
            TabIndex        =   13
            Top             =   2280
            Width           =   5205
         End
         Begin VB.ComboBox cmbParent 
            Height          =   315
            ItemData        =   "frmRegistros.frx":0264
            Left            =   1575
            List            =   "frmRegistros.frx":0266
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1875
            Width           =   5205
         End
         Begin VB.ComboBox cmbAuthor 
            Height          =   315
            ItemData        =   "frmRegistros.frx":0268
            Left            =   2250
            List            =   "frmRegistros.frx":026A
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1470
            Width           =   4530
         End
         Begin VB.Label lbl 
            Caption         =   "&Pertenece a"
            Height          =   195
            Index           =   21
            Left            =   195
            TabIndex        =   10
            Top             =   1905
            Width           =   1305
         End
         Begin VB.Label lbl 
            Caption         =   "&Sub-género"
            Height          =   195
            Index           =   12
            Left            =   195
            TabIndex        =   16
            Top             =   3120
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "ID archivo"
            Height          =   195
            Index           =   15
            Left            =   195
            TabIndex        =   0
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "A&utor"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   7
            Top             =   1500
            Width           =   1200
         End
         Begin VB.Label lbl 
            Caption         =   "&Nombre"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   5
            Top             =   1110
            Width           =   1230
         End
         Begin VB.Label lbl 
            Caption         =   "&Archivo"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   2
            Top             =   705
            Width           =   630
         End
         Begin VB.Label lbl 
            Caption         =   "P&rioridad"
            Height          =   195
            Index           =   7
            Left            =   195
            TabIndex        =   12
            Top             =   2310
            Width           =   1320
         End
         Begin VB.Label lbl 
            Caption         =   "Géne&ro"
            Height          =   195
            Index           =   8
            Left            =   195
            TabIndex        =   14
            Top             =   2715
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************
'* VARIABLES GLOBALES
'**************************************************************
Public gml_RowEditRegister As Long
Public gmb_DBAddRegister As Long

'**************************************************************
'* VARIABLES DE MODULO
'**************************************************************
Private mb_RegistroModificado As Boolean
Private mb_FormLoaded As Boolean

'**************************************************************
'* COMBOBOXS
'**************************************************************
Private Sub cmbFirstLetter_Click()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================
    
    On Error GoTo Handler
    
    If Trim(cmbFirstLetter.text) = "" Then
        cmbFirstLetter.text = "*"
    End If
    
    cmbAuthor.Clear
    
    'agregar el autor vacio
    cmbAuthor.AddItem "[Ninguno]"
    cmbAuthor.ItemData(0) = 0
    
    Select Case cmbFirstLetter.text
    
        Case "*"
            query = "SELECT id_author, author FROM author WHERE (id_author > 0) ORDER BY author"
            
        Case "#"
            query = "SELECT id_author, author FROM author WHERE ((author NOT LIKE '[a-z]%') AND (id_author > 0)) ORDER BY author"
            
        Case Else
            query = "SELECT id_author, author FROM author WHERE (author LIKE '" & Me.cmbFirstLetter.text & "%') ORDER BY author"
            
    End Select
    
    'cargar la tabla de autores
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbAuthor.AddItem rs!Author
        cmbAuthor.ItemData(cmbAuthor.NewIndex) = rs!id_author
        rs.MoveNext
    Wend
    rs.Close
    
    cmbAuthor.ListIndex = 1
    cmbAuthor.SetFocus
    
    Exit Sub
    
Handler:
    
    If Err.Number = 380 Then
        'el combo esta vacio
    Else
        If Err.Number = 5 Then
            'Me.cmbAuthor.SetFocus cargando formulario
            Resume Next
        Else
            MsgBox Err.Description, vbCritical, "cmbFirstLetter_Click()"
        End If
    End If
    
End Sub

Private Sub cmbFirstLetter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbFirstLetter_Click
        cmbFirstLetter.SetFocus
    End If
End Sub

Private Sub cmbGenre_Click()
Dim rs As ADODB.Recordset
    
    On Error GoTo Handler
    
    Me.cmbSubGenre.Clear
    
    'agregar el sub-genero vacio
    Me.cmbSubGenre.AddItem "[Ninguno]"
    Me.cmbSubGenre.ItemData(0) = 0
    
    If Me.cmbGenre.text = "[Ninguno]" Then
        Me.cmbSubGenre.ListIndex = 0
        Exit Sub
    End If
    
    'cargar la tabla de sub-generos
    Set rs = New ADODB.Recordset
    query = "SELECT id_sub_genre, sub_genre FROM sub_genre WHERE (id_genre=" & Me.cmbGenre.ItemData(Me.cmbGenre.ListIndex) & ") ORDER BY sub_genre"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        Me.cmbSubGenre.AddItem rs!sub_genre
        Me.cmbSubGenre.ItemData(Me.cmbSubGenre.NewIndex) = rs!id_sub_genre
        rs.MoveNext
    Wend
    rs.Close
    
    If mb_FormLoaded = True Then
        Me.cmbSubGenre.ListIndex = 1
    End If
    
    Exit Sub
    
Handler:
    
    If Err.Number = 380 Then
        'el combo esta vacio
        Me.cmbSubGenre.ListIndex = 0
    Else
        MsgBox Err.Description, vbCritical, "cmbGenre_Click()"
    End If

End Sub

Private Sub cmbStorage_Click()
    '===================================================
    Dim rs As ADODB.Recordset
    '===================================================
    
    On Error GoTo Handler
    
    txtEtiqueta = ""
    txtSerial = ""
    txtStorageObservation = ""
    txtStorageType = ""
    txtStorageCategory = ""
    dtFechaMedio.Enabled = True
    
    'leer los otros datos
    query = "SELECT storage.id_storage, storage.name, storage.fecha, storage.label, storage.serial, storage.comment, storage_type.storage_type, category.category "
    query = query & "FROM storage, storage_type, category "
    query = query & "WHERE ((storage.id_storage = " & cmbStorage.ItemData(cmbStorage.ListIndex) & ") AND "
    query = query & "(storage_type.id_storage_type=storage.id_storage_type) AND "
    query = query & "(category.id_category=storage.id_category))"
    
    Set rs = New ADODB.Recordset
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF = False Then
        txtIDStorage = rs!id_storage
        txtEtiqueta = rs!label
        txtSerial = rs!serial
        txtStorageObservation = rs!Comment
        txtStorageType = rs!storage_type
        txtStorageCategory = rs!category
        dtFechaMedio.value = rs!Fecha
    End If
    rs.Close
    
    Exit Sub
    
Handler:
    
    If Err.Number = 380 Then
        'la fecha esta vacia
        dtFechaMedio.Enabled = False
    Else
        If Err.Number = 94 Then
            'se intento llenar con cadena nula
            Resume Next
        Else
            MsgBox Err.Description, vbCritical, "cmbStorage_Click()"
        End If
    End If

End Sub

'**************************************************************
'* COMMAND BUTTONS
'**************************************************************
Private Sub cmdCancelar_Click()
    mb_RegistroModificado = False
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo Handler
    
    With Me.cmmdlg
        .CancelError = True
        .flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
        If .filename <> "" Then
            Me.txtArchivo = .filename
        End If
    End With
    
Handler:    Exit Sub
End Sub

Private Sub cmdGuardar_Click()
Dim rs As ADODB.Recordset
Dim cd As ADODB.Command
Dim bTransaccionIniciada As Boolean
Dim bEnVerificacion As Boolean
Dim Author As String
Dim Name As String
Dim File_Name As String
Dim TypeQuality As String
Dim TypeComment As String
Dim Comment As String


    On Error GoTo Handler

    bEnVerificacion = True

    '***********************************
    'verificaciones simples
    '***********************************
    If Me.txtArchivo = "" Then
        MsgBox "Archivo no ingresado"
        Me.txtArchivo.SetFocus
        Exit Sub
    End If

    If Me.txtNombre = "" Then
        MsgBox "Nombre no ingresado"
        Me.ssTab.Tab = 0
        Me.txtNombre.SetFocus
        Exit Sub
    End If

    If CInt(Me.txtPrioridad) < 0 Then
        MsgBox "La prioridad debe encontrarse entre [0 - 99]"
        Me.ssTab.Tab = 0
        Me.txtPrioridad.SetFocus
        Exit Sub
    End If

    If CLng(Me.txtSysLength) < 0 Then
        MsgBox "El tamaño debe ser >= 0"
        Me.ssTab.Tab = 1
        Me.txtSysLength.SetFocus
        Exit Sub
    End If

    If CDbl(Me.txtCalidad) < 0 Then
        MsgBox "La calidad debe ser >= 0"
        Me.ssTab.Tab = 1
        Me.txtCalidad.SetFocus
        Exit Sub
    End If

    If Me.cmbAuthor.ListIndex = -1 Then
        MsgBox "Autor no seleccionado"
        Me.ssTab.Tab = 0
        Me.cmbAuthor.SetFocus
        Exit Sub
    End If

    If Me.cmbFileType.ListIndex = -1 Then
        MsgBox "Tipo de archivo no seleccionado"
        Me.cmbFileType.SetFocus
        Exit Sub
    End If

    If Me.cmbGenre.ListIndex = -1 Then
        MsgBox "Género no seleccionado"
        Me.cmbGenre.SetFocus
        Exit Sub
    End If

    If Me.cmbParent.ListIndex = -1 Then
        MsgBox "Grupo de pertenencia no seleccionado"
        Me.cmbParent.SetFocus
        Exit Sub
    End If

    If Me.cmbStorage.ListIndex = -1 Then
        MsgBox "Medio no seleccionado"
        Me.cmbStorage.SetFocus
        Exit Sub
    End If

    If Me.cmbSubGenre.ListIndex = -1 Then
        MsgBox "Sub-género no seleccionado"
        Me.ssTab.Tab = 0
        Me.cmbSubGenre.SetFocus
        Exit Sub
    End If

    bEnVerificacion = False

    '***********************************
    'confirmacion
    '***********************************
    If gmb_DBAddRegister = True Then
        If MsgBox("¿Estas seguro de ingresar el nuevo registro?", vbExclamation + vbYesNo, "Confirme insercion") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("¿Estas seguro de modificar el registro?", vbExclamation + vbYesNo, "Confirme modificación") = vbNo Then
            Exit Sub
        End If
    End If

    '-------------------------------------
    'verificar cadenas
    gfnc_ParseString Me.txtArchivo, File_Name
    gfnc_ParseString Me.cmbAuthor.text, Author
    gfnc_ParseString Me.txtNombre, Name
    gfnc_ParseString Me.txtFileObservation, Comment
    gfnc_ParseString Me.txtFileTypeObserv, TypeComment
    gfnc_ParseString Me.txtTipoCalidad, TypeQuality
    
    '***********************************
    'modificacion DB
    '***********************************
    If gmb_DBAddRegister = True Then
        '***********************************
        'cuando se inserta
        '***********************************
        'iniciamos transaccion para insertar
        If gfnc_CrearConexionTransaccion(gs_DSN, gs_Pwd) = True Then

            cnTransaction.BeginTrans
            bTransaccionIniciada = True

            Set cd = New ADODB.Command
            Set cd.ActiveConnection = cnTransaction

            '-------------------------------------
            ' insertamos archivo
            query = "INSERT INTO file (id_file, id_storage, sys_name, "
            query = query & "id_file_type, sys_length, fecha, "
            query = query & "name, id_parent, priority, "
            query = query & "quality, type_quality, "
            query = query & "type_comment, comment, "
            query = query & "id_author, id_genre, id_sub_genre)"
            query = query & " VALUES ("
            query = query & Trim(Me.txtIDFile) & ", "
            query = query & Trim(Me.txtIDStorage) & ", '"
            query = query & File_Name & "', "
            query = query & Trim(str(Me.cmbFileType.ItemData(Me.cmbFileType.ListIndex))) & ", "
            query = query & Trim(str(CLng(Me.txtSysLength))) & ", #"
            query = query & Format(Me.dtFileDate.Month, "00") & "/" & Format(Me.dtFileDate.Day, "00") & "/" & Format(Me.dtFileDate.Year, "0000") & " " & Format(Me.dtFileDate.Hour, "00") & ":" & Format(Me.dtFileDate.Minute, "00") & ":" & Format(Me.dtFileDate.Second, "00") & "#, '"
            query = query & Name & "', "
            query = query & Trim(str(Me.cmbParent.ItemData(Me.cmbParent.ListIndex))) & ", "
            query = query & Trim(str(CInt(Me.txtPrioridad))) & ", "
            query = query & Trim(str(CDbl(Me.txtCalidad))) & ", '"
            query = query & TypeQuality & "', '"
            query = query & TypeComment & "', '"
            query = query & Comment & "', "
            query = query & Trim(str(Me.cmbAuthor.ItemData(Me.cmbAuthor.ListIndex))) & ", "
            query = query & Trim(str(Me.cmbGenre.ItemData(Me.cmbGenre.ListIndex))) & ", "
            query = query & Trim(str(Me.cmbSubGenre.ItemData(Me.cmbSubGenre.ListIndex))) & ")"
            cd.CommandText = query
            cd.Execute
        
            cnTransaction.CommitTrans
            bTransaccionIniciada = False
            
            gsub_CerrarConexionTransaccion
            
            '********************************
            'actualizar frmDataControl
            '********************************
            '(se hara despues de descargar este formulario)
            If frmDataControl.Visible = True Then

                gs_DBRegName = Me.txtNombre
                gs_DBRegAuthor = Me.cmbAuthor.text
                gs_DBRegStorage = Me.cmbStorage.text
                gs_DBRegGenre = Me.cmbGenre.text
                gs_DBRegPriority = Me.txtPrioridad
                gs_DBRegQuality = Me.txtCalidad
                gs_DBRegFileSize = Me.txtSysLength
                gd_DBRegFileDate = Me.dtFileDate.value
                gs_DBRegParent = Me.cmbParent.text
                gs_DBRegFileType = Me.cmbFileType.text
                
                gl_DBRegIDNew = CLng(Me.txtIDFile)
                
            End If
            
            mb_RegistroModificado = True
            Unload Me

        Else
            MsgBox "No se pudo crear la transaccion para la BD", vbExclamation, "Error de conexion"
            Exit Sub
        End If

    Else
        '***********************************
        'cuando se modifica
        '***********************************
        'iniciamos transaccion para modificar
        If gfnc_CrearConexionTransaccion(gs_DSN, gs_Pwd) = True Then

            cnTransaction.BeginTrans
            bTransaccionIniciada = True

            Set cd = New ADODB.Command
            Set cd.ActiveConnection = cnTransaction

            query = "UPDATE file SET id_storage=" & Trim(Me.txtIDStorage) & ", "
            query = query & "sys_name='" & File_Name & "', "
            query = query & "id_file_type=" & Trim(str(Me.cmbFileType.ItemData(Me.cmbFileType.ListIndex))) & ", "
            query = query & "sys_length=" & Trim(str(CLng(Me.txtSysLength))) & ", "
            query = query & "fecha=#" & Format(Me.dtFileDate.Month, "00") & "/" & Format(Me.dtFileDate.Day, "00") & "/" & Format(Me.dtFileDate.Year, "0000") & " " & Format(Me.dtFileDate.Hour, "00") & ":" & Format(Me.dtFileDate.Minute, "00") & ":" & Format(Me.dtFileDate.Second, "00") & "#, "
            query = query & "name='" & Name & "', "
            query = query & "id_parent=" & Trim(str(Me.cmbParent.ItemData(Me.cmbParent.ListIndex))) & ", "
            query = query & "priority=" & Trim(str(CInt(Me.txtPrioridad))) & ", "
            query = query & "quality=" & Trim(str(CDbl(Me.txtCalidad))) & ", "
            query = query & "type_quality='" & TypeQuality & "', "
            query = query & "type_comment='" & TypeComment & "', "
            query = query & "comment='" & Comment & "', "
            query = query & "id_author=" & Trim(str(Me.cmbAuthor.ItemData(Me.cmbAuthor.ListIndex))) & ", "
            query = query & "id_genre=" & Trim(str(Me.cmbGenre.ItemData(Me.cmbGenre.ListIndex))) & ", "
            query = query & "id_sub_genre=" & Trim(str(Me.cmbSubGenre.ItemData(Me.cmbSubGenre.ListIndex))) & " "
            query = query & "WHERE id_file=" & Trim(Me.txtIDFile) & ";"
            cd.CommandText = query

            cd.Execute

            cnTransaction.CommitTrans
            gsub_CerrarConexionTransaccion

            bTransaccionIniciada = False

            '********************************
            'actualizar frmDataControl
            '********************************
            '(se hara despues de descargar este formulario)
            If frmDataControl.Visible = True Then

                gs_DBRegName = Me.txtNombre
                gs_DBRegAuthor = Me.cmbAuthor.text
                gs_DBRegStorage = Me.cmbStorage.text
                gs_DBRegGenre = Me.cmbGenre.text
                gs_DBRegPriority = Me.txtPrioridad
                gs_DBRegQuality = Me.txtCalidad
                gs_DBRegFileSize = Me.txtSysLength
                gd_DBRegFileDate = Me.dtFileDate.value
                gs_DBRegParent = Me.cmbParent.text
                gs_DBRegFileType = Me.cmbFileType.text
                
            End If

            mb_RegistroModificado = True

            Unload Me

        End If

    End If

    Exit Sub

Handler:

    Select Case Err.Number

        Case 13
            'error de tipo
            If bEnVerificacion = True Then
                Resume Next
            Else
                GoTo Troubles
            End If

        Case Else
            
Troubles:
            If bTransaccionIniciada Then
                cnTransaction.RollbackTrans
                gsub_CerrarConexionTransaccion
            End If
            MsgBox Err.Description, vbCritical, "cmdGuardar_Click()"

    End Select

End Sub

'**************************************************************
'* GUI
'**************************************************************
Private Sub sstab_GotFocus()

    On Error Resume Next

    Select Case ssTab.Tab
        Case 0
            Me.txtArchivo.SetFocus
        Case 1
            Me.cmbFileType.SetFocus
        Case 2
            Me.cmbStorage.SetFocus
    End Select

End Sub

Private Sub cmbFileType_GotFocus()
    If Me.ssTab.Tab <> 1 Then
        Me.ssTab.Tab = 1
    End If
End Sub

Private Sub cmbStorage_GotFocus()
    If Me.ssTab.Tab <> 2 Then
        Me.ssTab.Tab = 2
    End If
End Sub

Private Sub txtArchivo_GotFocus()
    If Me.ssTab.Tab <> 0 Then
        Me.ssTab.Tab = 0
    End If
End Sub

Private Sub txtPrioridad_GotFocus()
    txtPrioridad.SelStart = 0
    txtPrioridad.SelLength = Len(txtPrioridad.text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And (Me.ActiveControl.Name <> "txtFileObservation") Then
        SendKeys "{TAB}"
    End If
End Sub

'**************************************************************
'* FORMULARIO
'**************************************************************
Private Sub Form_Load()
    '===================================================
    Dim rs As ADODB.Recordset
    Dim id_author As Long
    Dim id_genre As Long
    Dim id_sub_genre As Long
    Dim id_parent As Long
    Dim id_storage As Long
    Dim id_file_type As Long
    '===================================================

    On Error GoTo Handler

    Me.width = 7140
    Me.height = 4935

    txtArchivo.MaxLength = DB_MAX_LEN_FILE_SYS_NAME
    txtTipoCalidad.MaxLength = DB_MAX_LEN_FILE_TYPE_QUALITY
    txtNombre.MaxLength = DB_MAX_LEN_FILE_NAME
    txtFileObservation.MaxLength = DB_MAX_LEN_FILE_COMMENT
    txtFileTypeObserv.MaxLength = DB_MAX_LEN_FILE_TYPE_COMMENT
    txtStorageObservation.MaxLength = DB_MAX_LEN_STORAGE_COMMENT
    txtEtiqueta.MaxLength = DB_MAX_LEN_STORAGE_LABEL
    txtSerial.MaxLength = DB_MAX_LEN_STORAGE_SERIAL


    Set rs = New ADODB.Recordset

    '*****************************
    'cargar la tabla de generos
    '*****************************
    cmbGenre.AddItem "[Ninguno]"
    cmbGenre.ItemData(cmbGenre.NewIndex) = 0
    
    query = "SELECT id_genre, genre FROM genre WHERE (id_genre > 0) ORDER BY genre"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbGenre.AddItem rs!genre
        cmbGenre.ItemData(cmbGenre.NewIndex) = rs!id_genre
        rs.MoveNext
    Wend
    rs.Close

    '*****************************
    'cargar la tabla de medios
    '*****************************
    cmbStorage.AddItem "[Ninguno]"
    cmbStorage.ItemData(cmbStorage.NewIndex) = 0
    
    query = "SELECT id_storage, name FROM storage WHERE (id_storage > 0) ORDER BY name"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbStorage.AddItem rs!Name
        cmbStorage.ItemData(cmbStorage.NewIndex) = rs!id_storage
        rs.MoveNext
    Wend
    rs.Close

    '************************************
    'cargar la tabla de tipos de archivo
    '************************************
    cmbFileType.AddItem "[Ninguno]"
    cmbFileType.ItemData(cmbFileType.NewIndex) = 0
    
    query = "SELECT * FROM file_type WHERE (id_file_type > 0) ORDER BY file_type"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbFileType.AddItem rs!file_type
        cmbFileType.ItemData(cmbFileType.NewIndex) = rs!id_file_type
        rs.MoveNext
    Wend
    rs.Close

    '************************************
    'cargar la tabla parents
    '************************************
    cmbParent.AddItem "[Ninguno]"
    cmbParent.ItemData(cmbParent.NewIndex) = 0
    
    query = "SELECT * FROM parent WHERE (id_parent > 0) ORDER BY parent"
    rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    While rs.EOF = False
        cmbParent.AddItem rs!Parent
        cmbParent.ItemData(cmbParent.NewIndex) = rs!id_parent
        rs.MoveNext
    Wend
    rs.Close

    '*****************************
    'cargar otros campos
    '*****************************
    If gb_DBAddRegister = True Then

        '*****************************
        'en caso de agregar registro
        '*****************************
        gmb_DBAddRegister = True
        
        'calcular id insercion
        query = "SELECT MAX(id_file) AS max_id_file FROM file"
        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

        If IsNull(rs!max_id_file) Then
            txtIDFile = 1
        Else
            txtIDFile = rs!max_id_file + 1
        End If

        rs.Close
        Set rs = Nothing

        txtCalidad = "0"
        txtPrioridad = "0"
        txtSysLength = "0"
        mb_FormLoaded = True
        cmbFileType.ListIndex = 0
        cmbGenre.ListIndex = 0
        cmbStorage.ListIndex = 0
        cmbParent.ListIndex = 0
        cmbFirstLetter.text = "A"
        'el cambio de texto no desencadena el evento click
        cmbFirstLetter_Click

    Else
        '*****************************
        'en caso de modificar tabla
        '*****************************
        gmb_DBAddRegister = False

        query = "SELECT file.sys_name, file.name, file.fecha, "
        query = query & "file.sys_length, file.priority, "
        query = query & "file.quality, file.type_quality, "
        query = query & "file.comment, file.type_comment, "
        query = query & "author.author, parent.parent, storage.name AS storage, "
        query = query & "genre.genre, file_type.file_type, sub_genre.sub_genre, category.category "
        query = query & "FROM file, author, parent, storage, genre, sub_genre, file_type, category "
        query = query & "WHERE ((file.id_file=" & gl_DB_IDRegistroModificar & ") AND "
        query = query & "(author.id_author=file.id_author) AND "
        query = query & "(genre.id_genre=file.id_genre) AND "
        query = query & "(storage.id_storage=file.id_storage) AND "
        query = query & "(parent.id_parent=file.id_parent) AND "
        query = query & "(file_type.id_file_type=file.id_file_type) AND "
        query = query & "(sub_genre.id_sub_genre=file.id_sub_genre) AND  "
        query = query & "(category.id_category=storage.id_category))"

        rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

        If rs.EOF = True Then
            '-----------------------------------------------------------
            ' intentara cargar el registro dañado para poder corregirlo
            '-----------------------------------------------------------
            On Error GoTo Handler
            
            MsgBox "Se intentará cargar la información existente." & vbCrLf & "Deberá verificar cada campo y corregir el(los) erróneo(s).", vbExclamation, "[Error] :: [El registro se encuentra dañado]"
            rs.Close
        
            query = "SELECT * FROM file WHERE (file.id_file=" & gl_DB_IDRegistroModificar & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            txtIDFile = gl_DB_IDRegistroModificar
            txtArchivo = rs!Sys_Name
            txtNombre = rs!Name
            txtPrioridad = rs!Priority
            txtSysLength = rs!sys_length
            txtCalidad = rs!Quality
            txtTipoCalidad = rs!type_quality
            txtFileTypeObserv = rs!type_comment
            txtFileObservation = rs!Comment
            txtStorageCategory = rs!category
            
            dtFileDate.value = rs!Fecha
            
            id_author = rs!id_author
            id_parent = rs!id_parent
            id_genre = rs!id_genre
            id_sub_genre = rs!id_sub_genre
            id_storage = rs!id_storage
            id_file_type = rs!id_file_type
            
            rs.Close
            query = "SELECT author FROM author WHERE (id_author=" & Trim(str(id_author)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!Author <> "" Then
                    
                    Select Case UCase(Left(rs!Author, 1))
                        Case "A" To "Z"
                            cmbFirstLetter.text = UCase(Left(rs!Author, 1))
                        Case Else
                            cmbFirstLetter.text = "#"
                    End Select
                    'el cambio de texto no desencadena el evento click
                    cmbFirstLetter_Click
                    
                    cmbAuthor.text = rs!Author
                Else
                    cmbAuthor.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Autor] incorrecto"
            End If
            
            rs.Close
            query = "SELECT parent FROM parent WHERE (id_parent=" & Trim(str(id_parent)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!Parent <> "" Then
                    cmbParent.text = rs!Parent
                Else
                    cmbParent.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Pertenece a] incorrecto"
            End If
            
            rs.Close
            query = "SELECT name AS storage FROM storage WHERE (id_storage=" & Trim(str(id_storage)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!storage <> "" Then
                    cmbStorage.text = rs!storage
                Else
                    cmbStorage.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Medio] incorrecto"
            End If
            
            rs.Close
            query = "SELECT genre FROM genre WHERE (id_genre=" & Trim(str(id_genre)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!genre <> "" Then
                    mb_FormLoaded = False
                    cmbGenre.text = rs!genre
                    mb_FormLoaded = True
                Else
                    cmbGenre.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Género] incorrecto"
            End If
            
            rs.Close
            query = "SELECT file_type FROM file_type WHERE (id_file_type=" & Trim(str(id_file_type)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!file_type <> "" Then
                    cmbFileType.text = rs!file_type
                Else
                    cmbFileType.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Tipo de archivo] incorrecto"
            End If
            
            rs.Close
            query = "SELECT sub_genre FROM sub_genre WHERE (id_sub_genre=" & Trim(str(id_sub_genre)) & ")"
            rs.Open query, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rs.EOF = False Then
                If rs!sub_genre <> "" Then
                    cmbSubGenre.text = rs!sub_genre
                Else
                    cmbSubGenre.text = "[Ninguno]"
                End If
            Else
                MsgBox "[Sub Género] incorrecto"
            End If
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        Select Case UCase(Left(rs!Author, 1))
            Case "A" To "Z"
                cmbFirstLetter.text = UCase(Left(rs!Author, 1))
            Case Else
                cmbFirstLetter.text = "#"
        End Select
        'el cambio de texto no desencadena el evento click
        cmbFirstLetter_Click

        txtIDFile = gl_DB_IDRegistroModificar
        txtArchivo = rs!Sys_Name
        txtNombre = rs!Name
        txtPrioridad = rs!Priority
        txtSysLength = rs!sys_length
        txtCalidad = rs!Quality
        txtTipoCalidad = rs!type_quality
        txtFileTypeObserv = rs!type_comment
        txtFileObservation = rs!Comment
        txtStorageCategory = rs!category

        If rs!Author <> "" Then
            cmbAuthor.text = rs!Author
        Else
            cmbAuthor.text = "[Ninguno]"
        End If
        
        If rs!storage <> "" Then
            cmbStorage.text = rs!storage
        Else
            cmbStorage.text = "[Ninguno]"
        End If
        
        If rs!file_type <> "" Then
            cmbFileType.text = rs!file_type
        Else
            cmbFileType.text = "[Ninguno]"
        End If
        
        If rs!genre <> "" Then
            mb_FormLoaded = False
            cmbGenre.text = rs!genre
            mb_FormLoaded = True
        Else
            cmbGenre.text = "[Ninguno]"
        End If
        
        If rs!sub_genre <> "" Then
            cmbSubGenre.text = rs!sub_genre
        Else
            cmbSubGenre.text = "[Ninguno]"
        End If
        
        If rs!Parent <> "" Then
            cmbParent.text = rs!Parent
        Else
            cmbParent.text = "[Ninguno]"
        End If
        
        dtFileDate.value = rs!Fecha

        rs.Close
        Set rs = Nothing

    End If


    Exit Sub

Handler:

    Select Case Err.Number
        Case 94
            'se intento llenar los txts con una entrada nula
            Resume Next
        Case 380
            'se intento llenar el datepick con una entrada nula
            Resume Next
        Case 383
            'se intento llenar los combobox con una entrada nula
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical, "Form_Load()"
            Exit Sub
    End Select
    
    Exit Sub
    
End Sub

Private Sub Form_Terminate()
    
    On Error Resume Next

    If gmb_DBAddRegister Then
        gb_AddRegisterStarted = False
    End If
    
    If (True = frmDataControl.Visible) And (True = mb_RegistroModificado) Then
    
        If Not gmb_DBAddRegister Then
            '-----------------------------------------------------------
            ' [very big error] aqui continuaba:     Me.Visible = False
            ' lo que VOLVIA A CARGAR un formulario que se encontraba
            ' descargado y este se descargaba solo cuando frmDataControl
            ' se descargaba => despues de algunas ediciones el sistema
            ' se quedaba sin memoria....
            '
            frmDataControl.Show
            frmDataControl.Actualizar_Lista gml_RowEditRegister
        Else
            frmDataControl.Show
            frmDataControl.Agregar_Lista
        End If
        
    End If
    
End Sub

