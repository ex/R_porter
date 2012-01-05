VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DAC1C15E-A0D0-11D8-92BC-F3955AEE4860}#3.0#0"; "exHighLightCode.ocx"
Begin VB.Form frmSQL 
   Caption         =   "Ejecutar comando SQL"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6690
   Icon            =   "frmSQL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmmdlg 
      Left            =   5985
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin exHighLightCode.exCodeHighlight rchtxtSQL 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   4445
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RightMargin     =   16500
      SelRTF          =   $"frmSQL.frx":058A
      Language        =   4
      KeywordColor    =   14448760
      OperatorColor   =   255
      CommentColor    =   7895160
      LiteralColor    =   16711935
      ForeColor       =   0
      FunctionColor   =   0
      Author          =   "Esau R.O. [exe_q_tor] ...based in the DevDomainCodeHighlight control."
      BoldKeyword     =   -1  'True
      ItalicComment   =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Abrir SQL"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Salvar como..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuDatos 
      Caption         =   "&Datos"
      Begin VB.Menu mnuEjecutar 
         Caption         =   "&Ejecutar sentencia..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstructura 
         Caption         =   "&Ver estructura de la BD"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu nuConsultas 
      Caption         =   "R_porter"
      Begin VB.Menu mnuVerEspacio 
         Caption         =   "Ver el espacio ocupado por cada medio"
      End
      Begin VB.Menu mnuVerNumAutor 
         Caption         =   "Ver el número de archivos por autor"
      End
      Begin VB.Menu mnuVerNumGenero 
         Caption         =   "Ver el número de archivos por género"
      End
      Begin VB.Menu mnuVerSinAutorGenero 
         Caption         =   "Ver los archivos sin género o autor"
      End
      Begin VB.Menu VerGeneroVacio 
         Caption         =   "Ver los generos sin archivos asignados"
      End
      Begin VB.Menu mnuVerAutorVacio 
         Caption         =   "Ver los autores sin archivos asignados"
      End
      Begin VB.Menu mnuVerFielNameAs 
         Caption         =   "Ver los nombres de archivo que contienen la palabra..."
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "mnupopup"
      Visible         =   0   'False
      Begin VB.Menu mnuExecute 
         Caption         =   "Ejecutar"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLimpiar 
         Caption         =   "Limpiar"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "UPDATE"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "DELETE"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "INSERT"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "SELECT"
         Begin VB.Menu mnuSelect1 
            Caption         =   "SELECT *"
         End
         Begin VB.Menu mnuSelect2 
            Caption         =   "SELECT FROM"
         End
      End
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_def_ID_EX = 25647893

Private Sub Form_Load()
    rchtxtSQL.ExID = m_def_ID_EX
    gsub_SetRichTabs rchtxtSQL.RichHwnd, 4     '<- no funciona!
End Sub

Public Sub SetSQL(ByRef strSQL As String)
    With rchtxtSQL
        .Text = ""
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Consulta SQL de la ultima busqueda."
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & strSQL & vbCrLf
    End With
End Sub

'**************************************************************
' POPUP MENUS
'**************************************************************
Private Sub mnuEjecutar_Click()
    '===================================================
    Dim frmResultQuery As Form
    Dim rs As ADODB.Recordset
    Dim cd As ADODB.Command
    Dim bTransaccionIniciada As Boolean
    Dim b_select As Boolean
    Dim cad As String
    Dim sql_str As String
    Dim k As Long
    Dim chr As Integer
    Dim fld_tipo As ADODB.DataTypeEnum
    Dim n_fields As Long
    Dim b_Redraw As Boolean
    Dim n_state As Integer
    '===================================================
    
    On Error GoTo Handler
    
    '------------------------------------------------------
    ' ejecutar texto seleccionado o sino todo el texto
    If rchtxtSQL.SelLength > 0 Then
        sql_str = Trim(rchtxtSQL.SelText)
    Else
        sql_str = Trim(rchtxtSQL.Text)
    End If
    
    n_state = 0
    
    '------------------------------------------------------
    ' eliminar tabs y saltos de linea de la cadena sql
    ' ignorar las lineas de comentario --
    cad = ""
    For k = 1 To Len(sql_str)
    
        chr = Asc(Mid$(sql_str, k, 1))
        
        If (chr = 9) Then
            cad = cad & " "
        Else
            If (chr = 13) Or (chr = 10) Then
                cad = cad & " "
                n_state = 0
            Else
                If (chr = 45) Then
                    'se encontro caracter '-'
                    If n_state = 0 Then
                        n_state = 1
                    Else
                        If n_state = 1 Then
                            n_state = 2
                            ' eliminar primer signo de comentario
                            cad = Left$(cad, Len(cad) - 1)
                        End If
                    End If
                Else
                    If n_state = 1 Then
                        n_state = 0
                    End If
                End If
                
                If n_state <> 2 Then
                    cad = cad & Mid$(sql_str, k, 1)
                End If
                
            End If
        End If
        
    Next k
        
    sql_str = cad
    cad = UCase(Left(Trim(sql_str), 6))
    
    Select Case cad
    
        Case "SELECT"
            b_select = True
            
        Case "UPDATE"
            b_select = False
            
        Case "INSERT"
            b_select = False
            
        Case "DELETE"
            b_select = False
            
        Case Else
            MsgBox "La sintaxis parece ser incorrecta," & vbCrLf & "verifica el comando SQL por favor...", vbInformation, "Verifica SQL"
            rchtxtSQL.SetFocus
            Exit Sub
            
    End Select
    
    If vbNo = MsgBox("¿Estás seguro de continuar?" & vbCrLf & "(verifica el comando SQL a ejecutar)", vbExclamation + vbYesNo, "Advertencia") Then
        Exit Sub
    End If
        
    'iniciamos transaccion para ejecutar el comando
    If True = gfnc_CrearConexionTransaccion(gs_DSN, gs_Pwd) Then

        cnTransaction.BeginTrans
        bTransaccionIniciada = True

        Set cd = New ADODB.Command
        Set cd.ActiveConnection = cnTransaction

        cd.CommandText = sql_str
        
        Screen.MousePointer = vbHourglass
        
        If b_select Then
            
            Set rs = cd.Execute
            
            n_fields = rs.Fields.Count
            
            If n_fields > 0 Then
            
                '---------------------------------------
                ' crear nuevo formulario de resultado
                '
                Set frmResultQuery = New frmSelect
                frmResultQuery.Show vbModeless
                frmResultQuery.sql_str = sql_str
            
                With frmResultQuery.flxResults

                    .MousePointer = flexHourglass
                    
                    .Redraw = False
                    
                    .Cols = n_fields + 1
                    .Row = 0
                    
                    .Col = 0
                    .ColAlignment(0) = flexAlignRightCenter
                    .ColWidth(0) = 540
                    .Text = "Nº"
                    
                    '---------------------------------------
                    ' poner cabecera
                    '
                    For k = 1 To n_fields
                        .Col = k
                        .CellAlignment = flexAlignLeftCenter
                        .Text = rs.Fields(k - 1).Name
                        
                        '---------------------------------------
                        ' alineacion de columnas
                        '
                        fld_tipo = rs.Fields(k - 1).Type

                        Select Case fld_tipo
                            Case adVarWChar, adLongVarChar, adChar, adWChar, adBSTR, adDate, adDBDate, adLongVarWChar, adVarChar, adDBTimeStamp
                                .ColAlignment(k) = flexAlignLeftCenter
                                .ColWidth(k) = 1200
                            Case Else
                                .ColAlignment(k) = flexAlignRightCenter
                                .ColWidth(k) = 690
                        End Select
                        
                    Next k
                    
                    b_Redraw = True
                    
                    If rs.EOF = True Then
                        .RowHeight(1) = 0
                        GoTo SALIR
                    End If
                    
                    '---------------------------------------
                    ' llenar datos
                    '
                    While rs.EOF = False

                        .Row = .Rows - 1
                        'forzar visible
                        .RowHeight(.Row) = -1
                        
                        .Col = 0
                        .CellForeColor = RGB(100, 170, 255)
                        .Text = .Row

                        For k = 0 To (n_fields - 1)
                            .Col = k + 1
                            .Text = rs.Fields(k).value
                        Next k

                        rs.MoveNext

                        .Rows = .Rows + 1

                        If b_Redraw Then
                            If .Row >= CInt(((.height - .RowHeight(0)) / .RowHeight(.Row))) + 1 Then
                                .Redraw = True
                                .Refresh
                                .Redraw = False
                                b_Redraw = False
                            End If
                        End If

                    Wend

                    ' eliminar la ultima fila agregada que esta vacia
                    .Rows = .Rows - 1
SALIR:
                    .Row = 1
                    .Col = 0
                    .ColSel = n_fields
                    
                    .Redraw = True
                    .SetFocus
                    
                    .MousePointer = flexDefault
                    
                End With
                
                frmResultQuery.ZOrder 0
                frmResultQuery.Refresh
            
            End If
                
            rs.Close
            
        Else
            cd.Execute
        End If
    
        cnTransaction.CommitTrans
        bTransaccionIniciada = False
        
        gsub_CerrarConexionTransaccion
        
        Screen.MousePointer = vbDefault
        Me.MousePointer = vbDefault
        
        MsgBox "Comando SQL ejecutado exitosamente", vbExclamation, "Finalizado"
        
    Else
        MsgBox "No se pudo crear la transaccion para la BD", vbExclamation, "Error de conexion"
        Exit Sub
    End If
    
    Exit Sub
    
Handler:

    Select Case Err.Number
    
        Case 94, 13
            'uso no valido de NULL
            'cuando el campo esta vacio
            Resume Next
        
        Case Else
            
            MsgBox Err.Description, vbExclamation, "Error ejecutando comando"
            
            If bTransaccionIniciada Then
                cnTransaction.RollbackTrans
                gsub_CerrarConexionTransaccion
            End If
            
            Screen.MousePointer = vbDefault
            Me.MousePointer = vbDefault
        
    End Select

End Sub

Private Sub mnuEstructura_Click()
    
    If Not gb_DBConexionOK Then
        gsub_ShowMessageNoConection
        Unload Me
    Else
        frmReporte.Show
    End If

End Sub

Private Sub mnuExecute_Click()
    mnuEjecutar_Click
End Sub

Private Sub mnuLimpiar_Click()
    rchtxtSQL.Text = ""
End Sub

Private Sub mnuDelete_Click()
    
    With rchtxtSQL
        .SelText = "DELETE FROM nombre_tabla"
        .SelText = vbCrLf & "WHERE campo_busqueda=criterio_busqueda"
        .SelText = vbCrLf
    End With
    
End Sub

Private Sub mnuInsert_Click()
    
    With rchtxtSQL
        .SelText = "INSERT INTO  nombre_tabla"
        .SelText = vbCrLf & vbTab & "(campo_numerico, campo_cadena, campo_fecha)"
        .SelText = vbCrLf & "VALUES"
        .SelText = vbCrLf & vbTab & "(numero, 'cadena', #fecha#)"
        .SelText = vbCrLf
    End With
    
End Sub

Private Sub mnuOpen_Click()
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'esconde casilla de solo lectura y verifica que el archivo y el path existan
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .DialogTitle = "Indicar el archivo SQL a abrir:"
        .Filter = "Archivo SQL (*.sql)|*.sql|Todos los Archivos(*.*)|*.*"
        .InitDir = App.Path & "\sql"    ' de no existir el directorio usara el directorio activo
        'tipo predefinido VBS
        .FilterIndex = 1
        .ShowOpen
        If .filename <> "" Then
            'cargar el archivo SQL
            rchtxtSQL.LoadFile .filename
            gsub_SetRichTabs rchtxtSQL.RichHwnd, 4     '<- no funciona!
        End If
    End With
    
    Exit Sub
    
ErrorCancel:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Error Inesperado" & Err.Description, vbOKOnly + vbCritical, "mnuOpen_Click"
    End If
End Sub

Private Sub mnuSaveAs_Click()
    
    On Error GoTo ErrorCancel
    
    With cmmdlg
        'Para interceptar cuando se elige [cancelar]:
        .CancelError = True
        'avisa en caso de sobreescritura, esconde casilla solo lectura y verifica path
        .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .DialogTitle = "Exportar los resultados como:"
        .Filter = "Archivos SQL (*.sql)|*.sql|Todos los Archivos(*.*)|*.*"
        'necesario para controlar la extension con que se salvaran los archivos
        'sino si el usuario selecciona la opcion de ver todos los archivos sucede un error
        .DefaultExt = ""
        .InitDir = App.Path & "\sql"    ' de no existir el directorio usara el directorio activo
        'tipo predefinido VBS
        .FilterIndex = 1
        'nombre del archivo por defecto
        .filename = "new.sql"
        .ShowSave
        If .filename <> "" Then
            '---------------------------------------------
            ' salvar script
            rchtxtSQL.SaveFile .filename, rtfText
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

Private Sub mnuSelect1_Click()
    
    With rchtxtSQL
        .SelText = "SELECT * FROM tabla_1 "
        .SelText = vbCrLf & "ORDER BY campo_1"
        .SelText = vbCrLf
    End With

End Sub

Private Sub mnuSelect2_Click()
    
    With rchtxtSQL
        .SelText = "SELECT  tabla_1.campo_1, tabla_2.campo_2 "
        .SelText = vbCrLf & "FROM tabla_1, tabla_2 "
        .SelText = vbCrLf & "WHERE (tabla_1.campo_1=criterio_1) "
        .SelText = vbCrLf & "ORDER BY tabla_2.campo_2 DESC, tabla_1.campo_1"
        .SelText = vbCrLf
    End With

End Sub

Private Sub mnuUpdate_Click()
    With rchtxtSQL
        .SelText = "UPDATE nombre_tabla "
        .SelText = vbCrLf & "SET"
        .SelText = vbCrLf & vbTab & "campo_numerico=numero,"
        .SelText = vbCrLf & vbTab & "campo_cadena='cadena'," & vbCrLf
        .SelText = vbCrLf & vbTab & "campo_fecha=#fecha#"
        .SelText = vbCrLf & "WHERE"
        .SelText = vbCrLf & vbTab & "campo_busqueda=criterio_busqueda"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuVerEspacio_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los medios de "
        .SelText = vbCrLf & "-- almacenamiento y el espacio que ocupan (la suma del"
        .SelText = vbCrLf & "-- tamaño de todos los archivos de dicho medio en MB)."
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT" & vbTab & "storage.name AS Medio,"
        .SelText = vbCrLf & vbTab & vbTab & "INT(SUM(file.sys_length)/1048576) "
        .SelText = "AS [total MB]"
        .SelText = vbCrLf & "FROM" & vbTab & "file"
        .SelText = vbCrLf & "INNER JOIN storage ON (file.id_storage=storage.id_storage)"
        .SelText = vbCrLf & "GROUP BY storage.name, file.id_storage"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuVerAutorVacio_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los autores que"
        .SelText = vbCrLf & "-- no tienen asociado ningún archivo (están vacios)"
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT" & vbTab & "author.author "
        .SelText = "AS autor, "
        .SelText = vbCrLf & vbTab & vbTab & "COUNT(file.id_author) "
        .SelText = "AS total"
        .SelText = vbCrLf & "FROM file"
        .SelText = vbCrLf & "RIGHT JOIN author ON (file.id_author=author.id_author)"
        .SelText = vbCrLf & "GROUP BY author.author, file.id_author"
        .SelText = vbCrLf & "HAVING COUNT(file.id_author)=0"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuVerFielNameAs_Click()

    Dim sPalabra As String
    
    On Error Resume Next
    
    sPalabra = InputBox("Ingresa la palabra que quieres buscar:", "Comando SQL", "")
    
    If (Trim(sPalabra) <> "") Then
        With rchtxtSQL
            .SelText = "------------------------------------------------------"
            .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los registros"
            .SelText = vbCrLf & "-- con nombre de archivo que contienen la palabra:"
            .SelText = vbCrLf & "-- " & UCase$(sPalabra)
            .SelText = vbCrLf & "------------------------------------------------------"
            .SelText = vbCrLf & "SELECT * FROM file"
            .SelText = vbCrLf & "WHERE (sys_name LIKE '%" & LCase$(sPalabra) & "%')"
            .SelText = vbCrLf & "ORDER BY sys_name"
            .SelText = vbCrLf
        End With
    End If

End Sub

Private Sub mnuVerNumAutor_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de todos los autores"
        .SelText = vbCrLf & "-- y del numero de archivos que tienen asociados."
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT" & vbTab & "author.author "
        .SelText = "AS autor, "
        .SelText = vbCrLf & vbTab & vbTab & "COUNT(file.id_author) AS total"
        .SelText = vbCrLf & "FROM" & vbTab & "file"
        .SelText = vbCrLf & "RIGHT JOIN author ON (file.id_author=author.id_author)"
        .SelText = vbCrLf & "GROUP BY author.author, file.id_author"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuVerNumGenero_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los generos y del "
        .SelText = vbCrLf & "-- numero de archivos que tienen asociados."
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT" & vbTab & "genre.genre AS genero, "
        .SelText = vbCrLf & vbTab & vbTab & "COUNT(file.id_genre) AS total"
        .SelText = vbCrLf & "FROM" & vbTab & "file"
        .SelText = vbCrLf & "INNER JOIN genre ON (file.id_genre=genre.id_genre)"
        .SelText = vbCrLf & "GROUP BY genre.genre, file.id_genre"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuVerSinAutorGenero_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los archivos que"
        .SelText = vbCrLf & "-- no tienen asociado ningún género o autor"
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT" & vbTab & "id_file AS id, "
        .SelText = vbCrLf & vbTab & vbTab & "name AS archivo"
        .SelText = vbCrLf & "FROM" & vbTab & "file"
        .SelText = vbCrLf & "WHERE ((file.id_genre=0) OR (file.id_author=0))"
        .SelText = vbCrLf
    End With
End Sub

Private Sub VerGeneroVacio_Click()
    With rchtxtSQL
        .SelText = "------------------------------------------------------"
        .SelText = vbCrLf & "-- Este ejemplo muestra una lista de los generos que"
        .SelText = vbCrLf & "-- no tienen asociado ningún archivo (están vacios)"
        .SelText = vbCrLf & "------------------------------------------------------"
        .SelText = vbCrLf & "SELECT  genre.genre AS GENERO, genre.id_genre,"
        .SelText = vbCrLf & "COUNT(file.id_genre) AS TOTAL"
        .SelText = vbCrLf & "FROM file"
        .SelText = vbCrLf & "RIGHT JOIN genre ON (file.id_genre=genre.id_genre)"
        .SelText = vbCrLf & "GROUP BY genre.genre, genre.id_genre, file.id_genre"
        .SelText = vbCrLf & "HAVING COUNT(file.id_genre)=0"
        .SelText = vbCrLf
    End With
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

'**************************************************************
' RICHTEXTBOX
'**************************************************************
Private Sub rchtxtSQL_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 93 Then
        PopupMenu mnupopup, 2
    End If

End Sub

Private Sub rchtxtSQL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'el control richtext no refresca bien
    Me.ZOrder 0
    Me.Refresh
    
    If Button = vbRightButton Then
        PopupMenu Me.mnupopup, 2
    End If

End Sub

'**************************************************************
' FORM
'**************************************************************
Private Sub Form_Resize()

    On Error Resume Next

    If (Me.width < 2655) Then
        Me.width = 2655
    End If
    
    If (Me.height < 2100) Then
        Me.height = 2100
    End If
    
    rchtxtSQL.width = Me.ScaleWidth - 15
    rchtxtSQL.height = Me.ScaleHeight - 30
    
End Sub

