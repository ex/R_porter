VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "frmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   315
      Left            =   304
      TabIndex        =   13
      Top             =   5535
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2876
      TabIndex        =   0
      Top             =   5535
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   5535
      Width           =   1215
   End
   Begin VB.Frame fraTablas 
      Caption         =   "Tablas"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   4350
      Begin VB.ListBox lstTablas 
         Height          =   1425
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   4200
      End
   End
   Begin VB.Frame fraCampos 
      Caption         =   "Campos"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   3255
      Width           =   4350
      Begin MSComctlLib.ListView lstvwCampos 
         Height          =   1815
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraDSN 
      Height          =   1425
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4350
      Begin VB.TextBox txtPwd 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1245
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   645
         Width           =   1575
      End
      Begin VB.CommandButton cmdConectar 
         Caption         =   "&Conectar"
         Height          =   315
         Left            =   2865
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cmbDSN 
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   225
         Width           =   2865
      End
      Begin VB.Label lbl 
         Caption         =   "ms_ex_DSN"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   930
      End
      Begin VB.Label lbl 
         Caption         =   "ms_ex_PWD"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblStatusConexion 
         Caption         =   "Conexión no establecida"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1095
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mItem As ListItem

Dim m_TipeOK As Boolean

'*******************************************************************************
' TXTBOXS
'*******************************************************************************
Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdConectar.Value = True
    End If
End Sub

'*******************************************************************************
' LISTBOX
'*******************************************************************************
Private Sub lstTablas_Click()
Dim nType As ADODB.DataTypeEnum
Dim rs As ADODB.Recordset
Dim flds As ADODB.Fields
Dim k As Integer


    On Error GoTo Handler

    Me.lstvwCampos.ListItems.Clear
    
    gs_Form = lstTablas.Text
    
    'este metodo funciona aun cuando la tabla este vacia : (rs.EOF = True)
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM " & gs_Form, cn, adOpenForwardOnly, adLockReadOnly

    Set flds = rs.Fields

    For k = 0 To (flds.Count - 1)
       
        Set mItem = lstvwCampos.ListItems.Add()
        mItem.Text = flds(k).Name
        nType = flds(k).Type
        
        m_TipeOK = True
        
        Select Case nType
        
            Case ADODB.DataTypeEnum.adArray
                mItem.SubItems(1) = "adArray"
                
            Case ADODB.DataTypeEnum.adBigInt
                mItem.SubItems(1) = "adBigInt"
                
            Case ADODB.DataTypeEnum.adBinary
                mItem.SubItems(1) = "adBinary"
                
            Case ADODB.DataTypeEnum.adBoolean
                mItem.SubItems(1) = "adBoolean"
                
            Case ADODB.DataTypeEnum.adBSTR
                mItem.SubItems(1) = "adBSTR"
                
            Case ADODB.DataTypeEnum.adChapter
                mItem.SubItems(1) = "adChapter"
                
            Case ADODB.DataTypeEnum.adChar
                mItem.SubItems(1) = "adChar"
                
            Case ADODB.DataTypeEnum.adCurrency
                mItem.SubItems(1) = "adCurrency"
                
            Case ADODB.DataTypeEnum.adDate
                mItem.SubItems(1) = "adDate"
                
            Case ADODB.DataTypeEnum.adDBDate
                mItem.SubItems(1) = "adDBDate"
                
            Case ADODB.DataTypeEnum.adDBTime
                mItem.SubItems(1) = "adDBTime"
                
            Case ADODB.DataTypeEnum.adDBTimeStamp
                mItem.SubItems(1) = "adDBTimeStamp"
                
            Case ADODB.DataTypeEnum.adDecimal
                mItem.SubItems(1) = "adDecimal"
                
            Case ADODB.DataTypeEnum.adDouble
                mItem.SubItems(1) = "adDouble"
                
            Case ADODB.DataTypeEnum.adEmpty
                mItem.SubItems(1) = "adEmpty"
                
            Case ADODB.DataTypeEnum.adError
                mItem.SubItems(1) = "adError"
                
            Case ADODB.DataTypeEnum.adFileTime
                mItem.SubItems(1) = "adFileTime"
                
            Case ADODB.DataTypeEnum.adGUID
                mItem.SubItems(1) = "adGUID"
                
            Case ADODB.DataTypeEnum.adIDispatch
                mItem.SubItems(1) = "adIDispatch"
                
            Case ADODB.DataTypeEnum.adInteger
                mItem.SubItems(1) = "adInteger"
                
            Case ADODB.DataTypeEnum.adInteger
                mItem.SubItems(1) = "adInteger"
                
            Case ADODB.DataTypeEnum.adIUnknown
                mItem.SubItems(1) = "adIUnknown"
                
            Case ADODB.DataTypeEnum.adLongVarBinary
                mItem.SubItems(1) = "adLongVarBinary"
                
            Case ADODB.DataTypeEnum.adLongVarChar
                mItem.SubItems(1) = "adLongVarChar"
                
            Case ADODB.DataTypeEnum.adLongVarWChar
                mItem.SubItems(1) = "adLongVarWChar"
                
            Case ADODB.DataTypeEnum.adNumeric
                mItem.SubItems(1) = "adNumeric"
                
            Case ADODB.DataTypeEnum.adPropVariant
                mItem.SubItems(1) = "adPropVariant"
                
            Case ADODB.DataTypeEnum.adSingle
                mItem.SubItems(1) = "adSingle"
                
            Case ADODB.DataTypeEnum.adSmallInt
                mItem.SubItems(1) = "adSmallInt"
                
            Case ADODB.DataTypeEnum.adTinyInt
                mItem.SubItems(1) = "adTinyInt"
                
            Case ADODB.DataTypeEnum.adUnsignedBigInt
                mItem.SubItems(1) = "adUnsignedBigInt"
                
            Case ADODB.DataTypeEnum.adUnsignedInt
                mItem.SubItems(1) = "adUnsignedInt"
                
            Case ADODB.DataTypeEnum.adUnsignedSmallInt
                mItem.SubItems(1) = "adUnsignedSmallInt"
                
            Case ADODB.DataTypeEnum.adUnsignedTinyInt
                mItem.SubItems(1) = "adUnsignedTinyInt"
                
            Case ADODB.DataTypeEnum.adUserDefined
                mItem.SubItems(1) = "adUserDefined"
                
            Case ADODB.DataTypeEnum.adVarBinary
                mItem.SubItems(1) = "adVarBinary"
                
            Case ADODB.DataTypeEnum.adVarChar
                mItem.SubItems(1) = "adVarChar"
                
            Case ADODB.DataTypeEnum.adVariant
                mItem.SubItems(1) = "adVariant"
                
            Case ADODB.DataTypeEnum.adVarNumeric
                mItem.SubItems(1) = "adVarNumeric"
                
            Case ADODB.DataTypeEnum.adVarWChar
                mItem.SubItems(1) = "adVarWChar"
                
            Case ADODB.DataTypeEnum.adWChar
                mItem.SubItems(1) = "adWChar"
                
            Case Default
                MsgBox "Campo con tipo de dato no reconocido: " & flds(k).Name, vbExclamation, "Error"
                m_TipeOK = False
            
        End Select
        
        gdoc_usrdoc.AddTabla "Tabla: [" & UCase(gs_Form) & "]"
        
        mItem.SubItems(2) = flds(k).DefinedSize
       
    Next

    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "lstTablas_Click()"
    m_TipeOK = False
End Sub

'*******************************************************************************
' COMMAND BUTTONS
'*******************************************************************************
Private Sub cmdReporte_Click()
Dim nType As ADODB.DataTypeEnum
Dim rs As ADODB.Recordset
Dim flds As ADODB.Fields
Dim sType As String
Dim bIsText As Boolean
Dim k As Integer
Dim q As Integer

    
    On Error GoTo Handler
    
    Load frmReporte
    
    With frmReporte.rchtxt
    
        .Font = "Lucida Console"
        
        For k = 0 To Me.lstTablas.ListCount - 1
        
            .SelText = "**************************************************" & vbCrLf
            .SelText = "  " & Me.lstTablas.list(k) & vbCrLf
            .SelText = "  ------------------------------------------------" & vbCrLf
        
            'este metodo funciona aun cuando la tabla este vacia : (rs.EOF = True)
            Set rs = New ADODB.Recordset
            rs.Open "SELECT * FROM " & Me.lstTablas.list(k), cn, adOpenForwardOnly, adLockReadOnly
        
            Set flds = rs.Fields
        
            For q = 0 To (flds.Count - 1)
               
                .SelText = "    " & flds(q).Name & Space(20 - Len(flds(q).Name))

                nType = flds(q).Type
                
                m_TipeOK = True
                
                bIsText = False
                
                Select Case nType
                
                    Case ADODB.DataTypeEnum.adArray
                        sType = "adArray"
                        
                    Case ADODB.DataTypeEnum.adBigInt
                        sType = "adBigInt"
                        
                    Case ADODB.DataTypeEnum.adBinary
                        sType = "adBinary"
                        
                    Case ADODB.DataTypeEnum.adBoolean
                        sType = "adBoolean"
                        
                    Case ADODB.DataTypeEnum.adBSTR
                        sType = "adBSTR"
                        
                    Case ADODB.DataTypeEnum.adChapter
                        sType = "adChapter"
                        
                    Case ADODB.DataTypeEnum.adChar
                        sType = "adChar"
                        
                    Case ADODB.DataTypeEnum.adCurrency
                        sType = "adCurrency"
                        
                    Case ADODB.DataTypeEnum.adDate
                        sType = "adDate"
                        
                    Case ADODB.DataTypeEnum.adDBDate
                        sType = "adDBDate"
                        
                    Case ADODB.DataTypeEnum.adDBTime
                        sType = "adDBTime"
                        
                    Case ADODB.DataTypeEnum.adDBTimeStamp
                        sType = "adDBTimeStamp"
                        
                    Case ADODB.DataTypeEnum.adDecimal
                        sType = "adDecimal"
                        
                    Case ADODB.DataTypeEnum.adDouble
                        sType = "adDouble"
                        
                    Case ADODB.DataTypeEnum.adEmpty
                        sType = "adEmpty"
                        
                    Case ADODB.DataTypeEnum.adError
                        sType = "adError"
                        
                    Case ADODB.DataTypeEnum.adFileTime
                        sType = "adFileTime"
                        
                    Case ADODB.DataTypeEnum.adGUID
                        sType = "adGUID"
                        
                    Case ADODB.DataTypeEnum.adIDispatch
                        sType = "adIDispatch"
                        
                    Case ADODB.DataTypeEnum.adInteger
                        sType = "adInteger"
                        
                    Case ADODB.DataTypeEnum.adInteger
                        sType = "adInteger"
                        
                    Case ADODB.DataTypeEnum.adIUnknown
                        sType = "adIUnknown"
                        
                    Case ADODB.DataTypeEnum.adLongVarBinary
                        sType = "adLongVarBinary"
                        
                    Case ADODB.DataTypeEnum.adLongVarChar
                        sType = "adLongVarChar"
                        
                    Case ADODB.DataTypeEnum.adLongVarWChar
                        sType = "adLongVarWChar"
                        
                    Case ADODB.DataTypeEnum.adNumeric
                        sType = "adNumeric"
                        
                    Case ADODB.DataTypeEnum.adPropVariant
                        sType = "adPropVariant"
                        
                    Case ADODB.DataTypeEnum.adSingle
                        sType = "adSingle"
                        
                    Case ADODB.DataTypeEnum.adSmallInt
                        sType = "adSmallInt"
                        
                    Case ADODB.DataTypeEnum.adTinyInt
                        sType = "adTinyInt"
                        
                    Case ADODB.DataTypeEnum.adUnsignedBigInt
                        sType = "adUnsignedBigInt"
                        
                    Case ADODB.DataTypeEnum.adUnsignedInt
                        sType = "adUnsignedInt"
                        
                    Case ADODB.DataTypeEnum.adUnsignedSmallInt
                        sType = "adUnsignedSmallInt"
                        
                    Case ADODB.DataTypeEnum.adUnsignedTinyInt
                        sType = "adUnsignedTinyInt"
                        
                    Case ADODB.DataTypeEnum.adUserDefined
                        sType = "adUserDefined"
                        
                    Case ADODB.DataTypeEnum.adVarBinary
                        sType = "adVarBinary"
                        
                    Case ADODB.DataTypeEnum.adVarChar
                        sType = "adVarChar"
                        
                    Case ADODB.DataTypeEnum.adVariant
                        sType = "adVariant"
                        
                    Case ADODB.DataTypeEnum.adVarNumeric
                        sType = "adVarNumeric"
                        
                    Case ADODB.DataTypeEnum.adVarWChar
                        sType = "adVarWChar"
                        bIsText = True
                        
                    Case ADODB.DataTypeEnum.adWChar
                        sType = "adWChar"
                        
                End Select
                
                If bIsText = True Then
                    .SelText = sType & Space(19 - Len(sType))
                    .SelText = "[" & flds(q).DefinedSize & "]" & vbCrLf
                Else
                    .SelText = sType & Space(20 - Len(sType))
                    .SelText = flds(q).DefinedSize & vbCrLf
                End If
                
               
            Next q
            
        Next k

    End With
    
    frmReporte.Show
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "cmdReporte_Click()"
End Sub

Private Sub cmdAceptar_Click()
    Insertar_Formulario
End Sub

Private Sub cmdConectar_Click()
Dim rsSchema As ADODB.Recordset
Dim fld As ADODB.Field
Dim rCriteria As Variant
Dim list As ListBox
    
    On Error GoTo Handler
    
    If gb_DBConexionOK = True Then
        gp_CerrarConexionBaseDatos
    End If
    
    Me.lstTablas.Clear
    Me.lstvwCampos.ListItems.Clear
    gs_Form = ""
    
    If False = gf_CrearConexion(Trim(Me.cmbDSN), Trim(Me.txtPwd)) Then
        MsgBox "No se pudo entablar conexión con la base de datos" & vbCrLf & "Verifique el DSN y la contraseña. También es posible " & vbCrLf & "que la BD ya se encuentre abierta en modo exclusivo" & vbCrLf & "o que tenga atributo de sólo lectura.", vbExclamation, "Error en conexion"
        Me.lblStatusConexion.Caption = "Falla en la Conexión."
        gdoc_usrdoc.ClearList
        gb_DBConexionOK = False
    Else
        Me.lblStatusConexion.Caption = "Conexión exitosa."
        gb_DBConexionOK = True
        gs_ex_DSN = Trim(Me.cmbDSN)
        gs_ex_PWD = Trim(Me.txtPwd)
        
        gdoc_usrdoc.ClearList
        gdoc_usrdoc.AddString "Conexión: [" & UCase(gs_ex_DSN) & "]", True
        
        '----------------------------------------------
        'listar tablas de la base de datos
        '----------------------------------------------
        rCriteria = Array(Empty, Empty, Empty, "Table")
        Set rsSchema = cn.OpenSchema(adSchemaTables, rCriteria)
        
        While Not rsSchema.EOF
        
           For Each fld In rsSchema.Fields
              If fld.Name = "TABLE_NAME" Then
                Me.lstTablas.AddItem fld.Value
              End If
           Next
           
           rsSchema.MoveNext
        Wend
        
    End If
    
    Exit Sub
    
Handler:

    MsgBox Err.Description, vbCritical, "cmdConectar_Click()"
    gb_DBConexionOK = False
    Me.lblStatusConexion.Caption = "Falla en la Conexión."

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'*******************************************************************************
' FORMULARIO
'*******************************************************************************
Private Sub Form_Load()
    
    On Error Resume Next
    
    GetDSNsAndDrivers
    Me.cmbDSN.ListIndex = 0

    ' Borra la colección ColumnHeaders.
    lstvwCampos.ColumnHeaders.Clear
    ' Agrega cuatro objetos ColumnHeader.
    lstvwCampos.ColumnHeaders.Add , , "Nombre", 1410
    lstvwCampos.ColumnHeaders.Add , , "Tipo", 1650
    lstvwCampos.ColumnHeaders.Add , , "Tamaño", 795

End Sub

Private Sub Form_Unload(Cancel As Integer)
    gp_CerrarConexionBaseDatos
    gdoc_usrdoc.ClearList
End Sub

'*******************************************************************************
' FUNCIONES INTERNAS
'*******************************************************************************
Private Sub GetDSNsAndDrivers()
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String
Dim sDRV As String
Dim iDSNLen As Integer
Dim iDRVLen As Integer
Dim lHenv As Long       'controlador del entorno

    On Error Resume Next

    'obtener los DSN
    If SQLAllocEnv(lHenv) <> -1 Then
        
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = VBA.Left(sDSNItem, iDSNLen)
            sDRV = VBA.Left(sDRVItem, iDRVLen)
            
            If sDSN <> Space(iDSNLen) Then
                Me.cmbDSN.AddItem sDSN
            End If
        Loop

    End If
    
    'quitar los duplicados
    If Me.cmbDSN.ListCount > 0 Then
        With Me.cmbDSN
            If .ListCount > 1 Then
                i = 0
                While i < .ListCount
                    If .list(i) = .list(i + 1) Then
                        .RemoveItem (i)
                    Else
                        i = i + 1
                    End If
                Wend
            End If
        End With
    End If
End Sub

Private Sub Insertar_Formulario()
    
    If lstTablas.Text <> "" Then
        gs_Form = "frm" & UCase(Left(lstTablas.Text, 1)) & LCase(Mid(lstTablas.Text, 2))
        frmAddForm.Show vbModal
    End If
    
End Sub
