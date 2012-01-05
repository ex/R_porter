VERSION 5.00
Begin VB.Form frmConexion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmConexion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4035
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDSN 
      Height          =   1545
      Left            =   90
      TabIndex        =   6
      Top             =   15
      Width           =   3855
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   2505
         TabIndex        =   5
         Top             =   1065
         Width           =   1215
      End
      Begin VB.CommandButton cmdConectar 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtPwd 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   780
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   645
         Width           =   1680
      End
      Begin VB.ComboBox cmbDSN 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label lblStatusConexion 
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   1170
         Width           =   2295
      End
      Begin VB.Label lblPwd 
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   742
         Width           =   540
      End
      Begin VB.Label lblDSN 
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConectar_Click()
    
    On Error GoTo Handler
    If gb_DBConexionOK = True Then
        gsub_CerrarConexionBaseDatos
    End If
    
    If gfnc_CrearConexion(Trim(cmbDSN.text), Trim(txtPwd.text)) Then
        lblStatusConexion.Caption = FRM_CONEXION_1
        gb_DBConexionOK = True
        
        If gfnc_ValidateDB Then
            lblStatusConexion.Caption = FRM_CONEXION_2
            gb_DBFormatOK = True
        Else
            gsub_ShowMessageWrongDB
            gb_DBFormatOK = False
        End If
    Else
        lblStatusConexion.Caption = FRM_CONEXION_3
        gsub_ShowMessageFailedConection
        gb_DBConexionOK = False
        gb_DBFormatOK = False
    End If
    
    gs_DSN = Trim(cmbDSN.text)
    gs_Pwd = Trim(txtPwd.text)
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbExclamation, FRM_CONEXION_4
    gb_DBConexionOK = False
    gb_DBFormatOK = False
    lblStatusConexion.Caption = FRM_CONEXION_3
End Sub

Private Sub Form_Load()

    On Error GoTo Handler
    
    ' GUI Localization
    Me.Caption = FRM_CONEXION_GUI_1
    cmdCancelar.Caption = FRM_CONEXION_GUI_2
    cmdConectar.Caption = FRM_CONEXION_GUI_3
    lblStatusConexion.Caption = FRM_CONEXION_GUI_4
    lblPwd.Caption = FRM_CONEXION_GUI_5
    lblDSN.Caption = FRM_CONEXION_GUI_6
    
    ' Por si no esta iniciada la conexion...
    If Not gb_DBConexionOK Then
        MsgBox FRM_CONEXION_6, vbInformation, FRM_CONEXION_7
    End If

    GetDSNsAndDrivers

    cmbDSN.text = gs_DSN
    txtPwd.text = gs_Pwd
    
    If gb_DBConexionOK = True Then
        If gb_DBFormatOK Then
            lblStatusConexion.Caption = FRM_CONEXION_2
        Else
            lblStatusConexion.Caption = FRM_CONEXION_1
        End If
    Else
        lblStatusConexion.Caption = FRM_CONEXION_3
    End If
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "Form_Load"
   
End Sub

Private Sub cmbDSN_Change()
    lblStatusConexion.Caption = FRM_CONEXION_8
End Sub

Private Sub txtPwd_Change()
    lblStatusConexion.Caption = FRM_CONEXION_8
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdConectar.value = True
    End If
End Sub

Sub GetDSNsAndDrivers()
    Dim n As Integer
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
        
        sDSNItem = Space(1024)
        sDRVItem = Space(1024)
        n = SQLDataSources(lHenv, SQL_FETCH_FIRST_SYSTEM, sDSNItem, 1024, iDSNLen, _
                           sDRVItem, 1024, iDRVLen)
        
        If n = SQL_SUCCESS Then
            sDSN = VBA.Left(sDSNItem, iDSNLen)
            sDRV = VBA.Left(sDRVItem, iDRVLen)
            If sDSN <> Space(iDSNLen) Then
                cmbDSN.AddItem sDSN
            End If
            
            Do Until n <> SQL_SUCCESS
                sDSNItem = Space(1024)
                sDRVItem = Space(1024)
                n = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, _
                                   sDRVItem, 1024, iDRVLen)
                sDSN = VBA.Left(sDSNItem, iDSNLen)
                sDRV = VBA.Left(sDRVItem, iDRVLen)
                
                If sDSN <> Space(iDSNLen) Then
                    cmbDSN.AddItem sDSN
                End If
            Loop
        End If
    End If
End Sub

