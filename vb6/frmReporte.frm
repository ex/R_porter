VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReporte 
   Caption         =   "Reporte"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   Icon            =   "frmReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   -30
      Width           =   6810
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   315
         Left            =   5550
         TabIndex        =   1
         Top             =   165
         Width           =   1185
      End
      Begin VB.ComboBox cmbTablas 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   2730
      End
   End
   Begin RichTextLib.RichTextBox rchtxt 
      Height          =   4005
      Left            =   0
      TabIndex        =   2
      Top             =   570
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7064
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   15000
      TextRTF         =   $"frmReporte.frx":0442
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_lenTableName As Integer

'**************************************************************
' COMMANDBUTTONS
'**************************************************************
Private Sub cmbTablas_Click()
    rchtxt.find cmbTablas.text, 0, , rtfWholeWord Or rtfMatchCase
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'**************************************************************
' FORM
'**************************************************************
Private Sub Form_Load()

    Dim rsSchema As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim rCriteria As Variant
    Dim rs As ADODB.Recordset
    Dim flds As ADODB.Fields
    Dim sType As String
    Dim bIsText As Boolean
    Dim k As Integer
    Dim q As Integer
    Dim bIsDate As Boolean
    Dim lenTableName As Integer
    Dim arrayIndex() As Integer
    Dim lenField As Long
    On Error GoTo Handler
    
    '----------------------------------------------
    ' listar tablas de la base de datos
    rCriteria = Array(Empty, Empty, Empty, "Table")
    Set rsSchema = cn.OpenSchema(adSchemaTables, rCriteria)
    
    While Not rsSchema.EOF
       For Each fld In rsSchema.Fields
          If fld.Name = "TABLE_NAME" Then
            cmbTablas.AddItem fld.value
          End If
       Next
       rsSchema.MoveNext
    Wend
    
    cmbTablas.ListIndex = 0
    
    m_lenTableName = 30
    
    With rchtxt
    
        .Font = "Lucida Console"
        
        For k = 0 To cmbTablas.ListCount - 1

            .SelColor = RGB(128, 0, 192)
            .SelText = "************************************************************" & vbCrLf
            .SelBold = True
            .SelColor = RGB(128, 0, 192) 'XP
            .SelText = "  " & cmbTablas.List(k) & vbCrLf
            .SelBold = False
            .SelColor = RGB(128, 0, 192) 'XP
            .SelText = "  ----------------------------------------------------------" & vbCrLf
            .SelColor = vbBlack
        
            ' este metodo funciona aun cuando la tabla este vacia : (rs.EOF = True)
            Set rs = New ADODB.Recordset
            rs.Open "SELECT * FROM [" & cmbTablas.List(k) & "]", cn, adOpenForwardOnly, adLockReadOnly
            Set flds = rs.Fields
            
            For q = 0 To (flds.Count - 1)
            
                lenTableName = Len(flds(q).Name)
                If (lenTableName < m_lenTableName) Then
                    .SelText = "    " & flds(q).Name & Space(m_lenTableName - lenTableName)
                Else
                    .SelText = "    " & flds(q).Name & " "
                End If

                bIsText = False
                bIsDate = False
                
                If gfnc_getTypeDataAdoRecordset(flds(q).Type, bIsText, bIsDate) = "" Then
                    .SelColor = vbRed
                    .SelText = "Tipo no reconocido" & vbCrLf
                    .SelColor = vbBlack
                Else
                    If bIsText = True Then
                        .SelText = sType & Space(19 - Len(sType))
                        .SelText = "[" & flds(q).DefinedSize & "]" & vbCrLf
                    Else
                        If bIsDate = True Then
                            .SelText = sType & Space(19 - Len(sType))
                            .SelText = "<" & flds(q).DefinedSize & ">" & vbCrLf
                        Else
                            .SelText = sType & Space(20 - Len(sType))
                            .SelText = flds(q).DefinedSize & vbCrLf
                        End If
                    End If
                End If
            Next q
JMP_NEXT_TABLE:
        Next k
    End With

    Exit Sub
    
Handler:
    If Err.Number = 380 Then
        'cmbTablas.ListIndex = 0 cuando el combo esta vacio
        MsgBox "No se pudieron leer las tablas de la BD", vbExclamation, "Error"
    Else
        If Err.Number = -2147217900 Then
            'Error en la sintaxis SQL (nombre de objeto)
            Resume JMP_NEXT_TABLE
        Else
            MsgBox Err.Description, vbCritical, "frmReporte_Load()"
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If (Me.width < 4500) Then
        Me.width = 4500
        Exit Sub
    End If
    If (Me.height < 4500) Then
        Me.height = 4500
        Exit Sub
    End If
    
    rchtxt.width = Me.width - 75
    fra.width = rchtxt.width - 45
    rchtxt.height = Me.height - 960
    cmdSalir.Left = fra.width - 1260
End Sub

Private Sub rchtxt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Richtext control doesn't refresh well
    Me.ZOrder 0
    Me.Refresh
End Sub
