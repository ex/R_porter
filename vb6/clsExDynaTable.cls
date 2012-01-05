VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExDynaTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================
' clsExDynaTable.cls
'
' Clase de manipulacion dinamica de tablas
'
'-------------------------------------------------------------------
' PROGRAMADOR:  L. Esau Rodriguez Oscanoa
'
'-------------------------------------------------------------------
Option Explicit

'*******************************************************************
' VARIABLES PRIVADAS
'*******************************************************************
Private mz_DeleteConstrains() As exDynaTableDeleteConstrains
Private mn_Constrains As Integer
Private mz_Fields() As exDynaTableFields
Private mn_Fields As Integer

Private ms_SQL As String
Private ms_mainKey As String
Private ms_mainTable As String
Private ms_mainField As String
Private ms_mainFieldTitle As String
Private ms_SqlFROM As String
Private ms_SqlAdditionalWHERE As String
Private ms_SqlAdditionalORDER_BY As String
Private ms_caption As String
Private mb_dontSearchAtBeginning As Boolean

Private mb_Sql_Ok As Boolean
Private mb_SqlFROM_Ok As Boolean

'*******************************************************************
' PROPIEDADES
'*******************************************************************
Property Let SqlFROM(newSqlFROM As String)
        
    On Error GoTo RAISE_ERROR
    
    If Trim(newSqlFROM) = "" Then
        '-------------------------------------------------------------------
        ' lanzar error
RAISE_ERROR:

        mb_Sql_Ok = False
        Err.Raise Number:=vbObjectError + exDT_SQL_FROM_INVALID, Description:="Cadena SQL FROM no válida"
    Else
        ms_SqlFROM = Trim(newSqlFROM)
        mb_SqlFROM_Ok = True
    End If
End Property

Public Property Get SqlFROM() As String
    SqlFROM = ms_SqlFROM
End Property

Public Property Let Caption(ByVal sNewCaption As String)
    ms_caption = Trim(sNewCaption)
End Property

Public Property Get Caption() As String
    Caption = ms_caption
End Property

Public Property Let SqlAdditionalWHERE(ByVal sSqlAdditionalWHERE As String)
    ms_SqlAdditionalWHERE = Trim(sSqlAdditionalWHERE)
End Property

Public Property Get SqlAdditionalWHERE() As String
    SqlAdditionalWHERE = ms_SqlAdditionalWHERE
End Property

Public Property Let SqlAdditionalORDER_BY(ByVal sSqlAdditionalORDER_BY As String)
    ms_SqlAdditionalORDER_BY = Trim(sSqlAdditionalORDER_BY)
End Property

Public Property Get SqlAdditionalORDER_BY() As String
    SqlAdditionalORDER_BY = ms_SqlAdditionalORDER_BY
End Property

Public Property Let dontSearchAtBeginning(ByVal bDontSearch As Boolean)
    mb_dontSearchAtBeginning = bDontSearch
End Property

'*******************************************************************
' CLASE
'*******************************************************************
Private Sub Class_Initialize()
    mb_SqlFROM_Ok = False
    ms_caption = ""
    SqlAdditionalWHERE = ""
    SqlAdditionalORDER_BY = ""
    mn_Fields = 0
    mb_dontSearchAtBeginning = False
End Sub

'*******************************************************************
' FUNCIONES
'*******************************************************************
Public Sub AddFields(ByVal IsMain As Boolean, _
                     ByVal IsLinked As Boolean, _
                     ByVal IsPrimaryKey As Boolean, _
                     ByVal TableShowed As String, _
                     ByVal FieldShowed As String, _
                     Optional ByVal FlexTitle As String = "", _
                     Optional ByVal FlexLenght As Integer = 600, _
                     Optional ByVal AliasShowed As String = "", _
                     Optional ByVal TableLinked As String = "", _
                     Optional ByVal FieldLinked As String = "", _
                     Optional ByVal ShowedLink As String = "")
    
    On Error GoTo JMP_RAISE_ERROR
    
    If (Trim(TableShowed) = "") Or (Trim(FieldShowed) = "") Then
        '-------------------------------------------------------------------
        ' lanzar error
JMP_RAISE_ERROR:
        mb_Sql_Ok = False
        Err.Raise Number:=vbObjectError + exDT_FIELD_INVALID, Description:="El campo añadido no es válido"
    Else
        mn_Fields = mn_Fields + 1
        ReDim Preserve mz_Fields(1 To mn_Fields)
        
        mz_Fields(mn_Fields).bIsMain = IsMain
        mz_Fields(mn_Fields).bIsLinked = IsLinked
        mz_Fields(mn_Fields).bIsPrimaryKey = IsPrimaryKey
        mz_Fields(mn_Fields).sTableShowed = Trim(TableShowed)
        mz_Fields(mn_Fields).sFieldShowed = Trim(FieldShowed)
        mz_Fields(mn_Fields).sFlexTitle = Trim(FlexTitle)
        mz_Fields(mn_Fields).nFlexLenght = FlexLenght
        mz_Fields(mn_Fields).sAliasShowed = Trim(AliasShowed)
        mz_Fields(mn_Fields).sTableLinked = Trim(TableLinked)
        mz_Fields(mn_Fields).sFieldLinked = Trim(FieldLinked)
        mz_Fields(mn_Fields).sShowedLink = Trim(ShowedLink)
    End If
    
End Sub

Public Sub ShowForm(Optional ByVal Formwidth As Integer = 6060)
    
    Dim frmResultQuery As Form
    Dim rs As ADODB.Recordset
    Dim k As Long
    Dim n_fields As Long
    Dim fld_tipo As ADODB.DataTypeEnum
    Dim b_Redraw As Boolean
    Dim bErrIntern As Boolean
    On Error GoTo JMP_RAISE_ERROR
        
    bErrIntern = False
        
    If generateSqlString Then
    
        Screen.MousePointer = vbHourglass

        Set rs = New ADODB.Recordset
        rs.Open ms_SQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

        n_fields = rs.Fields.Count

        If n_fields > 0 Then
            '---------------------------------------
            ' crear nuevo formulario de resultado
            Set frmResultQuery = New frmExDynaTable
            frmResultQuery.Caption = ms_caption
            frmResultQuery.Show vbModeless
            frmResultQuery.gs_SQL = ms_SQL
            frmResultQuery.gs_mainTable = ms_mainTable
            frmResultQuery.gs_mainField = ms_mainField
            frmResultQuery.cmbFieldSearch.AddItem ms_mainFieldTitle
            frmResultQuery.cmbFieldSearch.ListIndex = 0
            frmResultQuery.width = Formwidth
            Set frmResultQuery.go_parent = Me

            For k = 1 To UBound(mz_DeleteConstrains)
                frmResultQuery.AddDeleteConstrains mz_DeleteConstrains(k).sTableLinked, mz_DeleteConstrains(k).sFieldLinked
            Next k
            
JMP_NO_RESTRICTIONS:
            With frmResultQuery.flxResults

                .MousePointer = flexHourglass
                .Redraw = False

                .Rows = 2
                .FixedRows = 1

                .Cols = n_fields + 1
                .FixedCols = 0

                .ColWidth(0) = 540
                .RowHeight(0) = 315

                .Row = 0
                .Col = 0

                .ColAlignment(0) = flexAlignRightCenter
                .text = "Nº"

                '---------------------------------------
                ' poner cabecera
                For k = 1 To n_fields
                
                    .Col = k
                    .CellAlignment = flexAlignLeftCenter
                    
                    If mz_Fields(k).sFlexTitle = "" Then
                        .text = rs.Fields(k - 1).Name
                    Else
                        .text = mz_Fields(k).sFlexTitle
                    End If
                    
                    '---------------------------------------
                    ' alineacion de columnas
                    fld_tipo = rs.Fields(k - 1).Type

                    Select Case fld_tipo
                        Case adVarWChar, adLongVarChar, adChar, adWChar, adBSTR, adDate, adDBDate, adLongVarWChar, adVarChar, adDBTimeStamp
                            .ColAlignment(k) = flexAlignLeftCenter
                        Case Else
                            .ColAlignment(k) = flexAlignRightCenter
                    End Select
                    
                    .ColWidth(k) = mz_Fields(k).nFlexLenght
                    
                    ' ver si es indice
                    If mz_Fields(k).bIsPrimaryKey Then
                        frmResultQuery.gs_indexField = mz_Fields(k).sFieldShowed
                        frmResultQuery.gl_IdRow = k
                    End If
                    
                Next k

                b_Redraw = True

                If (rs.EOF) Or mb_dontSearchAtBeginning Then
                    .RowHeight(1) = 0
                    GoTo JMP_FINISH
                End If
                
                '---------------------------------------
                ' llenar datos
                While rs.EOF = False

                    .Row = .Rows - 1
                    'forzar visible
                    .RowHeight(.Row) = -1

                    .Col = 0
                    .CellForeColor = RGB(100, 170, 255)
                    .text = .Row

                    For k = 0 To (n_fields - 1)
                        .Col = k + 1
                        .text = rs.Fields(k).value
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
JMP_FINISH:
                .Row = 1
                .Col = 0
                .ColSel = n_fields

                .Redraw = True
                .SetFocus
            End With

            frmResultQuery.ZOrder 0
            frmResultQuery.Refresh

        End If

        rs.Close
        Set rs = Nothing
        
        frmResultQuery.flxResults.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
    Else
        '-------------------------------------------------------------------
        ' lanzar error
        Err.Raise Number:=vbObjectError + exDT_SQL_NOT_ESTABLISHED, Description:="Cadena SQL no establecida"
        Exit Sub
    End If
    
    Exit Sub
   
JMP_RAISE_ERROR:
    Select Case Err.Number
        Case 9
            ' La matriz de restricciones esta vacia
            Resume JMP_NO_RESTRICTIONS
        Case 94, 13
            'uso no valido de NULL
            'cuando el campo esta vacio
            Resume Next
        Case Else
            Screen.MousePointer = vbDefault
            Err.Raise Number:=Err.Number, Description:=Err.Description
    End Select
End Sub

Private Function generateSqlString(Optional key As String = "") As Boolean ' [key] is used in editing

    Dim s_sql As String
    Dim s_sqlwhere As String
    Dim b_LinkFound As Boolean
    Dim k As Integer
    On Error GoTo Handler
    
    generateSqlString = False

    s_sqlwhere = ""
    b_LinkFound = False
    
    '--------------------------------------------------
    ' Add the SELECT
    s_sql = "SELECT "
    For k = 1 To UBound(mz_Fields)
    
        s_sql = s_sql & mz_Fields(k).sTableShowed & "." & mz_Fields(k).sFieldShowed
        
        If mz_Fields(k).sAliasShowed <> "" Then
            s_sql = s_sql & " AS " & mz_Fields(k).sAliasShowed
        End If
        
        If mz_Fields(k).bIsLinked Then
            If key <> "" Then
                ' If editing we need extra fields
                s_sql = s_sql & ", " & mz_Fields(k).sTableShowed & "." & mz_Fields(k).sShowedLink
            End If
            s_sqlwhere = s_sqlwhere & "(" & mz_Fields(k).sTableShowed & "." & _
                         mz_Fields(k).sShowedLink & "=" & mz_Fields(k).sTableLinked & "." & _
                         mz_Fields(k).sFieldLinked & ") AND "
            b_LinkFound = True
        End If
    
        If mz_Fields(k).bIsMain Then
            ms_mainFieldTitle = mz_Fields(k).sFlexTitle
            ms_mainField = mz_Fields(k).sTableShowed & "." & mz_Fields(k).sFieldShowed
            ms_mainTable = mz_Fields(k).sTableShowed
        End If
    
        If mz_Fields(k).bIsPrimaryKey Then
            ms_mainKey = mz_Fields(k).sFieldShowed
        End If
        
        s_sql = s_sql & ", "
        
    Next
    
    s_sql = Mid(s_sql, 1, Len(s_sql) - 2)                       ' quitamos coma final
    
    If b_LinkFound Then
        s_sqlwhere = Mid(s_sqlwhere, 1, Len(s_sqlwhere) - 5)    ' quitamos AND final
    End If
    
    '--------------------------------------------------
    ' Add the FROM
    s_sql = s_sql & " FROM " & ms_SqlFROM
    
    '--------------------------------------------------
    ' Add the WHERE
    If b_LinkFound Then
        If (ms_SqlAdditionalWHERE <> "") Then
            s_sql = s_sql & " WHERE " & s_sqlwhere & " AND " & ms_SqlAdditionalWHERE
        Else
            s_sql = s_sql & " WHERE " & s_sqlwhere
        End If
    Else
        If (ms_SqlAdditionalWHERE <> "") Then
            s_sql = s_sql & " WHERE  " & ms_SqlAdditionalWHERE
        End If
    End If
    If key <> "" Then
        s_sql = s_sql & " AND (" & ms_mainTable & "." & ms_mainKey & "=" & key & ")"
    End If
    
    '--------------------------------------------------
    ' Add the  ORDER BY
    If ms_SqlAdditionalORDER_BY <> "" Then
        s_sql = s_sql & " ORDER BY " & ms_SqlAdditionalORDER_BY
    End If
    
    ms_SQL = s_sql
    generateSqlString = True
    
    Exit Function
    
Handler:
End Function

Public Sub AddDeleteConstrains(ByRef strTable As String, ByRef strField As String)

    On Error GoTo JMP_RAISE_ERROR
    
    If (Trim(strTable) = "") Or (Trim(strField) = "") Then
        '-------------------------------------------------------------------
        ' lanzar error
JMP_RAISE_ERROR:

        mb_Sql_Ok = False
        Err.Raise Number:=vbObjectError + exDT_FIELD_INVALID, Description:="La restriccion añadida no es valida"
    Else
        mn_Constrains = mn_Constrains + 1
        ReDim Preserve mz_DeleteConstrains(1 To mn_Constrains)
        mz_DeleteConstrains(mn_Constrains).sTableLinked = strTable
        mz_DeleteConstrains(mn_Constrains).sFieldLinked = strField
    End If
    
End Sub

Public Sub initEditAction(key As String)
    
    Dim rs As ADODB.Recordset
    Dim frmEdit As Form
    On Error GoTo JMP_RAISE_ERROR
    
    If generateSqlString(key) Then
    
        Set rs = New ADODB.Recordset
        rs.Open ms_SQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        If Not (rs.EOF) And (rs.Fields.Count > 0) Then
            '---------------------------------------
            ' crear nuevo formulario
            Set frmEdit = New frmExDynaEdit
            frmEdit.Caption = "Editando en: " & ms_caption
            frmEdit.Show vbModeless
        
            With frmEdit.flxResults
        
                .Redraw = False
        
                .Rows = mn_Fields + 1
                .FixedRows = 1
        
                .Cols = 5
                .FixedCols = 2
        
                .ColWidth(0) = 0
                .ColWidth(1) = 1800
                .ColWidth(2) = 2700
                .ColWidth(3) = 2700
                .ColWidth(4) = 0
                .RowHeight(0) = 315
        
                .Row = 0
                .Col = 1
        
                .CellAlignment = flexAlignCenterCenter
                .text = "Campo"
        
                .Col = 2
                .ColAlignment(1) = flexAlignLeftCenter
                .CellAlignment = flexAlignCenterCenter
                .text = "Valor"
        
                .Col = 3
                .ColAlignment(2) = flexAlignLeftCenter
                .CellAlignment = flexAlignCenterCenter
                .text = "Dato asociado"
                
                '---------------------------------------
                ' poner campos
                Dim k As Integer
                For k = 1 To mn_Fields
                    .Row = k
                    If mz_Fields(k).bIsLinked Then
                        .Col = 0
                        .text = "SELECT " & mz_Fields(k).sFieldShowed & " FROM " & _
                                mz_Fields(k).sTableShowed & " WHERE " & mz_Fields(k).sShowedLink
                    End If
                    .Col = 1
                    If mz_Fields(k).sFlexTitle = "" Then
                        .text = rs.Fields(k - 1).Name
                    Else
                        .text = mz_Fields(k).sFlexTitle
                    End If
                Next k
        
                '---------------------------------------
                ' llenar datos
                Dim numLinkedFields As Integer
                Dim isText As Boolean
                Dim isDate As Boolean
                Dim sqlPrefix As String
                Dim sql As String
                numLinkedFields = 0
                sql = "UPDATE " & ms_mainTable & " SET " ' Initialize SQL instruction for updating
                For k = 0 To (mn_Fields - 1)
                    .Row = k + 1
                    .Col = 2
                    If mz_Fields(k + 1).bIsPrimaryKey Then
                        .CellForeColor = RGB(0, 0, 255)
                        .text = rs.Fields(k + numLinkedFields).value
                    Else
                        .text = rs.Fields(k + numLinkedFields).value
                    End If
                    If mz_Fields(k + 1).bIsLinked Then
                        .text = rs.Fields(k + numLinkedFields + 1).value
                        .Col = 3
                        .text = rs.Fields(k + numLinkedFields).value
                        numLinkedFields = numLinkedFields + 1
                    End If
                    ' Get type of data
                    .Col = 4
                    isText = False
                    isDate = False
                    sqlPrefix = ""
                    gfnc_getTypeDataAdoRecordset rs.Fields(k + numLinkedFields).Type, isText, isDate
                    If isText Then
                        .text = "Text"
                        sqlPrefix = "'"
                    End If
                    If isDate Then
                        .text = "Date"
                        sqlPrefix = "#"
                    End If
                    If mz_Fields(k + 1).bIsLinked Then
                        sql = sql & mz_Fields(k + 1).sFieldLinked & "=" & _
                              sqlPrefix & DYNATABLE_FIELD_TAG & sqlPrefix & ", "
                    Else
                        sql = sql & mz_Fields(k + 1).sFieldShowed & "=" & _
                              sqlPrefix & DYNATABLE_FIELD_TAG & sqlPrefix & ", "
                    End If
                    If mz_Fields(k + 1).bIsPrimaryKey Then
                        .text = "PK"
                    End If
                Next k
                sql = Mid(sql, 1, Len(sql) - 2)        ' quitamos coma final
                sql = sql & " WHERE " & ms_mainKey & "=" & key

                .Row = 1
                .Col = 2
                .ColSel = 3
                .Redraw = True
                .SetFocus
            End With
            
            ' Resizing
            frmEdit.width = 7620
            frmEdit.height = 1110 + mn_Fields * 270
            frmEdit.gs_SQL = sql
            Set frmEdit.go_parent = Me
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If

    Exit Sub
   
JMP_RAISE_ERROR:
    Select Case Err.Number
        Case 13
            'cuando el campo esta vacio
            Resume Next
        Case Else
            Err.Raise Number:=Err.Number, Description:=Err.Description
    End Select
End Sub

Public Sub initInsertAction()
    
    Dim id_insert As Long
    Dim rs As ADODB.Recordset
    Dim frmInsert As Form
    On Error GoTo JMP_RAISE_ERROR
    
    ' [WARNING] Just number PKs
    Set rs = New ADODB.Recordset
    rs.Open "SELECT MAX(" & ms_mainKey & ") AS max_id FROM " & ms_mainTable, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If IsNull(rs!max_id) Then
        id_insert = 1
    Else
        id_insert = rs!max_id + 1
    End If
    rs.Close
    Set rs = Nothing
    
    If generateSqlString(str(id_insert)) Then
    
        Set rs = New ADODB.Recordset
        rs.Open ms_SQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        If (rs.EOF) And (rs.Fields.Count > 0) Then
    
            '---------------------------------------
            ' crear nuevo formulario
            Set frmInsert = New frmExDynaInsert
            frmInsert.Caption = "Insertando en: " & ms_caption
            frmInsert.Show vbModeless
        
            With frmInsert.flxResults
        
                .Redraw = False
        
                .Rows = mn_Fields + 1
                .FixedRows = 1
        
                .Cols = 5
                .FixedCols = 2
        
                .ColWidth(0) = 0
                .ColWidth(1) = 1800
                .ColWidth(2) = 2700
                .ColWidth(3) = 2700
                .ColWidth(4) = 0
                .RowHeight(0) = 315
        
                .Row = 0
                .Col = 1
        
                .CellAlignment = flexAlignCenterCenter
                .text = "Campo"
        
                .Col = 2
                .ColAlignment(1) = flexAlignLeftCenter
                .CellAlignment = flexAlignCenterCenter
                .text = "Valor"
        
                .Col = 3
                .ColAlignment(2) = flexAlignLeftCenter
                .CellAlignment = flexAlignCenterCenter
                .text = "Dato asociado"
                
                '---------------------------------------
                ' poner campos
                Dim k As Integer
                For k = 1 To mn_Fields
                    .Row = k
                    If mz_Fields(k).bIsLinked Then
                        .Col = 0
                        .text = "SELECT " & mz_Fields(k).sFieldShowed & " FROM " & _
                                mz_Fields(k).sTableShowed & " WHERE " & mz_Fields(k).sShowedLink
                    End If
                    .Col = 1
                    .text = mz_Fields(k).sFlexTitle
                Next k
        
                '---------------------------------------
                ' llenar datos
                Dim numLinkedFields As Integer
                Dim isText As Boolean
                Dim isDate As Boolean
                Dim sqlPrefix As String
                Dim sql_fields As String
                Dim sql_values As String
                numLinkedFields = 0
                
                sql_fields = "INSERT INTO " & ms_mainTable & " (" ' Initialize SQL instruction for inserting
                sql_values = "VALUES ("
                For k = 0 To (mn_Fields - 1)
                    .Row = k + 1
                    If mz_Fields(k + 1).bIsLinked Then
                        numLinkedFields = numLinkedFields + 1
                    End If
                    ' Get type of data
                    .Col = 4
                    isText = False
                    isDate = False
                    sqlPrefix = ""
                    gfnc_getTypeDataAdoRecordset rs.Fields(k + numLinkedFields).Type, isText, isDate
                    If isText Then
                        .text = "Text"
                        sqlPrefix = "'"
                    End If
                    If isDate Then
                        .text = "Date"
                        sqlPrefix = "#"
                    End If
                    If mz_Fields(k + 1).bIsLinked Then
                        sql_fields = sql_fields & mz_Fields(k + 1).sFieldLinked & ", "
                        sql_values = sql_values & sqlPrefix & DYNATABLE_FIELD_TAG & sqlPrefix & ", "
                    Else
                        sql_fields = sql_fields & mz_Fields(k + 1).sFieldShowed & ", "
                        sql_values = sql_values & sqlPrefix & DYNATABLE_FIELD_TAG & sqlPrefix & ", "
                              
                    End If
                    If mz_Fields(k + 1).bIsPrimaryKey Then
                        .text = "PK"
                        .Col = 2
                        .CellForeColor = RGB(0, 0, 255)
                        .text = id_insert
        
                    End If
                Next k
                sql_fields = Mid(sql_fields, 1, Len(sql_fields) - 2)    ' quitamos coma final
                sql_values = Mid(sql_values, 1, Len(sql_values) - 2)
                sql_fields = sql_fields & ") " & sql_values & ")"
        
                .Row = 1
                .Col = 2
                .ColSel = 3
                .Redraw = True
                .SetFocus
            End With
            
            ' Resizing
            frmInsert.width = 7620
            frmInsert.height = 1110 + mn_Fields * 270
            frmInsert.gs_SQL = sql_fields
            Set frmInsert.go_parent = Me
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If

    Exit Sub
   
JMP_RAISE_ERROR:
    Err.Raise Number:=Err.Number, Description:=Err.Description
End Sub

