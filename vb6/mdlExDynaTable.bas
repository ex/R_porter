Attribute VB_Name = "mdlExDynaTable"

Option Explicit

Public Type exDynaTableFields
    bIsMain As Boolean
    bIsLinked As Boolean
    bIsPrimaryKey As Boolean
    
    sTableShowed As String
    sFieldShowed As String
    
    sFlexTitle As String
    nFlexLenght As Integer
    
    sAliasShowed As String
    
    sTableLinked As String
    sFieldLinked As String
    sShowedLink As String   ' Stores the index field of the table showed (when field is linked)
End Type

Public Type exDynaTableDeleteConstrains
    sTableLinked As String
    sFieldLinked As String
End Type

Public Enum errExDynaTable
    exDT_SQL_FROM_INVALID = 25000
    exDT_FIELD_INVALID
    exDT_SQL_NOT_ESTABLISHED
    exDT_NO_CONNECTION
    exDT_CRITICAL_ERR
End Enum

Public Const DYNATABLE_FIELD_TAG = "[$]"
