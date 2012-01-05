VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
' Clase de encriptacion simple (sin periodicidad)
'*******************************************************************************
' Creado:     Esau (Agosto 2003)
'*******************************************************************************
Option Explicit

Private ml_Cod As Long

Private Const EX_9 = 2
Private Const EX_8 = 3
Private Const EX_7 = 4
Private Const EX_6 = 6
Private Const EX_5 = 7
Private Const EX_4 = 8
Private Const EX_3 = 9
Private Const EX_2 = 11
Private Const EX_1 = 13
Private Const EX_0 = 14

Private Const EX_MAX = 126          ' termina en el 126
Private Const EX_MIN = 32           ' comienza del 32
Private Const EX_DOWN = EX_MIN - 1
Private Const EX_TOP = EX_MAX - EX_DOWN

Private Cods(1 To EX_TOP) As Byte
Private Codt(1 To EX_TOP) As Byte

Public Function SetCod(ByVal cod As Long) As Boolean
Dim num As Byte
Dim k As Byte
    
    If cod > 0 Then
                    
        ml_Cod = cod
        
        Do
            num = cod Mod 10
            
            Select Case num
                Case 0:
                    num = EX_0
                Case 1:
                    num = EX_1
                Case 2:
                    num = EX_2
                Case 3:
                    num = EX_3
                Case 4:
                    num = EX_4
                Case 5:
                    num = EX_5
                Case 6:
                    num = EX_6
                Case 7:
                    num = EX_7
                Case 8:
                    num = EX_8
                Case 9:
                    num = EX_9
            End Select
            
            Ord_Matrix num
            
            num = (cod Mod 100) / 2
            
            If num > 10 Then
                '--------------------------------------
                ' reordenamos para evitar periodicidad
                '--------------------------------------
                'en [Codt] se quedo una copia de [Cods]
                For k = num To EX_TOP
                    Cods(k) = Codt(EX_TOP + num - k)
                Next k
                
                cod = (cod - (cod Mod 10)) / 10
                
            Else
                Exit Do
            End If
        Loop
    End If
    
End Function

Public Function Decrypt(ByVal cad As String) As String
Dim char As Byte
Dim k As Integer

    On Error GoTo Handler
    
    Decrypt = ""
    
    While Len(cad) > 0
        
        char = Asc(cad)
        cad = Mid(cad, 2)
        
        If (char < EX_MIN) Or (char > EX_MAX) Then
            Decrypt = Decrypt & chr$(char)
        Else
            
            For k = 1 To EX_TOP
                If Cods(k) = char Then
                    Exit For
                End If
            Next k
            
            Decrypt = Decrypt & chr$(k + EX_DOWN)
            
        End If
        
    Wend
    
    Exit Function

Handler:
    Decrypt = ""
End Function

Public Function Encrypt(ByVal cad As String) As String
Dim char As Byte

    On Error GoTo Handler
    
    Encrypt = ""
    
    While Len(cad) > 0
        
        char = Asc(cad)
        cad = Mid(cad, 2)
        
        If (char < EX_MIN) Or (char > EX_MAX) Then
            Encrypt = Encrypt & chr$(char)
        Else
            Encrypt = Encrypt & chr$(Cods(char - EX_DOWN))
        End If
        
    Wend
    
    Exit Function

Handler:
    Encrypt = ""
End Function

Private Sub Class_Initialize()
    ml_Cod = 0
    Init_Matrix
End Sub

Private Sub Init_Matrix()
Dim k As Integer
   
    For k = 1 To EX_TOP
        Cods(k) = k + EX_DOWN
    Next k
End Sub

Private Sub Ord_Matrix(digit As Byte)
Dim k As Integer

    Codt(1) = Cods(EX_TOP)
        
    For k = 1 To EX_TOP - 1
        Codt(k + 1) = Cods((digit * k) Mod EX_TOP)
    Next k
        
    For k = 1 To EX_TOP
        Cods(k) = Codt(k)
    Next k
End Sub

