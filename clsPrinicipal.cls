VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinicipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Function retornaExtenso(dblValor As Double) As String

    Dim strExtenso As String
    Dim strValor As String
    Dim vetValor() As String
    Dim strCentavos As String
    
    strValor = CStr(dblValor)
    vetValor = Split(strValor, ",")
            
    If CDbl(vetValor(0)) = 1 And CDbl(vetValor(1)) = 0 Then
        strExtenso = "Um real."
    ElseIf vetValor(1) > 0 Then
        retornaDezena (vetValor(1))
    End
    

End Function

Private Function retornaDezena(dblValor As Double) As String
    
    If dblValor = 2 Then
        
        
    End If
End Function
