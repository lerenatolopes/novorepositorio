VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetornaValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RetornarExtenso(pdblValor As Double) As String
    
    Dim i As Integer
    Dim strExtenso As String
    Dim strValor As String
    Dim vetValor() As String
    Dim strCentavos As String
    Dim strValorDezena As String
    Dim dblValor As Double
    
    strValor = CStr(pdblValor)
    vetValor() = Split(strValor)
'    i = 0
'    While i < Len(strValor) - 1
'        vetValor(i) = Split(Left(strValor, Len(strValor) - 1))
'        i = i + 1
'    Wend
'
         
    If UBound(vetValor) > 0 Then
        If CDbl(vetValor(0)) = 1 And CDbl(vetValor(1)) = 0 Then
            strExtenso = "Um real."
        ElseIf vetValor(1) > 0 Then
            dblValor = vetValor(0)
            strValorDezena = RetornarDezena(dblValor)
                If vetValor(1) > 0 Then
                    dblValor = vetValor(1)
                    strValorDezena = RetornarFinal(dblValor)
                End If
        End If
    Else
        RetornarExtenso = "Um real."
    End If

End Function

Private Function RetornarDezena(pdblValor As Double) As String
    If pdblValor < 10 Then
        If pdblValor = 2 Then
            RetornarDezena = "vinte"
        ElseIf pdblValor = 3 Then
            RetornarDezena = "trinta"
        ElseIf pdblValor = 4 Then
            RetornarDezena = "quarenta"
        ElseIf pdblValor = 5 Then
            RetornarDezena = "cinquenta"
        ElseIf pdblValor = 6 Then
            RetornarDezena = "sessenta"
        ElseIf pdblValor = 7 Then
            RetornarDezena = "setenta"
        ElseIf pdblValor = 8 Then
            RetornarDezena = "oitenta"
        ElseIf pdblValor = 9 Then
            RetornarDezena = "noventa"
        End If
    End If
End Function

Private Function RetornarFinal(pdblValor As Double) As String
    If pdblValor < 10 Then
        If pdblValor = 1 Then
            RetornarFinal = "um"
        ElseIf pdblValor = 2 Then
            RetornarFinal = "dois"
        ElseIf pdblValor = 3 Then
            RetornarFinal = "três"
        ElseIf pdblValor = 4 Then
            RetornarFinal = "quatro"
        ElseIf pdblValor = 5 Then
            RetornarFinal = "cinco"
        ElseIf pdblValor = 6 Then
            RetornarFinal = "seis"
        ElseIf pdblValor = 7 Then
            RetornarFinal = "sete"
        ElseIf pdblValor = 8 Then
            RetornarFinal = "oito"
        ElseIf pdblValor = 9 Then
            RetornarFinal = "nove"
        End If
    End If
End Function
