Attribute VB_Name = "ModTipo"
Option Explicit

Public pUsuario As String

Public Enum enumAcao
    eIncluir
    eConsultar
    eAlterar
    eExcluir
    eConfirmar
    eCancelar
    eSair
End Enum

Public Function TratarVariavel(ByVal ParVariavel As Variant, ByVal ParTipo) As Variant
Dim sRetorno As Variant

If ParTipo = "T" Then
    If Not ParVariavel = Empty Then
        sRetorno = CStr(ParVariavel)
    Else
        sRetorno = Empty
    End If
End If

If ParTipo = "N" Then
    If Not ParVariavel = Empty Then
        sRetorno = CDbl(ParVariavel)
    Else
        sRetorno = 0
    End If
End If

If ParTipo = "D" Then
    If Not ParVariavel = Empty Then
        sRetorno = CDate(ParVariavel)
    Else
        sRetorno = "00/00/0000"
    End If
End If

If ParTipo = "B" Then
    If Not ParVariavel = Empty Then
        sRetorno = CByte(IIf(ParVariavel, 1, 0))
    Else
        sRetorno = 0
    End If
End If

TratarVariavel = sRetorno
End Function

Public Function AlimentarCombo(ByRef ParCombo As ComboBox, ByVal ParSql As String)
Dim sConsulta As New ADODB.Recordset, sLinhas As Long

Set sConsulta = Consulta(ParSql, sLinhas)
ParCombo.AddItem " ", 0

With sConsulta
    While Not .EOF
        ParCombo.AddItem TratarVariavel(!DESCRICAO, "T"), (TratarVariavel(!COR_CODIGO, "N"))
        .MoveNext
    Wend
End With

End Function
