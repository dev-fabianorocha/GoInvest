Attribute VB_Name = "ModVariavel"
Option Explicit

'©Copyright 2022 - Fabiano Gomes da Rocha

Public pClsUsuario As clsUsuarios

Public Enum EnumOpcao
    eIncluir
    eCosultar
    eAlterar
    eExcluir
    eConfirmar
    eCancelar
    eSair
End Enum

Public Enum EnumType
    eStringText
    eByteNumber
    eIntNumber
    eLongNumber
    eDoubleNumber
    eCurrencyNumber
    eDate
    eLogic
    eLogicSql
    eNumberSql
    eStringSql
End Enum

Public Function TratarVariavel(ByVal parVariavel As Variant, ByVal parTipo As EnumType) As Variant
Dim sRetorno As Variant

If parTipo = EnumType.eStringText Then
    If Not parVariavel = Empty Then
        sRetorno = CStr(parVariavel)
    Else
        sRetorno = Empty
    End If
ElseIf parTipo = EnumType.eByteNumber Then
    If Not parVariavel = Empty Then
        sRetorno = CByte(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eIntNumber Then
    If Not parVariavel = Empty Then
        sRetorno = CInt(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eLongNumber Then
    If Not parVariavel = Empty Then
        sRetorno = CLng(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eDoubleNumber Then
    If Not parVariavel = Empty Then
        sRetorno = CDbl(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eCurrencyNumber Then
    If Not parVariavel = Empty Then
        sRetorno = CCur(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eDate Then
    If Not parVariavel = Empty Then
        sRetorno = CDate(parVariavel)
    Else
        sRetorno = "00/00/0000"
    End If
ElseIf parTipo = EnumType.eLogic Then
    If Not parVariavel = Empty Then
        sRetorno = CBool(parVariavel)
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eLogicSql Then
    If Not parVariavel = Empty Then
        sRetorno = CByte(IIf(parVariavel, 1, 0))
    Else
        sRetorno = 0
    End If
ElseIf parTipo = EnumType.eNumberSql Then
    If Not parVariavel = Empty Then
        sRetorno = Replace(parVariavel, ",", ".")
    Else
        sRetorno = 0
    End If
ElseIf parTipo = eStringSql Then
    If Not parVariavel = Empty Then
        sRetorno = "'" & parVariavel & "'"
    Else
        sRetorno = Empty
    End If
End If

TratarVariavel = sRetorno

End Function

Public Function PreencherCombo(ByRef ParCombo As ComboBox, ByVal parSql As String)

Dim sRecordset As New ADODB.Recordset, sLinhas As Long

Set sRecordset = ConsultarSql(parSql, 0)
ParCombo.AddItem " ", 0

With sRecordset
    While Not .EOF
        ParCombo.AddItem TratarVariavel(!Descricao, eStringText), (TratarVariavel(!COR_CODIGO, eLongNumber))
        .MoveNext
    Wend
End With

End Function

