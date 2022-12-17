Attribute VB_Name = "ModuleVariableHandler"
Option Explicit

'Public Variable = g (Global)
'Private Variable = i (Internal)
'Form Variable = e (External)
'Parameter = Example_

Public gUser As String

Public gServer As String
Public gVersion As String
Public gWindowsConnection As Boolean
Public gConnection As String

Public Enum EnumOption
    Include
    Read
    Update
    Delete
    Confirm
    Cancel
    Leave
End Enum

Public Enum EnumType
    StringText
    ByteNumber
    IntNumber
    LongNumber
    DoubleNumber
    CurrencyNumber
    Date
    Logic
    LogicSql
    NumberSql
End Enum

Public Function VariableAdjust(ByVal Variable_ As Variant, ByVal Type_ As EnumType) As Variant
Dim iReturn As Variant, iPosition As Long, iFind As Boolean, iText As String

If Type_ = EnumType.StringText Then
    If Not Variable_ = Empty Then
        iReturn = CStr(Variable_)
    Else
        iReturn = Empty
    End If
ElseIf Type_ = EnumType.ByteNumber Then
    If Not Variable_ = Empty Then
        iReturn = CByte(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.IntNumber Then
    If Not Variable_ = Empty Then
        iReturn = CInt(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.LongNumber Then
    If Not Variable_ = Empty Then
        iReturn = CLng(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.DoubleNumber Then
    If Not Variable_ = Empty Then
        iReturn = CDbl(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.CurrencyNumber Then
    If Not Variable_ = Empty Then
        iReturn = CCur(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.Date Then
    If Not Variable_ = Empty Then
        iReturn = CDate(Variable_)
    Else
        iReturn = "00/00/0000"
    End If
ElseIf Type_ = EnumType.Logic Then
    If Not Variable_ = Empty Then
        iReturn = CBool(Variable_)
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.LogicSql Then
    If Not Variable_ = Empty Then
        iReturn = CByte(IIf(Variable_, 1, 0))
    Else
        iReturn = 0
    End If
ElseIf Type_ = EnumType.NumberSql Then
    If Not Variable_ = Empty Then
        For iPosition = 1 To Len(Variable_)
            iText = Mid(Variable_, iPosition, 1)
            If iText = "," Then iText = "."
            iReturn = iReturn & iText
        Next
    Else
        iReturn = 0
    End If
End If

VariableAdjust = iReturn
End Function

Public Function ComboBoxFill(ByRef ParCombo As ComboBox, ByVal Query_ As String)
Dim iRecordset As New ADODB.Recordset, iRows As Long

Set iRecordset = ReadQuery(Query_, iRows)
ParCombo.AddItem " ", 0

With iRecordset
    While Not .EOF
        ParCombo.AddItem VariableAdjust(!DESCRICAO, StringText), (VariableAdjust(!COR_CODIGO, LongNumber))
        .MoveNext
    Wend
End With

End Function

