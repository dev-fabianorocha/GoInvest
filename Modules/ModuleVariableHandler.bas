Attribute VB_Name = "ModuleVariableHandler"
Option Explicit

'Public Variable = g (Global)
'Private Variable = i (Internal)
'Form Variable = e (External)
'Parameter = Example_

'©Copyright 2022 - Fabiano Gomes da Rocha

Public gUserName As String

Public gServer As String
Public gVersion As String
Public gWindowsConnection As Boolean
Public gConnection As String

Public gClsUser As clsUsuarios

Public Enum EnumOption
    eInclude
    eRead
    Update
    eDelete
    eConfirm
    eCancel
    eLeave
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
End Enum

Public Function VariableAdjust(ByVal Variable_ As Variant, ByVal Type_ As EnumType) As Variant
    Dim iReturn As Variant, iPosition As Long, iFind As Boolean, iText As String
    
    If Type_ = EnumType.eStringText Then
        If Not Variable_ = Empty Then
            iReturn = CStr(Variable_)
        Else
            iReturn = Empty
        End If
    ElseIf Type_ = EnumType.eByteNumber Then
        If Not Variable_ = Empty Then
            iReturn = CByte(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eIntNumber Then
        If Not Variable_ = Empty Then
            iReturn = CInt(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eLongNumber Then
        If Not Variable_ = Empty Then
            iReturn = CLng(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eDoubleNumber Then
        If Not Variable_ = Empty Then
            iReturn = CDbl(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eCurrencyNumber Then
        If Not Variable_ = Empty Then
            iReturn = CCur(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eDate Then
        If Not Variable_ = Empty Then
            iReturn = CDate(Variable_)
        Else
            iReturn = "00/00/0000"
        End If
    ElseIf Type_ = EnumType.eLogic Then
        If Not Variable_ = Empty Then
            iReturn = CBool(Variable_)
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eLogicSql Then
        If Not Variable_ = Empty Then
            iReturn = CByte(IIf(Variable_, 1, 0))
        Else
            iReturn = 0
        End If
    ElseIf Type_ = EnumType.eNumberSql Then
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
            ParCombo.AddItem VariableAdjust(!Name, eStringText), (VariableAdjust(!ID, eLongNumber))
            .MoveNext
        Wend
    End With

End Function

