Attribute VB_Name = "ModuleSpread"
Option Explicit

Public Sub SpreadGetCode(ByRef Spread_ As fpSpread, ByVal Row_ As Long, ByRef Code_ As Integer)
Dim iPositionCol As Long, iPositionRow As Long
    With Spread_
        .col = 1
        .Row = Row_
        Code_ = .Text
    End With
End Sub

Public Sub SpreadClean(ByRef Spread_ As fpSpread)
Dim iPositionCol As Long, iPositionRow As Long

With Spread_
    For iPositionRow = 1 To .MaxRows
        For iPositionCol = 1 To .MaxCols
            .Row = iPositionRow
            .col = iPositionCol
            .Text = 0
            .RowHidden = True
        Next
    Next
End With

End Sub

Public Sub SpreadFill(ByRef Spread_ As fpSpread, ByVal Query_ As String)
Dim iRecordset As New ADODB.Recordset, iRows As Long, iPosition As Long

Set iRecordset = eReadQuery(Query_, iRows)

With Spread_
    For iPosition = 1 To iRows
        .Row = iPosition
        .RowHidden = False
    Next
    Set .DataSource = iRecordset
End With

End Sub

Public Function SpreadGetText(ByRef Spread_ As fpSpread, ByVal Col_ As Long, ByVal Row_ As Long) As String
Dim iReturn As String

    With Spread_
        .col = Col_
        .Row = Row_
        iReturn = .Text
    End With
    
SpreadGetText = iReturn
End Function
