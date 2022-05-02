Attribute VB_Name = "ModGrid"
Option Explicit

Public Sub MarcarLinha(ByRef ParGrid As fpSpread, ByVal ParRow As Long, ByRef ParCodigo As Integer)
Dim sContCol As Long, sContRow As Long
    With ParGrid
'        For sContRow = 1 To .MaxRows
'            For sContCol = 1 To .MaxCols
'                .Row = sContRow
'                .col = sContCol
'                .BackColor = &HE0E0E0
'                .ForeColor = &H0&
'            Next
'        Next
'        For sContCol = 1 To .MaxCols
'            .Row = ParRow
'            .col = sContCol
'            .BackColor = &H0&
'            .ForeColor = &HFFFFFF
'        Next
    .col = 1
    .Row = ParRow
    ParCodigo = .Text
    End With
End Sub

Public Sub LimparGrid(ByRef ParGrid As fpSpread)
Dim sContCol As Long, sContRow As Long

With ParGrid
    For sContRow = 1 To .MaxRows
        For sContCol = 1 To .MaxCols
            .Row = sContRow
            .col = sContCol
            .Text = 0
            .RowHidden = True
        Next
    Next
End With

End Sub

Public Sub PopularGrid(ByRef ParGrid As fpSpread, ByVal ParSql As String)
Dim sConsulta As New ADODB.Recordset, sLinhas As Long, sCont As Long

Set sConsulta = Consulta(ParSql, sLinhas)

With ParGrid
    For sCont = 1 To sLinhas
        .Row = sCont
        .RowHidden = False
    Next
    Set .DataSource = sConsulta
End With

End Sub

Public Function PegarTextoGrid(ByRef ParGrid As fpSpread, ByVal ParCol As Long, ByVal ParRow As Long) As String
Dim sRetorno As String

    With ParGrid
        .col = ParCol
        .Row = ParRow
        sRetorno = .Text
    End With
    
PegarTextoGrid = sRetorno
End Function
