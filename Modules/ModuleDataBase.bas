Attribute VB_Name = "ModuleDataBase"
Option Explicit

Private eUser As String
Private ePassword As String

Public Function ReadQuery(Query_ As String, Optional ByRef Rows_ As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

Dim iConnection As New ADODB.Connection, iRecordset As New ADODB.Recordset

iConnection = Connection
iConnection.Open

With iRecordset
    .Open Query_, iConnection, adOpenStatic
    Rows_ = .RecordCount
End With

Set ReadQuery = iRecordset

Exit Function
Resume
ErrorHandler:
MsgBox ErrorHandler(Err.Number, Err.Description, Query_), vbCritical, "ModBanco.ReadQuery"
End Function

Public Function Connection() As ADODB.Connection
On Error GoTo ErrorHandler

Dim iConnection As New ADODB.Connection

With iConnection
    If gWindowsConnection Then
    .Open "Provider=SQLOLEDB; " & _
        "Initial Catalog=" & gConnection & ";" & _
        "Data Source=" & gServer & ";" & _
        "integrated security=SSPI; persist security info=True;"
    Else
    .Open "Provider=SQLOLEDB; " _
        & " Initial Catalog=" & gConnection & ";" _
        & " Data Source=" & gServer & ";" _
        & " persist security info=True;", eUser, ePassword
    End If
End With

Set Connection = iConnection

Exit Function
Resume
ErrorHandler:
MsgBox ErrorHandler(Err.Number, Err.Description), vbCritical, "ModBanco.Connection"
End Function

Public Function QueryExecute(Query_ As String) As Boolean
On Error GoTo ErrorHandler
Dim iRecordset As New ADODB.Recordset

Dim iConnection As New ADODB.Connection

iConnection = Connection
iConnection.Open

With iConnection
    .Execute (Query_)
End With

QueryExecute = True
Exit Function
Resume
ErrorHandler:
MsgBox ErrorHandler(Err.Number, Err.Description, Query_), vbCritical, "ModBanco.QueryExecute"
End Function

Public Function ReadConfig() As Boolean
On Error GoTo ErrorHandler
Dim iServer As String, iDataBase As String, iUser As String, iPassword As String, iText As String, iReturn As Boolean, iWindowsConnection As Boolean

If Dir("C:\GoInvest\Config.ini") <> Empty Then
    Open "C:\GoInvest\Config.ini" For Input As #1
        Do While Not EOF(1)
            Input #1, iText
            If Mid(iText, 1, 2) = "01" Then iServer = Mid(iText, 4)
            If Mid(iText, 1, 2) = "02" Then iDataBase = Mid(iText, 4)
            If Mid(iText, 1, 2) = "03" Then iUser = Mid(iText, 4)
            If Mid(iText, 1, 2) = "04" Then iPassword = Mid(iText, 4)
            If Mid(iText, 1, 2) = "05" Then
                If Mid(iText, 4) = "Verdadeiro" Then
                    iWindowsConnection = True
                Else
                    iWindowsConnection = False
                End If
            End If
        Loop
    Close #1
    
    If iServer <> Empty And iDataBase <> Empty Then
        gServer = iServer
        gConnection = iDataBase
        eUser = iUser
        ePassword = iPassword
        gWindowsConnection = iWindowsConnection
    End If
    iReturn = True
Else
    iReturn = False
    frmConfig.Show
End If

ReadConfig = iReturn
Exit Function
Resume
ErrorHandler:
MsgBox ErrorHandler(Err.Number, Err.Description, ""), vbCritical, "ModBanco.ReadConfig"
End Function

Public Function GravarConfig(ParServidor As String, ParBanco As String, ParUsuario As String, ParSenha As String, ParConexaoWindows As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim sLinha As Long

Open "C:\GoInvest\Config.ini" For Output As #1
For sLinha = 1 To 5
    If sLinha = 1 Then Print #1, "0" & sLinha & "=" & ParServidor
    If sLinha = 2 Then Print #1, "0" & sLinha & "=" & ParBanco
    If sLinha = 3 Then Print #1, "0" & sLinha & "=" & ParUsuario
    If sLinha = 4 Then Print #1, "0" & sLinha & "=" & ParSenha
    If sLinha = 5 Then Print #1, "0" & sLinha & "=" & ParConexaoWindows
Next
Close #1

GravarConfig = True
Exit Function
Resume
ErrorHandler:
MsgBox ErrorHandler(Err.Number, Err.Description, ""), vbCritical, "ModBanco.QueryExecute"
End Function


