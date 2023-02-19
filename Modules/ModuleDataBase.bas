Attribute VB_Name = "ModuleDataBase"
Option Explicit

Private eUser As String
Private ePassword As String

Public Function eReadQuery(Query_ As String, Optional ByRef Rows_ As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

Dim iConnection As New ADODB.Connection, iRecordset As New ADODB.Recordset

iConnection = Connection
iConnection.Open

With iRecordset
    .Open Query_, iConnection, adOpenStatic
    Rows_ = .RecordCount
End With

Set eReadQuery = iRecordset

Exit Function
Resume
ErrorHandler:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.eReadQuery", Query_
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
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.Connection"
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
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.QueryExecute", Query_
End Function

Public Function eReadConfig() As Boolean
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

eReadConfig = iReturn
Exit Function
Resume
ErrorHandler:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.eReadConfig"
End Function

Public Function WriteConfig(Server_ As String, DataBase_ As String, User_ As String, Password_ As String, WindowsConnection_ As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim iPosition As Long

Open "C:\GoInvest\Config.ini" For Output As #1
For iPosition = 1 To 5
    If iPosition = 1 Then Print #1, "0" & iPosition & "=" & Server_
    If iPosition = 2 Then Print #1, "0" & iPosition & "=" & DataBase_
    If iPosition = 3 Then Print #1, "0" & iPosition & "=" & User_
    If iPosition = 4 Then Print #1, "0" & iPosition & "=" & Password_
    If iPosition = 5 Then Print #1, "0" & iPosition & "=" & WindowsConnection_
Next
Close #1

WriteConfig = True
Exit Function
Resume
ErrorHandler:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.WriteConfig"
End Function


