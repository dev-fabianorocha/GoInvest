Attribute VB_Name = "ModBanco"
Option Explicit
Dim pServidor As String
Dim pBanco As String

Public Function Consulta(ParSql As String, Optional ByRef ParLinhas As Long) As ADODB.Recordset
On Error GoTo Trata

Dim sConexao As New ADODB.Connection, sConsulta As New ADODB.Recordset

sConexao = Conexao
sConexao.Open

With sConsulta
    .Open ParSql, sConexao, adOpenStatic
    ParLinhas = .RecordCount
End With

Set Consulta = sConsulta

Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, ParSql), vbCritical, "ModBanco.Consulta"
End Function

Public Function Conexao() As ADODB.Connection
On Error GoTo Trata

Dim sConexao As New ADODB.Connection

pServidor = "DESKTOP-QO1NU53\PDVNET"
pBanco = "GoInvest"
   
sConexao.Open "Provider=SQLOLEDB; " & _
                  "Initial Catalog=" & pBanco & ";" & _
                  "Data Source=" & pServidor & ";" & _
                  "integrated security=SSPI; persist security info=True;"

Set Conexao = sConexao

Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description), vbCritical, "ModBanco.Conexao"
End Function

Public Function ExecutarInsert(ParSql As String) As Boolean
On Error GoTo Trata

Dim ExecutaSql As New ADODB.Command

Dim sConexao As ADODB.Connection
Set sConexao = New ADODB.Connection

sConexao = Conexao
sConexao.Open

With ExecutaSql
   .ActiveConnection = sConexao
   .CommandType = adCmdText
   .CommandText = ParSql
End With

ExecutarInsert = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, ParSql), vbCritical, "ModBanco.ExecutarInsert"
End Function

Public Function ExecutarSql(ParSql As String) As Boolean
On Error GoTo Trata
Dim sConsulta As New ADODB.Recordset

Dim sConexao As ADODB.Connection
Set sConexao = New ADODB.Connection

sConexao = Conexao
sConexao.Open

With sConexao
    .Execute (ParSql)
End With

ExecutarSql = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, ParSql), vbCritical, "ModBanco.ExecutarSql"
End Function

Public Function DescError(ByVal ParNumero As String, ByVal ParDescricao As String, Optional ByVal ParSql As String) As String
Dim sRetorno As String

sRetorno = "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & "                                         GoInvest" & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & "Error #" & ParNumero & ": '" & ParDescricao & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & ParSql

DescError = sRetorno
End Function

Public Function AlimentarRodape() As String
Dim sRetorno As String

sRetorno = "| Servidor: " & pServidor & " | DataBase: " & pBanco & " | "

AlimentarRodape = sRetorno
End Function

