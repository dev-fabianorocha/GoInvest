Attribute VB_Name = "ModBanco"
Option Explicit
Public pServidor As String
Public pVersao As String
Public pBanco As String
Dim fUsuario As String
Dim fSenha As String

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

Public Function Conexao(Optional ParConexaoWindows As Boolean) As ADODB.Connection
On Error GoTo Trata

Dim sConexao As New ADODB.Connection
   
'sConexao.Open "Provider=SQLOLEDB; " & _
'"Initial Catalog=" & pBanco & ";" & _
'"Data Source=" & pServidor & ";" & _
'"integrated security=SSPI; persist security info=True;", fUsuario, fSenha

With sConexao
    .Open "Provider=SQLOLEDB; " _
        & " Initial Catalog=" & pBanco & ";" _
        & " Data Source=" & pServidor & ";" _
        & IIf(ParConexaoWindows, "integrated security=SSPI;", "") _
        & " persist security info=True;", fUsuario, fSenha
End With

Set Conexao = sConexao

Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description), vbCritical, "ModBanco.Conexao"
End Function

'Public Function ExecutarInsert(ParSql As String) As Boolean
'On Error GoTo Trata

'Dim ExecutaSql As New ADODB.Command

'Dim sConexao As ADODB.Connection
'Set sConexao = New ADODB.Connection

'sConexao = Conexao
'sConexao.Open

'With ExecutaSql
'   .ActiveConnection = sConexao
'   .CommandType = adCmdText
'   .CommandText = ParSql
'End With

'ExecutarInsert = True
'Exit Function
'Resume
'Trata:
'MsgBox DescError(Err.Number, Err.Description, ParSql), vbCritical, "ModBanco.ExecutarInsert"
'End Function

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
sRetorno = sRetorno & ParSql & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf

DescError = sRetorno
End Function

Public Function AlimentarRodape() As String
Dim sRetorno As String

sRetorno = "| Servidor: " & pServidor & " | Banco de Dados: " & pBanco & " | Usuário: " & pUsuario & " | V." & pVersao & " | "

AlimentarRodape = sRetorno
End Function

Public Function LerConfig() As Boolean
On Error GoTo Trata
Dim sServidor As String, sBanco As String, sUsuario As String, sSenha As String, sTexto As String, sRetorno As Boolean

If Dir("C:\GoInvest\Config.ini") <> Empty Then
    Open "C:\GoInvest\Config.ini" For Input As #1
        Do While Not EOF(1)
            Input #1, sTexto
            If Mid(sTexto, 1, 2) = "01" Then sServidor = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "02" Then sBanco = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "03" Then sUsuario = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "04" Then sSenha = Mid(sTexto, 4)
        Loop
    Close #1
    
    If sServidor <> Empty And sBanco <> Empty And sUsuario <> Empty And sSenha <> Empty Then
        pServidor = sServidor
        pBanco = sBanco
        fUsuario = sUsuario
        fSenha = sSenha
    End If
    sRetorno = True
Else
    sRetorno = False
    frmConfig.Show
End If

LerConfig = sRetorno
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, ""), vbCritical, "ModBanco.ExecutarSql"
End Function

Public Function GravarConfig(ParServidor As String, ParBanco As String, ParUsuario As String, ParSenha As String) As Boolean
On Error GoTo Trata
Dim sLinha As Long

Open "C:\GoInvest\Config.ini" For Output As #1
For sLinha = 1 To 4
    If sLinha = 1 Then Print #1, "0" & sLinha & "=" & ParServidor
    If sLinha = 2 Then Print #1, "0" & sLinha & "=" & ParBanco
    If sLinha = 3 Then Print #1, "0" & sLinha & "=" & ParUsuario
    If sLinha = 4 Then Print #1, "0" & sLinha & "=" & ParSenha
Next
Close #1

GravarConfig = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, ""), vbCritical, "ModBanco.ExecutarSql"
End Function
