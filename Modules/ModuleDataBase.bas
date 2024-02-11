Attribute VB_Name = "ModBanco"
Option Explicit

Public pUsuarioBanco As String
Public pSenhaBanco As String
Public pServidorBanco As String
Public pVersao As String
Public pConexaoWindows As Boolean
Public pNomeBanco As String

Public Function ConsultarSql(parSql As String, ByRef parLinhas As Long) As ADODB.Recordset
On Error GoTo TratarErro

Dim sConexao As New ADODB.Connection, sRecordset As New ADODB.Recordset

sConexao = ConectarBanco
sConexao.Open
sRecordset.Open parSql, sConexao, adOpenStatic
parLinhas = sRecordset.RecordCount
Set ConsultarSql = sRecordset

Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.ReadQuery", parSql
End Function

Public Function ConectarBanco() As ADODB.Connection
On Error GoTo TratarErro

Dim sConexao As New ADODB.Connection

With sConexao
    If pConexaoWindows Then
    .Open "Provider=SQLOLEDB; " & _
        "Initial Catalog=" & pNomeBanco & ";" & _
        "Data Source=" & pServidorBanco & ";" & _
        "integrated security=SSPI; persist security info=True;"
    Else
    .Open "Provider=SQLOLEDB; " _
        & " Initial Catalog=" & pNomeBanco & ";" _
        & " Data Source=" & pServidorBanco & ";" _
        & " persist security info=True;", pUsuarioBanco, pSenhaBanco
    End If
End With

Set ConectarBanco = sConexao

Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.Connection"
End Function

Public Function ExecutarSql(parSql As String) As Boolean
On Error GoTo TratarErro
Dim sRecordset As New ADODB.Recordset
Dim sConexao As New ADODB.Connection

sConexao = ConectarBanco
sConexao.Open
sConexao.Execute parSql

ExecutarSql = True
Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.QueryExecute", parSql
End Function

Public Function LerConfig() As Boolean
On Error GoTo TratarErro
Dim sServidor As String, sNomeBanco As String, sUsuario As String, sSenha As String, sTexto As String, sRetorno As Boolean
Dim sConexaoWindows As Boolean, sClsCriptografia As New ClsCriptografia

If Dir("C:\GoInvest\Config.ini") <> Empty Then
    Open "C:\GoInvest\Config.ini" For Input As #1
        Do While Not EOF(1)
            Input #1, sTexto
            If Mid(sTexto, 1, 2) = "01" Then sServidor = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "02" Then sNomeBanco = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "03" Then sUsuario = Mid(sTexto, 4)
            If Mid(sTexto, 1, 2) = "04" Then sSenha = sClsCriptografia.Decrypt(Mid(sTexto, 4))
            If Mid(sTexto, 1, 2) = "05" Then
                If Mid(sTexto, 4) = "Verdadeiro" Then
                    sConexaoWindows = True
                Else
                    sConexaoWindows = False
                End If
            End If
        Loop
    Close #1
    
    If sServidor <> Empty And sNomeBanco <> Empty Then
        pServidorBanco = sServidor
        pNomeBanco = sNomeBanco
        pUsuarioBanco = sUsuario
        pSenhaBanco = sSenha
        pConexaoWindows = sConexaoWindows
    End If
    sRetorno = True
Else
    sRetorno = False
    frmConfig.Show
End If

Set sClsCriptografia = Nothing
LerConfig = sRetorno

Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.eReadConfig"
End Function

Public Function EscreverConfig(parServidor As String, parNomeBanco As String, parUsuario As String, parSenha As String, parConexaoWindows As Boolean) As Boolean
On Error GoTo TratarErro
Dim sCont As Long, sClsCriptografia As New ClsCriptografia

Open "C:\GoInvest\Config.ini" For Output As #1
For sCont = 1 To 5
    If sCont = 1 Then Print #1, "0" & sCont & "=" & parServidor
    If sCont = 2 Then Print #1, "0" & sCont & "=" & parNomeBanco
    If sCont = 3 Then Print #1, "0" & sCont & "=" & parUsuario
    If sCont = 4 Then Print #1, "0" & sCont & "=" & sClsCriptografia.Encrypt(parSenha)
    If sCont = 5 Then Print #1, "0" & sCont & "=" & parConexaoWindows
Next
Close #1

Set sClsCriptografia = Nothing
EscreverConfig = True
Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.WriteConfig"
End Function

Public Function PegarCampo(parSql As String, parTipo As EnumType) As Variant
On Error GoTo TratarErro
Dim sRetorno As Variant, sRecordset As New ADODB.Recordset

Set sRecordset = ConsultarSql(parSql, 0)

If Not sRecordset.EOF Then sRetorno = TratarVariavel(sRecordset.Fields(0), parTipo)

PegarCampo = sRetorno

Exit Function
Resume
TratarErro:
ErrorHandler Err.Number, Err.Description, "ModuleDataBase.FieldCollect"
End Function
