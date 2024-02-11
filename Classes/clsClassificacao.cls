VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsClassificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fCodigo As Integer
Dim fDescricao As String
Dim fUsuario As String
Dim fTipo As Boolean
Dim fImportancia As Byte
Dim fSenha As String
Dim fCadastro As Date
Dim fAtualizacao As Date
Dim fInativo As Byte
Dim fClsCriptografia As ClsCriptografia

Property Get Codigo() As Integer
Codigo = fCodigo
End Property
Property Let Codigo(ByVal ParCodigo As Integer)
fCodigo = ParCodigo
End Property
Property Get Descricao() As String
Descricao = fDescricao
End Property
Property Let Descricao(ByVal parDescricao As String)
fDescricao = parDescricao
End Property

Property Get Usuario() As String
Usuario = fUsuario
End Property
Property Let Usuario(ByVal parUsuario As String)
fUsuario = parUsuario
End Property

Property Get Senha() As String
Senha = fSenha
End Property
Property Let Senha(ByVal parSenha As String)
fSenha = parSenha
End Property
Property Get Cadastro() As Date
Cadastro = fCadastro
End Property
Property Let Cadastro(ByVal ParCadastro As Date)
fCadastro = ParCadastro
End Property
Property Get Atualizacao() As Date
Atualizacao = fAtualizacao
End Property
Property Let Atualizacao(ByVal ParAtualizacao As Date)
fAtualizacao = ParAtualizacao
End Property

Property Get Importancia() As Byte
Importancia = fImportancia
End Property
Property Let Importancia(ByVal parImportancia As Byte)
fImportancia = parImportancia
End Property

Property Get Tipo() As Byte
Tipo = fTipo
End Property
Property Let Tipo(ByVal parTipo As Byte)
fTipo = parTipo
End Property

Property Get Inativo() As Byte
Inativo = fInativo
End Property
Property Let Inativo(ByVal ParInativo As Byte)
fInativo = ParInativo
End Property
Public Function Consultar(ParCodigo) As Boolean
Dim sSql As String

sSql = "SELECT CLA_CODIGO, CLA_DESCRICAO, CLA_TIPO, CLA_IMPORTANCIA, CLA_USUARIO, CLA_CADASTRO, CLA_ATUALIZACAO, CLA_INATIVO FROM CLASSIFICACAO WHERE CLA_CODIGO = " _
    & TratarVariavel(ParCodigo, eStringSql)

Dim sRecordset As New Recordset

Set sRecordset = ConsultarSql(sSql, 0)

If Not sRecordset.EOF Then
    With sRecordset
        fCodigo = !CLA_CODIGO
        fDescricao = TratarVariavel(!CLA_DESCRICAO, eStringText)
        fTipo = TratarVariavel(!CLA_TIPO, eLogic)
        fImportancia = !CLA_IMPORTANCIA
        fUsuario = !CLA_USUARIO
        fCadastro = TratarVariavel(!CLA_CADASTRO, eStringText)
        fAtualizacao = TratarVariavel(!CLA_ATUALIZACAO, eStringText)
        fInativo = TratarVariavel(!CLA_INATIVO, eLogicSql)
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If


Consultar = True
End Function

Public Function Inserir() As Boolean
On Error GoTo Trata
Dim sSql As String


sSql = "INSERT INTO CLASSIFICACAO (CLA_DESCRICAO, CLA_TIPO, CLA_IMPORTANCIA, CLA_USUARIO, CLA_CADASTRO, CLA_ATUALIZACAO, CLA_INATIVO) VALUES (" _
    & TratarVariavel(fDescricao, eStringSql) & ", " & TratarVariavel(fTipo, eLogicSql) & "," & fImportancia & "," & pClsUsuario.Codigo & ", GETDATE(), GETDATE(), '" & fInativo & "')"
If Not ExecutarSql(sSql) Then GoTo Trata

Inserir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsUsuarios.Inserir", sSql
End Function

Public Function Atualizar() As Boolean
On Error GoTo Trata
Dim sSql As String

Dim eClsCipher As New ClsCriptografia

sSql = "UPDATE USUARIOS SET USU_NOME = '" & fDescricao & "', USU_SENHA = '" & eClsCipher.Encrypt(fSenha) & "', USU_ATUALIZACAO = GETDATE(), USU_INATIVO = '" & fInativo _
    & "' WHERE USU_CODIGO = '" & fCodigo & "'"
If Not ExecutarSql(sSql) Then GoTo Trata

Set eClsCipher = Nothing

Atualizar = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsUsuarios.Atualizar", sSql
End Function

Public Function Excluir(ParCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "UPDATE USUARIOS SET USU_INATIVO = '1' WHERE USU_CODIGO = '" & ParCodigo & "'"
If MsgBox("Deseja excluir esse usuário?", vbYesNo, "Corretoras") = 6 Then
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsUsuarios.Excluir", sSql
End Function


