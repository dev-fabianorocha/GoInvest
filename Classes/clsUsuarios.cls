VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fCodigo As Integer
Dim fNome As String
Dim fSenha As String
Dim fCadastro As Date
Dim fAtualizacao As Date
Dim fInativo As Byte
Dim eClsCipher As clsCipher

Property Get Codigo() As Integer
Codigo = fCodigo
End Property
Property Let Codigo(ByVal ParCodigo As Integer)
fCodigo = ParCodigo
End Property
Property Get Nome() As String
Nome = fNome
End Property
Property Let Nome(ByVal ParNome As String)
fNome = ParNome
End Property
Property Get Senha() As String
Senha = fSenha
End Property
Property Let Senha(ByVal ParSenha As String)
fSenha = ParSenha
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
Property Get Inativo() As Byte
Inativo = fInativo
End Property
Property Let Inativo(ByVal ParInativo As Byte)
fInativo = ParInativo
End Property
Public Function Consultar(Optional ByVal Id_ As Integer, Optional ByVal userName_ As String) As Boolean
Dim sSql As String

sSql = "SELECT USU_CODIGO, USU_NOME, USU_SENHA, USU_CADASTRO, USU_ATUALIZACAO, USU_INATIVO FROM USUARIOS WHERE "

If Id_ <> 0 Then
    sSql = sSql & " USU_CODIGO = '" & Id_ & "'"
Else
    sSql = sSql & " USU_NOME = '" & userName_ & "'"
End If

Set eClsCipher = New clsCipher

If Not ReadQuery(sSql).EOF Then
    With ReadQuery(sSql)
        fCodigo = VariableAdjust(!USU_CODIGO, eIntNumber)
        fNome = VariableAdjust(!USU_NOME, eStringText)
        fSenha = eClsCipher.Decrypt(VariableAdjust(!USU_SENHA, eStringText))
        fCadastro = VariableAdjust(!USU_CADASTRO, eStringText)
        fAtualizacao = VariableAdjust(!USU_ATUALIZACAO, eStringText)
        fInativo = VariableAdjust(!USU_INATIVO, eLogicSql)
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If

Set eClsCipher = Nothing

Consultar = True
End Function

Public Function Inserir() As Boolean
On Error GoTo Trata
Dim sSql As String

Set eClsCipher = New clsCipher

sSql = "INSERT INTO USUARIOS (USU_NOME, USU_SENHA, USU_CADASTRO, USU_ATUALIZACAO, USU_INATIVO) VALUES ('" _
    & fNome & "', '" & eClsCipher.Encrypt(fSenha) & "', GETDATE(), GETDATE(), '" & fInativo & "')"
If Not QueryExecute(sSql) Then GoTo Trata

Set eClsCipher = Nothing

Inserir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsUsuarios.Inserir", sSql
End Function

Public Function Atualizar() As Boolean
On Error GoTo Trata
Dim sSql As String

Set eClsCipher = New clsCipher

sSql = "UPDATE USUARIOS SET USU_NOME = '" & fNome & "', USU_SENHA = '" & eClsCipher.Encrypt(fSenha) & "', USU_ATUALIZACAO = GETDATE(), USU_INATIVO = '" & fInativo _
    & "' WHERE USU_CODIGO = '" & fCodigo & "'"
If Not QueryExecute(sSql) Then GoTo Trata

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
    If Not QueryExecute(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsUsuarios.Excluir", sSql
End Function
