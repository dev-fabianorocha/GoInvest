VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fCodigo As Integer
Dim fNome As String
Dim fAno As String
Dim fCorretora As Integer
Dim fCadastro As Date
Dim fAtualizacao As Date
Dim fInvestir As Byte
Dim fInativo As Byte
Property Get Codigo() As Integer
Codigo = fCodigo
End Property
Property Let Codigo(ByVal ParCodigo As Integer)
fCodigo = ParCodigo
End Property
Property Get Corretora() As Integer
Corretora = fCorretora
End Property
Property Let Corretora(ByVal ParCorretora As Integer)
fCorretora = ParCorretora
End Property
Property Get Nome() As String
Nome = fNome
End Property
Property Let Nome(ByVal ParNome As String)
fNome = ParNome
End Property
Property Get Ano() As String
Ano = fAno
End Property
Property Let Ano(ByVal ParAno As String)
fAno = ParAno
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
Property Get Investir() As Byte
Investir = fInvestir
End Property
Property Let Investir(ByVal ParInvestir As Byte)
fInvestir = ParInvestir
End Property
Property Get Inativo() As Byte
Inativo = fInativo
End Property
Property Let Inativo(ByVal ParInativo As Byte)
fInativo = ParInativo
End Property
Public Function Consultar(ParCodigo) As Boolean
Dim sSql As String
sSql = "SELECT APL_CODIGO, APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO " _
    & " FROM APLICACOES WHERE APL_CODIGO = '" & ParCodigo & "'"

If Not ReadQuery(sSql).EOF Then
    With ReadQuery(sSql)
        fCodigo = VariableAdjust(!APL_CODIGO, eDoubleNumber)
        fCorretora = VariableAdjust(!APL_CORRETORA, eDoubleNumber)
        fNome = VariableAdjust(!APL_NOME, eStringText)
        fAno = VariableAdjust(!APL_ANO, eDoubleNumber)
        fCadastro = VariableAdjust(!APL_CADASTRO, eStringText)
        fAtualizacao = VariableAdjust(!APL_ATUALIZACAO, eStringText)
        fInvestir = VariableAdjust(!APL_INVESTIR, eLogicSql)
        fInativo = VariableAdjust(!APL_INATIVO, eLogicSql)
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

sSql = "INSERT INTO APLICACOES (APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO) " _
    & " VALUES ('" & fNome & "','" & fCorretora & "','" & fAno & "','" & fInvestir & "',GETDATE(), GETDATE(), '" & fInativo & "')"
If Not QueryExecute(sSql) Then GoTo Trata

fCodigo = FieldCollect("SELECT MAX(APL_CODIGO) FROM APLICACOES (NOLOCK)", eIntNumber)

Inserir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Inserir", sSql
End Function

Public Function Atualizar() As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "UPDATE APLICACOES SET APL_NOME = '" & fNome & "', APL_CORRETORA = '" & fCorretora & "', APL_ANO = '" & fAno & "' ,APL_INVESTIR = '" _
    & fInvestir & "', APL_ATUALIZACAO = GETDATE(), APL_INATIVO = '" & fInativo _
    & "' WHERE APL_CODIGO = '" & fCodigo & "'"
If Not QueryExecute(sSql) Then GoTo Trata



Atualizar = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Atualizar", sSql
End Function

Public Function Excluir(ParCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "UPDATE APLICACOES SET APL_INATIVO = '1' WHERE APL_CODIGO = '" & ParCodigo & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not QueryExecute(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Excluir", sSql
End Function
