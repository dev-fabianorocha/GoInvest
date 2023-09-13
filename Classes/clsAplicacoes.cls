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

Dim eCodigo As Integer
Dim eNome As String
Dim eAno As String
Dim eCorretora As Integer
Dim eCadastro As Date
Dim eAtualizacao As Date
Dim eInvestir As Byte
Dim eInativo As Byte
Property Get Codigo() As Integer
Codigo = eCodigo
End Property
Property Let Codigo(ByVal ParCodigo As Integer)
eCodigo = ParCodigo
End Property
Property Get Corretora() As Integer
Corretora = eCorretora
End Property
Property Let Corretora(ByVal ParCorretora As Integer)
eCorretora = ParCorretora
End Property
Property Get Nome() As String
Nome = eNome
End Property
Property Let Nome(ByVal ParNome As String)
eNome = ParNome
End Property
Property Get Ano() As String
Ano = eAno
End Property
Property Let Ano(ByVal ParAno As String)
eAno = ParAno
End Property
Property Get Cadastro() As Date
Cadastro = eCadastro
End Property
Property Let Cadastro(ByVal ParCadastro As Date)
eCadastro = ParCadastro
End Property
Property Get Atualizacao() As Date
Atualizacao = eAtualizacao
End Property
Property Let Atualizacao(ByVal ParAtualizacao As Date)
eAtualizacao = ParAtualizacao
End Property
Property Get Investir() As Byte
Investir = eInvestir
End Property
Property Let Investir(ByVal ParInvestir As Byte)
eInvestir = ParInvestir
End Property
Property Get Inativo() As Byte
Inativo = eInativo
End Property
Property Let Inativo(ByVal ParInativo As Byte)
eInativo = ParInativo
End Property
Public Function Read(Codigo_) As Boolean
Dim iQuery As String
iQuery = "SELECT APL_CODIGO, APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO " _
    & " FROM APLICACOES WHERE APL_CODIGO = '" & Codigo_ & "'"

If Not ReadQuery(iQuery).EOF Then
    With ReadQuery(iQuery)
        eCodigo = VariableAdjust(!APL_CODIGO, eDoubleNumber)
        eCorretora = VariableAdjust(!APL_CORRETORA, eDoubleNumber)
        eNome = VariableAdjust(!APL_NOME, eStringText)
        eAno = VariableAdjust(!APL_ANO, eDoubleNumber)
        eCadastro = VariableAdjust(!APL_CADASTRO, eStringText)
        eAtualizacao = VariableAdjust(!APL_ATUALIZACAO, eStringText)
        eInvestir = VariableAdjust(!APL_INVESTIR, eLogicSql)
        eInativo = VariableAdjust(!APL_INATIVO, eLogicSql)
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If

Read = True
End Function

Public Function Include() As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = _
    "INSERT INTO APLICACOES (APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO) " _
    & " VALUES ('" & eNome & "','" & eCorretora & "','" & eAno & "','" & eInvestir _
    & "',GETDATE(), GETDATE(), '" & eInativo & "')"
    
If Not QueryExecute(iQuery) Then GoTo Trata

eCodigo = FieldCollect("SELECT MAX(APL_CODIGO) FROM APLICACOES (NOLOCK)", eIntNumber)

Include = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Include", iQuery
End Function

Public Function Update() As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = "UPDATE APLICACOES SET APL_NOME = '" & eNome & "', APL_CORRETORA = '" _
    & eCorretora & "', APL_ANO = '" & eAno & "' ,APL_INVESTIR = '" & eInvestir & _
    "', APL_ATUALIZACAO = GETDATE(), APL_INATIVO = '" & eInativo & _
    "' WHERE APL_CODIGO = '" & eCodigo & "'"
    
If Not QueryExecute(iQuery) Then GoTo Trata


Update = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Update", iQuery
End Function

Public Function Delete(Codigo_ As Integer) As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = "UPDATE APLICACOES SET APL_INATIVO = '1' WHERE APL_CODIGO = '" & Codigo_ & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not QueryExecute(iQuery) Then GoTo Trata
End If

Delete = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsAplicacoes.Excluir", iQuery
End Function
