VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCorretoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fCodigo As Integer
Dim fNome As String
Dim fCadastro As Date
Dim fAtualizacao As Date
Dim fInativo As Byte
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
Public Function Consultar(ParCodigo) As Boolean
Dim sSql As String

sSql = "SELECT COR_CODIGO, COR_NOME, COR_CADASTRO, COR_ATUALIZACAO, COR_INATIVO FROM CORRETORAS WHERE COR_CODIGO = '" & ParCodigo & "'"

If Not ConsultarSql(sSql).EOF Then
    With ConsultarSql(sSql)
        fNome = IIf(Null, "", !COR_NOME)
        fCadastro = !COR_CADASTRO
        fAtualizacao = !COR_ATUALIZACAO
        fInativo = IIf(!COR_INATIVO, 1, 0)
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

sSql = "INSERT INTO CORRETORAS (COR_NOME, COR_CADASTRO, COR_ATUALIZACAO, COR_INATIVO) VALUES ('" & fNome _
    & "', GETDATE(), GETDATE(), '" & fInativo & "')"
If Not ExecutarSql(sSql) Then GoTo Trata

Inserir = True
Exit Function
Resume
Trata:
MsgBox ExporErro(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Inserir"
End Function

Public Function Atualizar() As Boolean
On Error GoTo Trata
Dim sSql As String


sSql = "UPDATE CORRETORAS SET COR_NOME = '" & fNome & "', COR_ATUALIZACAO = GETDATE(), COR_INATIVO = '" & fInativo _
    & "' WHERE COR_CODIGO = '" & fCodigo & "'"
If Not ExecutarSql(sSql) Then GoTo Trata

Atualizar = True
Exit Function
Resume
Trata:
MsgBox ExporErro(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Atualizar"
End Function

Public Function Excluir(ParCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "UPDATE CORRETORAS SET COR_INATIVO = '1' WHERE COR_CODIGO = '" & ParCodigo & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
MsgBox ExporErro(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Excluir"
End Function
