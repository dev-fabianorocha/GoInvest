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
Dim sConsulta As ADODB.Recordset

sSql = "SELECT APL_CODIGO, APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO " _
    & " FROM APLICACOES WHERE APL_CODIGO = '" & ParCodigo & "'"

If Not Consulta(sSql).EOF Then
    Set sConsulta = Consulta(sSql)
    With sConsulta
        fCorretora = TratarVariavel(!APL_CORRETORA, "N")
        fNome = TratarVariavel(!APL_NOME, "T")
        fAno = TratarVariavel(!APL_ANO, "N")
        fCadastro = TratarVariavel(!APL_CADASTRO, "T")
        fAtualizacao = TratarVariavel(!APL_ATUALIZACAO, "T")
        fInvestir = TratarVariavel(!APL_INVESTIR, "B")
        fInativo = TratarVariavel(!APL_INATIVO, "B")
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If

Consultar = True
End Function

Public Function Atualizar(ParAcao) As Boolean
On Error GoTo Trata
Dim sSql As String

If ParAcao = enumAcao.eIncluir Then
    sSql = "INSERT INTO APLICACOES (APL_NOME, APL_CORRETORA, APL_ANO, APL_INVESTIR, APL_CADASTRO, APL_ATUALIZACAO, APL_INATIVO) " _
        & " VALUES ('" & fNome & "','" & fCorretora & "','" & fAno & "','" & fInvestir & "',GETDATE(), GETDATE(), '" & fInativo & "')"
    If Not ExecutarSql(sSql) Then GoTo Trata
Else
    sSql = "UPDATE APLICACOES SET APL_NOME = '" & fNome & "', APL_CORRETORA = '" & fCorretora & "', APL_ANO = '" & fAno & "' ,APL_INVESTIR = '" _
        & fInvestir & "', APL_ATUALIZACAO = GETDATE(), APL_INATIVO = '" & fInativo _
        & "' WHERE APL_CODIGO = '" & fCodigo & "'"
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

Atualizar = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Atualizar"
End Function

Public Function Excluir(ParCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "UPDATE APLICACOES SET APL_INATIVO = '1' WHERE APL_CODIGO = '" & ParCodigo & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Excluir"
End Function
