VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fRegistro As Integer
Dim fValor As Double
Dim fMes As String
Dim fTaxa As Double
Dim fSaque As Double
Dim fRendimento As Double
Dim fSaldo As Double
Property Get Registro() As Integer
Registro = fRegistro
End Property
Property Let Registro(ByVal ParRegistro As Integer)
fRegistro = ParRegistro
End Property
Property Get Valor() As Double
Valor = fValor
End Property
Property Let Valor(ByVal ParValor As Double)
fValor = ParValor
End Property
Property Get Mes() As String
Mes = fMes
End Property
Property Let Mes(ByVal ParMes As String)
fMes = ParMes
End Property
Property Get Taxa() As Double
Taxa = fTaxa
End Property
Property Let Taxa(ByVal ParTaxa As Double)
fTaxa = ParTaxa
End Property
Property Get Saque() As Double
Saque = fSaque
End Property
Property Let Saque(ByVal ParSaque As Double)
fSaque = ParSaque
End Property
Property Get Rendimento() As Double
Rendimento = fRendimento
End Property
Property Let Rendimento(ByVal ParRendimento As Double)
fRendimento = ParRendimento
End Property
Property Get Saldo() As Double
Saldo = fSaldo
End Property
Property Let Saldo(ByVal ParSaldo As Double)
fSaldo = ParSaldo
End Property
Public Function Consultar(ParCodigo) As Boolean
Dim sSql As String
Dim sConsulta As ADODB.Recordset

sSql = "SELECT EXT_VALOR, EXT_MES, EXT_TAXA, EXT_SAQUE, EXT_RENDIMENTO, EXT_SALDO FROM EXTRATO WHERE EXT_REGISTRO = '" & ParCodigo & "'"

If Not Consulta(sSql).EOF Then
    Set sConsulta = Consulta(sSql)
    With sConsulta
        fValor = TratarVariavel(EXT_VALOR, "N")
        fMes = TratarVariavel(EXT_VALOR, "T")
        fTaxa = TratarVariavel(EXT_VALOR, "N")
        fSaque = TratarVariavel(EXT_VALOR, "N")
        fRendimento = TratarVariavel(EXT_VALOR, "N")
        fSaldo = TratarVariavel(EXT_VALOR, "N")
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
    sSql = "INSERT INTO CORRETORAS (COR_NOME, COR_CADASTRO, COR_ATUALIZACAO, COR_INATIVO) VALUES ('" & fNome & "', GETDATE(), GETDATE(), '" & fInativo & "')"
    If Not ExecutarSql(sSql) Then GoTo Trata
Else
    sSql = "UPDATE CORRETORAS SET COR_NOME = '" & fNome & "', COR_ATUALIZACAO = GETDATE(), COR_INATIVO = '" & fInativo _
        & "' WHERE COR_CODIGO = '" & fCodigo & "'"
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

sSql = "UPDATE CORRETORAS SET COR_INATIVO = '1' WHERE COR_CODIGO = '" & ParCodigo & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Excluir"
End Function
