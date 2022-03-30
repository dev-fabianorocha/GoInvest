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
Dim fCodigoAplicacao As Integer
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
Property Get CodigoAplicacao() As Integer
Registro = fRegistro
End Property
Property Let CodigoAplicacao(ByVal ParCodigoAplicacao As Integer)
fCodigoAplicacao = ParCodigoAplicacao
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

If ParTaxa < 0 Then
    Err.Raise vbObjectError + 2001, "Conta Bancaria::TaxaInvestimento", "Definição de taxa inválida : " & ParTaxa
End If

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
Public Function ConsultarExtrato(ParCodigo) As Boolean
Dim sSql As String
Dim sConsulta As ADODB.Recordset

sSql = "SELECT EXT_VALOR, EXT_MES, EXT_TAXA, EXT_SAQUE, EXT_RENDIMENTO, EXT_SALDO FROM EXTRATO WHERE EXT_APLICACAO = '" & fCodigoAplicacao & "'"

If Not Consulta(sSql).EOF Then
    Set sConsulta = Consulta(sSql)
    With sConsulta
        fRegistro = TratarVariavel(!EXT_REGISTRO, "N")
        fValor = TratarVariavel(!EXT_VALOR, "N")
        fMes = TratarVariavel(!EXT_MES, "T")
        fTaxa = TratarVariavel(!EXT_TAXA, "N")
        fSaque = TratarVariavel(!EXT_SAQUE, "N")
        fRendimento = TratarVariavel(!EXT_RENDIMENTO, "N")
        fSaldo = TratarVariavel(!EXT_SALDO, "N")
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If

ConsultarExtrato = True
End Function

Public Function AtualizarExtrato(ParAcao) As Boolean
On Error GoTo Trata
Dim sSql As String

If ParAcao = enumAcao.eIncluir Then
    sSql = "INSERT INTO EXTRATO (EXT_VALOR, EXT_MES, EXT_TAXA, EXT_SAQUE, EXT_RENDIMENTO, EXT_SALDO) VALUES ('" & fValor & "', '" & fValor _
        & "', '" & fMes & "', '" & fTaxa & "', '" & fSaque & "', '" & fRendimento & "', '" & fSaldo & "')"
    If Not ExecutarSql(sSql) Then GoTo Trata
Else
    sSql = "UPDATE EXTRATO SET EXT_VALOR = '" & fValor & "', EXT_TAXA = '" & fTaxa & "', EXT_SAQUE = '" & fSaque _
        & "', EXT_RENDIMENTO = '" & fRendimento & "', EXT_SALDO = '" & fSaque & "' WHERE EXT_REGISTRO = '" & fRegistro & "'"
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

AtualizarExtrato = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsExtrato.AtualizarExtrato"
End Function

Public Function LimparExtrato(ParCodigoAplicacao As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "DELETE FROM EXTRATO WHERE EXT_APLICACAO = '" & fCodigoAplicacao & "'"
If MsgBox("Deseja realmente limpar essa aplicação?", vbYesNo, "Aplicacções") = 6 Then
    If Not ExecutarSql(sSql) Then GoTo Trata
End If

LimparExtrato = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsExtrato.LimparExtrato"
End Function

Public Sub Sacar(ByVal ParValorSaque As Double)

If ParValorSaque < 0 Then
    Err.Raise vbObjectError + 2002, "Conta Bancaria::Saque", "Valor de saque inválido : " & ParValorSaque
End If

If (fSaldo - ParValorSaque) < 0 Then
    Err.Raise vbObjectError + 2003, "Conta Bancaria::Saque", "Insuficiência de fundos"
End If

  fSaldo = fSaldo - ParValorSaque
End Sub

Public Sub Depositar(ByVal ParValorDeposito As Double)

If ParValorDeposito <= 0 Then
   Err.Raise vbObjectError + 2002, "Conta Bancária::Depósito", "Valor de depósito inválido : " & ParValorDeposito
End If

fSaldo = fSaldo + ParValorDeposito

End Sub

Private Sub CalculoInvestimento()

fSaldo = fSaldo + (fSaldo * (fTaxa / 100))

End Sub

Public Sub ProcessamentoMensal()

CalculoInvestimento

End Sub
