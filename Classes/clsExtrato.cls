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

Dim fRegistro As Long
Dim fCodigoAplicacao As Integer
Dim fValor As Double
Dim fMes As String
Dim fTaxa As Double
Dim fSaque As Double
Dim fRendimento As Double
Dim fSaldo As Double
Property Get Registro() As Long
Registro = fRegistro
End Property
Property Let Registro(ByVal ParRegistro As Long)
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
Public Function ConsultarExtrato(ByRef ParGrid As fpSpread, ByRef ParTextValor As fpCurrency, ByRef ParComboMes As ComboBox, ByRef ParTextTaxa As TextBox, ByRef ParTextSaque As fpCurrency, ByRef ParCmdAplicar As fpBtn) As Boolean
Dim sSql As String, sCont As Long, sLinha As Long
Dim sLinhas As Long

sSql = "SELECT EXT_REGISTRO, EXT_VALOR, EXT_MES, EXT_TAXA, EXT_SAQUE FROM EXTRATO WHERE EXT_APLICACAO = '" & fCodigoAplicacao & "' ORDER BY EXT_MES"
sCont = 1

If Not ReadQuery(sSql).EOF Then
    With ReadQuery(sSql, sLinhas)
        For sLinha = 1 To sLinhas
            ParTextValor.Text = VariableAdjust(!EXT_VALOR, eDoubleNumber)
            ParComboMes.ListIndex = VariableAdjust(!EXT_MES, eStringText)
            ParTextTaxa.Text = VariableAdjust(!EXT_TAXA, eDoubleNumber)
            ParTextSaque.Text = VariableAdjust(!EXT_SAQUE, eDoubleNumber)
            ParCmdAplicar.DoClick
            ParGrid.SetText 1, sCont, VariableAdjust(!EXT_REGISTRO, eStringText)
            ParGrid.SetText 2, sCont, MonthName(VariableAdjust(!EXT_MES, eByteNumber))
            sCont = sCont + 1
            .MoveNext
        Next
    End With
End If

ConsultarExtrato = True
End Function
Public Function AtualizarExtrato() As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "INSERT INTO EXTRATO (EXT_REGISTRO, EXT_VALOR, EXT_MES, EXT_TAXA, EXT_SAQUE, EXT_APLICACAO) VALUES ('" & fRegistro _
    & "', '" & VariableAdjust(fValor, eNumberSql) & "', '" & fMes & "', '" & VariableAdjust(fTaxa, eNumberSql) & "', '" & VariableAdjust(fSaque, eNumberSql) & "', '" _
    & fCodigoAplicacao & "')"
If Not QueryExecute(sSql) Then GoTo Trata

AtualizarExtrato = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsExtrato.AtualizarExtrato", sSql
End Function
Public Function LimparExtrato() As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "DELETE FROM EXTRATO WHERE EXT_APLICACAO = '" & fCodigoAplicacao & "'"
    If Not QueryExecute(sSql) Then GoTo Trata

LimparExtrato = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsExtrato.LimparExtrato", sSql
End Function

Public Function LimparSaldos() As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "DELETE FROM SALDOS WHERE SAL_APLICACAO = '" & fCodigoAplicacao & "'"
If Not QueryExecute(sSql) Then GoTo Trata

LimparSaldos = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsExtrato.LimparSaldos", sSql
End Function


Public Function InserirSaldos(ParMes As Byte, ParRendimento As Double, ParSaldo As Double) As Boolean
On Error GoTo Trata
Dim sSql As String

sSql = "INSERT INTO SALDOS (SAL_APLICACAO, SAL_MES, SAL_RENDIMENTO, SAL_SALDO) VALUES ('" & fCodigoAplicacao _
    & "', '" & VariableAdjust(ParMes, eDoubleNumber) & "', '" & VariableAdjust(ParRendimento, eNumberSql) & "', '" & VariableAdjust(ParSaldo, eNumberSql) & "')"
If Not QueryExecute(sSql) Then GoTo Trata

InserirSaldos = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsExtrato.InserirSaldos", sSql
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
