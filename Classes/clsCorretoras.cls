VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fCodigo As Integer
Dim fNome As String
Dim fCadastro As Date
Dim fAtualizacao As Date
Dim eUser As Integer
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

Property Get User() As Integer
User = eUser
End Property
Property Let User(ByVal User_ As Integer)
eUser = User_
End Property

Public Function Consultar(ParCodigo) As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = "SELECT BAN_ID, BAN_NAME, BAN_REGISTER, BAN_UPDATE, BAN_TRASH FROM BANK WHERE BAN_ID = '" & ParCodigo & "'"

If Not ReadQuery(iQuery).EOF Then
    With ReadQuery(iQuery)
        fCodigo = !BAN_ID
        fNome = IIf(Null, "", !BAN_NAME)
        fCadastro = !BAN_REGISTER
        fAtualizacao = !BAN_UPDATE
        fInativo = IIf(!BAN_TRASH, 1, 0)
    End With
Else
    MsgBox "Código não localizado"
    Exit Function
End If

Consultar = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsBank.Get", iQuery
End Function

Public Function Inserir() As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = "INSERT INTO BANK (BAN_NAME, BAN_REGISTER, BAN_UPDATE, BAN_TRASH, BAN_USER) VALUES ('" & fNome _
    & "', GETDATE(), GETDATE(), '" & fInativo & "', " & gClsUser.Codigo & ")"
If Not QueryExecute(iQuery) Then GoTo Trata

Inserir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsCorretoras.Inserir", iQuery
End Function

Public Function Atualizar() As Boolean
On Error GoTo Trata
Dim iQuery As String


iQuery = "UPDATE BANK SET BAN_NAME = '" & fNome & "', BAN_UPDATE = GETDATE(), BAN_TRASH = '" & fInativo & "', BAN_USER = " & gClsUser.Codigo _
    & " WHERE BAN_ID = '" & fCodigo & "'"
If Not QueryExecute(iQuery) Then GoTo Trata

Atualizar = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsCorretoras.Atualizar", iQuery
End Function

Public Function Excluir(Id_ As Integer) As Boolean
On Error GoTo Trata
Dim iQuery As String

iQuery = "UPDATE CORRETORAS SET COR_INATIVO = '1' WHERE COR_CODIGO = '" & Id_ & "'"
If MsgBox("Deseja excluir essa corretora?", vbYesNo, "Corretoras") = 6 Then
    If Not QueryExecute(iQuery) Then GoTo Trata
End If

Excluir = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "clsCorretoras.Excluir", iQuery
End Function
