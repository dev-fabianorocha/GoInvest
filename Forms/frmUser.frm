VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmUser 
   Appearance      =   0  'Flat
   Caption         =   "Usuários"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20385
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   20385
   WindowState     =   2  'Maximized
   Begin VB.Frame quadBotoes 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   10800
      Left            =   19440
      TabIndex        =   15
      Top             =   0
      Width           =   1335
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":680A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   4680
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":7AD6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   915
         _Version        =   131072
         _ExtentX        =   1614
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":8DA1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":A06F
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":B33B
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":C607
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
         BorderShowDefault=   0   'False
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   0
         DropShadowOffsetY=   0
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":D8D2
      End
   End
   Begin VB.Frame quadRodape 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   10320
      Width           =   19455
      Begin VB.Label lblRodape 
         Caption         =   "Footer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Frame quadPesquisa 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   19455
      Begin VB.CheckBox chkInativoPesquisa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I&nativos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   17280
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPesquisa 
         Height          =   495
         Left            =   8280
         TabIndex        =   2
         Top             =   360
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   873
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   0
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUser.frx":EB9D
      End
      Begin VB.TextBox txtPesquisa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7815
      End
      Begin FPSpreadADO.fpSpread gridPrincipal 
         Height          =   8805
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   18585
         _Version        =   458752
         _ExtentX        =   32782
         _ExtentY        =   15531
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   5
         MaxRows         =   1
         OperationMode   =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmUser.frx":FE74
         UserResize      =   0
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   260
         TabIndex        =   0
         Top             =   200
         Width           =   1335
      End
   End
   Begin VB.Frame quadCadastro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   10215
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   19455
      Begin VB.TextBox txtSenha 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   6240
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame quadDatas 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   975
         Left            =   15960
         TabIndex        =   23
         Top             =   9120
         Width           =   3255
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAtualizacao 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cadastro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Atualização"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkInativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I&nativo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   760
         Width           =   1095
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fOpcao As Integer
Dim fClsUsuarios As New clsUsuarios
Dim fCodigo As Integer
Dim fCondicao As String
Private Enum EnumGrid
    eCodigo = 1
    eNome
    eCadastro
    eAtualizacao
    eStatus
End Enum

Private Sub AlimentarGrid()
Dim sSql As String

sSql = "SELECT USU_CODIGO AS CODIGO, USU_NOME AS NOME, USU_CADASTRO AS CADASTRO, USU_ATUALIZACAO AS ATUALIZACAO, CASE WHEN USU_INATIVO = 0 THEN 'ATIVO'" _
    & "WHEN USU_INATIVO = 1 THEN 'INATIVO' END AS STATUS FROM USUARIOS WHERE 1 = 1"

If chkInativoPesquisa Then
    sSql = sSql & " AND USU_INATIVO = 1"
Else
    sSql = sSql & " AND USU_INATIVO = 0"
End If

If txtPesquisa.Text <> "" Then
    sSql = sSql & " AND USU_NOME LIKE '" & Trim(txtPesquisa.Text) & "%'"
End If

SpreadClean gridPrincipal
SpreadFill gridPrincipal, sSql

Exit Sub
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
On Error GoTo ErrorHandler

    If Index = EnumOption.Include Then
        fOpcao = Index
        DefinirTela True
        chkInativo.Visible = False
        quadDatas.Visible = False
        txtCodigo = "NOVO"
    ElseIf Index = EnumOption.Read Or Index = EnumOption.Update Or Index = EnumOption.Delete Then
        gridPrincipal_Click gridPrincipal.ActiveCol, gridPrincipal.ActiveRow
        fOpcao = Index
        DefinirTela True
        If Not ObterDados Then GoTo ErrorHandler
        If fOpcao = EnumOption.Delete Then cmdOpcao_Click (EnumOption.Confirm)
    ElseIf Index = EnumOption.Confirm Then
        If fOpcao = EnumOption.Include Or fOpcao = EnumOption.Update Then
            TransferirDados
            DefinirTela False
            ExpurgarDados
        ElseIf fOpcao = EnumOption.Delete Then
            If Not fClsUsuarios.Excluir(fCodigo) Then GoTo ErrorHandler
            AlimentarGrid
            DefinirTela False
            ExpurgarDados
        Else
            DefinirTela False
            ExpurgarDados
        End If
    ElseIf Index = EnumOption.Cancel Then
        DefinirTela False
        ExpurgarDados
    ElseIf Index = EnumOption.Leave Then
        Unload Me
    End If
Exit Sub

Exit Sub
Resume
ErrorHandler:
ErrorHandler Err.Number, Err.Description, "frmUser.cmdOption_Click"
End Sub

Private Sub cmdPesquisa_Click()
AlimentarGrid
End Sub

Private Sub Form_Load()
cmdOpcao(EnumOption.Confirm).Visible = False
cmdOpcao(EnumOption.Cancel).Visible = False
quadCadastro.Visible = False
quadPesquisa.Visible = True
AlimentarGrid
lblRodape = FillFooter
End Sub

Private Sub DefinirTela(ParCadastro As Boolean)
If ParCadastro = True Then
    quadPesquisa.Visible = False
    quadCadastro.Visible = True
    cmdOpcao(EnumOption.Include).Visible = False
    cmdOpcao(EnumOption.Read).Visible = False
    cmdOpcao(EnumOption.Update).Visible = False
    cmdOpcao(EnumOption.Delete).Visible = False
    cmdOpcao(EnumOption.Leave).Visible = False
    cmdOpcao(EnumOption.Confirm).Visible = True
    cmdOpcao(EnumOption.Cancel).Visible = True
Else
    quadPesquisa.Visible = True
    quadCadastro.Visible = False
    cmdOpcao(EnumOption.Include).Visible = True
    cmdOpcao(EnumOption.Read).Visible = True
    cmdOpcao(EnumOption.Update).Visible = True
    cmdOpcao(EnumOption.Delete).Visible = True
    cmdOpcao(EnumOption.Leave).Visible = True
    cmdOpcao(EnumOption.Confirm).Visible = False
    cmdOpcao(EnumOption.Cancel).Visible = False
End If

End Sub

Private Function ObterDados() As Boolean
If fClsUsuarios.Consultar(fCodigo) Then
    With fClsUsuarios
        txtNome = .Nome
        txtSenha = .Senha
        txtData = .Cadastro
        txtAtualizacao = .Atualizacao
        chkInativo = .Inativo
    End With
End If
txtCodigo = fCodigo
ObterDados = True
End Function

Private Function TransferirDados() As Boolean
On Error GoTo Trata

Dim sSql As String, sCont As Long

If fCodigo <> 0 Then fClsUsuarios.Consultar (fCodigo)
With fClsUsuarios
    .Codigo = fCodigo
    .Nome = Trim(txtNome)
    .Senha = txtSenha
    .Inativo = IIf(chkInativo.Value, 1, 0)
    If fOpcao = EnumOption.Include Then If Not .Inserir Then GoTo Trata
    If fOpcao = EnumOption.Update Then If Not .Atualizar Then GoTo Trata
End With


AlimentarGrid

TransferirDados = True
Exit Function
Resume
Trata:
MsgBox ErrorHandler(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Atualizar"
End Function

Private Sub ExpurgarDados()

txtCodigo = ""
txtNome = ""
chkInativo.Value = 0
txtData = ""
txtAtualizacao = ""
chkInativo.Visible = True
quadDatas.Visible = True
Set fClsUsuarios = Nothing

End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub gridPrincipal_Click(ByVal col As Long, ByVal Row As Long)
SpreadGetCode gridPrincipal, Row, fCodigo
End Sub

Private Sub gridPrincipal_DblClick(ByVal col As Long, ByVal Row As Long)
cmdOpcao_Click (EnumOption.Update)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
