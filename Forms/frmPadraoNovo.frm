VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmPadraoNovo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmPadraoNovo"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   405
   ClientWidth     =   19125
   Icon            =   "frmPadraoNovo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   19125
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame quadBotoes 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   10570
      Left            =   17670
      TabIndex        =   19
      Top             =   -90
      Width           =   1455
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   6360
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":680A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":7AD5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":8DA1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   4200
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":A06F
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   5160
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":B33A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   4440
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":C605
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOpcao 
         Height          =   855
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   5400
         Width           =   975
         _Version        =   131072
         _ExtentX        =   1720
         _ExtentY        =   1508
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
         ButtonDesigner  =   "frmPadraoNovo.frx":D8D1
      End
   End
   Begin VB.Frame quadPesquisa 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10575
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   19095
      Begin VB.CheckBox chkInativoPesquisa 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   16440
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
         ButtonDesigner  =   "frmPadraoNovo.frx":EB9D
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
         Width           =   17385
         _Version        =   458752
         _ExtentX        =   30665
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
         SpreadDesigner  =   "frmPadraoNovo.frx":FE74
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
      Height          =   10575
      Left            =   0
      TabIndex        =   12
      Top             =   -120
      Width           =   19095
      Begin VB.CheckBox SeePassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DownPicture     =   "frmPadraoNovo.frx":10359
         DragIcon        =   "frmPadraoNovo.frx":10DCB
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8640
         Picture         =   "frmPadraoNovo.frx":1183D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   860
         Width           =   400
      End
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
         Left            =   13920
         TabIndex        =   15
         Top             =   9000
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkInativo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9360
         TabIndex        =   8
         Top             =   780
         Width           =   1335
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
         TabIndex        =   18
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
Attribute VB_Name = "frmPadraoNovo"
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

If Index = EnumOpcao.eIncluir Then
    fOpcao = Index
    DefinirTela True
    chkInativo.Visible = False
    quadDatas.Visible = False
    txtCodigo = "NOVO"
ElseIf Index = EnumOpcao.eCosultar Or Index = EnumOpcao.eAlterar Or Index = EnumOpcao.eExcluir Then
    gridPrincipal_Click gridPrincipal.ActiveCol, gridPrincipal.ActiveRow
    fOpcao = Index
    DefinirTela True
    If Not ObterDados Then GoTo ErrorHandler
    If fOpcao = EnumOpcao.eExcluir Then cmdOpcao_Click (EnumOpcao.eConfirmar)
ElseIf Index = EnumOpcao.eConfirmar Then
    If fOpcao = EnumOpcao.eIncluir Or fOpcao = EnumOpcao.eAlterar Then
        If Not TransferirDados Then Exit Sub
        DefinirTela False
        ExpurgarDados
    ElseIf fOpcao = EnumOpcao.eExcluir Then
        If Not fClsUsuarios.Excluir(fCodigo) Then GoTo ErrorHandler
        AlimentarGrid
        DefinirTela False
        ExpurgarDados
    Else
        DefinirTela False
        ExpurgarDados
    End If
ElseIf Index = EnumOpcao.eCancelar Then
    DefinirTela False
    ExpurgarDados
ElseIf Index = EnumOpcao.eSair Then
    frmPrincipal.RemoverForm Me
    Unload Me
End If

Exit Sub
Resume
ErrorHandler:
ErrorHandler Err.Number, Err.Description, "frmUser.cmdOpcao_Click", ""
End Sub

Private Sub cmdPesquisa_Click()
AlimentarGrid
End Sub

Private Sub Form_Load()
cmdOpcao(EnumOpcao.eConfirmar).Visible = False
cmdOpcao(EnumOpcao.eCancelar).Visible = False
quadCadastro.Visible = False
quadPesquisa.Visible = True
AlimentarGrid
Me.Caption = Me.Caption & FillFooter
Me.Move ((frmPrincipal.Height - Me.Height) \ 2) + 1100
End Sub

Private Sub DefinirTela(ParCadastro As Boolean)
If ParCadastro = True Then
    quadPesquisa.Visible = False
    quadCadastro.Visible = True
    cmdOpcao(EnumOpcao.eIncluir).Visible = False
    cmdOpcao(EnumOpcao.eCosultar).Visible = False
    cmdOpcao(EnumOpcao.eAlterar).Visible = False
    cmdOpcao(EnumOpcao.eExcluir).Visible = False
    cmdOpcao(EnumOpcao.eSair).Visible = False
    cmdOpcao(EnumOpcao.eConfirmar).Visible = True
    cmdOpcao(EnumOpcao.eCancelar).Visible = True
    txtNome.SetFocus
Else
    quadPesquisa.Visible = True
    quadCadastro.Visible = False
    cmdOpcao(EnumOpcao.eIncluir).Visible = True
    cmdOpcao(EnumOpcao.eCosultar).Visible = True
    cmdOpcao(EnumOpcao.eAlterar).Visible = True
    cmdOpcao(EnumOpcao.eExcluir).Visible = True
    cmdOpcao(EnumOpcao.eSair).Visible = True
    cmdOpcao(EnumOpcao.eConfirmar).Visible = False
    cmdOpcao(EnumOpcao.eCancelar).Visible = False
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

If Not AnalisarDados Then Exit Function

If fCodigo <> 0 Then fClsUsuarios.Consultar (fCodigo)
With fClsUsuarios
    .Codigo = fCodigo
    .Nome = Trim(txtNome)
    .Senha = txtSenha
    .Inativo = IIf(chkInativo.value, 1, 0)
    If fOpcao = EnumOpcao.eIncluir Then If Not .Inserir Then GoTo Trata
    If fOpcao = EnumOpcao.eAlterar Then If Not .Atualizar Then GoTo Trata
End With


AlimentarGrid

TransferirDados = True
Exit Function
Resume
Trata:
ErrorHandler Err.Number, Err.Description, "frmUser.TransferirDados", sSql
End Function

Private Function AnalisarDados() As Boolean

If txtNome.Text = Empty Then
    MsgBox "Por favor informe o nome do usuário.", vbInformation, "GoInvest"
    txtNome.SetFocus
    Exit Function
End If

If txtSenha.Text = Empty Then
    MsgBox "Por favor informe uma senha para o usuário.", vbInformation, "GoInvest"
    txtSenha.SetFocus
    Exit Function
End If

AnalisarDados = True
End Function

Private Sub ExpurgarDados()
txtCodigo.Text = Empty
txtNome.Text = Empty
txtSenha.Text = Empty
chkInativo.value = 0
SeePassword.value = 0
txtData.Text = Empty
txtAtualizacao.Text = Empty
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
cmdOpcao_Click (EnumOpcao.eAlterar)
End Sub

Private Sub SeePassword_Click()
If SeePassword Then
    txtSenha.PasswordChar = ""
Else
    txtSenha.PasswordChar = "*"
End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
