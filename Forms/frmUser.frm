VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmUser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Usu�rios"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   19125
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   19125
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame quadBotoes 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   10800
      Left            =   17760
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
         ButtonDesigner  =   "frmUser.frx":680A
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
         ButtonDesigner  =   "frmUser.frx":7AD5
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
         ButtonDesigner  =   "frmUser.frx":8DA1
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
         ButtonDesigner  =   "frmUser.frx":A06F
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
         ButtonDesigner  =   "frmUser.frx":B33A
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
         ButtonDesigner  =   "frmUser.frx":C605
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
         ButtonDesigner  =   "frmUser.frx":D8D1
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
         SpreadDesigner  =   "frmUser.frx":FEAC
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
         DownPicture     =   "frmUser.frx":103B3
         DragIcon        =   "frmUser.frx":10E25
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8640
         Picture         =   "frmUser.frx":11897
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
            Caption         =   "Atualiza��o"
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
         Caption         =   "C�digo"
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

If Index = EnumOption.eInclude Then
    fOpcao = Index
    DefinirTela True
    chkInativo.Visible = False
    quadDatas.Visible = False
    txtCodigo = "NOVO"
ElseIf Index = EnumOption.eRead Or Index = EnumOption.Update Or Index = EnumOption.eDelete Then
    gridPrincipal_Click gridPrincipal.ActiveCol, gridPrincipal.ActiveRow
    fOpcao = Index
    DefinirTela True
    If Not ObterDados Then GoTo ErrorHandler
    If fOpcao = EnumOption.eDelete Then cmdOpcao_Click (EnumOption.eConfirm)
ElseIf Index = EnumOption.eConfirm Then
    If fOpcao = EnumOption.eInclude Or fOpcao = EnumOption.Update Then
        If Not TransferirDados Then Exit Sub
        DefinirTela False
        ExpurgarDados
    ElseIf fOpcao = EnumOption.eDelete Then
        If Not fClsUsuarios.Excluir(fCodigo) Then GoTo ErrorHandler
        AlimentarGrid
        DefinirTela False
        ExpurgarDados
    Else
        DefinirTela False
        ExpurgarDados
    End If
ElseIf Index = EnumOption.eCancel Then
    DefinirTela False
    ExpurgarDados
ElseIf Index = EnumOption.eLeave Then
    frmMain.FormRemove Me
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
cmdOpcao(EnumOption.eConfirm).Visible = False
cmdOpcao(EnumOption.eCancel).Visible = False
quadCadastro.Visible = False
quadPesquisa.Visible = True
AlimentarGrid
Me.Caption = Me.Caption & FillFooter
Me.Move ((frmMain.Height - Me.Height) \ 2) + 1100
End Sub

Private Sub DefinirTela(ParCadastro As Boolean)
If ParCadastro = True Then
    quadPesquisa.Visible = False
    quadCadastro.Visible = True
    cmdOpcao(EnumOption.eInclude).Visible = False
    cmdOpcao(EnumOption.eRead).Visible = False
    cmdOpcao(EnumOption.Update).Visible = False
    cmdOpcao(EnumOption.eDelete).Visible = False
    cmdOpcao(EnumOption.eLeave).Visible = False
    cmdOpcao(EnumOption.eConfirm).Visible = True
    cmdOpcao(EnumOption.eCancel).Visible = True
    txtNome.SetFocus
Else
    quadPesquisa.Visible = True
    quadCadastro.Visible = False
    cmdOpcao(EnumOption.eInclude).Visible = True
    cmdOpcao(EnumOption.eRead).Visible = True
    cmdOpcao(EnumOption.Update).Visible = True
    cmdOpcao(EnumOption.eDelete).Visible = True
    cmdOpcao(EnumOption.eLeave).Visible = True
    cmdOpcao(EnumOption.eConfirm).Visible = False
    cmdOpcao(EnumOption.eCancel).Visible = False
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
    If fOpcao = EnumOption.eInclude Then If Not .Inserir Then GoTo Trata
    If fOpcao = EnumOption.Update Then If Not .Atualizar Then GoTo Trata
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
    MsgBox "Por favor informe o nome do usu�rio.", vbInformation, "GoInvest"
    txtNome.SetFocus
    Exit Function
End If

If txtSenha.Text = Empty Then
    MsgBox "Por favor informe uma senha para o usu�rio.", vbInformation, "GoInvest"
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
cmdOpcao_Click (EnumOption.Update)
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
