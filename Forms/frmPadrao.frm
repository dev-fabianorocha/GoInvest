VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmPadrao 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form Padrão"
   ClientHeight    =   8175
   ClientLeft      =   2280
   ClientTop       =   2625
   ClientWidth     =   12360
   Icon            =   "frmPadrao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame quadBotoes 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   8280
      Left            =   11400
      TabIndex        =   9
      Top             =   0
      Width           =   975
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":680A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   4680
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":7ADA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   915
         _Version        =   131072
         _ExtentX        =   1614
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":8DA9
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":A07B
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
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
         BackStyle       =   1
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
         ButtonDesigner  =   "frmPadrao.frx":B34B
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":C61B
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdB 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   675
         _Version        =   131072
         _ExtentX        =   1191
         _ExtentY        =   1296
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   -2147483627
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
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   -2147483627
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmPadrao.frx":D8EA
      End
   End
   Begin VB.Frame quadRodape 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   7800
      Width           =   11895
      Begin VB.Label lblRodape 
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
         Left            =   0
         TabIndex        =   25
         Top             =   80
         Width           =   11895
      End
   End
   Begin VB.Frame quadPesquisa 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
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
         Left            =   9720
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPesquisa 
         Height          =   495
         Left            =   4320
         TabIndex        =   22
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
         ButtonDesigner  =   "frmPadrao.frx":EBB9
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
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin FPSpreadADO.fpSpread gridPrincipal 
         Height          =   6285
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   10545
         _Version        =   458752
         _ExtentX        =   18600
         _ExtentY        =   11086
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
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmPadrao.frx":FE90
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
         TabIndex        =   23
         Top             =   200
         Width           =   1335
      End
   End
   Begin VB.Frame quadCadastro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      Begin VB.Frame quadDatas 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   7560
         TabIndex        =   17
         Top             =   6600
         Width           =   3255
         Begin VB.TextBox txtData 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAtualizacao 
            Height          =   375
            Left            =   1680
            TabIndex        =   18
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
            TabIndex        =   21
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
            TabIndex        =   20
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
         Left            =   6360
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtCodigo 
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
         TabIndex        =   4
         Top             =   840
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fAcao As Integer
Dim fClsCorretoras As New clsCorretoras
Dim fCodigo As Integer
Dim fCondicao As String
Private Enum EnumGrid
    eCodigo = 1
    eNome
    eCadastro
    eAtualizacao
    eStatus
End Enum

Private Sub EncherGrid()
Dim sSql As String

sSql = "SELECT COR_CODIGO AS CODIGO, COR_NOME AS NOME, COR_CADASTRO AS CADASTRO, COR_ATUALIZACAO AS ATUALIZACAO, CASE WHEN COR_INATIVO = 0 THEN 'ATIVO'" _
    & "WHEN COR_INATIVO = 1 THEN 'INATIVO' END AS STATUS FROM CORRETORAS WHERE 1 = 1"

If chkInativoPesquisa Then
    sSql = sSql & " AND COR_INATIVO = 1"
Else
    sSql = sSql & " AND COR_INATIVO = 0"
End If

If txtPesquisa.Text <> "" Then
    sSql = sSql & " AND COR_NOME LIKE '" & Trim(txtPesquisa.Text) & "%'"
End If

LimparGrid gridPrincipal
PopularGrid gridPrincipal, sSql

Exit Sub
End Sub

Private Sub cmdB_Click(Index As Integer)
On Error GoTo Trata

    If Index = enumAcao.eIncluir Then
        fAcao = Index
        TrocarTela True
        txtCodigo = "NOVO"
    ElseIf Index = enumAcao.eConsultar Or Index = enumAcao.eAlterar Or Index = enumAcao.eExcluir Then
        gridPrincipal_Click gridPrincipal.ActiveCol, gridPrincipal.ActiveRow
        fAcao = Index
        TrocarTela True
        If Not ReceberDados Then GoTo Trata
        If fAcao = enumAcao.eExcluir Then cmdB_Click (enumAcao.eConfirmar)
    ElseIf Index = enumAcao.eConfirmar Then
        If fAcao = enumAcao.eIncluir Or fAcao = enumAcao.eAlterar Then
            PassarDados
            TrocarTela False
            LimparTela
        ElseIf fAcao = enumAcao.eExcluir Then
            If Not fClsCorretoras.Excluir(fCodigo) Then GoTo Trata
            EncherGrid
            TrocarTela False
            LimparTela
        Else
            TrocarTela False
            LimparTela
        End If
    ElseIf Index = enumAcao.eCancelar Then
        TrocarTela False
        LimparTela
    ElseIf Index = enumAcao.eSair Then
        Unload Me
    End If
Exit Sub

Exit Sub
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description), vbCritical, "clsCorretoras.Atualizar"
End Sub

Private Sub cmdPesquisa_Click()
EncherGrid
End Sub

Private Sub Form_Load()
cmdB(enumAcao.eConfirmar).Visible = False
cmdB(enumAcao.eCancelar).Visible = False
quadCadastro.Visible = False
quadPesquisa.Visible = True
EncherGrid
lblRodape = AlimentarRodape
End Sub

Private Sub TrocarTela(ParCadastro As Boolean)
If ParCadastro = True Then
    quadPesquisa.Visible = False
    quadCadastro.Visible = True
    cmdB(enumAcao.eIncluir).Visible = False
    cmdB(enumAcao.eConsultar).Visible = False
    cmdB(enumAcao.eAlterar).Visible = False
    cmdB(enumAcao.eExcluir).Visible = False
    cmdB(enumAcao.eSair).Visible = False
    cmdB(enumAcao.eConfirmar).Visible = True
    cmdB(enumAcao.eCancelar).Visible = True
Else
    quadPesquisa.Visible = True
    quadCadastro.Visible = False
    cmdB(enumAcao.eIncluir).Visible = True
    cmdB(enumAcao.eConsultar).Visible = True
    cmdB(enumAcao.eAlterar).Visible = True
    cmdB(enumAcao.eExcluir).Visible = True
    cmdB(enumAcao.eSair).Visible = True
    cmdB(enumAcao.eConfirmar).Visible = False
    cmdB(enumAcao.eCancelar).Visible = False
End If

End Sub

Private Function ReceberDados() As Boolean
If fClsCorretoras.Consultar(fCodigo) Then
    With fClsCorretoras
        txtNome = .Nome
        txtData = .Cadastro
        txtAtualizacao = .Atualizacao
        chkInativo = .Inativo
    End With
End If
txtCodigo = fCodigo
ReceberDados = True
End Function

Private Function PassarDados() As Boolean
On Error GoTo Trata

Dim sSql As String, sCont As Long

If fCodigo <> 0 Then fClsCorretoras.Consultar (fCodigo)
With fClsCorretoras
    .Codigo = fCodigo
    .Nome = txtNome
    .Inativo = IIf(chkInativo.value, 1, 0)
    If Not .Atualizar(fAcao) Then GoTo Trata
End With


EncherGrid

PassarDados = True
Exit Function
Resume
Trata:
MsgBox DescError(Err.Number, Err.Description, sSql), vbCritical, "clsCorretoras.Atualizar"
End Function

Private Sub LimparTela()

txtCodigo = ""
txtNome = ""
chkInativo.value = 0
txtData = ""
txtAtualizacao = ""
Set fClsCorretoras = Nothing

End Sub

Private Sub gridPrincipal_Click(ByVal col As Long, ByVal Row As Long)
MarcarLinha gridPrincipal, Row, fCodigo
End Sub

Private Sub gridPrincipal_DblClick(ByVal col As Long, ByVal Row As Long)
cmdB_Click (enumAcao.eAlterar)
End Sub

