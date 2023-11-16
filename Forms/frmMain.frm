VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "GoInvest"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10800
      Left            =   0
      TabIndex        =   2
      Top             =   -100
      Width           =   1280
      Begin fpBtnAtlLibCtl.fpBtn cmdTrocarUsuario 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   9960
         Width           =   1260
         _Version        =   131072
         _ExtentX        =   2222
         _ExtentY        =   1085
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
         ButtonDesigner  =   "frmMain.frx":19D2E
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdUsuario 
         Height          =   615
         Left            =   0
         TabIndex        =   0
         Top             =   9360
         Width           =   1260
         _Version        =   131072
         _ExtentX        =   2222
         _ExtentY        =   1085
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
         DrawFocusRect   =   1
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   0
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
         ButtonDesigner  =   "frmMain.frx":1B002
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim eForms As Dictionary
Dim fClsRedimenciona As clsRedimenciona

Private Sub cmdAplicacoes_Click()
Dim sWidth As Long

eForms.Add IIf(eForms.Exists(eForms.Count), eForms.Count + 1, eForms.Count), frmApplication

sWidth = frmMain.Width - FrameBarra.Width
SetForm frmApplication, frmMain
frmApplication.WindowState = 0
Centraliza frmMain, frmApplication, sWidth

frmApplication.Show
End Sub

Public Function FormRemove(parForm As Form) As Boolean
Dim sCont As Long

For sCont = 0 To eForms.Count
    If eForms.Exists(sCont) Then If eForms.Item(sCont).Name = parForm.Name Then eForms.Remove sCont
Next

FormRemove = True
End Function

Private Sub cmdCorretoras_Click()
Dim sWidth As Long

eForms.Add IIf(eForms.Exists(eForms.Count), eForms.Count + 1, eForms.Count), frmBank

sWidth = frmMain.Width - FrameBarra.Width
SetForm frmBank, frmMain
frmBank.WindowState = 0
Centraliza frmMain, frmBank, sWidth
frmBank.Show
End Sub

Private Sub cmdTrocarUsuario_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub cmdUsuario_Click()
Dim sWidth As Long, sId As Long

sId = IIf(eForms.Exists(eForms.Count), eForms.Count + 1, eForms.Count)
eForms.Add sId, frmUser

sWidth = frmMain.Width - FrameBarra.Width
SetForm frmUser, frmMain
frmUser.WindowState = 0
Centraliza frmMain, frmUser, sWidth, FrameBarra.Width / 2
frmUser.Show
End Sub

Private Sub Form_Load()
ConfigurarForm Me
Set eForms = New Dictionary
Set fClsRedimenciona = New clsRedimenciona

fClsRedimenciona.IniciarRedimencionamento Me

Me.Caption = Me.Caption & FillFooter
End Sub

Private Sub Form_Resize()
Dim sWidth As Long
sWidth = frmMain.Width - FrameBarra.Width


fClsRedimenciona.Redimencionar Me

If eForms.Count > 0 Then Centraliza frmMain, eForms.Item(eForms.Count - 1), sWidth, FrameBarra.Width / 2
End Sub

Private Sub fpBtn_Click()
Dim sWidth As Long, sId As Long

sId = IIf(eForms.Exists(eForms.Count), eForms.Count + 1, eForms.Count)
eForms.Add sId, frmUser

sWidth = frmMain.Width - FrameBarra.Width
SetForm frmAnalyze, frmMain
frmAnalyze.WindowState = 0
Centraliza frmMain, frmAnalyze, sWidth
frmAnalyze.Show
End Sub
