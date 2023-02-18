VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "GoInvest"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   20385
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8895
      Left            =   4080
      Picture         =   "frmMain.frx":680A
      ScaleHeight     =   8895
      ScaleWidth      =   13095
      TabIndex        =   7
      Top             =   720
      Width           =   13095
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   10800
      Left            =   0
      TabIndex        =   4
      Top             =   -100
      Width           =   1455
      Begin fpBtnAtlLibCtl.fpBtn fpBtn 
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
         _ExtentY        =   1931
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
         ButtonDesigner  =   "frmMain.frx":11AFA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAplicacoes 
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
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
         ButtonDesigner  =   "frmMain.frx":12DD4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdFechar 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   9360
         Width           =   1095
         _Version        =   131072
         _ExtentX        =   1931
         _ExtentY        =   1931
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   4210752
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
         DropShadowOffsetX=   0
         DropShadowOffsetY=   0
         DropShadowType  =   0
         DropShadowColor =   4210752
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmMain.frx":140A3
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdTrocarUsuario 
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   8160
         Width           =   1095
         _Version        =   131072
         _ExtentX        =   1931
         _ExtentY        =   1931
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
         ButtonDesigner  =   "frmMain.frx":15376
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdUsuario 
         Height          =   855
         Left            =   240
         TabIndex        =   0
         Top             =   240
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
         DrawFocusRect   =   1
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
         ButtonDesigner  =   "frmMain.frx":16649
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCorretoras 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
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
         ButtonDesigner  =   "frmMain.frx":17916
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicacoes_Click()
frmApplication.Show 1
End Sub

Private Sub cmdCorretoras_Click()
frmBroker.Show 1
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdTrocarUsuario_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub cmdUsuario_Click()
frmUser.Show 1
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & FillFooter
End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub fpBtn_Click()
frmAnalyze.Show 1
End Sub
