VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GoInvest"
   ClientHeight    =   8205
   ClientLeft      =   2760
   ClientTop       =   2535
   ClientWidth     =   11865
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrincipal.frx":680A
   ScaleHeight     =   8205
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00404040&
      Height          =   1215
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   12015
      Begin fpBtnAtlLibCtl.fpBtn cmdCorretoras 
         Height          =   855
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
         _ExtentY        =   1508
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
         BackStyle       =   0
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
         ButtonDesigner  =   "frmPrincipal.frx":11C48
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCorretoras_Click(Index As Integer)
    frmCorretoras.Show 1
End Sub
