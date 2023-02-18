VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   7755
   ClientLeft      =   6075
   ClientTop       =   1815
   ClientWidth     =   6570
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":680A
   ScaleHeight     =   7755
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1990
      Picture         =   "frmLogin.frx":E465
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   5460
      Width           =   495
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2000
      Picture         =   "frmLogin.frx":286BD
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   4280
      Width           =   495
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLogin 
      Height          =   615
      Left            =   2440
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   1085
      Enabled         =   -1  'True
      MouseIcon       =   "frmLogin.frx":30F8F
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
      ButtonDesigner  =   "frmLogin.frx":39871
   End
   Begin EditLib.fpText txtUsuario 
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   1085
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0,25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtSenha 
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   5400
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   1085
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   "*"
      IncHoriz        =   0,25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLogin_Click()
    If txtUsuario = "" Or Not VerificarUsuario(, True) Then
        MsgBox "Senha incorreta!!", vbInformation, "GoInvest"
        txtSenha.SetFocus
        txtSenha.Text = ""
    Else
        gUser = txtUsuario
        frmMain.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
gVersion = "26/03/2022"
If Not ReadConfig Then
    MsgBox "O banco de dados não esta configurado!", vbExclamation
    Unload Me
End If
End Sub

Private Sub txtSenha_GotFocus()
    txtUsuario_KeyPress (13)
End Sub

Private Function VerificarUsuario(Optional ByVal VerificaNome As Boolean, Optional ByVal VerificaSenha As Boolean) As Boolean
Dim sSql As String, sLinhas As Long, sRetorno As Boolean

If VerificaNome Then
    sSql = "SELECT USU_NOME FROM USUARIOS WHERE USU_NOME = '" & txtUsuario.Text & "'"
    ReadQuery sSql, sLinhas
    
    If sLinhas <> 0 Then
        sRetorno = True
    End If
ElseIf VerificaSenha Then
    sSql = "SELECT USU_NOME FROM USUARIOS WHERE USU_NOME = '" & txtUsuario.Text & "' AND USU_SENHA = '" & txtSenha.Text & "'"
    ReadQuery sSql, sLinhas
    
    If sLinhas <> 0 Then
        sRetorno = True
    End If
End If

VerificarUsuario = sRetorno
End Function

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    cmdLogin_Click
 End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 13 Then
    If txtUsuario = "" Or Not VerificarUsuario(True) Then
        MsgBox "Usuário não localizado!!", vbInformation, "GoInvest"
        txtUsuario.SetFocus
        txtUsuario.Text = ""
    Else
        txtSenha.SetFocus
    End If
End If
End Sub
