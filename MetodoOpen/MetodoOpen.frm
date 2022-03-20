VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtLinha 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim A As String
Dim Texto As String
Open "C:\teste\teste.txt" For Input As #1
   Do While Not EOF(1)
      Input #1, A
      Texto = Texto & A
   Loop
TxtLinha.Text = Texto
Close #1 '

End Sub
