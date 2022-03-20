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
   Begin VB.TextBox txtLinha 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
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
Open "c:\teste\teste.txt" For Input As #1 'Abre o arquivo para entrada.
 Do While Not EOF(1) 'Faz o loop até o fim do arquivo.
 Input #1, A
 Texto = Texto & A  'Concatena a variável Texto com a última linha lida
Loop
txtLinha.Text = Texto 'Joga o conteúdo da variável Texto para o TextBox
Close #1 ' Fecha o arquivo.

End Sub
