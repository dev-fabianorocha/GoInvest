VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAnaliseAplicacoes 
   Caption         =   "Análise de Aplicações"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   20385
   Icon            =   "frmAnaliseAplicacoes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   20385
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   10575
      Left            =   0
      OleObjectBlob   =   "frmAnaliseAplicacoes.frx":680A
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "frmAnaliseAplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim sConsulta As New Recordset, sSql As String, sLinhas As Long, sLinha As Long

sSql = "SELECT DISTINCT APL_NOME AS APLICACACAO, SUM(SAL_RENDIMENTO) AS RENDIMENTO FROM APLICACOES" _
    & " INNER JOIN SALDOS ON SAL_APLICACAO = APL_CODIGO GROUP BY  APL_NOME"
    
Set sConsulta = ReadQuery(sSql, sLinhas)

MSChart.chartType = 1
MSChart.ShowLegend = False
MSChart.Title = "Análise de Aplicacões"

MSChart.Column = 1
MSChart.RowCount = sLinhas
MSChart.Visible = True

If Not sConsulta.EOF Then
    With sConsulta
        For sLinha = 1 To sLinhas
            MSChart.Row = sLinha
            MSChart.RowLabel = VariableAdjust(!APLICACACAO, StringText)
            MSChart.Data = VariableAdjust(!Rendimento, DoubleNumber)
            .MoveNext
        Next
    End With
    
End If

End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub
