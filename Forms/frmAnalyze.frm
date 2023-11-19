VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAnalyze 
   Caption         =   "Análise de Aplicações"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   20385
   Icon            =   "frmAnalyze.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   20385
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   10575
      Left            =   0
      OleObjectBlob   =   "frmAnalyze.frx":680A
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim iRecordset As New Recordset, iQuery As String, iRows As Long, iPosition As Long

iQuery = "SELECT DISTINCT APL_NOME AS APLICACACAO, SUM(SAL_RENDIMENTO) AS RENDIMENTO FROM APLICACOES" _
    & " INNER JOIN SALDOS ON SAL_APLICACAO = APL_CODIGO AND APL_INATIVO = 0 GROUP BY  APL_NOME"
    
Set iRecordset = ReadQuery(iQuery, iRows)

MSChart.chartType = 1
MSChart.ShowLegend = False
MSChart.Title = "Análise de Aplicacões"

MSChart.Column = 1
MSChart.RowCount = iRows
MSChart.Visible = True

If Not iRecordset.EOF Then
    With iRecordset
        For iPosition = 1 To iRows
            MSChart.Row = iPosition
            MSChart.RowLabel = VariableAdjust(!APLICACACAO, eStringText)
            MSChart.Data = VariableAdjust(!Rendimento, eDoubleNumber)
            .MoveNext
        Next
    End With
End If

Me.Caption = Me.Caption & FillFooter
End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub
