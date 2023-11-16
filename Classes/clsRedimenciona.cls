VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRedimenciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ControlInfo_type
  Left As Single
  Top As Single
  Width As Single
  Height As Single
  FontSize As Single
End Type
Dim ControlInfos() As ControlInfo_type


Public Sub IniciarRedimencionamento(parForm As Form)

Dim ThisControl As Control
  
ReDim Preserve ControlInfos(0 To 0)
ControlInfos(0).Width = parForm.Width
ControlInfos(0).Height = parForm.Height
For Each ThisControl In parForm.Controls
  ReDim Preserve ControlInfos(0 To UBound(ControlInfos) + 1)
  On Error Resume Next  ' hack to bypass controls with no size or position properties
  With ControlInfos(UBound(ControlInfos))
    .Left = ThisControl.Left
    .Top = ThisControl.Top
    .Width = ThisControl.Width
    .Height = ThisControl.Height
    .FontSize = ThisControl.FontSize
  End With
  On Error GoTo 0
Next

End Sub


Public Sub Redimencionar(parForm As Form)
Dim ThisControl As Control, HorizRatio As Single, VertRatio As Single, Iter As Integer

If parForm.WindowState = vbMinimized Then Exit Sub

HorizRatio = parForm.Width / ControlInfos(0).Width
VertRatio = parForm.Height / ControlInfos(0).Height

Iter = 0
For Each ThisControl In parForm.Controls
  Iter = Iter + 1
  On Error Resume Next  ' hack to bypass controls
  With ThisControl
    .Left = ControlInfos(Iter).Left * HorizRatio
    .Top = ControlInfos(Iter).Top * VertRatio
    .Width = ControlInfos(Iter).Width * HorizRatio
    .Height = ControlInfos(Iter).Height * VertRatio
    .FontSize = ControlInfos(Iter).FontSize * HorizRatio
  End With
  On Error GoTo 0
Next

End Sub
