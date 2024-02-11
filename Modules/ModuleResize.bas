Attribute VB_Name = "ModResize"
Option Explicit

Public Type ctrObj
    Name As String
    Index As Long
    Parrent As String
    Top As Long
    Left As Long
    Height As Long
    Width As Long
    ScaleHeight As Long
    ScaleWidth As Long
End Type

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Const MF_BYPOSITION = &H400&

Dim hMenu As Long

Private FormRecord() As ctrObj
Private ControlRecord() As ctrObj
Private bRunning As Boolean
Private MaxForm As Long
Private MaxControl As Long

Public Sub RemoveMenus(frm As Form, Optional Remove_Restore As Boolean, Optional Remove_Move As Boolean, Optional Remove_Size As Boolean, _
    Optional Remove_Minimize As Boolean, Optional Remove_Maximize As Boolean, Optional Remove_Seperator As Boolean, Optional Remove_Close As Boolean)

hMenu = GetSystemMenu(frm.hWnd, False)

If Remove_Close Then DeleteMenu hMenu, 6, MF_BYPOSITION
If Remove_Seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
If Remove_Maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
If Remove_Minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
If Remove_Size Then DeleteMenu hMenu, 2, MF_BYPOSITION
If Remove_Move Then DeleteMenu hMenu, 1, MF_BYPOSITION
If Remove_Restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Private Function ActualPos(plLeft As Long) As Long
If plLeft < 0 Then
    ActualPos = plLeft + 75000
Else
    ActualPos = plLeft
End If
End Function


Public Sub Centraliza(Parent As Form, Child As Form, Width As Long)
Dim iTop As Integer
Dim iLeft As Integer
'If Parent.WindowState <> 0 Then Exit Sub
Child.Width = Width
Child.Height = Parent.Height
iTop = ((Parent.Height - Child.Height) \ 2)
iLeft = ((Parent.Width - Child.Width) \ 2) + (Parent.Width - Child.Width) - 730
Child.Move iLeft, iTop
Child.Refresh
End Sub

Public Function SetForm(childForm_ As Form, mdiForm_ As Form)
SetParent childForm_.hWnd, mdiForm_.hWnd
End Function

Private Function FindForm(pfrmIn As Form) As Long
Dim i As Long
FindForm = -1
If MaxForm > 0 Then
    For i = 0 To (MaxForm - 1)
        If FormRecord(i).Name = pfrmIn.Name Then
            FindForm = i
            Exit Function
        End If
    Next i
End If
End Function

Private Function AddForm(pfrmIn As Form) As Long
Dim FormControl As Control, i As Long
ReDim Preserve FormRecord(MaxForm + 1)
FormRecord(MaxForm).Name = pfrmIn.Name: FormRecord(MaxForm).Top = pfrmIn.Top: FormRecord(MaxForm).Left = pfrmIn.Left: FormRecord(MaxForm).Height = pfrmIn.Height: FormRecord(MaxForm).Width = pfrmIn.Width: FormRecord(MaxForm).ScaleHeight = pfrmIn.ScaleHeight: FormRecord(MaxForm).ScaleWidth = pfrmIn.ScaleWidth: AddForm = MaxForm: MaxForm = MaxForm + 1
For Each FormControl In pfrmIn
    i = FindControl(FormControl, pfrmIn.Name)
    If i < 0 Then i = AddControl(FormControl, pfrmIn.Name)
Next FormControl
End Function

Private Function FindControl(inControl As Control, inName As String) As Long
Dim i As Long
FindControl = -1
For i = 0 To (MaxControl - 1)
    If ControlRecord(i).Parrent = inName Then
        If ControlRecord(i).Name = inControl.Name Then
            On Error Resume Next
            
            If ControlRecord(i).Index = inControl.Index Then
                FindControl = i
                Exit Function
            End If
            On Error GoTo 0
        End If
    End If
Next i
End Function

Private Function AddControl(inControl As Control, inName As String) As Long
ReDim Preserve ControlRecord(MaxControl + 1)
On Error Resume Next
ControlRecord(MaxControl).Name = inControl.Name: ControlRecord(MaxControl).Index = inControl.Index: ControlRecord(MaxControl).Parrent = inName
If TypeOf inControl Is Line Then
    ControlRecord(MaxControl).Top = inControl.Y1
    ControlRecord(MaxControl).Left = ActualPos(inControl.X1)
    ControlRecord(MaxControl).Height = inControl.Y2
    ControlRecord(MaxControl).Width = ActualPos(inControl.X2)
Else
    ControlRecord(MaxControl).Top = inControl.Top
    ControlRecord(MaxControl).Left = ActualPos(inControl.Left)
    ControlRecord(MaxControl).Height = inControl.Height
    ControlRecord(MaxControl).Width = inControl.Width
End If
inControl.IntegralHeight = False
On Error GoTo 0
AddControl = MaxControl: MaxControl = MaxControl + 1
End Function

Private Function PerWidth(pfrmIn As Form) As Long
Dim i As Long
i = FindForm(pfrmIn)
If i < 0 Then i = AddForm(pfrmIn)
PerWidth = (pfrmIn.ScaleWidth * 100) \ FormRecord(i).ScaleWidth
End Function


Private Function PerHeight(pfrmIn As Form) As Single
Dim i As Long
i = FindForm(pfrmIn)
If i < 0 Then i = AddForm(pfrmIn)
PerHeight = (pfrmIn.ScaleHeight * 100) \ FormRecord(i).ScaleHeight
End Function

Private Sub ResizeControl(inControl As Control, pfrmIn As Form)
On Error Resume Next
Dim i As Long, widthfactor As Single, heightfactor As Single, minFactor As Single
Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
yRatio = PerHeight(pfrmIn): xRatio = PerWidth(pfrmIn): i = FindControl(inControl, pfrmIn.Name)
If inControl.Left < 0 Then
    lLeft = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
Else
    lLeft = CLng((ControlRecord(i).Left * xRatio) \ 100)
End If
lTop = CLng((ControlRecord(i).Top * yRatio) \ 100): lWidth = CLng((ControlRecord(i).Width * xRatio) \ 100): lHeight = CLng((ControlRecord(i).Height * yRatio) \ 100)
If TypeOf inControl Is Line Then
    If inControl.X1 < 0 Then
        inControl.X1 = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
    Else
        inControl.X1 = CLng((ControlRecord(i).Left * xRatio) \ 100)
    End If
    inControl.Y1 = CLng((ControlRecord(i).Top * yRatio) \ 100)
    If inControl.X2 < 0 Then
        inControl.X2 = CLng(((ControlRecord(i).Width * xRatio) \ 100) - 75000)
    Else
        inControl.X2 = CLng((ControlRecord(i).Width * xRatio) \ 100)
    End If
    inControl.Y2 = CLng((ControlRecord(i).Height * yRatio) \ 100)
Else
    inControl.Move lLeft, lTop, lWidth, lHeight: inControl.Move lLeft, lTop, lWidth: inControl.Move lLeft, lTop
End If
End Sub


Public Sub ResizeForm(pfrmIn As Form)
Dim FormControl As Control, isVisible As Boolean
Dim StartX, StartY, MaxX, MaxY As Long
Dim bNew As Boolean
If Not bRunning Then
    bRunning = True
    bNew = FindForm(pfrmIn) < 0
    If pfrmIn.Top < 30000 Then
        isVisible = pfrmIn.Visible
        On Error Resume Next
        If Not pfrmIn.MDIChild Then
            On Error GoTo 0
        Else
            If bNew Then
                StartY = pfrmIn.Height
                StartX = pfrmIn.Width
                On Error Resume Next
                For Each FormControl In pfrmIn
                    If FormControl.Left + FormControl.Width + 200 > MaxX Then MaxX = FormControl.Left + FormControl.Width + 200: If FormControl.Top + FormControl.Height + 500 > MaxY Then MaxY = FormControl.Top + FormControl.Height + 500: If FormControl.X1 + 200 > MaxX Then MaxX = FormControl.X1 + 200: If FormControl.Y1 + 500 > MaxY Then MaxY = FormControl.Y1 + 500: If FormControl.X2 + 200 > MaxX Then MaxX = FormControl.X2 + 200: If FormControl.Y2 + 500 > MaxY Then MaxY = FormControl.Y2 + 500
                Next FormControl
                On Error GoTo 0
                pfrmIn.Height = MaxY: pfrmIn.Width = MaxX
            End If

            On Error GoTo 0
        End If
        For Each FormControl In pfrmIn
            ResizeControl FormControl, pfrmIn
        Next FormControl
        On Error Resume Next
        If Not pfrmIn.MDIChild Then
            On Error GoTo 0
            pfrmIn.Visible = isVisible
        Else
            If bNew Then
                pfrmIn.Height = StartY: pfrmIn.Width = StartX
                For Each FormControl In pfrmIn
                    ResizeControl FormControl, pfrmIn
                Next FormControl
            End If
        End If
        On Error GoTo 0
    End If
bRunning = False
End If
End Sub

Public Sub SaveFormPosition(pfrmIn As Form)
Dim i As Long
If MaxForm > 0 Then
    For i = 0 To (MaxForm - 1)
        If FormRecord(i).Name = pfrmIn.Name Then
            FormRecord(i).Top = pfrmIn.Top: FormRecord(i).Left = pfrmIn.Left: FormRecord(i).Height = pfrmIn.Height: FormRecord(i).Width = pfrmIn.Width
            Exit Sub
        End If
    Next i
    AddForm (pfrmIn)
End If
End Sub

Public Sub RestoreFormPosition(pfrmIn As Form)
Dim i As Long
If MaxForm > 0 Then
    For i = 0 To (MaxForm - 1)
        If FormRecord(i).Name = pfrmIn.Name Then
            If FormRecord(i).Top < 0 Then
                pfrmIn.WindowState = 2
            ElseIf FormRecord(i).Top < 30000 Then
                pfrmIn.WindowState = 0
                pfrmIn.Move FormRecord(i).Left, FormRecord(i).Top, FormRecord(i).Width, FormRecord(i).Height
            Else
                pfrmIn.WindowState = 1
            End If
            Exit Sub
        End If
    Next i
End If
End Sub
