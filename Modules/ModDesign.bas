Attribute VB_Name = "ModDesign"
Option Explicit
Public Sub ConfigurarForm(parForm As Form)

Dim sItem As Variant

For Each sItem In parForm.Controls
    If TypeOf sItem Is fpBtn Then
        If sItem.Tag <> 1 Then
            sItem.BackColor = &H404040
            sItem.ThreeDHighlightColor = &H404040
            sItem.ThreeDShadowColor = &H404040
            sItem.ThreeDTextShadowColor = &H404040
            sItem.FontName = "Segoe UI"
        End If
    ElseIf TypeOf sItem Is Frame Then
        sItem.Font = "Segoe UI"
        sItem.FontSize = 10
        If sItem.Name = "FrameBarra" Then
            sItem.BackColor = &H404040
        Else
            sItem.BackColor = RGB(247, 247, 247)
        End If
    ElseIf TypeOf sItem Is Form Then
        sItem.FontSize = 10
        sItem.BackColor = RGB(247, 247, 247)
        sItem.Font = "Segoe UI"
    ElseIf TypeOf sItem Is Label Then
        If sItem.Tag = 0 Then
            sItem.FontSize = 13
            sItem.BackColor = RGB(247, 247, 247)
            sItem.Font = "Segoe UI"
        ElseIf sItem.Tag = 1 Then
            sItem.FontSize = 20
            sItem.BackColor = RGB(247, 247, 247)
            sItem.Font = "Segoe UI SemiBold"
        Else
            sItem.FontSize = 10
            sItem.BackColor = &H404040
            sItem.ForeColor = vbWhite
            sItem.Font = "Segoe UI"
        End If
    ElseIf TypeOf sItem Is CheckBox Then
        sItem.FontSize = 10
        sItem.BackColor = RGB(247, 247, 247)
        sItem.Font = "Segoe UI"
    ElseIf TypeOf sItem Is TextBox Then
        sItem.FontSize = 10
        sItem.BackColor = RGB(247, 247, 247)
        sItem.Font = "Segoe UI"
    End If
Next

parForm.BackColor = RGB(247, 247, 247)

End Sub
