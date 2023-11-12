Attribute VB_Name = "ModuleMenu"
Option Explicit

'API's do Windows para trabalhar com menus
Private Const MF_SEPARATOR As Long = &H800
Private Const MF_BYPOSITION As Long = &H400
Private Const MF_POPUP As Long = &H10
Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal uIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long

'API's do Windows para trabalhar com janelas e mensagens
Private Const WM_COMMAND As Long = &H111
Private Const GWL_WNDPROC As Long = -4
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Endereço do WndProc antigo do Form
Private oldWndProc As Long
Private frmOriginal As Form

Public Sub PreparaForm1(frm As Form)
    'Set frmOriginal = Form

    'Esse código todo tem que vir aqui em um módulo separado por causa
    'do operador AddressOf
    oldWndProc = SetWindowLong(frm.hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub AdicionaItem(ByVal indiceDoMenuPai As Long, ByVal id As Long, ByVal texto As String)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)

    AppendMenu menu, 0, id, texto
End Sub

Public Sub AdicionaItemSub(ByVal indiceDoMenuPai As Long, ByVal indiceDoSubMenu As Long, ByVal id As Long, ByVal texto As String)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)
    menu = GetSubMenu(menu, indiceDoSubMenu)

    AppendMenu menu, 0, id, texto
End Sub

Public Sub AdicionaSeparador(ByVal indiceDoMenuPai As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)

    AppendMenu menu, MF_SEPARATOR, 0, ""
End Sub

Public Sub AdicionaSeparadorSub(ByVal indiceDoMenuPai As Long, ByVal indiceDoSubMenu As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)
    menu = GetSubMenu(menu, indiceDoSubMenu)

    AppendMenu menu, MF_SEPARATOR, 0, ""
End Sub

Public Sub AdicionaSubMenu(ByVal indiceDoMenuPai As Long, ByVal texto As String)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)

    AppendMenu menu, MF_POPUP, CreatePopupMenu, texto
End Sub

Public Sub RemoveItemPorIndice(ByVal indiceDoMenuPai As Long, ByVal indiceDoItem As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)

    RemoveMenu menu, indiceDoItem, MF_BYPOSITION
End Sub

Public Sub RemoveItemPorIndiceSub(ByVal indiceDoMenuPai As Long, ByVal indiceDoSubMenu As Long, ByVal indiceDoItem As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)
    menu = GetSubMenu(menu, indiceDoSubMenu)

    RemoveMenu menu, indiceDoItem, MF_BYPOSITION
End Sub

Public Sub RemoveItemPorId(ByVal indiceDoMenuPai As Long, ByVal idDoItem As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)

    RemoveMenu menu, idDoItem, 0
End Sub

Public Sub RemoveItemPorIdSub(ByVal indiceDoMenuPai As Long, ByVal indiceDoSubMenu As Long, ByVal idDoItem As Long)
    Dim barraDeMenus As Long
    barraDeMenus = GetMenu(frmOriginal.hWnd)

    Dim menu As Long
    menu = GetSubMenu(barraDeMenus, indiceDoMenuPai)
    menu = GetSubMenu(menu, indiceDoSubMenu)

    RemoveMenu menu, idDoItem, 0
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If message = WM_COMMAND Then
        If frmOriginal.MenuClicado(wParam And &HFFFF) = True Then
            'Quando um dos nossos menus foi clicado, apenas retorna 0,
            'e para a função por aqui
            WndProc = 0
            Exit Function
        End If
    End If
    'Chama o WndProc antigo do Form
    WndProc = CallWindowProc(oldWndProc, hWnd, message, wParam, lParam)
End Function
