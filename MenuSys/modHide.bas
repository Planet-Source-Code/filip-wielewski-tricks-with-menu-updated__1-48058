Attribute VB_Name = "modHide"
Option Explicit

'======API Functions======
'Sends a message to the system to add, modify, or delete an icon from the...
'...taskbar status area.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias _
        "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) _
        As Boolean

'======Types======
'will store information for Shell_NotifyIcon function
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'======Dims======
Dim nidIcon As NOTIFYICONDATA

'======Consts======
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Sub HideToBar()
    
    'create icon and hide frmMain
    nidIcon.cbSize = Len(nidIcon)
    'hwnd to object which will be connected with icon on winbar
    nidIcon.hWnd = frmMain.picIco.hWnd
    nidIcon.uId = 1&
    nidIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nidIcon.ucallbackMessage = WM_MOUSEMOVE
    'icon
    nidIcon.hIcon = frmMain.Icon
    'tip's text
    nidIcon.szTip = "Pretty good, isn't it?" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, nidIcon
    'hide frmMain
    frmMain.Hide
    
End Sub

Public Sub ShowFromBar()
    
    frmMain.tmrMenuHide.Enabled = False
    'hide frmMain.mnuMain
    frmMain.mnuMain.Visible = False
    'destroy icon and show frmMain
    nidIcon.cbSize = Len(nidIcon)
    nidIcon.hWnd = frmMain.picIco.hWnd
    nidIcon.uId = 1&
    Shell_NotifyIcon NIM_DELETE, nidIcon
    frmMain.Show
    
End Sub


