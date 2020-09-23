Attribute VB_Name = "modAddressOf"
Option Explicit

'======API Functions======
'Returns handle to system menu (Restore, Maximize, Move, close etc.) ...
'...of specefied object
Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, _
        ByVal bRevert As Long) As Long

'Returns number of items of specifiied submenu
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) _
        As Long

'Adds menu item tu specified submenu
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" _
        (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, _
        lpmii As MENUITEMINFO) As Long

'This function passes message information to the specified window procedure
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

'This function changes an attribute of the specified window.
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Opens files (documents and executable files), opens folders, ftp and http...
'...pages and sends e-mails (very usefull function)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long


'======Types======
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


'======Dims======
Public lonOldProc As Long  'pointer to Form1's previous window procedure


'======Consts======
Private Const WM_SYSCOMMAND = &H112
Private Const SW_SHOWNORMAL = 1

'Be carefull with code using AddressOf!
'WindowProc function called by SetWindowLong using AddressOf is...
'...responsible for all prosesses handled with window (in this case...
'...with frmMain).

'The following function acts as Form1's window procedure to process messages.
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
        Case WM_SYSCOMMAND
            '"Hide" was pressed
            If wParam = 0 Then
                'Hide frmMain and create icon on status area
                HideToBar
            '"Vote !"was pressed
            ElseIf wParam = 2 Then
                'Show PSC
                ShellExecute frmMain.hWnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=48058&lngWId=1", _
                        vbNullString, vbNullString, SW_SHOWNORMAL
            Else
                'Other item was selected.  Let the previous window procedure
                'process it.
                WindowProc = CallWindowProc(lonOldProc, hWnd, uMsg, wParam, lParam)
            End If
        Case Else
            'If this is some other message, let the previous procedure handle it.
            WindowProc = CallWindowProc(lonOldProc, hWnd, uMsg, wParam, lParam)
    End Select
    
End Function


