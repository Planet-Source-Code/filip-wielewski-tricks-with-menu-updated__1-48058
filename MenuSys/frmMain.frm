VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Can you close me?"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenuHide 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   2280
   End
   Begin VB.PictureBox picExit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   960
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   675
      ScaleWidth      =   2535
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   960
      Picture         =   "frmMain.frx":114C
      ScaleHeight     =   675
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picIco 
      Height          =   375
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBmp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   600
      Picture         =   "frmMain.frx":15CE
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Whatever"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'=============================================================='
'=                                                            ='
'= ======AUTHOR======                                         ='
'= THIS IS A FREE CODE                                        ='
'= BY FILIP WIELEWSKI                                         ='
'= E-MAIL: WIELFILIST@WP.PL                                   ='
'=                                                            ='
'= ======SORRY FOR:======                                     ='
'= my bad english which I use in descriptions :]              ='
'=                                                            ='
'=============================================================='

'======API Functions======
'Draws menu (refreshes it and saves changes made to it)
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

'Removes specified menu item
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
        ByVal nPosition As Long, ByVal wFlags As Long) As Long

'This function sets a bitmap (small bitmap) to specified menu item
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, _
        ByVal nPosition As Long, ByVal wFlags As Long, _
        ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

'Returns handle to menu of specified object (window)
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

'Returns handle to submenu of specified menu
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
        ByVal nPos As Long) As Long

'Returns ID of menu item of specified menu or submenu
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
        ByVal nPos As Long) As Long

'Modifies menu item (adds even large bitmap to menu item)
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
        (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

'Ends a menu with one or more menu items
Private Declare Function EndMenu Lib "user32.dll" () As Long

'======Consts=====
Private Const MF_BYPOSITION = &H400&    'identify menu items by their position
Private Const MF_REMOVE = &H1000&       'action: remove menu item
'Private Consts for MENUITEMINFO type
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8
'Private Const for SetWindowLong function
Private Const GWL_WNDPROC = -4
'To recognize button pressed:
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
'To set large bitmaps
Private Const MF_BITMAP = 4

Private Sub cmdEnd_Click()
    
    'End program
    Unload frmMain
    
End Sub

Private Sub Form_Load()
    
    Dim lonSysMenu As Long      'Handle to system menu
    Dim lonItemCount As Long    'Value retrieved by functions
    Dim miiMii As MENUITEMINFO  'describes a menu item to add
    Dim lonRet As Long          'return value
    
    'Get handle to our form's system menu...
    '...(Restore, Maximize, Move, close etc.)
    lonSysMenu = GetSystemMenu(frmMain.hWnd, False)
    
    'Picture property returns graphic of object; image property returns...
    '...handle to persistent bitmap (provided by Microsoft). Thanks these...
    '...three lines of code below you can add to menu also icons. But:
    '...-1: icons won't be transparent on menu, so set backcolor property...
            '...of picturebox to menubar (then icons will look like...
            '...transparent)
    '...-2: ModifyMenu function won't set white color of bitmaps into...
            '...transparent (so if ModifyMenu function uses picturebox...
            '...with bitmaps - not with icons - then don't use those...
            '...three lines below)
            
    picBmp.Picture = picBmp.Image
    'notice that white color of bitmaps on menu will be set to transparent...
    '...if you delete these two lines below (ModifyMenu function uses these...
    '...two pictureboxes). But if picExit and picShow has icons - not bitmaps -...
    '...then don't delete these two lines!
    picExit.Picture = picExit.Image
    picShow.Picture = picShow.Image
    
    If lonSysMenu Then
        'Get System menu's menu count
        lonItemCount = GetMenuItemCount(lonSysMenu)
        'If theree is more than 0 items ("GetMenuItemCount" returns 0 when...
        '...error occurred.) then:
        If lonItemCount Then
            'Menu count is based on 0 (0, 1, 2, 3...)
            'Remove close button (X)
            RemoveMenu lonSysMenu, lonItemCount - 1, MF_BYPOSITION Or MF_REMOVE
            'Refresh frmMain's system menu
            DrawMenuBar frmMain.hWnd
        End If
    End If

    If lonSysMenu Then
        'Get System menu's menu count
        lonItemCount = GetMenuItemCount(lonSysMenu)
        'If theree is more than 0 items ("GetMenuItemCount" returns 0 when...
        '...error occurred.) then:
        If lonItemCount Then
            'Add a "Hide" to system menu
            With miiMii
                'Size of the structure.
                .cbSize = Len(miiMii)
                .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
                'This is a regular text item.
                .fType = MFT_STRING
                'The option is enabled.
                .fState = MFS_ENABLED
                'It has an ID of 1 (this identifies it in the window procedure).
                .wID = 0
                'Text to display in system menu
                .dwTypeData = "&Hide"
                .cch = Len(.dwTypeData)
            End With
            'Add "Hide" to the bottom of the system menu.
            lonRet = InsertMenuItem(lonSysMenu, lonItemCount, 1, miiMii)
            'Add a separator bar to the system menu.
            With miiMii
                'What parts of the structure to use.
                .fMask = MIIM_ID Or MIIM_TYPE
                'This is a separator.
                .fType = MFT_SEPARATOR
                'It has an ID of 0.
                .wID = 1
            End With
            'Add the separator to the end of the system menu.
            lonRet = InsertMenuItem(lonSysMenu, lonItemCount + 1, 1, miiMii)
            'Add a "&Vote" to system menu
            With miiMii
                'Size of the structure.
                .cbSize = Len(miiMii)
                .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
                'This is a regular text item.
                .fType = MFT_STRING
                'The option is enabled.
                .fState = MFS_ENABLED Or MFS_CHECKED
                'It has an ID of 1 (this identifies it in the window procedure).
                .wID = 2
                'Text to display in system menu
                .dwTypeData = "&Vote !"
                .cch = Len(.dwTypeData)
            End With
            'Add "Vote" to the bottom of the system menu.
            lonRet = InsertMenuItem(lonSysMenu, lonItemCount + 2, 1, miiMii)
            'set bitmap of "Hide" menu item in system menu
            SetMenuItemBitmaps lonSysMenu, lonItemCount, MF_BYPOSITION, picBmp.Picture, _
            picBmp.Picture
            lonOldProc = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, _
                    AddressOf WindowProc)
        End If
    End If
    
    'Print message
    frmMain.CurrentX = 30
    frmMain.CurrentY = 50
    frmMain.Print "Check system menu! (right click on title bar)"
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'set position of cmdEnd
    cmdEnd.Left = (frmMain.ScaleWidth - cmdEnd.Width) / 2
    cmdEnd.Top = (frmMain.ScaleHeight - cmdEnd.Height) / 2
    
End Sub

'Before unloading remove the custom window procedure.
Private Sub Form_Unload(Cancel As Integer)

    Dim lonRet As Long  'return value
    
    'Replace the previous window procedure to prevent crashing.
    lonRet = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, lonOldProc)
    'Remove the modifications made to the system menu.
    lonRet = GetSystemMenu(frmMain.hWnd, 1)

End Sub

Private Sub mnuExit_Click()
    
    'End
    Unload frmMain
    
End Sub

Private Sub mnuShow_Click()
    
    'Show frmMain and destroy icon on status area
    ShowFromBar
    
End Sub

Private Sub picIco_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Static booPressed As Boolean
    Static lonMsg As Long
    
    lonMsg = X / Screen.TwipsPerPixelX
    'if frmMain is not "popupping" at the moment then:
    If booPressed = False Then
        booPressed = True
        Select Case lonMsg
            Case WM_LBUTTONUP:
                'Modify menu
                ShowMenu
                'Enable timer (it will hide menu after 2 seconds)
                tmrMenuHide.Enabled = True
                'on left mouse button up"
                frmMain.PopupMenu mnuMain
            Case WM_RBUTTONUP:
                'Modify menu
                ShowMenu
                'Enable timer (it will hide menu after 2 seconds)
                tmrMenuHide.Enabled = True
                'on right mouse button up"
                frmMain.PopupMenu mnuMain
        End Select
        booPressed = False
    End If
    
End Sub

Private Sub ShowMenu()
    
    Dim lonMenu As Long         'Handle to frmMain's menu
    Dim lonSubMenu As Long      'Handle to submenu of mnuMain
    Dim lonMenuItemID As Long   'Handle to menu item of submenu of mnuMain
    
    If mnuMain.Visible = False Then
        'show frmMain.mnuMain (if it isn't visible then "ModifyMenu" function...
        '...won't work properly)
        mnuMain.Visible = True
        'Get handle to frmMain.mnuMain
        lonMenu = GetMenu(frmMain.hWnd)
        'Get handle to submenu of mnuMain
        lonSubMenu = GetSubMenu(lonMenu, 0)
        'Menu count is based on 0 (0, 1, 2, 3...)
        'Get ID of mnuShow
        lonMenuItemID = GetMenuItemID(lonSubMenu, 0)
        'set bitmap of picShow to mnuShow
        ModifyMenu lonMenu, lonMenuItemID, MF_BITMAP, lonMenuItemID, _
                CLng(picShow.Picture)
        'Get ID of mnuExit
        lonMenuItemID = GetMenuItemID(lonSubMenu, 2)
        'set bitmap of picExit to mnuExit
        ModifyMenu lonMenu, lonMenuItemID, MF_BITMAP, lonMenuItemID, _
                CLng(picExit.Picture)
    End If
    
End Sub

Private Sub tmrMenuHide_Timer()
    
    'Hide menu
    EndMenu
    'Disable this timer
    tmrMenuHide.Enabled = False
    
End Sub
