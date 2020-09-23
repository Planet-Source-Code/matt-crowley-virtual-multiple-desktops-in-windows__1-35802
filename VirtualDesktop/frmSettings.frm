VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Amoeba VirtualDesktop Settings"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   660
   ScaleWidth      =   2220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnu_1 
      Caption         =   "Menu"
      Begin VB.Menu mnu1 
         Caption         =   "1"
      End
      Begin VB.Menu mnu2 
         Caption         =   "2"
      End
      Begin VB.Menu mnu3 
         Caption         =   "3"
      End
      Begin VB.Menu mnu4 
         Caption         =   "4"
      End
      Begin VB.Menu mnu5 
         Caption         =   "5"
      End
      Begin VB.Menu mnu6 
         Caption         =   "6"
      End
      Begin VB.Menu mnu7 
         Caption         =   "7"
      End
      Begin VB.Menu mnu8 
         Caption         =   "8"
      End
      Begin VB.Menu mnu9 
         Caption         =   "9"
      End
      Begin VB.Menu mnu10 
         Caption         =   "10"
      End
      Begin VB.Menu mnuSeperatorB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//
'// frmSettings
'// -----------------------------
'// Manages the tray icon and menu
'//
'// Please comment and vote on PSC
'//

Private Sub Form_Load()
    
    '// Hide this form
    Me.Hide
    
    '// Initialize the current and past desktops
    currentDesktop = 1
    pastDesktop = 1
    
    '// Initialize the SystemTray Icon Info
    With NotifyIcon
        .cbSize = Len(NotifyIcon)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Click to edit VirtualDesktop" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, NotifyIcon
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Check if both isMinimized and Right Click on mouse
    Dim Result As Long
    Dim Message As Long
    If Me.ScaleMode = vbPixels Then
        Message = X
    Else
        Message = X / Screen.TwipsPerPixelX
    End If
    If Message = WM_RBUTTONUP Then
        Result = SetForegroundWindow(Me.hWnd)
        Me.PopupMenu Me.mnu_1
    End If

End Sub

Private Sub Form_Resize()

    '// Hide the form if it is minimized
    If frmSettings.WindowState = vbMinimized Then
        frmSettings.Hide
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Kill the SysTray icon
    Shell_NotifyIcon NIM_DELETE, NotifyIcon
    
End Sub

'// All these menus deal with access to the 10 desktops, Desktop 1 being
'// the original
Private Sub mnu1_Click()
    switchDesktop currentDesktop, 1
End Sub

Private Sub mnu2_Click()
    switchDesktop currentDesktop, 2
End Sub

Private Sub mnu3_Click()
    switchDesktop currentDesktop, 3
End Sub

Private Sub mnu4_Click()
    switchDesktop currentDesktop, 4
End Sub

Private Sub mnu5_Click()
    switchDesktop currentDesktop, 5
End Sub

Private Sub mnu6_Click()
    switchDesktop currentDesktop, 6
End Sub

Private Sub mnu7_Click()
    switchDesktop currentDesktop, 7
End Sub

Private Sub mnu8_Click()
    switchDesktop currentDesktop, 8
End Sub

Private Sub mnu9_Click()
    switchDesktop currentDesktop, 9
End Sub

Private Sub mnu10_Click()
    switchDesktop currentDesktop, 10
End Sub

Private Sub mnuExit_Click()
    Load frmExit
    frmExit.Show
    
End Sub
