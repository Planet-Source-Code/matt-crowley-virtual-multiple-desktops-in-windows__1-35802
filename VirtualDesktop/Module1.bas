Attribute VB_Name = "modWindowFunctions"
'//
'// modWindowFunctions
'// -----------------------------
'// Controls both functions to 'switch' desktops and control
'// the app's presence within the system tray
'//
'// Please comment and vote on PSC
'//

'// API Calls
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'// Public Types
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'// Public Constants
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const WM_CLOSE = &H10
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000

'// Desktop Information Arrays
'// These hold all the hWnd information for the 10 desktops
'// (0 to 9) and the bax number of windows per desktop (1024)
Public openWindows(0 To 10, 0 To 1023) As Long
Public openWindowsCount(0 To 10) As Long
Public currentDesktop As Integer
Public pastDesktop As Integer

'// Set var to type
Public NotifyIcon As NOTIFYICONDATA

'// Var to notifiy if it is a task
Public IsTask As Long

'// Switch Desktop
Public Function switchDesktop(fromDesktop As Integer, gotoDesktop As Integer)

    Dim hwndCurrentWindow As Long
    Dim intLen As Long
    Dim strWindowTitle As String
    Dim windowCounter As Integer
    
    '// Go through every window, check if it is a task, check if it is
    '// itself, then, if not, hide it
    IsTask = WS_VISIBLE Or WS_BORDER
    windowCounter = 0
    hwndCurrentWindow = GetWindow(frmSettings.hWnd, GW_HWNDFIRST)
    Do While hwndCurrentWindow
        If hwndCurrentWindow <> frmSettings.hWnd And TaskWindow(hwndCurrentWindow) Then
            intLen = GetWindowTextLength(hwndCurrentWindow) + 1
            strWindowTitle = Space$(intLen)
            intLen = GetWindowText(hwndCurrentWindow, strWindowTitle, intLen)
            If intLen > 0 Then
                If hwndCurrentWindow <> frmSettings.hWnd Then
                    RetVal = ShowWindow(hwndCurrentWindow, SW_HIDE)
                    openWindows(fromDesktop, windowCounter) = hwndCurrentWindow
                    windowCounter = windowCounter + 1
                End If
            End If
        End If
        hwndCurrentWindow = GetWindow(hwndCurrentWindow, GW_HWNDNEXT)
    Loop
    openWindowsCount(fromDesktop) = windowCounter

    '// Now, unhide the desktop that we want to have on top.  Go through
    '// the array information we collected from the last opening of this
    '// desktop.  By default, the array is blank, meaning no window will
    '// be opened if it the first time opening this desktop
    windowCounter = 0
    While windowCounter < openWindowsCount(gotoDesktop)
        RetVal = ShowWindow(openWindows(gotoDesktop, windowCounter), SW_SHOW)
        windowCounter = windowCounter + 1
    Wend
    
    '// Move the current to past, the new desktop as current
    pastDesktop = fromDesktop
    currentDesktop = gotoDesktop
    
End Function

Function TaskWindow(hwCurr As Long) As Long
    
    '// Determine if this is a task window
    Dim lngStyle As Long
    lngStyle = GetWindowLong(hwCurr, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then TaskWindow = True

End Function
