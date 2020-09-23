Attribute VB_Name = "SysTray"
'The necessary code for putting a program in the system tray

Option Explicit

Public Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uId As Long
     uFlags As Long
     uCallBackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, ptry As NOTIFYICONDATA) As Boolean
Public Prog As NOTIFYICONDATA

Public Sub ShowInSysTray()
    Prog.cbSize = Len(Prog)
    Prog.hwnd = frmMain.hwnd
    Prog.uId = vbNull
    Prog.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Prog.uCallBackMessage = WM_MOUSEMOVE
    Prog.hIcon = frmMain.Icon
    Prog.szTip = "Click here to record a command" & vbNullChar
    Shell_NotifyIcon NIM_ADD, Prog
    Shell_NotifyIcon NIM_MODIFY, Prog
End Sub

Public Sub CloseSysTray()
Shell_NotifyIcon NIM_DELETE, Prog
End Sub


