Attribute VB_Name = "modSysIcon"
Public Type NotifyIconData
    size As Long
    Handle As Long
    ID As Long
    Flags As Long
    CallBackMessage As Long
    Icon As Long
    Tip As String * 64
End Type

Public Const AddIcon = &H0
Public Const ModifyIcon = &H1
Public Const DeleteIcon = &H2

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Const MessageFlag = &H1
Public Const IconFlag = &H2
Public Const TipFlag = &H4

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NotifyIconData) As Boolean

Public Data As NotifyIconData


Public Function AddIconToTray(hWnd As Long, ico As Long)
    Data.size = Len(Data)
    Data.Handle = hWnd
    Data.ID = vbNull
    Data.Flags = IconFlag Or TipFlag Or MessageFlag
    Data.CallBackMessage = WM_MOUSEMOVE
    Data.Icon = ico
    Data.Tip = "KWQC Weather" & vbNullChar
    
    Call Shell_NotifyIcon(AddIcon, Data)
End Function

Public Function RemoveIconFromTray()
    Call Shell_NotifyIcon(DeleteIcon, Data)
End Function
