Attribute VB_Name = "modShortcut"
Public oShell As Object
Public oShortcut As Object

Public Function CreateDesktopShortcut()
    Set oShell = CreateObject("WScript.Shell")
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") & "\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function CreateStartMenuShortcut()
    Set oShell = CreateObject("WScript.Shell")
    If DirExists(oShell.SpecialFolders("StartMenu") & "\Programs\KWQC Weather") = False Then
        MkDir (oShell.SpecialFolders("StartMenu") & "\Programs\KWQC Weather")
    End If
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("StartMenu") & "\Programs\KWQC Weather\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function CreateStartupShortcut()
    Set oShell = CreateObject("WScript.Shell")
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("Startup") & "\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function CreateAllUsersDesktopShortcut()
    Set oShell = CreateObject("WScript.Shell")
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersDesktop") & "\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function CreateAllUsersStartMenuShortcut()
    Set oShell = CreateObject("WScript.Shell")
    If DirExists(oShell.SpecialFolders("AllUsersStartMenu") & "\Programs\KWQC Weather") = False Then
        MkDir (oShell.SpecialFolders("AllUsersStartMenu") & "\Programs\KWQC Weather")
    End If
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartMenu") & "\Programs\KWQC Weather\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function CreateAllUsersStartupShortcut()
    Set oShell = CreateObject("WScript.Shell")
    Set oShortcut = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartup") & "\KWQC Weather.lnk")
    oShortcut.TargetPath = App.Path & "\KWQC Weather.exe"
    oShortcut.IconLocation = App.Path & "\KWQC Weather.exe, 0"
    oShortcut.Description = "KWQC Weather"
    oShortcut.WorkingDirectory = App.Path
    oShortcut.Save
End Function

Public Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
