VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KWQC Weather"
   ClientHeight    =   5100
   ClientLeft      =   14265
   ClientTop       =   1845
   ClientWidth     =   7575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin VB.Timer tmrSetCustomLocData 
      Interval        =   100
      Left            =   3840
      Top             =   2400
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   3240
      Top             =   2400
   End
   Begin VB.Image imgCur 
      Height          =   5100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Visible         =   0   'False
      Begin VB.Menu mnuPictureRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuPictureRefreshAll 
         Caption         =   "Refresh All"
      End
      Begin VB.Menu mnuPictureOpts 
         Caption         =   "Options"
         Begin VB.Menu mnuPictureOptsShortcuts 
            Caption         =   "Shortcuts..."
            Begin VB.Menu mnuPictureOptsShortcutsDesktop 
               Caption         =   "Desktop"
            End
            Begin VB.Menu mnuPictureOptsShortcutsStartMenu 
               Caption         =   "Start Menu"
            End
            Begin VB.Menu mnuPictureOptsShortcutsStartUp 
               Caption         =   "Start-Up"
            End
            Begin VB.Menu mnuPictureOptsShortcutsAllUsers 
               Caption         =   "All Users"
               Begin VB.Menu mnuPictureOptsShortcutsAllUsersDesktop 
                  Caption         =   "Desktop"
               End
               Begin VB.Menu mnuPictureOptsShortcutsAllUsersStartMenu 
                  Caption         =   "Start Menu"
               End
               Begin VB.Menu mnuPictureOptsShortcutsAllUsersStartUp 
                  Caption         =   "Start-Up"
               End
            End
         End
         Begin VB.Menu mnuPictureOptsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPictureOptsStayOnTop 
            Caption         =   "Stay On Top"
         End
         Begin VB.Menu mnuPictureOptsOnExit 
            Caption         =   "On Exit..."
            Begin VB.Menu mnuPictureOptsOnExitChoice 
               Caption         =   "Close Application"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuPictureOptsOnExitChoice 
               Caption         =   "To System Tray"
               Index           =   1
            End
         End
         Begin VB.Menu mnuPictureOptsRefreshOnChange 
            Caption         =   "Refresh On Change"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPictureOptsRefreshOnClick 
            Caption         =   "Refresh On Click"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPictureOptsRefreshTime 
            Caption         =   "Refresh Time"
            Begin VB.Menu mnuPictureOptsRefreshTimeDur 
               Caption         =   "1 Minute"
               Index           =   0
            End
            Begin VB.Menu mnuPictureOptsRefreshTimeDur 
               Caption         =   "5 Minute"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuPictureOptsRefreshTimeDur 
               Caption         =   "10 Minute"
               Index           =   2
            End
            Begin VB.Menu mnuPictureOptsRefreshTimeDur 
               Caption         =   "30 Minute"
               Index           =   3
            End
         End
         Begin VB.Menu mnuPictureOptsDock 
            Caption         =   "Docking Position"
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Top Left"
               Index           =   0
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Top Center"
               Index           =   1
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Top Right"
               Index           =   2
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Left"
               Index           =   3
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Center"
               Index           =   4
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Right"
               Index           =   5
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Bottom Left"
               Index           =   6
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Bottom Center"
               Index           =   7
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Bottom Right"
               Index           =   8
            End
            Begin VB.Menu mnuPictureOptsDockLoc 
               Caption         =   "Custom"
               Checked         =   -1  'True
               Index           =   9
            End
         End
      End
      Begin VB.Menu mnuPictureSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPictureType 
         Caption         =   "7-Day Forecast"
         Index           =   0
      End
      Begin VB.Menu mnuPictureType 
         Caption         =   "Radar"
         Index           =   1
      End
      Begin VB.Menu mnuPictureType 
         Caption         =   "Radar (Loop)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPictureType 
         Caption         =   "Current"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuPictureType 
         Caption         =   "Current Area"
         Index           =   4
      End
      Begin VB.Menu mnuPictureSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPictureAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuPictureSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPictureExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "KWQC Weather already running!", vbOKOnly, "Error"
        End
    End If

    ' Options
    '   1 - Picture Type
    '   2 - Refresh Type
    '   3 - Refresh Counter
    '   4 - Dock Type
    '   5 - Dock Custom Left
    '   6 - Dock Custom Top
    '   7 - Refresh On Click
    '   8 - Refresh On Change
    '   9 - Stay On Top
    '   10 - Sys Tray On Exit
    
    Opts(1) = 3
    Opts(2) = 1
    Opts(3) = 300
    Opts(4) = 2
    Opts(5) = Me.Left
    Opts(6) = Me.Top
    Opts(7) = 1
    Opts(8) = 1
    Opts(9) = 0
    Opts(10) = 0
    
    
    LoadOpts
    modWeatherData.PopulateInfo
        
    mnuPictureType_Click (Opts(1))
    mnuPictureOptsRefreshTimeDur_Click (Opts(2))
    mnuPictureOptsDockLoc_Click (Opts(4))
    If Opts(7) = 0 Then
        mnuPictureOptsRefreshOnClick_Click
    End If
    If Opts(8) = 0 Then
        mnuPictureOptsRefreshOnChange_Click
    End If
    If Opts(9) = 1 Then
        mnuPictureOptsStayOnTop_Click
    End If
    mnuPictureOptsOnExitChoice_Click (Opts(10))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X
        Case WM_RBUTTONDOWN
            PopupMenu mnuSysTray
        Case WM_LBUTTONDBLCLK
            Me.Visible = Not Me.Visible
            modSysIcon.RemoveIconFromTray
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Opts(10) = 1 Then
        Cancel = 1
        modSysIcon.AddIconToTray Me.hWnd, Me.Icon
        Me.Hide
    End If
End Sub

Private Sub imgCur_Click()
    If Opts(7) = 1 Then
        ObtainImage (Opts(1))
    End If
End Sub

Private Sub imgCur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then
        PopupMenu mnuPicture
    End If
End Sub

Private Sub mnuPictureAbout_Click()
    MsgBox "Program created by Anthoni Wiese.", vbOKOnly, "About"
End Sub

Private Sub mnuPictureExit_Click()
    End
End Sub

Private Sub mnuPictureOptsDockLoc_Click(Index As Integer)
    For i = mnuPictureOptsDockLoc.LBound To mnuPictureOptsDockLoc.UBound
        mnuPictureOptsDockLoc(i).Checked = False
    Next
    mnuPictureOptsDockLoc(Index).Checked = True
    Call SetOptVal(Index, 4)
    '----
    Select Case Opts(4)
        Case 0:
            ' Top Left
            Me.Top = 0
            Me.Left = 0
        Case 1:
            Me.Top = 0
            Me.Left = Screen.Width / 2 - Me.Width / 2
        Case 2:
            Me.Top = 0
            Me.Left = Screen.Width - Me.Width
        Case 3:
            Me.Left = 0
            Me.Top = Screen.Height / 2 - Me.Height / 2
        Case 4:
            Me.Left = Screen.Width / 2 - Me.Width / 2
            Me.Top = Screen.Height / 2 - Me.Height / 2
        Case 5:
            Me.Left = Screen.Width - Me.Width
            Me.Top = Screen.Height / 2 - Me.Height / 2
        Case 6:
            Me.Top = Screen.Height - Me.Height
            Me.Left = 0
        Case 7:
            Me.Top = Screen.Height - Me.Height
            Me.Left = Screen.Width / 2 - Me.Width / 2
        Case 8:
            Me.Top = Screen.Height - Me.Height
            Me.Left = Screen.Width - Me.Width
        Case 9:
            Me.Left = Opts(5)
            Me.Top = Opts(6)
    End Select
End Sub

Private Sub mnuPictureOptsOnExitChoice_Click(Index As Integer)
    For i = mnuPictureOptsOnExitChoice.LBound To mnuPictureOptsOnExitChoice.UBound
        mnuPictureOptsOnExitChoice(i).Checked = False
    Next
    mnuPictureOptsOnExitChoice(Index).Checked = True
    Call SetOptVal(Index, 10)
End Sub

Private Sub mnuPictureOptsRefreshOnChange_Click()
    mnuPictureOptsRefreshOnChange.Checked = Not mnuPictureOptsRefreshOnChange.Checked
    If mnuPictureOptsRefreshOnChange.Checked = True Then
        Call SetOptVal(1, 8)
    Else
        Call SetOptVal(0, 8)
    End If
End Sub

Private Sub mnuPictureOptsRefreshOnClick_Click()
    mnuPictureOptsRefreshOnClick.Checked = Not mnuPictureOptsRefreshOnClick.Checked
    If mnuPictureOptsRefreshOnClick.Checked = True Then
        Call SetOptVal(1, 7)
    Else
        Call SetOptVal(0, 7)
    End If
End Sub

Private Sub mnuPictureOptsRefreshTimeDur_Click(Index As Integer)
    For i = mnuPictureOptsRefreshTimeDur.LBound To mnuPictureOptsRefreshTimeDur.UBound
        mnuPictureOptsRefreshTimeDur(i).Checked = False
    Next
    mnuPictureOptsRefreshTimeDur(Index).Checked = True
    Call SetOptVal(Index, 2)
    '----
    Select Case Opts(2)
        Case 0:
            Call SetOptVal(1 * 60, 3)
        Case 1:
            Call SetOptVal(5 * 60, 3)
        Case 2:
            Call SetOptVal(10 * 60, 3)
        Case 3:
            Call SetOptVal(30 * 60, 3)
    End Select
End Sub

Private Sub mnuPictureOptsShortcutsAllUsersDesktop_Click()
    modShortcut.CreateAllUsersDesktopShortcut
End Sub

Private Sub mnuPictureOptsShortcutsAllUsersStartMenu_Click()
    modShortcut.CreateAllUsersStartMenuShortcut
End Sub

Private Sub mnuPictureOptsShortcutsAllUsersStartUp_Click()
    modShortcut.CreateAllUsersStartupShortcut
End Sub

Private Sub mnuPictureOptsShortcutsDesktop_Click()
    modShortcut.CreateDesktopShortcut
End Sub

Private Sub mnuPictureOptsShortcutsStartMenu_Click()
    modShortcut.CreateStartMenuShortcut
End Sub

Private Sub mnuPictureOptsShortcutsStartUp_Click()
    modShortcut.CreateStartupShortcut
End Sub

Private Sub mnuPictureOptsStayOnTop_Click()
    mnuPictureOptsStayOnTop.Checked = Not mnuPictureOptsStayOnTop.Checked
    If mnuPictureOptsStayOnTop.Checked = True Then
        Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2)
        Call SetOptVal(1, 9)
    Else
        Call SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2)
        Call SetOptVal(0, 9)
    End If
End Sub

Private Sub mnuPictureRefreshAll_Click()
    modWeatherData.ObtainImages
End Sub

Private Sub mnuPictureType_Click(Index As Integer)
    For i = mnuPictureType.LBound To mnuPictureType.UBound
        mnuPictureType(i).Checked = False
    Next
    mnuPictureType(Index).Checked = True
    Call SetOptVal(Index, 1)
    '----
    If Opts(8) = 1 Then
        ObtainImage (Opts(1))
    End If
    LoadImg (Opts(1))
End Sub

Private Sub mnuPictureRefresh_Click()
    modWeatherData.ObtainImage (Opts(1))
End Sub

Private Sub mnuSysTrayExit_Click()
    modSysIcon.RemoveIconFromTray
    End
End Sub

Private Sub mnuSysTrayShow_Click()
    Me.Visible = True
    modSysIcon.RemoveIconFromTray
End Sub

Private Sub tmrRefresh_Timer()
    refreshCounter = refreshCounter + 1
    If refreshCounter >= Opts(3) Then
        modWeatherData.ObtainImage (Opts(1))
        refreshCounter = 0
    End If
End Sub

Private Sub tmrSetCustomLocData_Timer()
    Call SetOptVal(Me.Left, 5)
    Call SetOptVal(Me.Top, 6)
End Sub

Private Sub SetOptVal(ByVal val As Long, code As Long)
    Opts(code) = val
    SaveOpts
End Sub

Private Sub SaveOpts()
    Dim fNum As Long, fPath As String
    Dim sOut As String
    
    For i = LBound(Opts) To UBound(Opts)
        sOut = sOut & Opts(i) & "|"
    Next
    sOut = Left$(sOut, Len(sOut) - 1)
    
    fNum = FreeFile()
    fPath = App.Path & "\opts.kdk"
    Open fPath For Output As fNum
    Print #fNum, sOut
    Close #fNum
End Sub

Private Sub LoadOpts()
    Dim fNum As Long, fPath As String
    Dim sIn As String, sSplit() As String
        
    fNum = FreeFile()
    fPath = App.Path & "\opts.kdk"
    
    If FileExists(fPath) = False Then Exit Sub
    
    Open fPath For Input As fNum
    sIn = Input(LOF(fNum), fNum)
    Close #fNum
    
    sSplit() = Split(sIn, "|")
    
    If UBound(sSplit) < 7 Then
        MsgBox "Option File Bad. Options Reset To Default."
        Kill App.Path & "\opts.kdk"
    End If
    
    For i = LBound(sSplit) To UBound(sSplit)
        Opts(i) = sSplit(i)
    Next
End Sub
