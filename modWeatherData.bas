Attribute VB_Name = "modWeatherData"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwRserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Info(0 To 4) As PageInfo

Public Type PageInfo
    url As String
    local As String
    canGet As Boolean
End Type

Public Opts(0 To 10) As Long

Public refreshCounter As Long

Public Function PopulateInfo()
    Info(0).canGet = True
    Info(0).local = "cur7Day.kdk"
    Info(0).url = "http://ftpcontent.worldnow.com/kwqc/WEATHER_7day.jpg"
    
    Info(1).canGet = True
    Info(1).local = "curRadarStill.kdk"
    Info(1).url = "http://ftpcontent.worldnow.com/kwqc/WEATHER_radar.jpg"
    
    Info(2).canGet = False
    Info(2).local = "curRadarLoop.kdk"
    Info(2).url = "http://ftpcontent.worldnow.com/kwqc/WEATHER_radar_loop.gif"
    
    Info(3).canGet = True
    Info(3).local = "curNow.kdk"
    Info(3).url = "http://ftpcontent.worldnow.com/kwqc/WEATHER_currents.jpg"
    
    Info(4).canGet = True
    Info(4).local = "curNowArea.kdk"
    Info(4).url = "http://ftpcontent.worldnow.com/kwqc/WEATHER_Temps.jpg"
End Function

Public Function ObtainImages()
    Dim errCode As Long
    For i = LBound(Info) To UBound(Info)
        If Info(i).canGet = True Then
            Call DeleteUrlCacheEntry(Info(i).url)
            errCode = URLDownloadToFile(0, Info(i).url, App.Path & "\" & Info(i).local, 0, 0)
            If errCode <> 0 Then
                Call WriteError(i + 1, errCode)
            End If
        End If
    Next
    
    LoadImg (Opts(1))
End Function

Public Function ObtainImage(selCode As Integer)
    Dim errCode As Long
    
    If Info(selCode).canGet = True Then
        Call DeleteUrlCacheEntry(Info(selCode).url)
        errCode = URLDownloadToFile(0, Info(selCode).url, App.Path & "\" & Info(selCode).local, 0, 0)
        If errCode <> 0 Then
            Call WriteError(selCode, errCode)
        End If
    End If
    
    LoadImg (selCode)
End Function

Public Function LoadImg(selCode As Integer)
    If FileExists(App.Path & "\" & Info(selCode).local) = False Then
        ObtainImage (selCode)
        Exit Function
    End If
    frmMain.imgCur.Picture = LoadPicture(App.Path & "\" & Info(selCode).local)
End Function

Public Function FileExists(fileName As String) As Boolean
    On Error GoTo ErrorHandler
    FileExists = (GetAttr(fileName) And vbDirectory) = 0
ErrorHandler:
    Exit Function
End Function

Public Function WriteError(class As Long, msg As String)
    Dim fNum As Long, fPath As String
    fNum = FreeFile
    fPath = App.Path & "\err.kdk"
    Open fPath For Append As #fNum
    Print #fNum, Now & vbTab & class & vbTab & msg
    Close #fNum
End Function
