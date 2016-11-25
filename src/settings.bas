Attribute VB_Name = "SettingsMd"
Function Limit(ByVal Setting As Integer, ByVal Min As Integer, ByVal Max As Integer, ByVal Default As Integer)
    If Setting < Min Or Setting > Max Then
        Limit = Default
    Else
        Limit = Setting
    End If
End Function

Sub LoadSettings()
    Dim fn As String
    Dim ln As String
    Dim Setting() As String
    Dim SettingValue As Integer
    
    ' Defaults
    Spacing = 1
    LiveRefresh = True
    AltBubbles = True
    OpenLast = True
    GroupAlpha = 31
    TicksAlpha = 15
    ColorScheme = 0     ' "Default"
    ColorSat = 50
    AntiAliasing = True
    UISplitY = 256
    
    fn = App.Path & "\settings.ini"
    
    If FSO.FileExists(fn) = True Then
        Open fn For Input As #1
            Do
                Line Input #1, ln
                Setting = Split(LCase(Trim(ln)), "=")
                Setting(1) = Trim(Setting(1))
                If UBound(Setting) >= 0 Then
                    SettingValue = Val(Setting(1))
                    If Setting(0) = "spacing" Then Spacing = Limit(SettingValue, 5, 40, 10) / 10
                    If Setting(0) = "liverefresh" Then LiveRefresh = S2B(Setting(1))
                    If Setting(0) = "altbubbles" Then AltBubbles = S2B(Setting(1))
                    If Setting(0) = "openlast" Then OpenLast = S2B(Setting(1))
                    If Setting(0) = "lastopened" Then LastOpened = Setting(1)
                    If Setting(0) = "groupalpha" Then GroupAlpha = Limit(SettingValue, 5, 127, 31)
                    If Setting(0) = "ticksalpha" Then TicksAlpha = Limit(SettingValue, 0, 127, 31)
                    If Setting(0) = "colorscheme" Then ColorScheme = Limit(SettingValue, 0, SettingsFrm.Combo1.ListCount, 0)
                    If Setting(0) = "colorsat" Then ColorSat = Limit(SettingValue, 10, 100, 50)
                    If Setting(0) = "antialiasing" Then AntiAliasing = S2B(Setting(1))
                    If Setting(0) = "split" Then UISplitY = Limit(SettingValue, 16, MainFrm.ScaleHeight - 48, 256)
                End If
            Loop While Not EOF(1)
        Close #1
    End If
    
    LoadColorScheme
End Sub

Sub SaveSettings()
    Dim fn As String
    Dim ln As String
    Dim Setting() As String
    
    fn = App.Path & "\settings.ini"
    
    Open fn For Output As #1
        ln = "spacing=" & Int(Spacing * 10)
        Print #1, ln
        ln = "liverefresh=" & B2S(LiveRefresh)
        Print #1, ln
        ln = "altbubbles=" & B2S(AltBubbles)
        Print #1, ln
        ln = "openlast=" & B2S(OpenLast)
        Print #1, ln
        ln = "lastopened=" & FilePath
        Print #1, ln
        ln = "groupalpha=" & GroupAlpha
        Print #1, ln
        ln = "ticksalpha=" & TicksAlpha
        Print #1, ln
        ln = "colorscheme=" & ColorScheme
        Print #1, ln
        ln = "colorsat=" & ColorSat
        Print #1, ln
        ln = "antialiasing=" & B2S(AntiAliasing)
        Print #1, ln
        ln = "split=" & UISplitY
        Print #1, ln
    Close #1
End Sub
