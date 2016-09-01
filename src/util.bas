Attribute VB_Name = "UtilMd"

Sub SetGLColor(Color As TGLByteColor)
    glColor4b Color.Red, Color.Green, Color.Blue, Color.Alpha
End Sub

Sub InitColor(Color As TGLByteColor, Red As Byte, Green As Byte, Blue As Byte, Alpha As Byte, ApplySat As Boolean)
    If ApplySat = True Then
        Color.Red = Saturate(Red, ColorSat)
        Color.Green = Saturate(Green, ColorSat)
        Color.Blue = Saturate(Blue, ColorSat)
    Else
        Color.Red = Red
        Color.Green = Green
        Color.Blue = Blue
    End If
    Color.Alpha = Alpha
End Sub

Function Saturate(ByVal Inp As Single, ByVal Sat As Single) As Integer
    Sat = Sat * 1.27    ' 100 to 127
    Saturate = ((Inp - 63) * Sat / 128) + 63
End Function

Sub LoadColorScheme()
    If ColorScheme = 0 Then
        ' Default
        InitColor Color_Ticks, 0, 0, 0, CByte(TicksAlpha), False
        InitColor Color_Names, 0, 0, 0, 127, False
        InitColor Color_Waves, 0, 0, 0, 127, False
        InitColor Color_Background, 127, 127, 127, 127, False
    ElseIf ColorScheme = 1 Then
        ' Inverted
        InitColor Color_Ticks, 127, 127, 127, CByte(TicksAlpha), False
        InitColor Color_Names, 127, 127, 127, 127, False
        InitColor Color_Waves, 127, 127, 127, 127, False
        InitColor Color_Background, 0, 0, 0, 127, False
    End If
End Sub

Sub Redraw()
    Render
    UpdateDisplay
End Sub

Function Confirm() As Boolean
    Confirm = False
    If MsgBox("Discard unsaved data ?", vbYesNo + vbQuestion) = vbYes Then Confirm = True
End Function

Function GetFileName(fn As String)
    Dim pos As Integer
    
    pos = 1
    While InStr(pos, fn, "\") > 0
        pos = InStr(pos, fn, "\") + 1
    Wend
    
    GetFileName = Right(fn, Len(fn) - pos + 1)
End Function

Public Sub SetDataColor(DataColor As Integer, Alpha As Integer)
    Dim SatC As Integer
    Dim SatZ As Integer
    
    SatC = Saturate(127, ColorSat)
    SatZ = Saturate(0, ColorSat)
    
    If DataColor = 0 Then
        glColor4b SatC, SatC, SatC, Alpha       ' Grey (neutral)
    ElseIf DataColor = 1 Then
        glColor4b SatC, SatZ, SatZ, Alpha       ' Red
    ElseIf DataColor = 2 Then
        glColor4b SatZ, SatC, SatZ, Alpha       ' Green
    ElseIf DataColor = 3 Then
        glColor4b SatZ, SatZ, SatC, Alpha       ' Blue
    ElseIf DataColor = 4 Then
        glColor4b SatC, SatC, SatZ, Alpha       ' Yellow
    ElseIf DataColor = 5 Then
        glColor4b SatZ, SatC, SatC, Alpha       ' Cyan
    ElseIf DataColor = 6 Then
        glColor4b SatC, SatZ, SatC, Alpha       ' Purple
    ElseIf DataColor = 7 Then
        glColor4b SatC, SatC, SatC, Alpha       ' Grey
    Else
        SetGLColor Color_Waves
    End If
End Sub

Function NextIsDot(FieldData As String, c As Integer) As Boolean
    If c <= Len(FieldData) Then
        If Mid(FieldData, c + 2, 1) = "." Then
            NextIsDot = True
        Else
            NextIsDot = False
        End If
    Else
        NextIsDot = False
    End If
End Function

Function MatchT(ByVal s As String) As Integer
    If s = "SP" Then MatchT = 0     ' Start point
    If s = "EP" Then MatchT = 1     ' End point
    If s = "L" Then MatchT = 2      ' Line
    If s = "LS" Then MatchT = 3     ' Line strip
    If s = "SH" Then MatchT = 4     ' Polygon
    If s = "LC" Then MatchT = 6     ' Line color
End Function

Function S2B(Ins As String) As Boolean
    If Ins = "1" Or LCase(Ins) = "true" Then
        S2B = True
    Else
        S2B = False
    End If
End Function

Function B2S(Ins As Boolean) As String
    If Ins = True Then
        B2S = "1"
    Else
        B2S = "0"
    End If
End Function

Function B2C(Ins As Boolean) As CheckBoxConstants
    If Ins = True Then
        B2C = vbChecked
    Else
        B2C = vbUnchecked
    End If
End Function

Function C2B(Ins As CheckBoxConstants) As Boolean
    If Ins = vbChecked Then
        C2B = True
    Else
        C2B = False
    End If
End Function

Function ErrorBox(Msg As String, Quit As Boolean)
    If Quit = False Then
        MsgBox Msg, vbExclamation, "Error"
    Else
        MsgBox Msg & vbCrLf & "Waimea will close.", vbCritical, "Error"
        End
    End If
End Function
