Attribute VB_Name = "UtilMd"
Sub Redraw()
    Render
    Display
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
    If DataColor = 0 Then
        glColor4b 0, 0, 0, Alpha    ' Black
    ElseIf DataColor = 1 Then
        glColor4b 127, 0, 0, Alpha  ' Red
    ElseIf DataColor = 2 Then
        glColor4b 0, 127, 0, Alpha  ' Green
    ElseIf DataColor = 3 Then
        glColor4b 0, 0, 127, Alpha  ' Blue
    ElseIf DataColor = 4 Then
        glColor4b 127, 127, 0, Alpha    ' Yellow
    ElseIf DataColor = 5 Then
        glColor4b 0, 127, 127, Alpha    ' Cyan
    ElseIf DataColor = 6 Then
        glColor4b 63, 0, 127, Alpha     ' Purple
    ElseIf DataColor = 7 Then
        glColor4b 100, 100, 100, Alpha  ' Grey
    Else
        glColor4b 0, 0, 0, Alpha    ' Default
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
End Function

