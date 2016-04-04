Attribute VB_Name = "RenderMd"

Sub RenderTicks()
    Dim c As Integer
    Dim xpos As Single
    
    If glIsList(TicksDL) = 1 Then glDeleteLists TicksDL, 1
    TicksDL = glGenLists(1)
    glNewList TicksDL, GL_COMPILE
        glBegin bmLines
            xpos = XMargin
            While (xpos < MainFrm.ScaleWidth)
                glVertex2d xpos, 0
                glVertex2d xpos, MainFrm.ScaleHeight
                xpos = xpos + 15
            Wend
        glEnd
    glEndList
End Sub

Public Sub Render()
    Dim c As Integer
    Dim xpos As Single
    Dim d As Integer
    Dim WaveDef As String
    Dim waves() As String
    Dim fields() As String
    Dim w As Integer

    WaveDef = MainFrm.Text1.Text
    
    ' For each line...
    waves = Split(WaveDef, vbCrLf)
    nWaves = UBound(waves)
    nPins = 0
    For w = 0 To nWaves
        ReDim DataTxt(1)
        DataTxt(0) = ""
        HasName = False
        fields = Split(waves(w), ";")
        If glIsList(WaveDL(w)) = 1 Then glDeleteLists WaveDL(w), 1
        WaveDL(w) = glGenLists(1)
        glNewList WaveDL(w), GL_COMPILE
            ProcessFields fields, "name", w
            ProcessFields fields, "data", w
            ProcessFields fields, "wave", w
            ProcessFields fields, "ruler", w
            ProcessFields fields, "pin", w
            glTranslatef 0#, 20#, 0#    ' Next line
        glEndList
    Next w
End Sub


Sub RenderRuler(FieldData As String)
    Dim DF() As String
    
    DF = Split(FieldData, ",")
    If UBound(DF) = 1 Then
        glPopMatrix
        glPushMatrix
        SetDataColor Val(DF(1)), 63
        glBegin bmLines
            glVertex2d Val(DF(0)) * 15 + XMargin, 0
            glVertex2d Val(DF(0)) * 15 + XMargin, MainFrm.ScaleHeight
        glEnd
    End If
End Sub

Sub RenderData(FieldData As String)
    Dim DF() As String
    Dim f As Integer
    
    DF = Split(FieldData, ",")
    If UBound(DF) >= 0 Then
        ReDim DataTxt(UBound(DF))
        For f = 0 To UBound(DF)
            DataTxt(f) = DF(f)
        Next f
    End If
End Sub

Sub RenderPin(FieldData As String, YPos As Integer)
    Dim DF() As String
    Dim f As Integer
    
    DF = Split(FieldData, ",")
    If UBound(DF) > 1 Then
        If HasName = True Then  ' Not necessary
            PinList(nPins).X = 15 * DF(0) + XMargin
            PinList(nPins).Y = YPos - 9
            PinList(nPins).Color = Val(DF(1))
            PinList(nPins).Txt = DF(2)
            nPins = nPins + 1
        End If
    End If
End Sub

Sub RenderName(FieldData As String)
    glColor4b 0, 0, 0, 127
    RenderText FieldData, -((Len(FieldData) * 8) - XMargin + 4), 0
    HasName = True
End Sub

Sub RenderText(Txt As String, xofs As Integer, yofs As Integer)
    Dim pch As Integer
    Dim sx, sy, ex, ey As Single
    Dim c As Integer
    
    glPushMatrix
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glEnable glcTexture2D
    glBindTexture glTexture2D, FontTex
    
    glTranslatef xofs, yofs, 0
    
    For c = 0 To Len(Txt) - 1
        pch = Asc(Mid(Txt, c + 1, 1)) - 32
        glCallList CharDL(pch)
        glTranslatef 8, 0, 0
    Next c
    
    glDisable glcTexture2D
    
    glPopMatrix
    
End Sub
