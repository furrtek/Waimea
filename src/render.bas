Attribute VB_Name = "RenderMd"

Sub RenderTicks()
    Dim c As Integer
    Dim XPos As Single
    
    If glIsList(TicksDL) = GL_TRUE Then glDeleteLists TicksDL, 1
    TicksDL = glGenLists(1)
    glNewList TicksDL, GL_COMPILE
        glBegin bmLines
            XPos = XMargin
            While (XPos < MainFrm.ScaleWidth)
                glVertex2d XPos, 0
                glVertex2d XPos, MainFrm.ScaleHeight
                XPos = XPos + 15
            Wend
        glEnd
    glEndList
End Sub

Public Sub Render()
    Dim c As Integer
    Dim w As Integer
    Dim WaveDef As String
    Dim Lines() As String
    Dim Fields() As String

    WaveDef = MainFrm.Text1.Text
    
    ' For each line...
    Lines = Split(WaveDef, vbCrLf)
    nPins = 0
    nWaves = 0
    nGroups = 0
    GIdx = 0
    For w = 0 To UBound(Lines)
        ReDim DataTxt(1)
        DataTxt(0) = ""
        UsedWave = False
        Fields = Split(Lines(w), ";")
        If glIsList(Waves(w).DL) = GL_TRUE Then
            glDeleteLists Waves(w).DL, 1
        End If
        If UBound(Fields) >= 0 Then
            Waves(w).DL = glGenLists(1)
            glNewList Waves(w).DL, GL_COMPILE   ' Parse priority is important here
                ProcessFields Fields, "group", nWaves
                ProcessFields Fields, "groupend", nWaves
                ProcessFields Fields, "name", nWaves
                ProcessFields Fields, "data", nWaves
                ProcessFields Fields, "wave", nWaves
                ProcessFields Fields, "ruler", nWaves
                ProcessFields Fields, "pin", nWaves
                If UsedWave = True Then
                    Waves(nWaves).Used = True
                    nWaves = nWaves + 1
                    glTranslatef 0#, 20#, 0#    ' Next line
                End If
            glEndList
        End If
    Next w
    
    ' Fix groups if needed
    If GIdx > 0 Then
        For c = GIdx - 1 To 0 Step -1
            GroupStack(c).Stop = w - 1
        Next c
        GIdx = 0
    End If
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
        PinList(nPins).X = 15 * DF(0) + XMargin
        PinList(nPins).Y = YPos - 9
        PinList(nPins).Color = Val(DF(1))
        PinList(nPins).Txt = DF(2)
        nPins = nPins + 1
    End If
End Sub

Sub RenderName(FieldData As String)
    glColor4b 0, 0, 0, 127
    RenderText FieldData, -((Len(FieldData) * 8) - XMargin + 4), 0
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
