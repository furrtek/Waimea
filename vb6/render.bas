Attribute VB_Name = "RenderMd"

Sub RenderTicks()
    Dim XPos As Single

    glNewList TicksDL, GL_COMPILE
        glBegin bmLines
            XPos = 0
            While (XPos < MaxWidth)
                glVertex2d XPos, 0
                glVertex2d XPos, MaxHeight
                XPos = XPos + 15
            Wend
        glEnd
    glEndList
    
    ' Render colored group stripes
    ' This belongs elsewhere
    MakeNewDL BackDL
    glNewList BackDL, lstCompile
        SetGLColor Color_Names
        For w = 0 To GIdxAdd - 1
            ' Shapes
            glPopMatrix                         ' Restore
            glPushMatrix
            glTranslatef GroupStack(w).Level * 14, 0, 0
            'glPushMatrix                        ' Level offset
            SetDataColor GroupStack(w).Color, GroupAlpha
            glBegin bmPolygon
                glVertex2f 0, GroupStack(w).Start
                glVertex2f MaxWidth, GroupStack(w).Start
                glVertex2f MaxWidth, GroupStack(w).Stop
                glVertex2f 0, GroupStack(w).Stop
            glEnd
        Next w
    glEndList
End Sub

Public Sub Render()
    Dim w As Integer
    Dim WaveDef As String
    Dim Lines() As String
    Dim Fields() As String

    WaveDef = MainFrm.EditBox.Text
    
    ' For each line...
    Lines = Split(WaveDef, vbCrLf)
    nPins = 0
    nWaves = 0
    nRulers = 0
    GIdxAdd = 0
    GLevel = 0
    MaxWidth = XMargin + 32
    MaxHeight = YMargin
    For w = 0 To UBound(Lines)
        ReDim DataTxt(1)
        DataTxt(0) = ""
        UsedWave = False
        Waves(nWaves).Used = False
        Fields = Split(Lines(w), ";")
        'If glIsList(Waves(nWaves).DL) = GL_TRUE Then
        '    glDeleteLists Waves(nWaves).DL, 1
        'End If
        If UBound(Fields) >= 0 Then
            glNewList Waves(nWaves).DL, lstCompile   ' Parse priority is important here
                ProcessFields Fields, "group", nWaves
                ProcessFields Fields, "groupend", nWaves
                ProcessFields Fields, "name", nWaves
                ProcessFields Fields, "data", nWaves
                ProcessFields Fields, "ana", nWaves
                ProcessFields Fields, "wave", nWaves
                ProcessFields Fields, "ruler", nWaves
                ProcessFields Fields, "pin", nWaves
                If UsedWave = True Then
                    Waves(nWaves).Used = True
                    Waves(nWaves).Name = WaveName
                    nWaves = nWaves + 1
                    glTranslatef 0#, 20#, 0#    ' Next line
                    MaxHeight = MaxHeight + 20
                End If
            glEndList
        End If
    Next w
    
    MaxWidth = MaxWidth + 15
    
    ' Regen names
    RenderNames
    
    ' Regen ticks
    RenderTicks
End Sub


Sub RenderRuler(FieldData As String)
    Dim DF() As String
    
    DF = Split(FieldData, ",")
    If UBound(DF) = 1 Then
        Rulers(nRulers).X = Val(DF(0)) * 15
        Rulers(nRulers).Color = Val(DF(1))
        nRulers = nRulers + 1
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
    
    DF = Split(FieldData, ",")
    If UBound(DF) > 1 Then
        PinList(nPins).X = 15 * Val(DF(0)) + XMarginCompensated
        PinList(nPins).Y = YPos - 9
        PinList(nPins).Color = Val(DF(1))
        PinList(nPins).Txt = DF(2)
        nPins = nPins + 1
    End If
End Sub

Function GetTextDisplayWidth(ByVal Text As String) As Integer
    Dim c As Integer
    Dim Width As Integer
    Dim max_width As Integer
    
    Width = 0
    max_width = 0
    For c = 0 To Len(Text) - 1
        If Mid(Text, c + 1, 1) = "\" Then
            If Width > max_width Then max_width = Width
            Width = 0
        Else
            Width = Width + 1
        End If
    Next c
    
    If Width > max_width Then max_width = Width
    
    GetTextDisplayWidth = max_width
End Function

Function GetTextDisplayHeight(ByVal Text As String) As Integer
    Dim c As Integer
    Dim Height As Integer
    
    Height = 0
    For c = 0 To Len(Text) - 1
        If Mid(Text, c + 1, 1) = "\" Then Height = Height + 1
    Next c
    
    GetTextDisplayHeight = Height
End Function

Sub RenderNames()
    Dim w As Integer
    Dim WName As String

    MakeNewDL NamesDL
    glNewList NamesDL, lstCompile
        
        ' Render wave names
        glTranslatef 0, YMargin, 0
        glPushMatrix
        SetGLColor Color_Names
        For w = 0 To nWaves - 1
            If Waves(w).Used = True Then
                WName = Waves(w).Name
                RenderText WName, -((GetTextDisplayWidth(WName) * 8) - XMargin + 4), 0, 1
                glTranslatef 0, 20, 0
            End If
        Next w
    
        ' Render group names
        SetGLColor Color_Names
        For w = 0 To GIdxAdd - 1
            ' Shapes
            glPopMatrix                         ' Restore
            glPushMatrix
            glTranslatef GroupStack(w).Level * 14, 0, 0
        
            ' Text
            SetDataColor GroupStack(w).Color, 127
            glTranslatef 0, ((GroupStack(w).Stop + GroupStack(w).Start) / 2) + ((Len(GroupStack(w).Txt) * 7) / 2), 0
            glRotatef -90, 0, 0, 1
            RenderText GroupStack(w).Txt, 2, 0, 0.8
        Next w
        
        glPopMatrix
    
    glEndList
End Sub

Sub RenderText(Txt As String, Xofs As Integer, YOfs As Integer, Coef As Single)
    Dim pch As Integer
    Dim sx, sy, ex, ey As Single
    Dim c As Integer
    Dim ch As String
    
    glPushMatrix
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glEnable glcTexture2D
    If ColorScheme = 1 Then glTexEnvi tetTextureEnv, tenTextureEnvMode, GL_ADD
    glBindTexture glTexture2D, FontTex
    
    glTranslatef Xofs, YOfs, 0
    glScalef Coef, Coef, 1
    
    glPushMatrix
    
    For c = 0 To Len(Txt) - 1
        ch = Mid(Txt, c + 1, 1)
        If ch = "\" Then
            glPopMatrix
            glTranslatef 0, 14, 0
            glPushMatrix
        Else
            pch = Asc(ch) - 32
            If pch < 128 Then glCallList CharDL(pch)
            glTranslatef 8, 0, 0
        End If
    Next c
    
    glPopMatrix
    
    glDisable glcTexture2D
    glTexEnvi tetTextureEnv, tenTextureEnvMode, GL_MODULATE
    
    glPopMatrix
    
End Sub
