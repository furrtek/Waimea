Attribute VB_Name = "MainMd"
Option Explicit

Public Type TGLByteColor
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
End Type

Private Type DCoord
    X As Integer
    Y As Integer
End Type

Private Type WDispList
    DL As GLuint
    Char As String * 1
    SP As DCoord
    EP As DCoord
End Type

Private Type TPin
    X As Integer
    Y As Integer
    Color As Integer
    Txt As String
    Show As Boolean
End Type

Private Type TGroup
    Start As Integer
    Stop As Integer
    Level As Integer
    Color As Integer
    Txt As String
End Type

Private Type TWave
    DL As GLuint
    Used As Boolean
    Name As String
End Type

Private Type TRuler
    X As Integer
    Color As Integer
End Type

Public FSO As FileSystemObject

Public Loaded As Boolean

Public FontTex As GLuint
Public PinTex As GLuint

Public PinList(255) As TPin
Public GroupStack(256) As TGroup
Public Rulers(256) As TRuler

Public UsedWave As Boolean

Public nPins As Integer
Public nRulers As Integer
Public nWaves As Integer
Public GIdxAdd As Integer
Public GLevel As Integer

Public WaveName As String

' Settings
Public Spacing As Single
Public LiveRefresh As Boolean
Public AltBubbles As Boolean
Public OpenLast As Boolean
Public GroupAlpha As Integer
Public TicksAlpha As Integer
Public ColorScheme As Integer
Public AntiAliasing As Boolean
Public LastOpened As String

Public Color_Background As TGLByteColor
Public Color_Ticks As TGLByteColor
Public Color_Waves As TGLByteColor
Public Color_Names As TGLByteColor

Public MaxWidth As Integer
Public MaxHeight As Integer

Public Nav_X As Integer
Public Nav_Y As Integer

' Display Lists
Public EverythingDL As GLuint
Public PinDL As GLuint
Public TicksDL As GLuint
Public Waves(256) As TWave    ' Blocks
Public CharDL(128) As GLuint    ' Characters
Public DispLists(256) As WDispList  ' For blocks

Public FilePath As String

Public XMargin As Integer
Public YMargin As Integer

Public Saved As Boolean
Public DataTxt() As String

Public Keys(255) As Boolean

Private hrc As Long

Public Sub Display()
    glClearColor Color_Background.Red / 127, Color_Background.Green / 127, Color_Background.Blue / 127, 1
    glClear clrColorBufferBit Or clrDepthBufferBit
    
    If AntiAliasing = True Then
        glEnable glcLineSmooth
    Else
        glDisable glcLineSmooth
    End If
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    
    glMatrixMode mmModelView
    glLoadIdentity
    glTranslatef Nav_X, Nav_Y, 0#
    
    glCallList EverythingDL
    
    SwapBuffers MainFrm.Picture1.hDC
End Sub

Public Sub UpdateDisplay()


Dim w As Integer
    Dim PinTextLen As Integer
    
    If Loaded = False Then Exit Sub
    
    ' Fill EverythingDL
    If glIsList(EverythingDL) = GL_TRUE Then glDeleteLists EverythingDL, 1
    EverythingDL = glGenLists(1)
    glNewList EverythingDL, lstCompile
    
    glScalef Spacing, 1, 1
    glPushMatrix                            ' Initial matrix (nav, spacing)
    
    ' Draw ticks
    SetGLColor Color_Ticks
    glCallList TicksDL
    
    ' Draw rulers
    glPopMatrix                             ' Restore initial matrix
    glPushMatrix
    For w = 0 To nRulers - 1
        SetDataColor Rulers(w).Color, 63
        glBegin bmLines
            glVertex2d Rulers(w).X, 0
            glVertex2d Rulers(w).X, MaxHeight
        glEnd
    Next w
    
    glPopMatrix                             ' Restore initial matrix
    glTranslatef 0, YMargin, 0
    glPushMatrix                            ' Add Y margin, new origin
    
    ' Draw groups
    For w = 0 To GIdxAdd - 1
        ' Shapes
        glPopMatrix                         ' Restore
        glPushMatrix
        glTranslatef GroupStack(w).Level * 14, 0, 0
        glPushMatrix                        ' Level offset
        SetDataColor GroupStack(w).Color, GroupAlpha
        glBegin bmPolygon
            glVertex2f 0, GroupStack(w).Start
            glVertex2f MaxWidth, GroupStack(w).Start
            glVertex2f MaxWidth, GroupStack(w).Stop
            glVertex2f 0, GroupStack(w).Stop
        glEnd
        
        ' Text
        glPopMatrix
        SetDataColor GroupStack(w).Color, 127
        glTranslatef 0, ((GroupStack(w).Stop + GroupStack(w).Start) / 2) + ((Len(GroupStack(w).Txt) * 7) / 2), 0
        glRotatef -90, 0, 0, 1
        RenderText GroupStack(w).Txt, 2, 0, 0.8
    Next w

    ' Draw waves
    glPopMatrix                             ' Restore
    glPushMatrix
    For w = 0 To nWaves - 1
        If Waves(w).Used = True Then glCallList Waves(w).DL
    Next w

    ' Draw pins
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glEnable glcTexture2D
    glBindTexture glTexture2D, PinTex
    For w = 0 To nPins - 1
        glPopMatrix                         ' Restore
        glPushMatrix
        glTranslatef PinList(w).X - 10, PinList(w).Y, 0
        SetDataColor PinList(w).Color, 127
        glCallList PinDL
    Next w
    glDisable glcTexture2D
    
    ' Draw pin bubble if needed
    For w = 0 To nPins - 1
        If PinList(w).Show = True Then
            glPopMatrix                     ' Restore
            glPushMatrix
            PinTextLen = Len(PinList(w).Txt) * 8 + 16
            glTranslatef PinList(w).X + 8, PinList(w).Y + 10, 0
            glColor4b 120, 120, 120, 90
            glBegin bmPolygon
                glVertex2f 0, 0
                glVertex2f 8, -8
                glVertex2f PinTextLen, -8
                glVertex2f PinTextLen, 8
                glVertex2f 8, 8
                glVertex2f 0, 0
            glEnd
            SetDataColor PinList(w).Color, 70
            glBegin bmLineStrip
                glVertex2f 0, 0
                glVertex2f 8, -8
                glVertex2f PinTextLen, -8
                glVertex2f PinTextLen, 8
                glVertex2f 8, 8
                glVertex2f 0, 0
            glEnd
            SetDataColor PinList(w).Color, 127
            RenderText PinList(w).Txt, 12, -8, 1
        End If
    Next w
    
    glPopMatrix
    
    glEndList

    Display
End Sub

Sub ProcessFields(Fields() As String, TypeMatch As String, w As Integer)
    Dim XDraw As Integer
    Dim YDraw As Integer
    Dim c As Integer
    Dim st As Integer
    Dim blk As String * 1
    Dim ch As String * 1
    Dim AscP As Integer
    Dim LastBlk As String * 1
    Dim PBlk As String * 1
    Dim Steps As Integer
    Dim WaveDef As String
    Dim DataColor As Integer
    Dim DataState As Integer
    Dim DStart As Integer
    Dim eq, f, d As Integer
    Dim FieldType As String
    Dim FieldData As String
    Dim Found As Boolean
    Dim l As Integer
    Dim dti As Integer
    Dim DF() As String
    Dim LastD As Integer
    Dim DAlpha As Integer
    Dim sx, sy, ex, ey As Integer
    
    Found = False
    For f = 0 To UBound(Fields)
        eq = InStr(1, Fields(f), ":")   ' Field has type:data pair ?
        If eq > 0 Then
            Fields(f) = Replace(Fields(f), Chr(9), " ")    ' Tab to space
            FieldType = Trim(LCase(Left(Fields(f), eq - 1)))
            FieldData = Trim(Right(Fields(f), Len(Fields(f)) - eq))
        Else
            ' Line only has field with no data
            FieldType = Trim(LCase(Fields(f)))
        End If
        If FieldType = TypeMatch Then
            Found = True
            Exit For
        End If
    Next f
        
    If Found = False Then Exit Sub
    
    If FieldType = "ruler" Then
        RenderRuler FieldData
    ElseIf FieldType = "name" Then
        RenderName FieldData
        UsedWave = True
    ElseIf FieldType = "data" Then
        RenderData FieldData
    ElseIf FieldType = "pin" Then
        RenderPin FieldData, w * 20
    ElseIf FieldType = "group" Then
        DF = Split(FieldData, ",")
        If UBound(DF) >= 1 Then
            With GroupStack(GIdxAdd)
                .Level = GLevel
                .Start = (w * 20) - 4
                .Txt = DF(0)
                .Color = Val(DF(1))
                .Stop = -1
            End With
            GIdxAdd = GIdxAdd + 1
            GLevel = GLevel + 1
        End If
    ElseIf FieldType = "groupend" Then
        ' Go back up the groupstack to see what was the last started group
        For c = GIdxAdd - 1 To 0 Step -1
            If GroupStack(c).Stop = -1 Then
                GroupStack(c).Stop = (w * 20) - 2
                GLevel = GLevel - 1
                Exit For
            End If
        Next c
    ElseIf FieldType = "wave" Then
        UsedWave = True
        LastBlk = "z"   ' Default block
        DataState = -1
        dti = 0

        glPushMatrix
        glTranslatef XMargin, 0, 0
        
        For c = 0 To Len(FieldData) - 1
            ' Draw
            XDraw = XMargin + (c * 15)
            
            PBlk = Mid(FieldData, c + 1, 1)
            blk = PBlk
            
            If PBlk = "." Then
                If LastBlk = "H" Then
                    PBlk = "h"          ' Adapt
                ElseIf LastBlk = "L" Then
                    PBlk = "l"          ' Adapt
                Else
                    PBlk = LastBlk      ' Repeat
                End If
                blk = PBlk
            End If
            
            AscP = Asc(PBlk)
            If (AscP >= &H30 And AscP <= &H36) Or PBlk = "=" Then   ' Start data
                If DataState = -1 Then
                    If PBlk = "=" Then
                        DataColor = 0
                    Else
                        DataColor = Asc(PBlk) - &H30
                    End If
                    If Not NextIsDot(FieldData, c) Then
                        blk = "u"
                        DStart = XDraw
                        DataState = -2
                    Else
                        DataState = 0
                    End If
                    DAlpha = 31
                End If
            Else
                DataColor = -1
                DAlpha = 127
            End If
            
            If DataState = 0 Then
                DStart = XDraw
                blk = "s"
            End If
            If DataState = 1 Then blk = "d"
            If DataState > -1 And Not NextIsDot(FieldData, c) Then
                blk = "e"
                DataState = -2
            End If
            
            ' Look for block def
            If PBlk <> "." Then
                d = 0
                Do While True
                    ch = DispLists(d).Char
                    If ch = " " Then
                        d = 0
                        Exit Do
                    End If
                    If ch = blk Then Exit Do
                    d = d + 1
                Loop
            End If
            
            SetGLColor Color_Waves
            
            ' Transition
            If c > 0 Then
                sx = DispLists(LastD).EP.X - 15
                sy = DispLists(LastD).EP.Y
                ex = DispLists(d).SP.X
                ey = DispLists(d).SP.Y
                If (sx <> ex) Or (sy <> ey) Then
                    glBegin bmLines
                        glVertex2d sx, sy
                        glVertex2d ex, ey
                    glEnd
                End If
            End If
            
            SetDataColor DataColor, DAlpha
            glCallList DispLists(d).DL
            
            If DataState = 0 Then DataState = 1
            If DataState = -2 Then
                If (DataTxt(0) <> "") Then
                    If dti <= UBound(DataTxt) Then
                        SetDataColor DataColor, 127
                        RenderText DataTxt(dti), -(((XDraw - DStart) / 2) + (Len(DataTxt(dti)) * 4)) + 7, 0, 1
                        dti = dti + 1
                    End If
                End If
                DataState = -1
            End If
            
            LastBlk = PBlk
            LastD = d
            
            glTranslatef 15, 0, 0
        Next c
        
        If XDraw > MaxWidth Then MaxWidth = XDraw
        
        glPopMatrix
    End If
End Sub
