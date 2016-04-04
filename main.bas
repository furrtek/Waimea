Attribute VB_Name = "MainMd"
Option Explicit

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
    Show As Boolean
    Txt As String
End Type

Public FontTex As GLuint
Public PinTex As GLuint

Public HasName As Boolean
Public PinList(255) As TPin
Public nPins As Integer

' Settings
Public LiveRefresh As Boolean
Public Spacing As Single

Public xdraw As Integer
Public ydraw As Integer

Public nWaves As Integer

Public Nav_X As Integer
Public Nav_Y As Integer

' Display Lists
Public PinDL As GLuint
Public TicksDL As GLuint
Public WaveDL(256) As GLuint    ' Blocks
Public CharDL(128) As GLuint    ' Characters

Public DispLists(256) As WDispList  ' For blocks

Public FilePath As String

Public XMargin As Integer

Public Saved As Boolean
Public datatxt() As String


Public Keys(255) As Boolean             ' used to keep track of key_downs

Private hrc As Long

Public Sub ReSizeGLScene(ByVal Width As GLsizei, ByVal Height As GLsizei)
    If Height = 0 Then Height = 1
    If Width = 0 Then Width = 1
    
    glViewport 0, 150, Width, Height - 150 ' Reset The Current Viewport
    glMatrixMode mmProjection       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix

    glOrtho 0#, Width, Height - 150, 0#, -1, 1

    glMatrixMode mmModelView        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
    
    RenderTicks
End Sub

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

Public Function InitGL() As Boolean
    glEnable glcTexture2D               ' Enable Texture Mapping ( NEW )
    glShadeModel smFlat

    glLineWidth 1

    glShadeModel GL_SMOOTH

    glClearDepth 1#                     ' Depth Buffer Setup
    glHint GL_LINE_SMOOTH_HINT, GL_NICEST
    glEnable GL_LINE_SMOOTH
    
    glEnable GL_BLEND

    InitGL = True                       ' Initialization Went OK
End Function

Public Sub LoadFont()
    Dim FontData() As GLbyte
    Dim bd As Byte
    Dim h, w As Integer
    
    ReDim FontData(3, 511, 255)
    
    Open "font.tga" For Binary As #1
        Seek #1, 19
        Get #1, , FontData
    Close #1

    glGenTextures 1, FontTex
    glBindTexture glTexture2D, FontTex
    glTexImage2D glTexture2D, 0, 4, 512, 256, 0, tiRGBA, GL_UNSIGNED_BYTE, FontData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear Filtering
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR     ' Linear Filtering
    
    Erase FontData
End Sub

Public Sub LoadPin()
    Dim PinData() As GLbyte
    Dim bd As Byte
    Dim h, w As Integer
    
    ReDim PinData(3, 63, 63)
    
    Open "pin.tga" For Binary As #1
        Seek #1, 19
        Get #1, , PinData
    Close #1

    glGenTextures 1, PinTex
    glBindTexture glTexture2D, PinTex
    glTexImage2D glTexture2D, 0, 4, 64, 64, 0, tiRGBA, GL_UNSIGNED_BYTE, PinData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear Filtering
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR     ' Linear Filtering
    
    Erase PinData
End Sub

Public Sub KillGLWindow()
    If hrc Then                                     ' Do We Have A Rendering Context?
        If wglMakeCurrent(0, 0) = 0 Then             ' Are We Able To Release The DC And RC Contexts?
            MsgBox "Release Of DC And RC Failed.", vbInformation, "SHUTDOWN ERROR"
        End If

        If wglDeleteContext(hrc) = 0 Then           ' Are We Able To Delete The RC?
            MsgBox "Release Rendering Context Failed.", vbInformation, "SHUTDOWN ERROR"
        End If
        hrc = 0                                     ' Set RC To NULL
    End If

    ' Note
    ' The form owns the device context (hDC) window handle (hWnd) and class (RTThundermain)
    ' so we do not have to do all the extra work

End Sub

Public Function CreateGLWindow(frm As Form, Width As Integer, Height As Integer, Bits As Integer) As Boolean
    Dim PixelFormat As GLuint                       ' Holds The Results After Searching For A Match
    Dim pfd As PIXELFORMATDESCRIPTOR                ' pfd Tells Windows How We Want Things To Be

    pfd.cColorBits = Bits
    pfd.cDepthBits = 16
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    
    PixelFormat = ChoosePixelFormat(GetDC(frm.hWnd), pfd)
    If PixelFormat = 0 Then                     ' Did Windows Find A Matching Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Find A Suitable PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If SetPixelFormat(GetDC(frm.hWnd), PixelFormat, pfd) = 0 Then ' Are We Able To Set The Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Set The PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                           ' Return FALSE
    End If
    
    hrc = wglCreateContext(GetDC(frm.hWnd))
    If (hrc = 0) Then                           ' Are We Able To Get A Rendering Context?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Create A GL Rendering Context.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If wglMakeCurrent(GetDC(frm.hWnd), hrc) = 0 Then    ' Try To Activate The Rendering Context
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Activate The GL Rendering Context.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If
    frm.Show                                    ' Show The Window
    SetForegroundWindow frm.hWnd                ' Slightly Higher Priority
    frm.SetFocus                                ' Sets Keyboard Focus To The Window
    ReSizeGLScene frm.ScaleWidth, frm.ScaleHeight ' Set Up Our Perspective GL Screen

    If Not InitGL() Then                        ' Initialize Our Newly Created GL Window
        KillGLWindow                            ' Reset The Display
        MsgBox "Initialization Failed.", vbExclamation, "ERROR"
        CreateGLWindow = False                   ' Return FALSE
    End If

    CreateGLWindow = True                       ' Success
End Function

Public Sub Display()
    Dim w As Integer
    Dim PinTextLen As Integer
    
    glMatrixMode mmModelView        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
    
    glClearColor 1, 1, 1, 1
    glClear clrColorBufferBit Or clrDepthBufferBit
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glHint GL_LINE_SMOOTH_HINT, GL_NICEST
    
    glTranslatef Nav_X, Nav_Y, 0#
    glPushMatrix
    
    ' Draw ticks
    glColor4b 0, 0, 0, 31
    glScalef Spacing, 1, 1
    glCallList TicksDL
    
    For w = 0 To nWaves
        If glIsList(WaveDL(w)) = 1 Then glCallList WaveDL(w)
    Next w
    
    glPopMatrix
    glPushMatrix
    
    ' Draw pins
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glEnable glcTexture2D
    glBindTexture glTexture2D, PinTex
    For w = 0 To nPins - 1
        glTranslatef PinList(w).X, PinList(w).Y, 0
        SetDataColor PinList(w).Color, 127
        glCallList PinDL
    Next w
    glDisable glcTexture2D
    
    glPopMatrix
    ' Draw pin bubble if needed
    For w = 0 To nPins - 1
        If PinList(w).Show = True Then
            PinTextLen = Len(PinList(w).Txt) * 8 + 16
            glTranslatef PinList(w).X + 10, PinList(w).Y + 10, 0
            glColor4b 120, 120, 120, 120
            glBegin bmPolygon
                glVertex2f 0, 0
                glVertex2f 8, -4
                glVertex2f 8, -8
                glVertex2f PinTextLen, -8
                glVertex2f PinTextLen, 8
                glVertex2f 8, 8
                glVertex2f 8, 4
                glVertex2f 0, 0
            glEnd
            SetDataColor PinList(w).Color, 100
            glBegin bmLineStrip
                glVertex2f 0, 0
                glVertex2f 8, -4
                glVertex2f 8, -8
                glVertex2f PinTextLen, -8
                glVertex2f PinTextLen, 8
                glVertex2f 8, 8
                glVertex2f 8, 4
                glVertex2f 0, 0
            glEnd
            SetDataColor PinList(w).Color, 127
            dr_text PinList(w).Txt, 12, -8
            Exit For    ' Only one can be shown at a time
        End If
    Next w

    SwapBuffers MainFrm.hDC
End Sub

Public Sub Render(frm As Form)
    Dim c As Integer
    Dim xpos As Single
    Dim d As Integer
    Dim wavedef As String
    Dim waves() As String
    Dim fields() As String
    Dim w As Integer

    wavedef = MainFrm.Text1.Text
    
    ' For each line...
    waves = Split(wavedef, vbCrLf)
    nWaves = UBound(waves)
    nPins = 0
    For w = 0 To nWaves
        ReDim datatxt(1)
        datatxt(0) = ""
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
        ReDim datatxt(UBound(DF))
        For f = 0 To UBound(DF)
            datatxt(f) = DF(f)
        Next f
    End If
End Sub

Sub RenderPin(FieldData As String, YPos As Integer)
    Dim DF() As String
    Dim f As Integer
    
    DF = Split(FieldData, ",")
    If UBound(DF) > 1 Then
        If HasName = True Then  ' Not necessary
            PinList(nPins).X = 15 * DF(0) - 10 + XMargin
            PinList(nPins).Y = YPos - 9
            PinList(nPins).Color = Val(DF(1))
            PinList(nPins).Txt = DF(2)
            nPins = nPins + 1
        End If
    End If
End Sub

Sub RenderName(FieldData As String)
    glColor4b 0, 0, 0, 127
    dr_text FieldData, -((Len(FieldData) * 8) - XMargin + 4), 0
    HasName = True
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

Sub ProcessFields(fields() As String, TypeMatch As String, w As Integer)
    Dim c As Integer
    Dim st As Integer
    Dim blk As String * 1
    Dim ch As String * 1
    Dim AscP As Integer
    Dim lastblk As String * 1
    Dim pblk As String * 1
    Dim steps As Integer
    Dim wavedef As String
    Dim DataColor As Integer
    Dim DataState As Integer
    Dim dstart As Integer
    Dim eq, f, d As Integer
    Dim FieldType As String
    Dim FieldData As String
    Dim Found As Boolean
    Dim l As Integer
    Dim dti As Integer
    Dim DF() As String
    Dim lastd As Integer
    Dim DAlpha As Integer
    Dim sx, sy, ex, ey As Integer
    
    Found = False
    For f = 0 To UBound(fields)
        eq = InStr(1, fields(f), ":")   ' Field has type:data pair ?
        If eq > 0 Then
            fields(f) = Replace(fields(f), Chr(9), " ")    ' Tab to space
            FieldType = Trim(LCase(Left(fields(f), eq - 1)))
            FieldData = Trim(Right(fields(f), Len(fields(f)) - eq))
            If FieldType = TypeMatch Then
                Found = True
                Exit For
            End If
        End If
    Next f
        
    If Found = False Then Exit Sub
    
    If FieldType = "ruler" Then
        RenderRuler FieldData
    ElseIf FieldType = "name" Then
        RenderName FieldData
    ElseIf FieldType = "data" Then
        RenderData FieldData
    ElseIf FieldType = "pin" Then
        RenderPin FieldData, w * 20
    ElseIf FieldType = "wave" Then
        lastblk = "z"   ' Default block
        DataState = -1
        dti = 0
        
        glPushMatrix
        glTranslatef XMargin, 0#, 0#
        
        For c = 0 To Len(FieldData) - 1
            ' Draw
            xdraw = XMargin + (c * 15)
            
            pblk = Mid(FieldData, c + 1, 1)
            blk = pblk
            
            If pblk = "." Then
                pblk = lastblk     ' Repeat
                blk = pblk
            End If
            
            AscP = Asc(pblk)
            If (AscP >= &H30 And AscP <= &H36) Or pblk = "=" Then   ' Start data
                If DataState = -1 Then
                    If pblk = "=" Then
                        DataColor = 0
                    Else
                        DataColor = Asc(pblk) - &H30
                    End If
                    If Not NextIsDot(FieldData, c) Then
                        blk = "u"
                        dstart = xdraw
                        DataState = -2
                    Else
                        DataState = 0
                    End If
                    DAlpha = 31
                End If
            Else
                DataColor = 0
                DAlpha = 127
            End If
            
            If DataState = 0 Then
                dstart = xdraw
                blk = "s"
            End If
            If DataState = 1 Then blk = "d"
            If DataState > -1 And Not NextIsDot(FieldData, c) Then    ' Dangerous (+1 into void ?)
                blk = "e"
                DataState = -2
            End If
            
            ' Look for block def
            If pblk <> "." Then
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
            
            glColor4b 0, 0, 0, 127
            
            ' Transition
            If c > 0 Then
                sx = DispLists(lastd).EP.X - 15
                sy = DispLists(lastd).EP.Y
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
                If (datatxt(0) <> "") Then
                    If dti <= UBound(datatxt) Then
                        SetDataColor DataColor, 127
                        dr_text datatxt(dti), -(((xdraw - dstart) / 2) + (Len(datatxt(dti)) * 4)) + 7, 0
                        dti = dti + 1
                    End If
                End If
                DataState = -1
            End If
            
            lastblk = pblk
            lastd = d
            
            glTranslatef 15, 0, 0
        Next c
        
        glPopMatrix
    End If
End Sub

Sub dr_text(Txt As String, xofs As Integer, yofs As Integer)
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

Function MatchT(ByVal s As String) As Integer
    If s = "SP" Then MatchT = 0     ' Start point
    If s = "EP" Then MatchT = 1     ' End point
    If s = "L" Then MatchT = 2      ' Line
    If s = "LS" Then MatchT = 3     ' Line strip
    If s = "SH" Then MatchT = 4     ' Polygon
End Function
