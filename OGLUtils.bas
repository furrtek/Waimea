Attribute VB_Name = "OGLUtils"
Option Explicit

Public Texture(1) As GLuint

Public xdraw As Integer
Public ydraw As Integer

Public RenderTex As GLuint

Public nWaves As Integer

Public Nav_X As Integer
Public Nav_Y As Integer

Public TicksDL As GLuint
Public WaveDL(256) As GLuint
Public CharDL(128) As GLuint

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

Public DispLists(256) As WDispList  ' For blocks

Public FilePath As String

Public xmargin As Integer

Public WaveName As String
Public Saved As Boolean
Public datatxt() As String


Public Keys(255) As Boolean             ' used to keep track of key_downs

Private hrc As Long

Public Sub ReSizeGLScene(ByVal Width As GLsizei, ByVal Height As GLsizei)
    Dim c As Integer
    Dim xpos As Integer
    
    If Height = 0 Then Height = 1
    If Width = 0 Then Width = 1
    
    glViewport 0, 150, Width, Height - 150 ' Reset The Current Viewport
    glMatrixMode mmProjection       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix

    glOrtho 0#, Width, Height - 150, 0#, -1, 1

    glMatrixMode mmModelView        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
    
    If glIsList(TicksDL) = 1 Then glDeleteLists TicksDL, 1
    TicksDL = glGenLists(1)
    glNewList TicksDL, GL_COMPILE
        glBegin bmLines
            For c = 0 To Form1.ScaleWidth - 1 Step 15
                xpos = c + xmargin
                glVertex2d xpos, 0
                glVertex2d xpos, Form1.ScaleHeight
            Next c
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

    ' Font stuff
    glGenTextures 1, Texture(0)
    glBindTexture glTexture2D, Texture(0)
    glTexImage2D glTexture2D, 0, 4, 512, 256, 0, tiRGBA, GL_UNSIGNED_BYTE, FontData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear Filtering
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR     ' Linear Filtering
    
    Erase FontData
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
    pfd.dwflags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
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
    
    glMatrixMode mmModelView        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
    
    glClearColor 1, 1, 1, 1
    glClear clrColorBufferBit Or clrDepthBufferBit
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glHint GL_LINE_SMOOTH_HINT, GL_NICEST
    
    glTranslatef Nav_X, Nav_Y, 0#
    
    ' Draw ticks
    glColor4b 1, 0, 0, 15
    glCallList TicksDL
    
    For w = 0 To nWaves
        If glIsList(WaveDL(w)) = 1 Then glCallList WaveDL(w)
    Next w
    
    SwapBuffers Form1.hDC
End Sub

Public Sub Render(frm As Form)
    Dim c As Integer
    Dim xpos As Integer
    Dim d As Integer
    Dim wavedef As String
    Dim waves() As String
    Dim fields() As String
    Dim w As Integer

    wavedef = Form1.Text1.Text
    
    ' Split lines
    waves = Split(wavedef, vbCrLf)

    nWaves = UBound(waves)
    
    For w = 0 To nWaves
        WaveName = ""
        ReDim datatxt(1)
        datatxt(0) = ""
        fields = Split(waves(w), ";")
        If glIsList(WaveDL(w)) = 1 Then glDeleteLists WaveDL(w), 1
        WaveDL(w) = glGenLists(1)
        glNewList WaveDL(w), GL_COMPILE
            ProcessFields fields, "name", w
            ProcessFields fields, "data", w
            ProcessFields fields, "wave", w
            ProcessFields fields, "ruler", w
            glTranslatef 0#, 20#, 0#
        glEndList
    Next w
    
    'Display
End Sub

Public Sub SetDataColor(DataColor As Integer, Alpha As Integer)
    If DataColor = 0 Then
        glColor4b 0, 0, 0, Alpha
    ElseIf DataColor = 1 Then
        glColor4b 127, 0, 0, Alpha
    ElseIf DataColor = 2 Then
        glColor4b 0, 127, 0, Alpha
    ElseIf DataColor = 3 Then
        glColor4b 0, 0, 127, Alpha
    ElseIf DataColor = 4 Then
        glColor4b 127, 127, 0, Alpha
    ElseIf DataColor = 5 Then
        glColor4b 0, 127, 127, Alpha
    ElseIf DataColor = 6 Then
        glColor4b 63, 0, 127, Alpha
    ElseIf DataColor = 7 Then
        glColor4b 100, 100, 100, Alpha
    Else
    End If
End Sub

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
    Dim df() As String
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
    
    ydraw = 32 + (w * 24)
    
    If FieldType = "ruler" Then
        df = Split(FieldData, ",")
        If UBound(df) = 1 Then
            SetDataColor Val(df(1)), 63
            glBegin bmLines
                glVertex2d Val(df(0)) * 15 + xmargin, 0
                glVertex2d Val(df(0)) * 15 + xmargin, Form1.ScaleHeight
            glEnd
        End If
    End If
    
    If FieldType = "name" Then
        WaveName = FieldData
        glColor4f 0, 0, 0, 1
        dr_text WaveName, -((Len(WaveName) * 8) - xmargin + 4), 0
    End If

    If FieldType = "data" Then
        df = Split(FieldData, ",")
        If UBound(df) >= 0 Then
            ReDim datatxt(UBound(df))
            For f = 0 To UBound(df)
                datatxt(f) = df(f)
            Next f
        End If
    End If

    If FieldType = "wave" Then
        lastblk = "z"   ' Default block
        DataState = -1
        dti = 0
        
        glPushMatrix
        glTranslatef xmargin, 0#, 0#
        
        For c = 0 To Len(FieldData) - 1
            ' Draw
            xdraw = 64 + (c * 15)
            
            pblk = Mid(FieldData, c + 1, 1)
            blk = pblk
            
            If pblk = "." Then
                pblk = lastblk     ' Repeat
                blk = pblk
            End If
            
            AscP = Asc(pblk)
            If (AscP >= &H30 And AscP <= &H35) Or pblk = "=" Then   ' Start data
                If DataState = -1 Then
                    If pblk = "=" Then
                        DataColor = 0
                    Else
                        DataColor = Asc(pblk) - &H31
                    End If
                    If Mid(FieldData, c + 2, 1) <> "." Then    ' Dangerous (+1 into void ?)
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
            If DataState > -1 And Mid(FieldData, c + 2, 1) <> "." Then    ' Dangerous (+1 into void ?)
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

Sub dr_text(txt As String, xofs As Integer, yofs As Integer)
    Dim pch As Integer
    Dim sx, sy, ex, ey As Single
    Dim c As Integer
    
    glPushMatrix
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glEnable glcTexture2D
    glBindTexture glTexture2D, Texture(0)
    
    glTranslatef xofs, yofs, 0
    
    For c = 0 To Len(txt) - 1
        pch = Asc(Mid(txt, c + 1, 1)) - 32
        glCallList CharDL(pch)
        glTranslatef 8, 0, 0
    Next c
    
    glDisable glcTexture2D
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    
    glPopMatrix
    
End Sub

Function MatchT(ByVal s As String) As Integer
    If s = "SP" Then MatchT = 0     ' Start point
    If s = "EP" Then MatchT = 1     ' End point
    If s = "L" Then MatchT = 2      ' Line
    If s = "LS" Then MatchT = 3     ' Line strip
    If s = "SH" Then MatchT = 4     ' Polygon
End Function
