Attribute VB_Name = "OGLUtils"
Option Explicit

Public xdraw As Integer
Public ydraw As Integer
Public lastx As Integer
Public lasty As Integer

Public Texture(1) As GLuint

Private Type DCoord
    x As Integer
    y As Integer
End Type

Private Type DStep
    t As Integer
    P(32) As Integer
    PCount As Integer
End Type

Private Type Lay
    Ch As String * 1
    Drawstep(32) As DStep
    DCount As Integer
    SP As DCoord
    EP As DCoord
End Type

Public Layout(256) As Lay

' a couple of declares to work around some deficiencies of the type library
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
    dmDeviceName        As String * CCDEVICENAME
    dmSpecVersion       As Integer
    dmDriverVersion     As Integer
    dmSize              As Integer
    dmDriverExtra       As Integer
    dmFields            As Long
    dmOrientation       As Integer
    dmPaperSize         As Integer
    dmPaperLength       As Integer
    dmPaperWidth        As Integer
    dmScale             As Integer
    dmCopies            As Integer
    dmDefaultSource     As Integer
    dmPrintQuality      As Integer
    dmColor             As Integer
    dmDuplex            As Integer
    dmYResolution       As Integer
    dmTTOption          As Integer
    dmCollate           As Integer
    dmFormName          As String * CCFORMNAME
    dmUnusedPadding     As Integer
    dmBitsPerPel        As Integer
    dmPelsWidth         As Long
    dmPelsHeight        As Long
    dmDisplayFlags      As Long
    dmDisplayFrequency  As Long
End Type

Public Keys(255) As Boolean             ' used to keep track of key_downs

Private hrc As Long
Private fullscreen As Boolean

Private OldWidth As Long
Private OldHeight As Long
Private OldBits As Long
Private OldVertRefresh As Long

Private mPointerCount As Integer

Private Sub HidePointer()
    ' hide the cursor (mouse pointer)
    mPointerCount = ShowCursor(False) + 1
    Do While ShowCursor(False) >= -1
    Loop
    Do While ShowCursor(True) <= -1
    Loop
    ShowCursor False
End Sub

Private Sub ShowPointer()
    ' show the cursor (mouse pointer)
    Do While ShowCursor(False) >= mPointerCount
    Loop
    Do While ShowCursor(True) <= mPointerCount
    Loop
End Sub

Public Sub ReSizeGLScene(ByVal Width As GLsizei, ByVal Height As GLsizei)
' Resize And Initialize The GL Window
    If Height = 0 Then              ' Prevent A Divide By Zero By
        Height = 1                  ' Making Height Equal One
    End If
    glViewport 0, 0, Width, Height  ' Reset The Current Viewport
    glMatrixMode mmProjection       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix

    ' Calculate The Aspect Ratio Of The Window
    glOrtho 0#, Width, Height, 0#, -1, 1

    glMatrixMode mmModelView        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
End Sub

Public Function InitGL() As Boolean
    glEnable glcTexture2D               ' Enable Texture Mapping ( NEW )
    glShadeModel smSmooth               ' Enables Smooth Shading

    glLineWidth 1
    glClearColor 1#, 1#, 1#, 0.5

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
    
    ReDim FontData(2, 511, 255)
    
    Open "font.tga" For Binary As #1
        Seek #1, 19
        Get #1, , FontData
    Close #1

    ' Font stuff
    glGenTextures 1, Texture(0)
    glBindTexture glTexture2D, Texture(0)
    glTexImage2D glTexture2D, 0, 3, 512, 256, 0, GL_RGB, GL_UNSIGNED_BYTE, FontData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear Filtering
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR     ' Linear Filtering
    
    Erase FontData
End Sub

Public Sub KillGLWindow()
' Properly Kill The Window
    If fullscreen Then                              ' Are We In Fullscreen Mode?
        ResetDisplayMode                            ' If So Switch Back To The Desktop
        ShowPointer                                 ' Show Mouse Pointer
    End If

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

Private Sub SaveCurrentScreen()
    ' Save the current screen resolution, bits, and Vertical refresh
    Dim ret As Long
    ret = CreateIC("DISPLAY", "", "", 0&)
    OldWidth = GetDeviceCaps(ret, HORZRES)
    OldHeight = GetDeviceCaps(ret, VERTRES)
    OldBits = GetDeviceCaps(ret, BITSPIXEL)
    OldVertRefresh = GetDeviceCaps(ret, VREFRESH)
    ret = DeleteDC(ret)
End Sub

Private Function FindDEVMODE(ByVal Width As Integer, ByVal Height As Integer, ByVal Bits As Integer, Optional ByVal VertRefresh As Long = -1) As DEVMODE
    ' locate a DEVMOVE that matches the passed parameters
    Dim ret As Boolean
    Dim i As Long
    Dim dm As DEVMODE
    i = 0
    Do  ' enumerate the display settings until we find the one we want
        ret = EnumDisplaySettings(0&, i, dm)
        If dm.dmPelsWidth = Width And _
            dm.dmPelsHeight = Height And _
            dm.dmBitsPerPel = Bits And _
            ((dm.dmDisplayFrequency = VertRefresh) Or (VertRefresh = -1)) Then Exit Do ' exit when we have a match
        i = i + 1
    Loop Until (ret = False)
    FindDEVMODE = dm
End Function

Private Sub ResetDisplayMode()
    Dim dm As DEVMODE             ' Device Mode
    
    dm = FindDEVMODE(OldWidth, OldHeight, OldBits, OldVertRefresh)
    dm.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    If OldVertRefresh <> -1 Then
        dm.dmFields = dm.dmFields Or DM_DISPLAYFREQUENCY
    End If
    ' Try To Set Selected Mode And Get Results.  NOTE: CDS_FULLSCREEN Gets Rid Of Start Bar.
    If (ChangeDisplaySettings(dm, CDS_FULLSCREEN) <> DISP_CHANGE_SUCCESSFUL) Then
    
        ' If The Mode Fails, Offer Two Options.  Quit Or Run In A Window.
        MsgBox "The Requested Mode Is Not Supported By Your Video Card", , "NeHe GL"
    End If

End Sub

Private Sub SetDisplayMode(ByVal Width As Integer, ByVal Height As Integer, ByVal Bits As Integer, ByRef fullscreen As Boolean, Optional VertRefresh As Long = -1)
    Dim dmScreenSettings As DEVMODE             ' Device Mode
    Dim P As Long
    SaveCurrentScreen                           ' save the current screen attributes so we can go back later
    
    dmScreenSettings = FindDEVMODE(Width, Height, Bits, VertRefresh)
    dmScreenSettings.dmBitsPerPel = Bits
    dmScreenSettings.dmPelsWidth = Width
    dmScreenSettings.dmPelsHeight = Height
    dmScreenSettings.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    If VertRefresh <> -1 Then
        dmScreenSettings.dmDisplayFrequency = VertRefresh
        dmScreenSettings.dmFields = dmScreenSettings.dmFields Or DM_DISPLAYFREQUENCY
    End If
    ' Try To Set Selected Mode And Get Results.  NOTE: CDS_FULLSCREEN Gets Rid Of Start Bar.
    If (ChangeDisplaySettings(dmScreenSettings, CDS_FULLSCREEN) <> DISP_CHANGE_SUCCESSFUL) Then
    
        ' If The Mode Fails, Offer Two Options.  Quit Or Run In A Window.
        If (MsgBox("The Requested Mode Is Not Supported By" & vbCr & "Your Video Card. Use Windowed Mode Instead?", vbYesNo + vbExclamation, "NeHe GL") = vbYes) Then
            fullscreen = False                  ' Select Windowed Mode (Fullscreen=FALSE)
        Else
            ' Pop Up A Message Box Letting User Know The Program Is Closing.
            MsgBox "Program Will Now Close.", vbCritical, "ERROR"
            End                   ' Exit And Return FALSE
        End If
    End If
End Sub

Public Function CreateGLWindow(frm As Form, Width As Integer, Height As Integer, Bits As Integer, fullscreenflag As Boolean) As Boolean
    Dim PixelFormat As GLuint                       ' Holds The Results After Searching For A Match
    Dim pfd As PIXELFORMATDESCRIPTOR                ' pfd Tells Windows How We Want Things To Be


    fullscreen = fullscreenflag                     ' Set The Global Fullscreen Flag


    pfd.cColorBits = Bits
    pfd.cDepthBits = 16
    pfd.dwflags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    
    PixelFormat = ChoosePixelFormat(frm.hDC, pfd)
    If PixelFormat = 0 Then                     ' Did Windows Find A Matching Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Find A Suitable PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If SetPixelFormat(frm.hDC, PixelFormat, pfd) = 0 Then ' Are We Able To Set The Pixel Format?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Set The PixelFormat.", vbExclamation, "ERROR"
        CreateGLWindow = False                           ' Return FALSE
    End If
    
    hrc = wglCreateContext(frm.hDC)
    If (hrc = 0) Then                           ' Are We Able To Get A Rendering Context?
        KillGLWindow                            ' Reset The Display
        MsgBox "Can't Create A GL Rendering Context.", vbExclamation, "ERROR"
        CreateGLWindow = False                  ' Return FALSE
    End If

    If wglMakeCurrent(frm.hDC, hrc) = 0 Then    ' Try To Activate The Rendering Context
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

Sub dr_point(P() As Integer, ofs As Integer)
    glVertex2f P(ofs) + xdraw, P(ofs + 1) + ydraw
End Sub

Public Function DrawGLScene(frm As Form) As Boolean
    Dim c As Integer
    Dim xpos As Integer
    Dim blk As String * 1
    Dim lastblk As String * 1
    Dim pblk As String * 1
    Dim wavedef As String
    Dim d As Integer
    Dim st As Integer
    Dim l As Integer
    Dim steps As Integer
    Dim waves() As String
    Dim w As Integer
    Dim datacolor As Integer
    Dim datastate As Integer
    Dim dstart As Integer
    
    glClear clrColorBufferBit Or clrDepthBufferBit  ' Clear The Screen And The Depth Buffer
    glLoadIdentity                                  ' Reset The View

    glTranslatef -Form1.HScroll1.Value, 0#, 0#
    
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glHint GL_LINE_SMOOTH_HINT, GL_NICEST

    wavedef = frm.Text1.Text

    glColor4b 0, 0, 0, 10
    glBegin bmLines
        For c = 0 To frm.ScaleWidth - 1 Step 15
            xpos = c + 64
            glVertex2d xpos, 0
            glVertex2d xpos, 128
        Next c
    glEnd
    
    waves = Split(wavedef, vbCrLf)
    
    For w = 0 To UBound(waves)
        
        lastblk = "z"   ' Default block
        datastate = -1
    
        For c = 0 To Len(waves(w)) - 1
            ' Draw
            xdraw = 64 + (c * 15)
            ydraw = 32 + (w * 24)
            
            pblk = Mid(waves(w), c + 1, 1)
            blk = pblk
            
            If pblk = "." Then pblk = lastblk     ' Repeat
            
            If (Asc(pblk) >= &H30 And Asc(pblk) <= &H35) Or pblk = "=" Then   ' Start data
                If datastate = -1 Then
                    If pblk = "=" Then
                        datacolor = 6
                    Else
                        datacolor = Asc(pblk) - &H30
                    End If
                    If Mid(waves(w), c + 2, 1) <> "." Then    ' Dangerous (+1 into void ?)
                        blk = "u"
                        dstart = xdraw
                        datastate = -2
                    Else
                        datastate = 0
                    End If
                End If
            End If
            
            If datastate = 0 Then
                dstart = xdraw
                blk = "s"
            End If
            If datastate = 1 Then blk = "d"
            If datastate > -1 And Mid(waves(w), c + 2, 1) <> "." Then    ' Dangerous (+1 into void ?)
                blk = "e"
                datastate = -2
            End If
            
            If datastate = -2 Then
                dr_text Form1.Text2.Text, (xdraw + dstart) / 2, ydraw - 1
                datastate = -1
            End If
            
            ' Look for block def
            d = 0
            Do While 1
                If Layout(d).DCount = 0 Then Exit Do
                If Layout(d).Ch = blk Then Exit Do
                d = d + 1
            Loop
            
            ' Transition
            If c > 0 Then
                glBegin bmLines
                    glVertex2d lastx, lasty
                    glVertex2d Layout(d).SP.x + xdraw, Layout(d).SP.y + ydraw
                glEnd
            End If
            
            For steps = 0 To Layout(d).DCount - 1
            
                glColor4b 0, 0, 0, 127
                
                st = Layout(d).Drawstep(steps).t
                
                If st = 2 Then
                    ' Line
                    glBegin bmLines
                        dr_point Layout(d).Drawstep(steps).P, 0
                        dr_point Layout(d).Drawstep(steps).P, 2
                    glEnd
                ElseIf st = 3 Then
                    ' Line strip
                    glBegin bmLineStrip
                    With Layout(d).Drawstep(steps)
                        dr_point Layout(d).Drawstep(steps).P, 0
                        For l = 2 To Layout(d).Drawstep(steps).PCount Step 2
                            dr_point Layout(d).Drawstep(steps).P, l
                        Next l
                    End With
                    glEnd
                ElseIf st = 4 Then
                    ' Polygon
                    If datacolor = 0 Then glColor4b 127, 0, 0, 32
                    If datacolor = 1 Then glColor4b 0, 127, 0, 32
                    If datacolor = 2 Then glColor4b 0, 0, 127, 32
                    If datacolor = 3 Then glColor4b 127, 127, 0, 32
                    If datacolor = 4 Then glColor4b 0, 127, 127, 32
                    If datacolor = 5 Then glColor4b 63, 0, 127, 32
                    If datacolor = 6 Then glColor4b 100, 100, 100, 32
                    glBegin bmPolygon
                    With Layout(d).Drawstep(steps)
                        dr_point Layout(d).Drawstep(steps).P, 0
                        For l = 2 To Layout(d).Drawstep(steps).PCount Step 2
                            dr_point Layout(d).Drawstep(steps).P, l
                        Next l
                    End With
                    glEnd
                End If
            Next steps
            
            lastx = xdraw + Layout(d).EP.x
            lasty = ydraw + Layout(d).EP.y
            
            If datastate = 0 Then datastate = 1
            
            lastblk = pblk
    
        Next c
    Next w
    
    DrawGLScene = True                              ' Everything Went OK
End Function

Sub dr_text(txt As String, xdraw As Integer, ydraw As Integer)
    Dim pch As Integer
    Dim sx, sy, ex, ey As Single
    Dim c As Integer
    Dim xofs As Integer
    
    glBlendFunc sfDstColor, dfZero
    glEnable GL_TEXTURE_2D
    glColor3f 1#, 1#, 1#    ' 2 hours were lost here
    glTexEnvf GL_TEXTURE_ENV, GL_TEXTURE_ENV_MODE, GL_REPLACE
    glBindTexture glTexture2D, Texture(0)
    
    xdraw = xdraw - (Len(txt) * 8 / 2) + 8
    
    For c = 0 To Len(txt) - 1
        xofs = (c * 8) + xdraw
        pch = Asc(Mid(txt, c + 1, 1)) - 32
        sx = ((pch Mod 16) / 16)
        sy = 1 - ((pch \ 16) / 8)
        ex = sx + (1 / 16)
        ey = sy - (1 / 8)
        
        glBegin bmQuads
            glTexCoord2f sx, sy
            glVertex2f xofs, ydraw
            glTexCoord2f ex, sy
            glVertex2f 16 + xofs, ydraw
            glTexCoord2f ex, ey
            glVertex2f 16 + xofs, 16 + ydraw
            glTexCoord2f sx, ey
            glVertex2f xofs, 16 + ydraw
        glEnd
    Next c
    
    glDisable GL_TEXTURE_2D
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
End Sub

Function MatchT(ByVal s As String) As Integer
    If s = "SP" Then MatchT = 0      ' Start point
    If s = "EP" Then MatchT = 1     ' End point
    If s = "L" Then MatchT = 2      ' Line
    If s = "LS" Then MatchT = 3     ' Line strip
    If s = "SH" Then MatchT = 4     ' Polygon
End Function
