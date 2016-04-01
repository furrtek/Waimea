Attribute VB_Name = "OGLUtils"
Option Explicit

Public xdraw As Integer
Public ydraw As Integer
Public lastx As Integer
Public lasty As Integer

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
' All Setup For OpenGL Goes Here
    glShadeModel smSmooth               ' Enables Smooth Shading

    glLineWidth 1
    glClearColor 1#, 1#, 1#, 0.5

    glClearDepth 1#                     ' Depth Buffer Setup
    glHint GL_LINE_SMOOTH_HINT, GL_NICEST
    glEnable GL_LINE_SMOOTH
    
    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA


    InitGL = True                       ' Initialization Went OK
End Function


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
    glVertex2d P(ofs) + xdraw, P(ofs + 1) + ydraw
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
    Dim datacolor As Integer
    Dim datastate As Integer
    
    glClear clrColorBufferBit Or clrDepthBufferBit  ' Clear The Screen And The Depth Buffer
    glLoadIdentity                                  ' Reset The View

    glTranslatef 0#, 0#, 0#
    
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
    
    lastx = 64
    lasty = 64
    
    lastblk = "z"   ' Default block
    datastate = -1

    For c = 0 To Len(wavedef) - 1
        pblk = Mid(wavedef, c + 1, 1)
        blk = pblk
        
        If pblk = "." Then pblk = lastblk     ' Repeat
        
        If (Asc(pblk) >= &H30 And Asc(pblk) <= &H35) Or pblk = "=" Then   ' Start data
            If datastate = -1 Then
                If pblk = "=" Then
                    datacolor = 6
                Else
                    datacolor = Asc(pblk) - &H30
                End If
                datastate = 0
            End If
        End If
        
        If datastate = 0 Then blk = "s"
        If datastate = 1 Then blk = "d"
        If datastate > -1 And Mid(wavedef, c + 2, 1) <> "." Then    ' Dangerous (+1 into void ?)
            blk = "e"
            datastate = -1
        End If
        
        ' Look for block def
        d = 0
        Do While 1
            If Layout(d).DCount = 0 Then Exit Do
            If Layout(d).Ch = blk Then Exit Do
            d = d + 1
        Loop
        
        ' Draw
        xdraw = 64 + (c * 15)
        ydraw = 64
        
        ' Transition
        glBegin bmLines
            glVertex2d lastx, lasty
            glVertex2d Layout(d).SP.x + xdraw, Layout(d).SP.y + ydraw
        glEnd
        
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

    DrawGLScene = True                              ' Everything Went OK
End Function

Function MatchT(ByVal s As String) As Integer
    If s = "SP" Then MatchT = 0      ' Start point
    If s = "EP" Then MatchT = 1     ' End point
    If s = "L" Then MatchT = 2      ' Line
    If s = "LS" Then MatchT = 3     ' Line strip
    If s = "SH" Then MatchT = 4     ' Polygon
End Function
