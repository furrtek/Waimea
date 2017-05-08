Attribute VB_Name = "GLMd"

Public Sub ReSizeGLScene()
    Dim Height, Width As Integer
    
    Width = MainFrm.Vis.ScaleWidth
    Height = MainFrm.Vis.ScaleHeight
    
    If Height = 0 Then Height = 1
    If Width = 0 Then Width = 1
    
    glViewport 0, 0, Width, Height
    glMatrixMode mmProjection
    glLoadIdentity

    glOrtho 0#, Width, Height, 0#, -1, 1

    glMatrixMode mmModelView
    glLoadIdentity
End Sub

Public Sub InitGL()
    glLineWidth 1
    glShadeModel smSmooth
    glClearDepth 1#
    glHint htLineSmoothHint, hmNicest
    glEnable glcBlend
End Sub

Public Sub KillGLWindow()
    If hrc Then
        If wglMakeCurrent(0, 0) = 0 Then
            ErrorBox "Release of DC and RC failed.", False
        End If

        If wglDeleteContext(hrc) = 0 Then
            ErrorBox "Release of the rendering context failed.", False
        End If
        hrc = 0
    End If
End Sub

Public Function CreateGLWindow(Width As Integer, Height As Integer, Bits As Integer) As Boolean
    Dim PixelFormat As GLuint
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim CanvasDc As Long
    
    CanvasDc = MainFrm.Vis.hDC

    pfd.cColorBits = Bits
    pfd.cDepthBits = 16
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    
    PixelFormat = ChoosePixelFormat(CanvasDc, pfd)
    If PixelFormat = 0 Then
        KillGLWindow
        ErrorBox "Can't find a suitable PixelFormat.", True
        CreateGLWindow = False
    End If

    If SetPixelFormat(CanvasDc, PixelFormat, pfd) = 0 Then
        KillGLWindow
        ErrorBox "Can't set the PixelFormat.", True
        CreateGLWindow = False
    End If
    
    hrc = wglCreateContext(CanvasDc)
    If (hrc = 0) Then
        KillGLWindow
        ErrorBox "Can't create a GL rendering context.", True
        CreateGLWindow = False
    End If

    If wglMakeCurrent(CanvasDc, hrc) = 0 Then
        KillGLWindow
        ErrorBox "Can't activate the GL rendering context.", True
        CreateGLWindow = False
    End If
    
    InitGL
    
    MainFrm.Show
    SetForegroundWindow MainFrm.hWnd
    MainFrm.SetFocus
    
    CreateGLWindow = True
End Function
