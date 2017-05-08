Attribute VB_Name = "LoadMd"

Public Sub LoadFont()
    Dim fn As String
    Dim FontData() As GLbyte
    Dim bd As Byte
    Dim h, w As Integer
    
    fn = App.Path & "\font.tga"
    
    If FSO.FileExists(fn) = False Then ErrorBox "The file font.tga is missing.", True
    
    ReDim FontData(3, 511, 255)
    
    Open fn For Binary As #1
        Seek #1, 19
        Get #1, , FontData
    Close #1

    glGenTextures 1, FontTex
    glBindTexture glTexture2D, FontTex
    glTexImage2D glTexture2D, 0, 4, 512, 256, 0, tiRGBA, GL_UNSIGNED_BYTE, FontData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR     ' Linear filtering
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
    
    Erase FontData
End Sub

Public Sub LoadPin()
    Dim fn As String
    Dim PinData() As GLbyte
    Dim bd As Byte
    Dim h, w As Integer
    
    fn = App.Path & "\pin.tga"
    
    If FSO.FileExists(fn) = False Then ErrorBox "The file pin.tga is missing.", True
    
    ReDim PinData(3, 63, 63)
    
    Open fn For Binary As #1
        Seek #1, 19
        Get #1, , PinData
    Close #1

    glGenTextures 1, PinTex
    glBindTexture glTexture2D, PinTex
    glTexImage2D glTexture2D, 0, 4, 64, 64, 0, tiRGBA, GL_UNSIGNED_BYTE, PinData(0, 0, 0)
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
    
    Erase PinData
End Sub

Sub LoadLayout()
    Dim fn As String
    Dim lidx As Integer
    Dim didx As Integer
    Dim lline As String
    Dim a() As String
    Dim b() As String
    Dim pidx As Integer
    Dim t As String
    Dim DataColor As Integer
    Dim c As Integer
    Dim d As Integer
    
    Dim sx, sy, ex, ey As Single
    
    fn = App.Path & "\layout.txt"
    
    If FSO.FileExists(fn) = False Then ErrorBox "The file layout.txt is missing.", True
    
    ' Pin display list
    PinDL = glGenLists(1)
    glNewList PinDL, GL_COMPILE
        glBegin bmQuads
            glTexCoord2f 0, 1
            glVertex2f 0, 0
            glTexCoord2f 1, 1
            glVertex2f 20, 0
            glTexCoord2f 1, 0
            glVertex2f 20, 20
            glTexCoord2f 0, 0
            glVertex2f 0, 20
        glEnd
    glEndList
    
    ' Generate font (characters) display lists
    For c = 0 To 128 - 1
        sx = ((c Mod 16) / 16)
        sy = 1 - ((c \ 16) / 8)
        ex = sx + (1 / 16)
        ey = sy - (1 / 8)
    
        CharDL(c) = glGenLists(1)
        glNewList CharDL(c), GL_COMPILE
            glBegin bmQuads
                glTexCoord2f sx, sy
                glVertex2f 0, 0
                glTexCoord2f ex, sy
                glVertex2f 16, 0
                glTexCoord2f ex, ey
                glVertex2f 16, 16
                glTexCoord2f sx, ey
                glVertex2f 0, 16
            glEnd
        glEndList
    Next c
    
    lidx = -1
    Open App.Path & "\layout.txt" For Input As #1
        Do
            Line Input #1, lline
            If lline <> "" Then
                If InStr(1, UCase(lline), "DEF") Then
                    lidx = lidx + 1
                    If lidx > 0 Then glEndList
                    DispLists(lidx).DL = glGenLists(1)
                    glNewList DispLists(lidx).DL, GL_COMPILE
                    
                    a = Split(lline, " ")
                    DispLists(lidx).Char = Left(a(1), 1)
                Else
                    a = Split(UCase(lline), " ")
                    t = a(0)
                    
                    If t = "SP" Or t = "EP" Then
                        b = Split(a(1), ",")
                        If UBound(b) > 0 Then
                            If t = "SP" Then
                                DispLists(lidx).SP.X = b(0)
                                DispLists(lidx).SP.Y = b(1)
                            ElseIf t = "EP" Then
                                DispLists(lidx).EP.X = b(0)
                                DispLists(lidx).EP.Y = b(1)
                            End If
                        End If
                    Else
                        If UBound(a) > 0 Then
                            lline = a(1)
                            a = Split(lline, ":")
                        End If
                        
                        If t = "L" Then
                            ' Line
                            glBegin bmLines
                                b = Split(a(0), ",")
                                glVertex2f b(0), b(1)
                                b = Split(a(1), ",")
                                glVertex2f b(0), b(1)
                            glEnd
                        ElseIf t = "LS" Then
                            ' Line strip
                            glBegin bmLineStrip
                            For c = 0 To UBound(a)
                                b = Split(a(c), ",")
                                glVertex2f b(0), b(1)
                            Next c
                            glEnd
                        ElseIf t = "SH" Then
                            ' Polygon
                            glBegin bmPolygon
                            For c = 0 To UBound(a)
                                b = Split(a(c), ",")
                                glVertex2f b(0), b(1)
                            Next c
                            glEnd
                        ElseIf t = "LC" Then
                            ' Line color
                            SetGLColor Color_Waves
                        End If
                    End If
                End If
            End If
        Loop While Not EOF(1)
        glEndList
    Close #1
    
    DispLists(lidx + 1).Char = " "
End Sub

