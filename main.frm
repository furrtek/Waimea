VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainFrm 
   Caption         =   "Waimea"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14235
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text files|*.txt"
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans Mono"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   4560
      Width           =   13935
   End
   Begin VB.Menu menu_sheet 
      Caption         =   "Sheet"
      Begin VB.Menu menu_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu menu_save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_saveas 
         Caption         =   "Save as"
      End
      Begin VB.Menu menu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu menu_export 
         Caption         =   "Export"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu menu_tools 
      Caption         =   "Tools"
      Begin VB.Menu menu_settings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "?"
      Begin VB.Menu menu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dragging As Boolean
Dim Drag_X As Integer
Dim Drag_Y As Integer
Dim PrevNav_X As Integer
Dim PrevNav_Y As Integer

Sub LoadLayout()
    Dim lidx As Integer
    Dim didx As Integer
    Dim lline As String
    Dim a() As String
    Dim b() As String
    Dim pidx As Integer
    Dim t As Integer
    Dim DataColor As Integer
    
    Dim c As Integer
    Dim d As Integer
    
    Dim sx, sy, ex, ey As Single
    
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
    Open "layout.txt" For Input As #1
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
                    a = Split(lline, " ")
                    t = MatchT(a(0))
                    
                    If t < 2 Then
                        b = Split(a(1), ",")
                        If t = 0 Then
                            DispLists(lidx).SP.X = b(0)
                            DispLists(lidx).SP.Y = b(1)
                        ElseIf t = 1 Then
                            DispLists(lidx).EP.X = b(0)
                            DispLists(lidx).EP.Y = b(1)
                        End If
                    Else
                        lline = a(1)
                        a = Split(lline, ":")

                        If t = 2 Then
                            ' Line
                            glBegin bmLines
                                b = Split(a(0), ",")
                                glVertex2f b(0), b(1)
                                b = Split(a(1), ",")
                                glVertex2f b(0), b(1)
                            glEnd
                        ElseIf t = 3 Then
                            ' Line strip
                            glBegin bmLineStrip
                            For c = 0 To UBound(a)
                                b = Split(a(c), ",")
                                glVertex2f b(0), b(1)
                            Next c
                            glEnd
                        ElseIf t = 4 Then
                            ' Polygon
                            glBegin bmPolygon
                            For c = 0 To UBound(a)
                                b = Split(a(c), ",")
                                glVertex2f b(0), b(1)
                            Next c
                            glEnd
                        End If
                    End If
                End If
            End If
        Loop While Not EOF(1)
        glEndList
    Close #1
    
    DispLists(lidx + 1).Char = " "
End Sub

Private Sub Form_Load()
    On Error GoTo skipdebug
    
    Dim ln As String
    
    XMargin = 64
    Nav_X = 0
    Nav_Y = 16
    Spacing = 1
    LiveRefresh = True
    
    FilePath = ""
    SetSaveState False
    
    If Not CreateGLWindow(Me, 640, 480, 16) Then End    ' 24 ?

    LoadLayout
    LoadFont
    LoadPin
    
    ' FOR DEBUG ONLY !
    Dim dbgload As String
    FilePath = App.Path & "\waveform.txt"
    Open "waveform.txt" For Input As #1
        dbgload = ""
        While Not EOF(1)
            Line Input #1, ln
            dbgload = dbgload & ln & vbCrLf
        Wend
        Text1.Text = dbgload
        DoEvents
        SetSaveState True
    Close #1

    Exit Sub

skipdebug:
    If Err.Number <> 53 Then MsgBox "Error in load.", vbCritical
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    PrevNav_X = Nav_X
    PrevNav_Y = Nav_Y
    Drag_X = X
    Drag_Y = Y
    MainFrm.MousePointer = vbSizeAll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Integer
    Dim PopupFlag As Boolean
    
    If Dragging = True Then
        Nav_X = PrevNav_X - (Drag_X - X)
        Nav_Y = PrevNav_Y - (Drag_Y - Y)
        Display
    Else
        PopupFlag = False
        For c = 0 To nPins - 1
            If (X > PinList(c).X + Nav_X) And _
                (X < PinList(c).X + 20 + Nav_X) And _
                (Y > PinList(c).Y + Nav_Y) And _
                (Y < PinList(c).Y + 20 + Nav_Y) Then
                If PinList(c).Show = False Then
                    PopupFlag = True
                    PinList(c).Show = True
                End If
            Else
                If PinList(c).Show = True Then
                    PopupFlag = True
                    PinList(c).Show = False
                End If
            End If
        Next c
        If PopupFlag = True Then Redraw
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
    MainFrm.MousePointer = vbCrosshair
End Sub

Private Sub Form_Paint()
    Redraw
End Sub

Function GetFileName(fn As String)
    Dim pos As Integer
    
    pos = 1
    While InStr(pos, fn, "\") > 0
        pos = InStr(pos, fn, "\") + 1
    Wend
    
    GetFileName = Right(fn, Len(fn) - pos + 1)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Done As Boolean
    
    If Saved = False Then
        Done = Confirm
    Else
        Done = True
    End If
    
    If Done = True Then
        SettingsFrm.Hide
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    ReSizeGLScene ScaleWidth, ScaleHeight
    If (MainFrm.ScaleWidth > 16) And (MainFrm.ScaleHeight > 32) Then
        Text1.Top = MainFrm.ScaleHeight - Text1.Height - 8
        Text1.Width = MainFrm.ScaleWidth - 16
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Sub MoveView()
    glMatrixMode mmModelView
    glTranslatef 2, 1, 1
End Sub

Private Sub menu_about_Click()
    MsgBox "Waimea " & App.Major & "." & App.Minor & vbCrLf & "By furrtek - 2016" & vbCrLf & vbCrLf & "https://github.com/furrtek/Waimea", vbInformation
End Sub

Private Sub menu_export_Click()
    MsgBox "Not implemented yet"
End Sub

Function Confirm() As Boolean
    Confirm = False
    If MsgBox("Discard unsaved data ?", vbYesNo + vbQuestion) = vbYes Then Confirm = True
End Function

Private Sub menu_open_Click()
    On Error GoTo Abort
    
    Dim ln As String
    
    CommonDialog1.DialogTitle = "Open waveform file"
    CommonDialog1.ShowOpen
    
    If Saved = False Then
        If Confirm = False Then Exit Sub
    End If
    
    Open CommonDialog1.FileName For Input As #1
        Text1.Text = ""
        While Not EOF(1)
            Line Input #1, ln
            Text1.Text = Text1.Text & ln & vbCrLf
        Wend
        SetSaveState True
    Close #1
    
Abort:
End Sub

Private Sub menu_save_Click()
    SaveFile False
End Sub

Sub SaveFile(Force As Boolean)
    On Error GoTo Abort
    
    Dim ln As String
    
    If FilePath = "" Or Force = True Then
        CommonDialog1.DialogTitle = "Save waveform file"
        CommonDialog1.ShowSave
    Else
        CommonDialog1.FileName = FilePath
    End If
    
    ln = Text1.Text
    
    Open CommonDialog1.FileName For Output As #1
        If Len(ln) >= 2 Then
            If Right(ln, 2) = vbCrLf Then ln = Left(ln, Len(ln) - 2)
        End If
        Print #1, Text1.Text;
        SetSaveState True
    Close #1
Abort:
End Sub

Private Sub menu_saveas_Click()
    SaveFile True
End Sub

Private Sub menu_settings_Click()
    SettingsFrm.Show
End Sub

Private Sub Text1_Change()
    SetSaveState False
End Sub

Sub Redraw()
    Render Me
    Display
End Sub

Sub SetSaveState(v As Boolean)
    If Saved <> v Then
        Saved = v
        SetFormTitle
    End If
End Sub

Sub SetFormTitle()
    Dim title As String
    
    title = GetFileName(FilePath)
    
    If Saved = False Then title = title & "*"
    
    title = title & " - Waimea " & App.Major & "." & App.Minor
    
    MainFrm.Caption = title
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = False
    Redraw
End Sub
