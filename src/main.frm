VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainFrm 
   Caption         =   "Waimea"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11895
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
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
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans Mono"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3000
      Width           =   11775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   0
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   2
      Top             =   0
      Width           =   11880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Measure:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   5220
      Width           =   1455
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
      Begin VB.Menu menu_extend 
         Caption         =   "Extend waves (beta)"
      End
      Begin VB.Menu menu_sep2 
         Caption         =   "-"
      End
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

Dim JustLoaded As Boolean   ' Not used anymore ?
Dim Dragging As Boolean
Dim Measuring As Boolean
Dim Drag_X As Integer
Dim Drag_Y As Integer
Dim PrevNav_X As Integer
Dim PrevNav_Y As Integer

Private Sub Form_Activate()
    Dim w As Integer
    
    Set FSO = New FileSystemObject
    
    XMargin = 64
    YMargin = 20
    Nav_X = 0
    Nav_Y = 0
    
    FilePath = ""
    SetSaveState False
    
    If Not CreateGLWindow(640, 480, 16) Then End    ' 24 ?
    
    TicksDL = glGenLists(1)
    For w = 0 To 255
        Waves(w).DL = glGenLists(1)
    Next w
    
    LoadSettings
    LoadLayout
    LoadFont
    LoadPin
    
    If OpenLast = True Then
        If LoadWaveDef(LastOpened) = False Then
            LoadWaveDef App.Path & "\demo.txt"
        End If
    Else
        LoadWaveDef App.Path & "\demo.txt"
    End If
    
    Loaded = True
    
    Redraw
End Sub

Function LoadWaveDef(fn As String)
    Dim ln As String
    Dim LoadStr As String

    If FSO.FileExists(fn) = False Then
        LoadWaveDef = False
        Exit Function
    End If

    Open fn For Input As #1
        LoadStr = ""
        While Not EOF(1)
            Line Input #1, ln
            LoadStr = LoadStr & ln & vbCrLf
        Wend
        JustLoaded = True       ' Not used anymore ?
        Text1.Text = LoadStr
        DoEvents
        FilePath = fn
        SetSaveState True
    Close #1
    
    ReSizeGLScene
    Redraw
    
    LoadWaveDef = True
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Done As Boolean
    
    If Saved = False Then
        Done = Confirm
    Else
        Done = True
    End If
    
    If Done = True Then
        SaveSettings
        SettingsFrm.Hide
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    If (MainFrm.ScaleWidth > 16) And (MainFrm.ScaleHeight > 32) Then
        Text1.Top = MainFrm.ScaleHeight - Text1.Height - 4
        Text1.Width = MainFrm.ScaleWidth - 8
        Picture1.Width = MainFrm.ScaleWidth
        Picture1.Height = MainFrm.ScaleHeight - Text1.Height - 8
        ReSizeGLScene
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Private Sub Image1_DblClick()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub menu_about_Click()
    MsgBox "Waimea " & App.Major & "." & App.Minor & vbCrLf & "By furrtek - 2016" & vbCrLf & vbCrLf & "https://github.com/furrtek/Waimea", vbInformation
End Sub

Private Sub menu_export_Click()
    MsgBox "Not implemented yet", vbInformation
End Sub

Private Sub menu_extend_Click()
    If nWaves > 0 Then ExtendFrm.Show 1
End Sub

Private Sub menu_open_Click()
    Dim fn As String
    
    CommonDialog1.DialogTitle = "Open waveform file"
    CommonDialog1.ShowOpen
    
    If Saved = False Then
        If Confirm = False Then Exit Sub
    End If
    
    fn = CommonDialog1.FileName
    
    If FSO.FileExists(fn) = True Then LoadWaveDef fn
End Sub

Private Sub menu_save_Click()
    SaveFile False
End Sub

Sub SaveFile(Force As Boolean)
    On Error GoTo Abort
    
    Dim fn As String
    Dim ln As String
    
    If FilePath = "" Or Force = True Then
        CommonDialog1.DialogTitle = "Save waveform file"
        CommonDialog1.ShowSave
    Else
        CommonDialog1.FileName = FilePath
    End If
    
    fn = CommonDialog1.FileName
    
    If FSO.FileExists(fn) = True Then
        ln = Text1.Text
        Open CommonDialog1.FileName For Output As #1
            If Len(ln) >= 2 Then
                If Right(ln, 2) = vbCrLf Then ln = Left(ln, Len(ln) - 2)
            End If
            Print #1, Text1.Text;
            SetSaveState True
        Close #1
    End If
    
Abort:
End Sub

Private Sub menu_saveas_Click()
    SaveFile True
End Sub

Private Sub menu_settings_Click()
    SettingsFrm.Show 1
End Sub


Private Sub Picture1_DblClick()
    Nav_X = 0
    Nav_Y = 0
    Display
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dragging = True
        PrevNav_X = Nav_X
        PrevNav_Y = Nav_Y
        Drag_X = X
        Drag_Y = Y
        Picture1.MousePointer = vbSizeAll
    ElseIf Button = 2 Then
        Measuring = True
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Integer
    Dim PopupFlag As Boolean
    Dim PLX, PLY As Single
    
    If Dragging = True Then
        Nav_X = PrevNav_X - (Drag_X - X)
        Nav_Y = PrevNav_Y - (Drag_Y - Y)
        Display
    ElseIf Keys(18) = False Then
        PopupFlag = False
        For c = 0 To nPins - 1
            PLX = PinList(c).X * Spacing
            PLY = PinList(c).Y * Spacing
            If (X > PLX - 10 + Nav_X) And _
                (X < PLX + 10 + Nav_X) And _
                (Y > PLY + YMargin + Nav_Y) And _
                (Y < PLY + YMargin + 20 + Nav_Y) Then
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

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
    If Button = 2 Then Measuring = False
    Picture1.MousePointer = vbCrosshair
    Display
End Sub

Private Sub Picture1_Paint()
    If Loaded = True Then Redraw
End Sub

Private Sub Text1_Change()
    If JustLoaded = False Then SetSaveState False
    JustLoaded = False
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
    Dim w As Integer
    
    If KeyCode = 116 Then Redraw      ' F5 key

    If KeyCode = 18 And Keys(18) = False And AltBubbles = True Then   ' Prevents Alt retrig
        ' Show all bubbles
        For w = 0 To nPins - 1
            PinList(w).Show = True
        Next w
        Redraw
    End If
    
    Keys(KeyCode) = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim w As Integer
    
    Keys(KeyCode) = False
    
    If KeyCode = 18 Then
        ' Hide all bubbles
        For w = 0 To nPins - 1
            PinList(w).Show = False
        Next w
        Redraw
    ElseIf KeyCode < 112 Or KeyCode > 123 Then
        If LiveRefresh = True Then Redraw
    End If
End Sub

