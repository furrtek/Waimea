VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainFrm 
   Caption         =   "Waimea"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14175
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   945
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
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   4560
      Width           =   13935
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   0
      MousePointer    =   2  'Cross
      Top             =   0
      Width           =   14175
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

Dim JustLoaded As Boolean
Dim Dragging As Boolean
Dim Drag_X As Integer
Dim Drag_Y As Integer
Dim PrevNav_X As Integer
Dim PrevNav_Y As Integer

Private Sub Form_Load()
    On Error GoTo skipdebug
    
    Dim ln As String
    
    XMargin = 64
    Nav_X = 0
    Nav_Y = 0
    Spacing = 1
    LiveRefresh = True
    Loaded = False
    
    FilePath = ""
    SetSaveState False
    
    If Not CreateGLWindow(640, 480, 16) Then End    ' 24 ?
    
    LoadLayout
    LoadFont
    LoadPin
    
    ' FOR DEBUG ONLY !
    Dim dbgload As String
    FilePath = App.Path & "\waveform.txt"
    Open FilePath For Input As #1
        dbgload = ""
        While Not EOF(1)
            Line Input #1, ln
            dbgload = dbgload & ln & vbCrLf
        Wend
        JustLoaded = True
        Text1.Text = dbgload
        DoEvents
        SetSaveState True
    Close #1
    
    Loaded = True
    
    Redraw

    Exit Sub

skipdebug:
    If Err.Number <> 53 Then MsgBox "Error in load.", vbCritical
End Sub

Private Sub Form_Paint()
    Redraw
End Sub


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
        Image1.Width = MainFrm.ScaleWidth
        Image1.Height = MainFrm.ScaleHeight - Text1.Height - 8
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dragging = True
        PrevNav_X = Nav_X
        PrevNav_Y = Nav_Y
        Drag_X = X / 15
        Drag_Y = Y / 15
        Image1.MousePointer = vbSizeAll
    ElseIf Button = 2 Then
        Nav_X = 0
        Nav_Y = 0
        Redraw
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Integer
    Dim PopupFlag As Boolean
    
    X = X / 15
    Y = Y / 15
    
    If Dragging = True Then
        Nav_X = PrevNav_X - (Drag_X - X)
        Nav_Y = PrevNav_Y - (Drag_Y - Y)
        Display
    Else
        PopupFlag = False
        For c = 0 To nPins - 1
            If (X > PinList(c).X - 10 + Nav_X) And _
                (X < PinList(c).X + 10 + Nav_X) And _
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

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
    Image1.MousePointer = vbCrosshair
End Sub

Private Sub menu_about_Click()
    MsgBox "Waimea " & App.Major & "." & App.Minor & vbCrLf & "By furrtek - 2016" & vbCrLf & vbCrLf & "https://github.com/furrtek/Waimea", vbInformation
End Sub

Private Sub menu_export_Click()
    MsgBox "Not implemented yet"
End Sub

Private Sub menu_open_Click()
    On Error GoTo Abort
    
    Dim ln As String
    Dim LoadStr As String
    
    CommonDialog1.DialogTitle = "Open waveform file"
    CommonDialog1.ShowOpen
    
    If Saved = False Then
        If Confirm = False Then Exit Sub
    End If
    
    Open CommonDialog1.FileName For Input As #1
        LoadStr = ""
        While Not EOF(1)
            Line Input #1, ln
            LoadStr = LoadStr & ln & vbCrLf
        Wend
        JustLoaded = True
        Text1.Text = LoadStr
        DoEvents
        FilePath = CommonDialog1.FileName
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

    If KeyCode = vbKeyControl And Keys(vbKeyControl) = False Then   ' Prevents retrig
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
    
    If KeyCode = vbKeyControl Then
        ' Hide all bubbles
        For w = 0 To nPins - 1
            PinList(w).Show = False
        Next w
        Redraw
    ElseIf KeyCode < 112 Or KeyCode > 123 Then
        If LiveRefresh = True Then Redraw
    End If
End Sub

