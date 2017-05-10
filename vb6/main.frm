VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainFrm 
   Caption         =   "Waimea"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13200
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text files|*.txt"
   End
   Begin VB.TextBox EditBox 
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
      Height          =   3405
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3000
      Width           =   13095
   End
   Begin VB.PictureBox Vis 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   13200
      Begin VB.Timer RefreshTmr 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   120
      End
   End
   Begin VB.Menu menu_sheet 
      Caption         =   "Sheet"
      Begin VB.Menu mwnu_new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menu_sep3 
         Caption         =   "-"
      End
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
         Enabled         =   0   'False
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

' Order of stuff, from back to front:
' Ticks, rulers (panDL)
' Group horizontal stripes (fixed)
' Waves, pins, bubbles... (panDL)
' Names (fixed)

Dim Dragging As Boolean
Dim LastFile As String      ' Last saved filename
Dim Drag_X As Integer
Dim Drag_Y As Integer
Dim PrevNav_X As Integer
Dim PrevNav_Y As Integer
Dim ResizeMode As Boolean   ' Edit box resize mode
Dim Resizing As Boolean

Dim RefreshCountdown As Integer

Private Sub EditBox_KeyPress(KeyAscii As Integer)
    SetSaveState False
End Sub

Private Sub Form_Activate()
    Dim w As Integer
    
    If Loaded = True Then Exit Sub
    
    ' Startup initializations
    Set FSO = New FileSystemObject
    
    XMargin = 100
    YMargin = 20
    
    FilePath = ""
    SetSaveState False
    
    If Not CreateGLWindow(640, 480, 16) Then
        MsgBox "Could not create OpenGL view.", vbCritical
        End
    End If
    
    ' Ask for some displaylists
    TicksDL = glGenLists(1)
    For w = 0 To 255
        Waves(w).DL = glGenLists(1)
    Next w
    
    LoadSettings
    UIResize
    LoadLayout
    LoadFont
    LoadPin
    
    ' Try opening last opened file if needed
    If OpenLast = True Then
        If LoadWaveDef(LastOpened) = False Then
            LoadWaveDef App.Path & "\demo.txt"
        End If
    Else
        LoadWaveDef App.Path & "\demo.txt"
    End If
    
    ' Init done
    Loaded = True
    
    Redraw
End Sub

Function LoadWaveDef(File As String) As Boolean
    Dim FileLine As String
    Dim LoadStr As String

    If FSO.FileExists(File) = False Then
        LoadWaveDef = False
        Exit Function
    End If

    Open File For Input As #1
        LoadStr = ""
        While Not EOF(1)
            Line Input #1, FileLine
            LoadStr = LoadStr & FileLine & vbCrLf
        Wend
        EditBox.Text = LoadStr
        DoEvents
        FilePath = File
        SetSaveState True
        LastFile = File
    Close #1
    
    ' Reset pan
    Nav_X = 0
    Nav_Y = 0
    
    ReSizeGLScene
    Redraw
    
    LoadWaveDef = True
End Function

Private Sub Form_Load()
    ' Init must be done
    Loaded = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Click & drag in the right spot means resizing
    If ResizeMode = True Then Resizing = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim VisHeight As Integer
    
    If Resizing = False Then
        VisHeight = Vis.Height
        
        ' Change cursor to North-South resize arrows when in between views
        If Y > VisHeight And Y < VisHeight + 4 Then
            If ResizeMode = False Then
                MainFrm.MousePointer = vbSizeNS
                ResizeMode = True
            End If
        Else
            If ResizeMode = True Then
                MainFrm.MousePointer = vbDefault
                ResizeMode = False
                Resizing = False
            End If
        End If
    Else
        ' Resize limits
        If Y > 16 And Y < MainFrm.ScaleHeight - 48 Then UISplitY = Y
        UIResize
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Resizing = False
End Sub

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
    UIResize
End Sub

Private Sub UIResize()
    ' Resize controls with limits
    If (MainFrm.ScaleWidth > 256) And (MainFrm.ScaleHeight > 128) Then
        If MainFrm.ScaleHeight < UISplitY + 8 Then
            UISplitY = MainFrm.ScaleHeight - 8
        End If
        Vis.Width = MainFrm.ScaleWidth
        Vis.Height = UISplitY
        EditBox.Top = UISplitY + 4
        EditBox.Width = MainFrm.ScaleWidth - 8
        EditBox.Height = MainFrm.ScaleHeight - UISplitY - 8
        ReSizeGLScene
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Private Sub menu_about_Click()
    MsgBox "Waimea " & App.Major & "." & App.Minor & App.Revision & vbCrLf & "By furrtek - 2017" & vbCrLf & vbCrLf & "https://github.com/furrtek/Waimea", vbInformation
End Sub

Private Sub menu_export_Click()
    MsgBox "Not implemented yet", vbInformation
End Sub

Private Sub menu_extend_Click()
    If nWaves > 0 Then
        ExtendFrm.Show 1
    Else
        MsgBox "No waves to extend.", vbExclamation
    End If
End Sub

Private Sub menu_open_Click()
    Dim File As String
    
    CommonDialog1.DialogTitle = "Open waveform sheet"
    CommonDialog1.ShowOpen
    
    ' Ask to save current file before opening new one if needed
    If Saved = False Then
        If Confirm = False Then Exit Sub
    End If
    
    File = CommonDialog1.FileName
    
    LoadWaveDef File
End Sub

Private Sub menu_save_Click()
    SaveFile False
End Sub

Sub SaveFile(Force As Boolean)
    On Error GoTo Abort
    
    Dim File As String
    Dim FileLine As String
    
    If FilePath = "" Or Force = True Then
        ' "Save as"
        CommonDialog1.DialogTitle = "Save waveform sheet"
        CommonDialog1.ShowSave
        FilePath = CommonDialog1.FileName
    Else
        ' "Save"
        CommonDialog1.FileName = FilePath
    End If
    
    File = CommonDialog1.FileName
    
    If FSO.FileExists(File) = True And File <> LastFile Then
        If MsgBox("Overwrite existing file ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    FileLine = EditBox.Text
    Open CommonDialog1.FileName For Output As #1
        ' Remove trailing CRLF
        If Len(FileLine) >= 2 Then
            If Right(FileLine, 2) = vbCrLf Then FileLine = Left(FileLine, Len(FileLine) - 2)
        End If
        Print #1, EditBox.Text;
        SetSaveState True
        LastFile = File
    Close #1
    
Abort:
End Sub

Private Sub menu_saveas_Click()
    SaveFile True
End Sub

Private Sub menu_settings_Click()
    SettingsFrm.Show 1
End Sub

Private Sub EditBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ResizeMode = True Then
        MainFrm.MousePointer = vbDefault
        ResizeMode = False
        Resizing = False
    End If
End Sub

Private Sub mwnu_new_Click()
    Dim File As String
    
    ' Ask to save current file before clearing if needed
    If Saved = False Then
        If Confirm = False Then Exit Sub
    End If
    
    EditBox.Text = ""
    FilePath = "new_sheet.txt"
    Saved = False
    SetFormTitle
    Redraw
End Sub

Private Sub vis_DblClick()
    ' Reset pan
    Nav_X = 0
    Nav_Y = 0
    Display
End Sub

Private Sub vis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SnapC As Single
    Dim Str_Start As Integer
    Dim Line_Count As Integer
    
    If Button = 1 Then
        ' Left mouse button
        If X > XMargin Then
            ' Drag-pan
            Dragging = True
            PrevNav_X = Nav_X
            PrevNav_Y = Nav_Y
            Drag_X = X
            Drag_Y = Y
            Vis.MousePointer = vbSizeAll
        Else
            ' Select line in editbox
            ' Find where each line starts
            Str_Start = 0
            Line_Count = 0
            Do
                If Line_Count = (Y - Nav_Y) \ 20 Then Exit Do
                Str_Start = InStr(Str_Start + 1, LCase(EditBox.Text), "name:")
                Line_Count = Line_Count + 1
            Loop While (Str_Start <> 0)
            If Str_Start > 1 Then
                EditBox.SelStart = Str_Start - 1
                EditBox.SetFocus
            End If
        End If
    ElseIf Button = 2 And X > (XMargin * Spacing) Then
        ' Right mouse button: measure time
        SnapC = Spacing * 15
        Meas_X = (((X - Nav_X - (XMargin * Spacing) + 8) \ SnapC) * SnapC) + (XMargin * Spacing)
        Meas_Y = (((Y - Nav_Y - (YMargin * Spacing) + 2) \ 20) * 20) + 8
        Measuring = True
    End If
End Sub

Private Sub vis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SnapC As Single
    Dim c As Integer
    Dim PopupFlag As Boolean
    Dim PLX, PLY As Single
    Dim New_X, New_Y As Integer
    
    If ResizeMode = True Then
        MainFrm.MousePointer = vbDefault
        ResizeMode = False
        Resizing = False
    End If
    
    If Dragging = True Then
        New_X = PrevNav_X - (Drag_X - X)
        If New_X <= 0 Then
            Nav_X = New_X
        Else
            Nav_X = 0
        End If
        New_Y = PrevNav_Y - (Drag_Y - Y)
        If New_Y <= 0 Then
            Nav_Y = New_Y
        Else
            Nav_Y = 0
        End If

        Display
    ElseIf Measuring = True Then
        SnapC = Spacing * 15
        New_X = (((X - Nav_X - (XMargin * Spacing) + 8) \ SnapC) * SnapC) + (XMargin * Spacing)
        New_Y = (((Y - Nav_Y - (YMargin * Spacing) + 2) \ 20) * 20) + 8
        If (New_X <> Cur_X) Or (New_Y <> Cur_Y) Then
            Cur_X = New_X
            Cur_Y = New_Y
            Display
        End If
    End If
    
    If X < XMargin Then
        Vis.MousePointer = vbArrowQuestion
    Else
        Vis.MousePointer = vbDefault
        
        If Keys(18) = False Then
            ' Pin bubble popup if cursor is hovering over
            PopupFlag = False
            For c = 0 To nPins - 1
                PLX = PinList(c).X * Spacing
                PLY = PinList(c).Y
                If (X > PLX - 10 + XMargin + Nav_X) And _
                    (X < PLX + 10 + XMargin + Nav_X) And _
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
    End If
End Sub

Private Sub vis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ' Left mouse button
        Dragging = False
    ElseIf Button = 2 Then
        ' Right mouse button
        Measuring = False
    End If
    Vis.MousePointer = vbCrosshair
    Display
End Sub

Private Sub vis_Paint()
    If Loaded = True Then Redraw
End Sub

Private Sub EditBox_Change()
    If ResizeMode = True Then
        MainFrm.MousePointer = vbDefault
        ResizeMode = False
        Resizing = False
    End If
End Sub

Sub SetSaveState(v As Boolean)
    Saved = v
    SetFormTitle
End Sub

Sub SetFormTitle()
    Dim title As String
    
    title = GetFileName(FilePath)
    
    If Saved = False Then title = title & "*"
    
    title = title & " - Waimea " & App.Major & "." & App.Minor & App.Revision
    
    MainFrm.Caption = title
End Sub

Private Sub EditBox_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub EditBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim w As Integer
    
    Keys(KeyCode) = False
    
    If KeyCode = 18 Then
        ' Hide all bubbles
        For w = 0 To nPins - 1
            PinList(w).Show = False
        Next w
        Redraw
    ElseIf KeyCode < 112 Or KeyCode > 123 Then
        ' Retrig refresh timeout
        If LiveRefresh = True Then ResetRT
    End If
End Sub

Sub ResetRT()
    RefreshCountdown = 4    ' 400ms
    RefreshTmr.Enabled = True
End Sub

Private Sub RefreshTmr_Timer()
    If RefreshCountdown = 0 Then
        RefreshTmr.Enabled = False
        Redraw
    Else
        RefreshCountdown = RefreshCountdown - 1
    End If
End Sub
