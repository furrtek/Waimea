VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Waimea"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
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
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   32
      Left            =   120
      Max             =   1024
      SmallChange     =   8
      TabIndex        =   1
      Top             =   4200
      Width           =   13935
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
      Text            =   "Form1.frx":0E42
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
   Begin VB.Menu menu_help 
      Caption         =   "?"
      Begin VB.Menu menu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Done As Boolean

Private Sub Form_Load()
    Dim frm As Form
    Done = False
    
    Dim lidx As Integer
    Dim didx As Integer
    Dim lline As String
    Dim a() As String
    Dim b() As String
    Dim pidx As Integer
    Dim t As Integer
    Dim ln As String
    
    Dim c As Integer
    Dim d As Integer
    
    xmargin = 64
    
    FilePath = ""
    SetSaveState False
    
    ' Parse layout file
    lidx = -1
    Open "layout.txt" For Input As #1
        Do
            Line Input #1, lline
            If lline <> "" Then
                If InStr(1, UCase(lline), "DEF") Then
                    If (lidx > -1) Then
                        Layout(lidx).DCount = didx
                    End If
                    a = Split(lline, " ")
                    lidx = lidx + 1
                    didx = 0
                    Layout(lidx).Ch = a(1)
                Else
                    a = Split(lline, " ")
                    t = MatchT(a(0))
                    Layout(lidx).Drawstep(didx).t = t    ' Type
                    If t < 2 Then
                        b = Split(a(1), ",")
                        If t = 0 Then
                            Layout(lidx).SP.x = b(0)
                            Layout(lidx).SP.y = b(1)
                        ElseIf t = 1 Then
                            Layout(lidx).EP.x = b(0)
                            Layout(lidx).EP.y = b(1)
                        End If
                    Else
                        lline = a(1)
                        pidx = 0
                        a = Split(lline, ":")
                        For c = 0 To UBound(a)
                            b = Split(a(c), ",")
                            For d = 0 To UBound(b)
                                Layout(lidx).Drawstep(didx).P(pidx) = b(d)
                                pidx = pidx + 1
                            Next d
                        Next c
                        Layout(lidx).Drawstep(didx).PCount = pidx - 1
                        didx = didx + 1
                    End If
                End If
            End If
        Loop While Not EOF(1)
        Layout(lidx).DCount = didx
    Close #1
    
    Layout(lidx + 1).DCount = 0
    
    If Not CreateGLWindow(Me, 640, 480, 16, False) Then Done = True
    
    LoadFont
    
    ' DEBUG ONLY !!!
    FilePath = App.Path & "\waveform.txt"
    Open "waveform.txt" For Input As #1
        Text1.Text = ""
        While Not EOF(1)
            Line Input #1, ln
            Text1.Text = Text1.Text & ln & vbCrLf
        Wend
        DoEvents
        SetSaveState True
    Close #1

    Do While Not Done
        If (Keys(vbKeyEscape)) Then
            Unload Me
            Done = True
        Else
            DoEvents
        End If
    Loop

    End
End Sub

Private Sub Form_Paint()
    Text1_Change
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
    If Saved = False Then
        Done = Confirm
    Else
        Done = True
    End If
    
    Cancel = 1
End Sub

Private Sub Form_Resize()
    ReSizeGLScene ScaleWidth, ScaleHeight
    Text1.Top = Form1.ScaleHeight - Text1.Height - 8
    Text1.Width = Form1.ScaleWidth - 16
    HScroll1.Top = Text1.Top - 24
    HScroll1.Width = Form1.ScaleWidth - 18
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Private Sub HScroll1_Change()
    Text1_Change
End Sub

Private Sub HScroll1_Scroll()
    Text1_Change
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

Private Sub Text1_Change()
    SetSaveState False
    DrawGLScene Form1
    SwapBuffers Form1.hDC
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
    
    Form1.Caption = title
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = False
End Sub

Private Sub Text2_Change()
    Text1_Change
End Sub
