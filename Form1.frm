VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Waimea"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   579
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   32
      Left            =   120
      Max             =   1024
      SmallChange     =   8
      TabIndex        =   1
      Top             =   2280
      Width           =   8415
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
      Text            =   "Form1.frx":000C
      Top             =   2640
      Width           =   8415
   End
   Begin VB.Menu menu_sheet 
      Caption         =   "Sheet"
      Begin VB.Menu menu_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu open_save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_export 
         Caption         =   "Export"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Done As Boolean
    Dim frm As Form
    Done = False
    
    Dim lidx As Integer
    Dim didx As Integer
    Dim lline As String
    Dim a() As String
    Dim b() As String
    Dim pidx As Integer
    Dim t As Integer
    
    Dim c As Integer
    Dim d As Integer
    
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
    
    If Not CreateGLWindow(Me, 640, 480, 16, False) Then
        Done = True                             ' Quit If Window Was Not Created
    End If

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

Private Sub Form_Resize()
    ReSizeGLScene ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillGLWindow
End Sub

Private Sub HScroll1_Change()
    Text1_Change
End Sub

Private Sub menu_open_Click()
    Dim ln As String
    
    If Text1.Text <> "" Then Exit Sub
    
    Open "waveform.txt" For Input As #1
        While Not EOF(1)
            Line Input #1, ln
            Text1.Text = Text1.Text & ln & vbCrLf
        Wend
    Close #1
End Sub

Private Sub open_save_Click()
    Open "waveform.txt" For Output As #1
        Print #1, Text1.Text
    Close #1
End Sub

Private Sub Text1_Change()
    DrawGLScene Form1
    SwapBuffers Form1.hDC
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = True
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Keys(KeyCode) = False
End Sub
