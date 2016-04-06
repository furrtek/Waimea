VERSION 5.00
Begin VB.Form SettingsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Caption         =   "Anti-aliasing"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "settings.frx":0E42
      Left            =   2760
      List            =   "settings.frx":0E4C
      TabIndex        =   9
      Text            =   "Color scheme"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   127
      Min             =   5
      TabIndex        =   7
      Top             =   960
      Value           =   5
      Width           =   4455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Load last opened file"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Alt shows all pin notes"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Live refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   360
      Value           =   1
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Groups opacity:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Scaling:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "SettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LocalSpacing As Single
Dim LocalGroupAlpha As Integer
Dim NeedRefresh As Boolean

Private Sub Check4_Click()
    NeedRefresh = True
End Sub

Private Sub Command1_Click()
    LiveRefresh = C2B(Check1.Value)
    AltBubbles = C2B(Check2.Value)
    OpenLast = C2B(Check3.Value)
    AntiAliasing = C2B(Check4.Value)
    
    If (Combo1.ListIndex <> ColorScheme) And Combo1.ListIndex >= 0 Then
        ColorScheme = Combo1.ListIndex
        LoadColorScheme
        NeedRefresh = True
    End If
    
    If LocalSpacing <> Spacing Then
        Spacing = LocalSpacing
        RenderTicks
        NeedRefresh = True
    End If
    
    If LocalGroupAlpha <> GroupAlpha Then
        GroupAlpha = LocalGroupAlpha
        NeedRefresh = True
    End If
    
    If NeedRefresh = True Then Redraw
    
    SaveSettings
    
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    NeedRefresh = False
    
    Check1.Value = B2C(LiveRefresh)
    Check2.Value = B2C(AltBubbles)
    Check3.Value = B2C(OpenLast)
    Check4.Value = B2C(AntiAliasing)
    
    LocalSpacing = Spacing
    LocalGroupAlpha = GroupAlpha
    
    Combo1.ListIndex = ColorScheme
    
    HScroll1.Value = LocalSpacing * 10
    HScroll2.Value = LocalGroupAlpha
End Sub

Private Sub HScroll1_Change()
    LocalSpacing = HScroll1.Value / 10
    Label1.Caption = "Scaling: " & LocalSpacing
End Sub

Private Sub HScroll2_Change()
    LocalGroupAlpha = HScroll2.Value
    Label2.Caption = "Groups opacity: " & LocalGroupAlpha
End Sub
