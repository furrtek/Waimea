VERSION 5.00
Begin VB.Form SettingsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Load last opened file"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Alt shows all pin notes"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Live refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
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

Private Sub Command1_Click()
    LiveRefresh = C2B(Check1.Value)
    AltBubbles = C2B(Check2.Value)
    OpenLast = C2B(Check3.Value)
    
    If LocalSpacing <> Spacing Then
        Spacing = LocalSpacing
        RenderTicks
        Redraw
    End If
    
    SaveSettings
    
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Check1.Value = B2C(LiveRefresh)
    Check2.Value = B2C(AltBubbles)
    Check3.Value = B2C(OpenLast)
    
    LocalSpacing = Spacing
    
    HScroll1.Value = LocalSpacing * 10
End Sub

Private Sub HScroll1_Change()
    LocalSpacing = HScroll1.Value / 10
    Label1.Caption = "Scaling: " & LocalSpacing
End Sub
