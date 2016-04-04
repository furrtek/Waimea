VERSION 5.00
Begin VB.Form SettingsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Live refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Scaling:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
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
    If Check1.Value = vbChecked Then
        LiveRefresh = True
    Else
        LiveRefresh = False
    End If
    If LocalSpacing <> Spacing Then
        Spacing = LocalSpacing
        RenderTicks
        MainFrm.Redraw
    End If
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    If LiveRefresh = True Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    
    LocalSpacing = Spacing
    
    HScroll1.Value = LocalSpacing * 2
End Sub

Private Sub HScroll1_Change()
    LocalSpacing = HScroll1.Value / 2
    Label1.Caption = "Scaling: " & LocalSpacing
End Sub
