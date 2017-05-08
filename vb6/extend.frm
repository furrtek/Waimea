VERSION 5.00
Begin VB.Form ExtendFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extend waves"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "extend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "extend.frx":0E42
      Left            =   3240
      List            =   "extend.frx":0E44
      TabIndex        =   8
      Text            =   "Block"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "8"
      Top             =   330
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "extend.frx":0E46
      Left            =   240
      List            =   "extend.frx":0E48
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Before"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "After"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "With block:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Extend by:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "ExtendFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim WaveDefs As String
    Dim WaveDefsLn() As String
    Dim WDSplit() As String
    Dim c, d As Integer
    Dim WDStart, WDEnd As Integer
    Dim WD As String
    Dim OldWD As String
    Dim ToAdd As String
    Dim nFill As Integer
    
    ' Can be very buggy !
    
    nFill = Val(Text1.Text)
    
    If nFill <> 0 Then
    
        WaveDefs = MainFrm.EditBox.Text
        WaveDefsLn = Split(WaveDefs, vbCrLf)
        
        If UBound(WaveDefsLn) >= 0 Then
            If Combo1.ListIndex > 0 Then ToAdd = String(nFill, Combo1.Text)
        
            For c = 0 To nWaves - 1
                If List1.Selected(c) = True Then
                    List1.ListIndex = c
                    For d = 0 To UBound(WaveDefsLn)
                        If InStr(1, WaveDefsLn(d), "name:" & List1.Text & ";") > 0 Then
                            ' Find wave def in line
                            WDStart = InStr(1, WaveDefsLn(d), "wave:")
                            If WDStart > 0 Then
                                WDStart = WDStart + 5
                                WDEnd = InStr(WDStart, WaveDefsLn(d), ";")
                                If WDEnd = 0 Then WDEnd = Len(WaveDefsLn(d)) + 1
                                ' Found
                                WD = Mid(WaveDefsLn(d), WDStart, WDEnd - WDStart)
                                OldWD = WD
                                If Option1.Value = True Then
                                    ' After
                                    If Combo1.ListIndex = 0 Then ToAdd = String(nFill, Right(WD, 1))
                                    WD = WD & ToAdd
                                Else
                                    ' Before
                                    If Combo1.ListIndex = 0 Then ToAdd = String(nFill, Left(WD, 1))
                                    WD = ToAdd & WD
                                End If
                                
                                WaveDefsLn(d) = Replace(WaveDefsLn(d), OldWD, WD)
                            End If
                        End If
                    Next d
                End If
            Next c
            
            ToAdd = ""
            For c = 0 To UBound(WaveDefsLn) - 1
                ToAdd = ToAdd & WaveDefsLn(c) & vbCrLf
            Next c
            MainFrm.EditBox.Text = ToAdd & WaveDefsLn(c)
        End If
    
    End If
    
    Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim c, i As Integer
    
    List1.Clear
    
    i = 0
    For c = 0 To nWaves - 1
        If Waves(c).Used = True Then
            List1.AddItem Waves(c).Name
            'List1.ItemData(i) = c
            List1.Selected(i) = True
            i = i + 1
        End If
    Next c
    
    List1.ListIndex = 0
    
    Combo1.AddItem "Last one"
    Combo1.AddItem "."
    Combo1.AddItem "x"
    Combo1.AddItem "z"
    Combo1.AddItem "h"
    Combo1.AddItem "l"
    
    Combo1.ListIndex = 1
End Sub

