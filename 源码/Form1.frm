VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Paint by Morgan"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   14670
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll5 
      Height          =   255
      Left            =   5280
      Max             =   2000
      Min             =   1
      TabIndex        =   66
      Top             =   120
      Value           =   1
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3240
      TabIndex        =   64
      Text            =   "line"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退    出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清    除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   1440
      Max             =   20
      Min             =   1
      TabIndex        =   7
      Top             =   2160
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   1440
      Max             =   255
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   1440
      Max             =   255
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1440
      Max             =   255
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   4
      X1              =   3120
      X2              =   0
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label8 
      Caption         =   "半径："
      Height          =   255
      Left            =   4800
      TabIndex        =   65
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "自定义颜色——单击此处存储"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   63
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400040&
      Height          =   255
      Index           =   47
      Left            =   2640
      TabIndex        =   61
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      Height          =   255
      Index           =   46
      Left            =   2280
      TabIndex        =   60
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      Height          =   255
      Index           =   45
      Left            =   1920
      TabIndex        =   59
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004000&
      Height          =   255
      Index           =   44
      Left            =   1560
      TabIndex        =   58
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004040&
      Height          =   255
      Index           =   43
      Left            =   1200
      TabIndex        =   57
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Height          =   255
      Index           =   42
      Left            =   840
      TabIndex        =   56
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000040&
      Height          =   255
      Index           =   41
      Left            =   480
      TabIndex        =   55
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   54
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Height          =   255
      Index           =   39
      Left            =   2640
      TabIndex        =   53
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Height          =   255
      Index           =   38
      Left            =   2280
      TabIndex        =   52
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Height          =   255
      Index           =   37
      Left            =   1920
      TabIndex        =   51
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Height          =   255
      Index           =   36
      Left            =   1560
      TabIndex        =   50
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008080&
      Height          =   255
      Index           =   35
      Left            =   1200
      TabIndex        =   49
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004080&
      Height          =   255
      Index           =   34
      Left            =   840
      TabIndex        =   48
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   255
      Index           =   33
      Left            =   480
      TabIndex        =   47
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   46
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
      Height          =   255
      Index           =   31
      Left            =   2640
      TabIndex        =   45
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Height          =   255
      Index           =   30
      Left            =   2280
      TabIndex        =   44
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   29
      Left            =   1920
      TabIndex        =   43
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   28
      Left            =   1560
      TabIndex        =   42
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   27
      Left            =   1200
      TabIndex        =   41
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   255
      Index           =   26
      Left            =   840
      TabIndex        =   40
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   25
      Left            =   480
      TabIndex        =   39
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   38
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   23
      Left            =   2640
      TabIndex        =   37
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   22
      Left            =   2280
      TabIndex        =   36
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   21
      Left            =   1920
      TabIndex        =   35
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   20
      Left            =   1560
      TabIndex        =   34
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   1200
      TabIndex        =   33
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   255
      Index           =   18
      Left            =   840
      TabIndex        =   32
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   480
      TabIndex        =   31
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   29
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   28
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Index           =   13
      Left            =   1920
      TabIndex        =   27
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Height          =   255
      Index           =   12
      Left            =   1560
      TabIndex        =   26
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Index           =   11
      Left            =   1200
      TabIndex        =   25
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   24
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   23
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   21
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   20
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   19
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   17
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   16
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      BorderWidth     =   4
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   7200
   End
   Begin VB.Label Label4 
      Caption         =   "粗细："
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "蓝(B)："
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "绿(G)："
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "红(R)："
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Cls
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
HScroll4.Min = 1
HScroll4.Max = 20
Combo1.AddItem "line"
'Combo1.AddItem "circle"
'Combo1.AddItem "straight line"
Label8.Visible = False
HScroll5.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Combo1.Text = "line" Then
Label8.Visible = False
HScroll5.Visible = False
If Button = 1 Then Line -(X, Y), Line1.BorderColor
End If
'If Combo1.Text = "circle" Then
'Label8.Visible = True
'HScroll5.Visible = True
'If Button = 1 Then Circle (X, Y), ((M - X) ^ 2 + (N - Y) ^ 2) ^ (1 / 2), Line1.BorderColor
'End If
'If Combo1.Text = "line" Then
'Label8.Visible = False
'HScroll5.Visible = False
'If Button = 1 Then Line Step(X, Y)-Step(M, N), Line1.BorderColor
'End If
End Sub
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form1.CurrentX = M
'Form1.CurrentY = N
'End Sub
Private Sub HScroll1_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = HScroll1.Value
End Sub
Private Sub HScroll2_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.Text = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
Line1.BorderColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text3.Text = HScroll3.Value
End Sub
Private Sub HScroll4_Change()
Form1.DrawWidth = HScroll4.Value
Line1.BorderWidth = HScroll4.Value
Text4.Text = HScroll4.Value
End Sub
Private Sub Label5_Click(Index As Integer)
Line1.BorderColor = Label5(Index).BackColor
End Sub
Private Sub label6_click()
Line1.BorderColor = Label6.BackColor
End Sub
Private Sub Label7_Click()
Label6.BackColor = Line1.BorderColor
End Sub
Private Sub Text1_Change()
If Text1.Text = "" Then
Text1.Text = 0
End If
If Text1.Text > 255 Then
Text1.Text = 255
End If
HScroll1.Value = Text1.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text2_Change()
If Text2.Text = "" Then
Text2.Text = 0
End If
If Text2.Text > 255 Then
Text2.Text = 255
End If
HScroll2.Value = Text2.Text
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text3_Change()
If Text3.Text = "" Then
Text3.Text = 0
End If
If Text3.Text > 255 Then
Text3.Text = 255
End If
HScroll3.Value = Text3.Text
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text4_Change()
If Text4 = "" Then
Text4.Text = 1
End If
If Text4.Text = 0 Then
Text4.Text = 1
End If
If Text4.Text > 20 Then
Text4.Text = 20
End If
HScroll4.Value = Text4.Text
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
