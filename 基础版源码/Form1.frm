VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Paint Copyright (c) 2022"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   12225
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   1920
      Min             =   1
      TabIndex        =   7
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "À¶(B)£º"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ÂÌ(G)£º"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ºì(R)£º"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = 255
HScroll2.Min = 0
HScroll2.Max = 255
HScroll3.Min = 0
HScroll3.Max = 255
VScroll1.Min = 1
VScroll1.Max = 15
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.CurrentX = X
Form1.CurrentY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Line -(X, Y), RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll1_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll2_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub HScroll3_Change()
Label4.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
Private Sub VScroll1_Change()
Form1.DrawWidth = VScroll1.Value
End Sub
