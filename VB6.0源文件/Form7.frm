VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   Caption         =   "彩蛋"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI = 3.141592
Private Sub Form_Load()
  Dim i As Double
  Form7.Width = 6000
  Form7.Height = 6000
  Line (0, 3000)-(ScaleWidth, 3000)
  Line (3000, 0)-(3000, ScaleHeight)
  For i = 0 To 2 * PI Step 0.01
     PSet (3000 + 2000 * Sin(2 * i), 3000 + 2000 * Sin(3 * i))
  Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub


