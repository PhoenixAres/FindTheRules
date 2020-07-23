VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "游戏说明"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   2850
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(10); "游戏说明"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(6); "输入你喜欢的两个自然数"
  Print
  Print Tab(8); "根据结果寻找规律"
  Print
  Print Tab(6); "完成挑战，成为课间娱乐"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

