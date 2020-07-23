VERSION 5.00
Begin VB.Form Form10 
   AutoRedraw      =   -1  'True
   Caption         =   "挑战1说明"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form10"
   ScaleHeight     =   2925
   ScaleWidth      =   2880
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(12); "挑战1说明"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(8); "请输入两个自然数"
  Print
  Print Tab(8); "使结果为“2016”！"
  Print
  Print Tab(1); "注：若输出结果为彩蛋，则挑战失败"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

