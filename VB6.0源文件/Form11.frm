VERSION 5.00
Begin VB.Form Form11 
   AutoRedraw      =   -1  'True
   Caption         =   "挑战2说明"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   LinkTopic       =   "Form11"
   ScaleHeight     =   2805
   ScaleWidth      =   2835
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(12); "挑战2说明"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print
  Print Tab(3); "请输入你所发现的结果的最大值"
  Print
  Print
  Print Tab(12); "彩蛋除外！"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub


