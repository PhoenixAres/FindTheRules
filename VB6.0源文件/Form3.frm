VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "彩蛋系统说明"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form3"
   ScaleHeight     =   2940
   ScaleWidth      =   2850
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(10); "彩蛋系统说明"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(6); "根据你输入的不同的数"
  Print
  Print Tab(3); "会出现不同的意想不到的结果"
  Print
  Print Tab(8); "这便是“彩蛋”！"
End Sub
  
Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub
