VERSION 5.00
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   Caption         =   "彩蛋"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form9"
   ScaleHeight     =   1875
   ScaleWidth      =   2730
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim a(1 To 10, 1 To 10) As Integer
  Dim i As Integer
  Dim j As Integer
  For i = 1 To 10
     For j = 1 To 10
        If i = j Or i + j = 11 Then
           a(i, j) = 1
        Else
           a(i, j) = 0
        End If
        Print a(i, j);
    Next j
    Print
  Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub



