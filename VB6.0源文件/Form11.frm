VERSION 5.00
Begin VB.Form Form11 
   AutoRedraw      =   -1  'True
   Caption         =   "��ս2˵��"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   LinkTopic       =   "Form11"
   ScaleHeight     =   2805
   ScaleWidth      =   2835
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(12); "��ս2˵��"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print
  Print Tab(3); "�������������ֵĽ�������ֵ"
  Print
  Print
  Print Tab(12); "�ʵ����⣡"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub


