VERSION 5.00
Begin VB.Form Form12 
   AutoRedraw      =   -1  'True
   Caption         =   "��ս3˵��"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2805
   LinkTopic       =   "Form12"
   ScaleHeight     =   2805
   ScaleWidth      =   2805
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(12); "��ս3˵��"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(6); "�ʵ����Բ�ͬ����ʽ����"
  Print
  Print
  Print Tab(8); "������ʵ��е�����"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub

