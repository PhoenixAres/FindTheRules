VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "��Ϸ˵��"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   2850
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(10); "��Ϸ˵��"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(6); "������ϲ����������Ȼ��"
  Print
  Print Tab(8); "���ݽ��Ѱ�ҹ���"
  Print
  Print Tab(6); "�����ս����Ϊ�μ�����"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub

