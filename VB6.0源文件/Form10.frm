VERSION 5.00
Begin VB.Form Form10 
   AutoRedraw      =   -1  'True
   Caption         =   "��ս1˵��"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form10"
   ScaleHeight     =   2925
   ScaleWidth      =   2880
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(12); "��ս1˵��"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(8); "������������Ȼ��"
  Print
  Print Tab(8); "ʹ���Ϊ��2016����"
  Print
  Print Tab(1); "ע����������Ϊ�ʵ�������սʧ��"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub

