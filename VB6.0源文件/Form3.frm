VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "�ʵ�ϵͳ˵��"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form3"
   ScaleHeight     =   2940
   ScaleWidth      =   2850
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Print
  Print Tab(2); String(30, "*")
  Print Tab(10); "�ʵ�ϵͳ˵��"
  Print Tab(2); String(30, "*")
  Print
  Print
  Print Tab(6); "����������Ĳ�ͬ����"
  Print
  Print Tab(3); "����ֲ�ͬ�����벻���Ľ��"
  Print
  Print Tab(8); "����ǡ��ʵ�����"
End Sub
  
Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub
