VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "��ս2"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form5"
   ScaleHeight     =   3150
   ScaleWidth      =   6135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "˵��"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "���"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "���������ֵ"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text1.Text = "" Then
    MsgBox "�������ݲ���Ϊ�գ�"
    Text1.SetFocus
    Exit Sub
  ElseIf Text1.Text = 2041 Then
        Text2.Text = "��ս�ɹ���"
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
  Else
    Text2.Text = "��սʧ�ܣ�"
    Text1.Text = ""
    Text1.SetFocus
  End If
End Sub

Private Sub Command2_Click()
  Load Form11
  Form11.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "���Ҫ�رմ��ڣ�"
  If MsgBox(msg, vbYesNo, "�˳�") = vbNo Then
    Cancel = True
  End If
End Sub

