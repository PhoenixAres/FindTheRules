VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "挑战2"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form5"
   ScaleHeight     =   3150
   ScaleWidth      =   6135
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "说明"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运行"
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
      Caption         =   "结果"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "请输入最大值"
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
    MsgBox "输入数据不能为空！"
    Text1.SetFocus
    Exit Sub
  ElseIf Text1.Text = 2041 Then
        Text2.Text = "挑战成功！"
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
  Else
    Text2.Text = "挑战失败！"
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
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

