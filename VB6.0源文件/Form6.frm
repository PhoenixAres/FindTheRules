VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "挑战3"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form6"
   ScaleHeight     =   3300
   ScaleWidth      =   6045
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "说明"
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运行"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "结果"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "请输入彩蛋中的中文"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Text1.Text = "" Then
     MsgBox "输入数据不能为空！"
     Text1.SetFocus
     Exit Sub
   ElseIf Text1.Text = "辉耀" Then
         Text2.Text = "挑战成功！"
         Text1.Text = ""
         Text1.SetFocus
   Else
     Text2.Text = "挑战失败！"
     Text1.Text = ""
     Text1.SetFocus
   End If
End Sub

Private Sub Command2_Click()
   Load Form12
   Form12.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

