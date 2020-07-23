VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "挑战1"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form4"
   ScaleHeight     =   3285
   ScaleWidth      =   6315
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "说明"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运行 "
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "结果"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "请输入自然数b的值"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "请输入自然数a的值"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "输入数据不能为空！"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
  ElseIf Text1.Text = Text2.Text Then
        Text3.Text = "挑战失败！"
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
  ElseIf Text1.Text = 0 Or Text2.Text = 0 Then
        Text3.Text = "挑战成功！"
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
  Else
      Text3.Text = "挑战失败！"
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
  End If
End Sub

Private Sub Command2_Click()
  Load Form10
  Form10.Show
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要关闭窗口？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

