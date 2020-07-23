VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "趣味找规律"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   7140
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "挑战3"
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "挑战2"
      Height          =   615
      Left            =   2760
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "挑战1"
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "彩蛋系统说明"
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "游戏说明"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运行"
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   2760
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "结果"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "请输入自然数b的值"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "请输入自然数a的值"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim a As Integer
   Dim b As Integer
   Dim c As Integer
   Dim d As String
   If Text1.Text = "" Or Text2.Text = "" Then
     Text1.Text = ""
     Text2.Text = ""
     Text1.SetFocus
     MsgBox "输入数据不能为空！"
     Exit Sub
   Else
     a = Text1.Text
     b = Text2.Text
     c = a * b * (10 - a * b) + 2016
     d = "输入数据a=" & a & "  " & "输入数据b=" & b & "  " & "输出数据c=" & c
  End If
  If a = 0 And b = 0 Then
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Load Form7
    Form7.Show
    Exit Sub
  ElseIf a = 5 And b = 3 Then
      Text3.Text = "辉耀"
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
      Exit Sub
  ElseIf a = 3 And b = 5 Then
      Text3.Text = "辉耀"
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
      Exit Sub
  ElseIf a = b Then
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
      Load Form8
      Form8.Show
      Exit Sub
  ElseIf a = 2 And b = 1 Then
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
      Load Form9
      Form9.Show
      Exit Sub
  ElseIf a = 1 And b = 2 Then
      Text1.Text = ""
      Text2.Text = ""
      Text1.SetFocus
      Load Form9
      Form9.Show
      Exit Sub
  Else
    Text3.Text = d
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
  End If
End Sub

Private Sub Command2_Click()
   Load Form2
   Form2.Show
End Sub

Private Sub Command3_Click()
   Load Form3
   Form3.Show
End Sub

Private Sub Command4_Click()
   Load Form4
   Form4.Show
End Sub

Private Sub Command5_Click()
   Load Form5
   Form5.Show
End Sub

Private Sub Command6_Click()
   Load Form6
   Form6.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, Unloadmode As Integer)
  Dim msg As String
  msg = "真的要退出程序？"
  If MsgBox(msg, vbYesNo, "退出") = vbNo Then
    Cancel = True
  End If
End Sub

