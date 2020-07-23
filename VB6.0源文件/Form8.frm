VERSION 5.00
Begin VB.Form Form8 
   AutoRedraw      =   -1  'True
   Caption         =   "²Êµ°"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim i As Integer
  For i = 1 To 20
     DrawWidth = i
     PSet (i * 300, 1000), vbBlue
  Next i
End Sub
 
