VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Skin Picture"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ControlBox      =   0   'False
   Icon            =   "pic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture11 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3015
      ScaleHeight     =   15.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   13.5
      TabIndex        =   0
      Top             =   6420
      Width           =   270
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   3240
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   3960
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public firstclick As Boolean
Public startdraw As Boolean
Public firstX As Single
Public firstY As Single
Public lastX As Single
Public lastY As Single
Public counter As Integer
Public CodeStr As String



Private Sub Form_Load()
firstclick = True

End Sub



'
Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If firstclick = True Then
   Me.Picture2.Left = x - (Picture2.Width / 2)
   Me.Picture2.Top = y - (Picture2.Height / 2)
   Me.Picture2.Visible = True
   firstX = x
   firstY = y
   counter = 0
   CodeStr = "point(0).x = " & x * (1.33) & vbCrLf
   CodeStr = CodeStr & "point(0).y = " & y * (1.33) & vbCrLf

Else



End If


End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If firstclick = False And Picture2.Visible = True Then
   Picture11.AutoRedraw = False
   Picture11.Cls
   Picture11.Line (lastX, lastY)-(x, y)
   Picture11.AutoRedraw = True
End If
End Sub
'
Private Sub Picture11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If firstclick = False And Picture2.Visible = True Then
   Picture11.Line (lastX, lastY)-(x, y)
   counter = counter + 1
   CodeStr = CodeStr & "point(" & counter & ").x = " & x * (1.33) & vbCrLf
   CodeStr = CodeStr & "point(" & counter & ").y = " & y * (1.33) & vbCrLf
Else
 

End If
    firstclick = False
    lastX = x
    lastY = y
End Sub



Private Sub Picture11_Resize()
Picture11.Top = 10
Picture11.Left = 10
Height = Picture11.Picture.Height * (1 / 1.61) + 20
Width = Picture11.Picture.Width * (1 / 1.61) + 20
If Height < 4695 Then Height = 4695

End Sub
Private Sub Form_Resize()
On Error Resume Next
'Shape1.Height = Picture11.Picture.Width * (1 / 1.6) - 20
End Sub

Private Sub Picture2_Click()
If firstclick = False And Picture2.Visible = True Then
   Picture11.Line (lastX, lastY)-(firstX, firstY)
   Picture2.Visible = False
   MDIForm1.mnuGensrcipt.Enabled = True
End If

End Sub
