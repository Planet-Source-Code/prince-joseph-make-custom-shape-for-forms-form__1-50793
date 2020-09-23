VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Code"
   ClientHeight    =   6045
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "codefrm.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   7320
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   375
      Width           =   3765
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveAsFRM 
         Caption         =   "Save As Form"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "Close Script"
      End
   End
   Begin VB.Menu mnuEdt 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopyall 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flpath As String

Private Sub Form_Load()
Dim balcode As String
balcode = "Private Declare Function SetWindowRgn Lib ""user32"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long" & vbCrLf
balcode = balcode & "Private Declare Function CreatePolygonRgn Lib ""gdi32"" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long" & vbCrLf
balcode = balcode & "Dim down As Boolean" & vbCrLf
balcode = balcode & "Dim t As Integer '" & vbCrLf
balcode = balcode & "Dim w As Integer" & vbCrLf
balcode = balcode & "Private Type POINTAPI" & vbCrLf
balcode = balcode & "    x As Long" & vbCrLf
balcode = balcode & "    y As Long" & vbCrLf
balcode = balcode & "End Type" & vbCrLf
balcode = balcode & "Private Sub Form_Load()" & vbCrLf
balcode = balcode & "    me.picture=loadpicture(""" & flpath & """)" & vbCrLf
balcode = balcode & "    down = False" & vbCrLf
balcode = balcode & "    MakeShape Me.hWnd" & vbCrLf
balcode = balcode & "End Sub" & vbCrLf
balcode = balcode & "Private Function MakeShape(MHWND As Long)" & vbCrLf
balcode = balcode & "Dim point(" & Form1.counter & ") As POINTAPI" & vbCrLf
balcode = balcode & Form1.CodeStr & vbCrLf
balcode = balcode & "SetWindowRgn MHWND, CreatePolygonRgn(point(0), " & Form1.counter + 1 & ", 1), True" & vbCrLf
balcode = balcode & "End Function" & vbCrLf
balcode = balcode & "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)" & vbCrLf
balcode = balcode & "    down = True" & vbCrLf
balcode = balcode & "    w = x" & vbCrLf
balcode = balcode & "    t = y" & vbCrLf
balcode = balcode & "End Sub" & vbCrLf
balcode = balcode & "Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)" & vbCrLf
balcode = balcode & "    If down Then" & vbCrLf
balcode = balcode & "        Top = Top + y - t" & vbCrLf
balcode = balcode & "        Left = Left + x - w" & vbCrLf
balcode = balcode & "    End If" & vbCrLf
balcode = balcode & "End Sub" & vbCrLf
balcode = balcode & "Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)" & vbCrLf
balcode = balcode & "    down = False" & vbCrLf
balcode = balcode & "End Sub" & vbCrLf

Me.Text1.Text = balcode
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Text1.Top = 120
Me.Text1.Left = 120
Me.Text1.Height = Me.Height - 580
Me.Text1.Width = Me.Width - 340
End Sub

Private Sub mnuCopyall_Click()
Clipboard.SetText (Text1.Text)
End Sub

Private Sub mnuExt_Click()
Unload Me
End Sub
