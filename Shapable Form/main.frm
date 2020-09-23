VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "MDIForm1"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8370
   Icon            =   "main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
            MinWidth        =   7832
            Text            =   "Skin Maker"
            TextSave        =   "Skin Maker"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1305
            MinWidth        =   1305
            Text            =   "X:"
            TextSave        =   "X:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1305
            MinWidth        =   1305
            Text            =   "Y:"
            TextSave        =   "Y:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:53 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuskinimg 
         Caption         =   "Open Skin Image"
      End
      Begin VB.Menu mnuggggg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuext 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tool"
      Begin VB.Menu mnureset 
         Caption         =   "Reset/Start"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGensrcipt 
         Caption         =   "Genarate Script"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuabt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isopen, isdra As Boolean

Private Sub MDIForm_Load()
isopen = False
isdra = False
End Sub

Private Sub mnuabt_Click()
frmAbout.Show 1
End Sub

Private Sub mnuExt_Click()
End
End Sub

Private Sub mnuGensrcipt_Click()
Form2.Show
End Sub

Private Sub mnureset_Click()
Form1.Picture2.Visible = False
Form1.firstclick = True
Form1.Picture11.Cls
mnuGensrcipt.Enabled = False
End Sub

Private Sub mnuskinimg_Click()
With Me.CommonDialog1
    .DialogTitle = "Open Skin Picture"
    .Filter = "All Picture Formats|*.jpg;*.gif;*.bmp"
    .ShowOpen
    If .FileName <> "" Then
      Form1.Picture11.Picture = LoadPicture(.FileName)
      Form2.flpath = .FileName
        isopen = True
        isdra = False
        mnureset.Enabled = True
        mnuGensrcipt.Enabled = False
    End If
    
End With
End Sub
