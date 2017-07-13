VERSION 5.00
Begin VB.Form frmBGM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmBGM"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmBGM"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      TabIndex        =   2
      Text            =   "127.0.0.1 www.domain.com"
      Top             =   6600
      Width           =   10575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5580
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   10575
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      title           =   "µÚ·Î"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»çÀÌÆ® Â÷´ÜÇÏ±â"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmBGM.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmBGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
    
    Update
End Sub
 
 
 
Sub Update(): On Error Resume Next
    Dim old As VbFileAttribute
    old = GetAttr("C:\Windows\System32\drivers\etc\hosts")
    SetAttr "C:\Windows\system32\drivers\etc\hosts", vbNormal
    
    List1.Clear
    Call OpenFile("C:\Windows\System32\drivers\etc\hosts", True, List1, vbCrLf)
    SetAttr "C:\Windows\system32\drivers\etc\hosts", old
    
    
End Sub

Private Sub List1_DblClick()
    If MsgBox(List1.List(List1.ListIndex) & vbCrLf & "»èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbYesNo, "") = vbNo Then Exit Sub
    
    Dim old As VbFileAttribute, a As String, ff%: ff = FreeFile
    old = GetAttr("C:\Windows\System32\drivers\etc\hosts")
    SetAttr "C:\Windows\system32\drivers\etc\hosts", vbNormal
    
    a = OpenFile("C:\Windows\System32\drivers\etc\hosts")
    a = Replace$(a, List1.List(List1.ListIndex), vbNullString)
    
    Open "C:\Windows\System32\drivers\etc\hosts" For Output As #ff
        Print #ff, a
    Close #ff
    SetAttr "C:\Windows\system32\drivers\etc\hosts", old
    
    Update
End Sub

Private Sub newButton1_Click()
    frmSet.Show
    Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim old As VbFileAttribute, a As String
        old = GetAttr("C:\Windows\System32\drivers\etc\hosts")
        SetAttr "C:\Windows\system32\drivers\etc\hosts", vbNormal
        
        PrintFile "C:\Windows\System32\drivers\etc\hosts", Text1.Text
        SetAttr "C:\Windows\system32\drivers\etc\hosts", old
    End If
    Update
End Sub
