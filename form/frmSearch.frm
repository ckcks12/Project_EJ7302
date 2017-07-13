VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmSearch"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmSearch"
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
      Height          =   735
      Left            =   1328
      TabIndex        =   0
      Top             =   3953
      Width           =   7935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "°Ë»ö¾î¸¦ ÀÔ·ÂÇÏ¼¼¿ä"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   3795
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Load()
    FadeIN Me
End Sub

Private Sub Image3_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ShellExecute 0&, "open", Replace$(value_Search_Url, "°Ë»ö¾î", Text1.Text), vbNullString, vbNullString, vbNormalFocus
        frmMain.Show
        Unload Me
    End If
End Sub
