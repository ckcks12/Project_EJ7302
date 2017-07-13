VERSION 5.00
Begin VB.Form frmZoomSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmZoomSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmZoomSet"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin projectEJ7302.newButton newButton1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "돋보기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   9360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "확대율"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmZoomSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmZoomSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
    value_Load
    Label1 = "확대율 " & value_Zoom_Zoom
End Sub

Private Sub Image4_Click()

End Sub

Private Sub Label1_DblClick()
    Dim tmp$, a%
    tmp = InputBox("1~100사이의 값을 입력해주세요", "")
    a = CInt(tmp)
    
    If a >= 1 And a <= 100 Then
        value_Zoom_Zoom = a
        value_Save
        Label1 = "확대율 " & value_Zoom_Zoom
    End If
End Sub

Private Sub newButton1_Click()
    frmSet.Show
    Unload Me
End Sub
