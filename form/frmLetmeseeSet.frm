VERSION 5.00
Begin VB.Form frmLetmeseeSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmLetmeseeSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmLetmeseeSet"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "창투명화"
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
      Left            =   9000
      TabIndex        =   3
      Top             =   120
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "항상위"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "투명도"
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
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmLetmeseeSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmLetmeseeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
    value_Load
    Label1 = "투명도 " & value_SetAlpha_Alpha
    Label2 = "항상위 " & value_SetAlpha_AlwaysTop
End Sub

Private Sub Image2_Click()
End Sub

Private Sub Label1_DblClick()
    Dim tmp$, a As Double
    tmp = InputBox("투명도를 입력해주세요." & vbCrLf & "0부터 1사이의 소수점 둘째자리까지 유효합니다.", "", "0.5")
    tmp = Left$(tmp, 3)
    
    a = CDbl(tmp)
    If Not a = 0 And a < 1 And a > 0 Then
        value_SetAlpha_Alpha = a
        value_Save
    End If
    Label1 = "투명도 " & value_SetAlpha_Alpha
End Sub

Private Sub Label2_DblClick()
    value_SetAlpha_AlwaysTop = value_SetAlpha_AlwaysTop Xor True
    value_Save
    Label2 = "항상위 " & value_SetAlpha_AlwaysTop
End Sub

Private Sub newButton1_Click()
    frmSet.Show
    Unload Me
End Sub
