VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmSet"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin projectEJ7302.newButton newButton1 
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
      title           =   "창투명화"
   End
   Begin projectEJ7302.newButton newButton8 
      Height          =   1935
      Left            =   3960
      TabIndex        =   5
      Top             =   3240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3413
      title           =   "사이트차단"
   End
   Begin projectEJ7302.newButton newButton7 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3201
      title           =   "돋보기"
   End
   Begin projectEJ7302.newButton newButton5 
      Height          =   1815
      Left            =   6360
      TabIndex        =   3
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3201
      title           =   "검색"
   End
   Begin projectEJ7302.newButton newButton4 
      Height          =   1815
      Left            =   2760
      TabIndex        =   2
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      title           =   "클립보드확장"
   End
   Begin projectEJ7302.newButton newButton2 
      Height          =   3735
      Left            =   8640
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6588
      title           =   "스크립트"
   End
   Begin projectEJ7302.newButton newButton3 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "설정"
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
      Left            =   9720
      TabIndex        =   7
      Top             =   120
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
End Sub

Private Sub newButton1_Click()
    frmLetmeseeSet.Show
    Unload Me
End Sub

Private Sub newButton2_Click() '스크립트
    frmBGMSet.Show
    Unload Me
End Sub

Private Sub newButton3_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub newButton4_Click()
    frmCBSet.Show
    Unload Me
End Sub

Private Sub newButton5_Click()
    frmSearchSet.Show
    Unload Me
End Sub

Private Sub newButton6_Click() '사전
    
End Sub

Private Sub newButton7_Click()
    frmZoomSet.Show
    Unload Me
End Sub

Private Sub newButton8_Click()
    frmBGM.Show
    Unload Me
End Sub
