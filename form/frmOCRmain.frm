VERSION 5.00
Begin VB.Form frmOCRmain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmOCRmain"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmOCRmain"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin projectEJ7302.newButton newButton3 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
      title           =   "삭제하기"
   End
   Begin projectEJ7302.newButton newButton2 
      Height          =   3255
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5741
      title           =   "새로만들기"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "필기인식"
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
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmOCRmain.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmOCRmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub Form_Load()
    FadeIN Me
End Sub
 

Private Sub newButton1_Click()
    frmOCRList.Show
    Unload Me
End Sub

Private Sub newButton2_Click()
    frmOCRMake.Show
    Unload Me
End Sub

Private Sub newButton3_Click()
    Form1.Show
    Unload Me
End Sub
