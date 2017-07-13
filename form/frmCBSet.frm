VERSION 5.00
Begin VB.Form frmCBSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmCBSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmCBSet"
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "그림자동저장"
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
      TabIndex        =   4
      Top             =   3840
      Width           =   2430
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "클립보드확장"
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
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "스크립트사용"
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
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "최대버퍼"
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
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmCBSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmCBSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
    
    value_Load
    Label1 = "최대버퍼 " & value_CB_Max
    Label2 = "스크립트사용 " & value_CB_Script
    Label4 = "그림자동저장 " & value_CB_PictureAutoSave
    
End Sub

Private Sub Image4_Click()
End Sub

Private Sub Label1_DblClick()
    Dim tmp$, a%
    tmp = InputBox("1~100 사이 값을 입력하세요")
    a = CInt(tmp)
    If a >= 1 And a <= 100 Then
        value_CB_Max = a
        value_Save
        Label1 = "최대버퍼 " & value_CB_Max
    End If
    
    If isDir(App.Path & "\cbex\", True) = False Then
        MkDir App.Path & "\cbex\"
    End If

    For a = 0 To value_CB_Max - 1
        If isDir(CBEXFolder & "\" & a & "\", True) = False Then
            MkDir CBEXFolder & "\" & a & "\"
        End If
    Next a
End Sub

Private Sub Label2_DblClick()
    value_CB_Script = value_CB_Script Xor True
    value_Save
    Label2 = "스크립트사용 " & value_CB_Script
End Sub

Private Sub Label4_Click()
    value_CB_PictureAutoSave = value_CB_PictureAutoSave Xor True
    value_Save
    Label4 = "그림자동저장 " & value_CB_PictureAutoSave
End Sub

Private Sub newButton1_Click()
    frmSet.Show
    Unload Me
End Sub
