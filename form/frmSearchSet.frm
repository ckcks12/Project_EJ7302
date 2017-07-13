VERSION 5.00
Begin VB.Form frmSearchSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmSearchSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmSearchSet"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   645
      Left            =   1920
      TabIndex        =   6
      Top             =   5040
      Width           =   7935
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "기타"
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
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "구글"
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
      Top             =   4440
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "네이트"
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
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "다음"
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
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "네이버"
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
      Width           =   1935
   End
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "검색"
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
      Top             =   240
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmSearchSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmSearchSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
    
    value_Load
    
    asdf
End Sub

Private Sub Image4_Click()
End Sub

Private Sub Image3_Click()
End Sub

Private Sub newButton1_Click()
    If InStr(value_Search_Url, "검색어") = 0 And Option5.Value = True Then
        MsgBox "검색어가 입력될 위치에 '검색어' 라고 입력해주십시오.", vbCritical + vbSystemModal + vbOKOnly, ""
    Else
        frmSet.Show
        value_Save
        Unload Me
    End If
End Sub

Private Sub Option1_Click()
    
    value_Search_Url = "http://search.naver.com/search.naver?where=nexearch&query=검색어"
    
    asdf
    
End Sub

Private Sub Option2_Click()
    value_Search_Url = "http://search.daum.net/search?w=tot&t__nil_searchbox=btn&sug=&q=검색어"
    
    asdf
End Sub

Private Sub Option3_Click()
    value_Search_Url = "http://search.nate.com/search/all.html?s=&sc=&afc=&j=&thr=sbma&nq=&q=검색어"
    asdf
End Sub

Private Sub Option4_Click()
    value_Search_Url = "http://www.google.co.kr/?gws_rd=cr#newwindow=1&output=search&sclient=psy-ab&q=검색어&oq=검색어"
    asdf
End Sub

Private Sub Option5_Click()
    value_Search_Url = Text1.Text
    asdf
End Sub

Sub asdf()
    
    Dim s As String
    s = value_Search_Url
    
    If InStr(s, "naver") Then
        Option1.Value = True
    ElseIf InStr(s, "daum") Then
        Option2.Value = True
    ElseIf InStr(s, "nate") Then
        Option3.Value = True
    ElseIf InStr(s, "google") Then
        Option4.Value = True
    Else
        Option5.Value = True
        Text1 = s
    End If
End Sub
