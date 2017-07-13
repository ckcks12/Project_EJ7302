VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmOCRMake 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmOCRMake"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmOCRMake"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin projectEJ7302.newButton newButton1 
      Height          =   1215
      Left            =   5160
      TabIndex        =   4
      Top             =   7440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2143
      title           =   "만들기"
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   5760
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmOCRMake.frx":0000
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "필기의 이름을 입력하세요."
      Text            =   "필기의 이름을 입력하세요."
      Top             =   1560
      Width           =   5415
   End
   Begin projectEJ7302.newOCR OCR 
      Height          =   4815
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
   End
   Begin projectEJ7302.newButton newButton3 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin projectEJ7302.newOCR OCR 
      Height          =   4815
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
   End
   Begin projectEJ7302.newOCR OCR 
      Height          =   4815
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
   End
   Begin projectEJ7302.newOCR OCR 
      Height          =   4815
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
   End
   Begin projectEJ7302.newOCR OCR 
      Height          =   4815
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "필기인식 새로만들기"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1번째 필기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   1995
   End
   Begin VB.Image rightarrow 
      Height          =   1050
      Left            =   4200
      Picture         =   "frmOCRMake.frx":0040
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Image leftarrow 
      Height          =   1050
      Left            =   0
      Picture         =   "frmOCRMake.frx":3A7A
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmOCRMake.frx":74B4
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmOCRMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

Dim nowIndex As Integer
 

Private Sub Form_Load()
    FadeIN Me
    rightarrow.Left = OCR(0).Width - 1200
    nowIndex = 0
End Sub
 
Private Sub leftarrow_Click()
    If nowIndex > 0 Then
        OCR(nowIndex).Visible = False
        nowIndex = nowIndex - 1
        lb = nowIndex + 1 & "번째 필기"
        OCR(nowIndex).Visible = True
    End If
End Sub

Private Sub newButton1_Click()
    'input 검사
    '---제목검사
        If Trim$(txtTitle.Text) = Trim$(txtTitle.Tag) Then
            MsgBox "필기의 이름을 입력해주세요", vbCritical + vbSystemModal + vbOKOnly, "오류"
            Exit Sub
        End If
    '---스크립트검사
        On Error GoTo h
        SC.AddCode txtSource.Text
        If SC.Error.Number <> 0 Then
            GoTo h
        End If
        
    '---만들기
        '폴더검사
        'thumb만들고
        'data파일만들고
        'script.txt만들고
        Dim FolderPath As String
        
        If Not isDir(App.Path & "\ocr\", True) Then
            On Error GoTo w
            MkDir App.Path & "\ocr\"
        End If
        
        FolderPath = App.Path & "\ocr\" & txtTitle.Text
        If Not mod_IO.isDir(FolderPath, True) Then
            MkDir FolderPath
        End If
        
        OCR(0).Save App.Path & "\ocr\" & txtTitle.Text & ".dat" '대표dat파일
        
        OCR(0).SaveTo FolderPath & "\" & txtTitle.Text & ".bmp"  'thumbnail
        OCR(0).Save FolderPath & "\0.dat" 'data1
        OCR(1).Save FolderPath & "\1.dat" 'data2
        OCR(2).Save FolderPath & "\2.dat" 'data3
        OCR(3).Save FolderPath & "\3.dat"
        OCR(4).Save FolderPath & "\4.dat"
        
        mod_IO.PrintFile FolderPath & "\command.txt", txtSource.Text
        
        
        MsgBox "성공적으로 저장되었습니다", vbInformation + vbSystemModal + vbOKOnly, "성공"
        
        Call newButton3_Click '빨리 튀어야햄 안그러면 사용자 나븐자식이 두번클릭하면 오류 폭팔 ㅡㅡ
        
        Exit Sub
    
h:
    MsgBox SC.Error.Description, vbCritical + vbSystemModal + vbOKOnly, "구문 오류"
    Exit Sub
w:
    MsgBox Err.Description, vbCritical + vbSystemModal + vbOKOnly, "파일 쓰기 오류"
End Sub

Private Sub newButton3_Click()
    frmOCRmain.Show
    Unload Me
End Sub

Private Sub rightarrow_Click()
    If nowIndex < 4 Then
        OCR(nowIndex).Visible = False
        nowIndex = nowIndex + 1
        lb = nowIndex + 1 & "번째 필기"
        OCR(nowIndex).Visible = True
    End If
End Sub
