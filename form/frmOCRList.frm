VERSION 5.00
Begin VB.Form frmOCRList 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmOCRList"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmOCRList"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5760
      Width           =   7215
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   6735
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin projectEJ7302.newButton newButton3 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   1215
      Left            =   5280
      TabIndex        =   3
      Top             =   7440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      title           =   "삭제"
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1560
      Pattern         =   "*.dat"
      TabIndex        =   5
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "필기인식 삭제하기"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   3390
   End
   Begin VB.Image img 
      Height          =   4815
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmOCRList.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmOCRList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


Private Sub Form_Load()
    FadeIN Me
    
    
    '리스트 불러오기
        lstUpdate
End Sub

Private Sub lst_DblClick()
    '불량확인
    
        Dim str As String
        str = lst.List(lst.ListIndex)
        
        Debug.Print File1.Path
        If Not isDir(File1.Path & "\" & str, True) Then
            '폴더가없다
            MsgBox "OCR 데이터를 찾을 수 없습니다. 폴더명을 다시 확인해주세요.", vbCritical + vbSystemModal + vbOKOnly, "오류"
            Exit Sub
        ElseIf Not isDir(File1.Path & "\" & str & "\" & str & ".bmp") Then
            'no thumbnail
            MsgBox "OCR 데이터를 찾을 수 없습니다. 썸네일을 다시 확인해주세요.", vbCritical + vbSystemModal + vbOKOnly, "오류"
            Exit Sub
        ElseIf Not isDir(File1.Path & "\" & str & "\command.txt") Then
            'no command.text
            MsgBox "OCR 데이터를 찾을 수 없습니다. command.txt 파일을 다시 확인해주세요.", vbCritical + vbSystemModal + vbOKOnly, "오류"
            Exit Sub
        End If
        
    'thumbnail불러오기
        
        img.Picture = LoadPicture(File1.Path & "\" & str & "\" & str & ".bmp")
        
    'command불러오기
        Text1.Text = OpenFile(File1.Path & "\" & str & "\command.txt")
End Sub

Private Sub newButton1_Click()
    '삭제버튼
    Dim str$
    
    On Error GoTo h
    
    str = lst.List(lst.ListIndex)
    
    If isDir(File1.Path & "\" & str & ".dat") Then
        DeleteFile File1.Path & "\" & str & ".dat"
    End If
    If isDir(File1.Path & "\" & str, True) Then
        
        DeleteFile File1.Path & "\" & str & "\0.dat"
        DeleteFile File1.Path & "\" & str & "\1.dat"
        DeleteFile File1.Path & "\" & str & "\2.dat"
        DeleteFile File1.Path & "\" & str & "\3.dat"
        DeleteFile File1.Path & "\" & str & "\4.dat"
        DeleteFile File1.Path & "\" & str & "\" & str & ".bmp"
        DeleteFile File1.Path & "\" & str & "\command.txt"
        RmDir File1.Path & "\" & str
    End If
    
    
    If Not isDir(File1.Path & "\" & str & ".dat") And Not isDir(File1.Path & "\" & str, True) Then
        MsgBox "성공적으로 삭제되었습니다", vbInformation + vbSystemModal + vbOKOnly, "성공"
        lstUpdate
        img.Picture = Nothing
        Text1.Text = vbNullString
    Else
h:
    End If
End Sub

Private Sub newButton3_Click()
    frmOCRmain.Show
    Unload Me
End Sub


Sub lstUpdate()
    If Not isDir(App.Path & "\ocr\", True) Then Exit Sub
    
    'Dir1.Path = App.Path & "\ocr\"
    File1.Path = App.Path & "\ocr\"
    File1.Refresh
    lst.Clear
    
    Dim i%
    
    For i = 0 To File1.ListCount - 1&
        lst.AddItem Replace$(File1.List(i), ".dat", vbNullString)
    Next i
End Sub
