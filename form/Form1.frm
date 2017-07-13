VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin projectEJ7302.newButton newButton4 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      title           =   "종료"
   End
   Begin projectEJ7302.newButton newButton3 
      Height          =   3255
      Left            =   7080
      TabIndex        =   2
      Top             =   4200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      title           =   "설정"
   End
   Begin projectEJ7302.newButton newButton2 
      Height          =   3255
      Left            =   3600
      TabIndex        =   1
      Top             =   4200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      title           =   "시작"
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      title           =   "필기인식"
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   7200
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   3600
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   0
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Image Image5 
      Height          =   870
      Left            =   0
      Picture         =   "Form1.frx":2FAA
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit


 Private Declare Function GetTickCount Lib "kernel32" () As Long
 

Private Sub Form_Load()
    FadeIN Me
     
    '--초기설정값불러와!
    value_Load
    'cbex folder
    If isDir(App.Path & "\cbex\", True) = False Then
        MkDir App.Path & "\cbex\"
    End If
    Dim a As Long
    If Len(CBEXFolder) Then Exit Sub
    a = GetTickCount
    If isDir(App.Path & "\cbex\" & a & "\", True) = False Then
        MkDir App.Path & "\cbex\" & a & "\"
    End If
    CBEXFolder = App.Path & "\cbex\" & a

    For a = 0 To value_CB_Max - 1
        If isDir(CBEXFolder & "\" & a & "\", True) = False Then
            MkDir CBEXFolder & "\" & a & "\"
        End If
    Next a
End Sub

Private Sub Image5_Click()
    FadeOUT Me
    End
End Sub

Private Sub newButton1_Click()
    frmOCRmain.Show
    Unload Me
End Sub

Private Sub newButton2_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub newButton3_Click()
    frmSet.Show
    Unload Me
End Sub

Private Sub newButton4_Click()
    FadeOUT Me
    End
End Sub

