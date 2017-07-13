VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmBGMSet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmBGMSet"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmBGMSet"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   7920
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin projectEJ7302.newButton newButton2 
      Height          =   1215
      Left            =   5280
      TabIndex        =   2
      Top             =   7440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      title           =   "실행"
   End
   Begin VB.TextBox Text1 
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
      Height          =   6495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   960
      Width           =   10575
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      title           =   "뒤로"
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmBGMSet.frx":0000
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmBGMSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim source As String

Private Sub Form_Load()
    FadeIN Me
    
    source = OpenFile(App.Path & "\scriptex.dll")
    'source = source & vbCrLf & OpenFile(App.Path & "\scriptdiy.dll")
    SC.AddCode source
    
    Text1 = OpenFile(App.Path & "\scriptdiy.dll")
    
    
End Sub

Private Sub newButton1_Click(): On Error GoTo h
    
    SC.AddCode source
    SC.AddCode Text1.Text
    
    Dim ff%
    ff = FreeFile
    Open App.Path & "\scriptdiy.dll" For Output As #ff
        Print #ff, Text1.Text
    Close #ff
    
    
    frmSet.Show
    Unload Me
    
    Exit Sub
h:
    If SC.Error.Number <> 0 Then
        MsgBox SC.Error.Description, vbCritical + vbSystemModal + vbOKOnly, SC.Error.Number
    End If
End Sub

Private Sub newButton2_Click(): On Error GoTo h
    SC.Reset
    SC.AddCode source
    SC.AddCode Text1.Text
    SC.Run InputBox("실행시킬 함수명을 입력해주세요", , "")
    
    Exit Sub
    
h:
    If SC.Error.Number <> 0 Then
        MsgBox SC.Error.Description, vbCritical + vbSystemModal + vbOKOnly, SC.Error.Number
    End If
End Sub

