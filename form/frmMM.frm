VERSION 5.00
Begin VB.Form frmMM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmMM"
   ClientHeight    =   5595
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   3645
   LinkTopic       =   "frmMM"
   ScaleHeight     =   5595
   ScaleWidth      =   3645
   Tag             =   "8640 11535"
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim cp As POINTAPI
Dim oldcp As POINTAPI

Private Sub Form_Load()
    'FadeIN Me
    SetAlpha Me, 0.3
    AlwaysTop Me, True
    GetCursorPos oldcp
    
    Me.Left = oldcp.X * 15 + (50 * 15)
    Me.Top = oldcp.Y * 15 + (50 * 15)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
End Sub

Private Sub Text1_DblClick()
    Me.Hide
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Text1.Text = Clipboard.GetText
    End If
End Sub

Private Sub Timer1_Timer()

    GetCursorPos cp
        
    If GetAsyncKeyState(162) And GetAsyncKeyState(164) Then
    Else
        Me.Left = Me.Left + ((cp.X - oldcp.X) * 15)
        Me.Top = Me.Top + ((cp.Y - oldcp.Y) * 15)
    End If
    oldcp = cp
End Sub
