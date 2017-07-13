VERSION 5.00
Begin VB.Form frmFindWindow 
   Caption         =   "FindWindow 1.0"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "frmFindWindow"
   ScaleHeight     =   795
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3000
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmFindWindow.frx":0000
      MousePointer    =   1  'Arrow
      Picture         =   "frmFindWindow.frx":030A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "창 이름:"
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   420
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "창 핸들:"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmFindWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function InvertRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private bFind As Boolean
Private Last As RECT, LastHandle As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub Form_Load()
    bFind = True
    Picture1.Picture = Nothing
    Picture1.MousePointer = vbCustom
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bFind Then Exit Sub
    
    Dim pt As POINTAPI, Handle As Long, lLen&, sTitle$

    pt.X = X: pt.Y = Y
    
    ClientToScreen Picture1.hwnd, pt
    Handle = WindowFromPoint(pt.X, pt.Y)
    If CheckMe(Handle) Then Exit Sub
    If LastHandle > 0 Then
        InvertWindow LastHandle
    End If
    LastHandle = Handle
    InvertWindow Handle
End Sub

Sub InvertWindow(ByVal lHwnd As Long)
    Dim rc As RECT, hDCWindow As Long, hRgn1&, hRgn2&, hRgn3&
    GetWindowRect lHwnd, rc
    hDCWindow = GetWindowDC(lHwnd)
    hRgn1 = CreateRectRgn(0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top)
    hRgn2 = CreateRectRgn(3, 3, rc.Right - rc.Left - 3, rc.Bottom - rc.Top - 3)
    CombineRgn hRgn1, hRgn1, hRgn2, 4
    InvertRgn hDCWindow, hRgn1
    DeleteObject hRgn2
    DeleteObject hRgn1
    ReleaseDC lHwnd, hDCWindow
End Sub

Function CheckMe(ByVal lHwnd As Long) As Boolean
    Dim Root As Long
    Root = lHwnd
    If Root = Me.hwnd Then CheckMe = True: Exit Function
    
    Do While Root
        lHwnd = Root
        
        Root = GetParent(Root)
        If Root = Me.hwnd Then CheckMe = True: Exit Function
    Loop
End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bFind = False
    
    InvertWindow LastHandle
    
    Last.Left = 0
    Last.Top = 0
    Last.Right = 0
    Last.Bottom = 0
    Picture1.Picture = Picture1.MouseIcon
    Picture1.MousePointer = vbArrow
    
    SetAlphaEX LastHandle
    
    Unload Me
End Sub

Private Sub Timer1_Timer()
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyZ) Then
        Dim a As POINTAPI
        GetCursorPos a
        SetAlphaEX WindowFromPoint(a.X, a.Y)
    ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyMenu) And GetAsyncKeyState(vbKeyX) Then
        GetCursorPos a
        SetAlphaEX WindowFromPoint(a.X, a.Y), 1
    End If
End Sub
