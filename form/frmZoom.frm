VERSION 5.00
Begin VB.Form frmZoom 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmZoom"
   ClientHeight    =   3435
   ClientLeft      =   2310
   ClientTop       =   60
   ClientWidth     =   4305
   DrawMode        =   6  'Mask Pen Not
   LinkTopic       =   "frmZoom"
   MinButton       =   0   'False
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   Tag             =   "8640 11535"
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5160
      Top             =   3720
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private dHDC As Long, pAPI As POINTAPI
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private xVal As Long
Private Sub Form_DblClick()
    'On Error Resume Next
      'xVal = CLng(InputBox("Please enter a number, which will be the base size to be zoomed from the screen. (Example: 45)", , "45"))
End Sub

Private Sub Form_Load()
    xVal = value_Zoom_Zoom
    dHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    SetWindowPos Me.hwnd, -1, 0, 0, 150, 150, &H1 Or &H2

End Sub
Private Sub Form_Paint()
    SetWindowPos Me.hwnd, -1, 0, 0, 150, 150, &H1 Or &H2
End Sub

Private Sub Form_Resize(): On Error Resume Next
    Me.Width = Me.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteDC dHDC
End Sub


Private Sub Timer1_Timer()
    GetCursorPos pAPI
    StretchBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleWidth, dHDC, pAPI.X - (xVal / 2), pAPI.Y - (xVal / 2), xVal, xVal, &HCC0020
    Me.ZOrder

End Sub


