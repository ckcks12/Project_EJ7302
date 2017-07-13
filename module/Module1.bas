Attribute VB_Name = "mod_Form"
Option Explicit

 
Public Const LWA_ALPHA As Long = &H2
Public Const GCL_STYLE As Long = -26&
Public Const GWL_STYLE As Long = -16&
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_LAYERED As Long = &H80000
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const TOPMOST_FLAGS = &H2 Or &H1



Function FadeIN(Frm As Form): On Error Resume Next
        
   Const timeFadeIn As Long = 6 '점점 불투명해지는 간격(밀리초)
   Dim lastTimer As Single, i As Long

    '---align
    Dim tmpHeight As Long, tmpWidth As Long
    Dim dHeight As Double, dWidth As Double
    Dim obj As Object
    
    '화면늘어난비율구해서
    'tmpHeight = Split(Frm.Tag, " ")(0)
    'tmpWidth = Split(Frm.Tag, " ")(1)
    dHeight = Screen.Height / Frm.Height
    dWidth = Screen.Width / Frm.Width
    
    '위치조정하고 다시 비지블
    For Each obj In Frm
        'obj.Left = (Val(Split(obj.Tag, " ")(0)) + (obj.Width / 2)) * dWidth - (obj.Width / 2)
        'obj.Top = (Val(Split(obj.Tag, " ")(1)) + (obj.Height / 2)) * dHeight - (obj.Height / 2)
        obj.Left = obj.Left * dWidth
        obj.Top = obj.Top * dHeight
        obj.Width = obj.Width * dWidth
        obj.Height = obj.Height * dHeight
        'obj.Left = obj.Left * dWidth
        'obj.Top = obj.Top * dHeight
        obj.Visible = True
        '---페이드효과와 컬러링효과 노이즈방지
        obj.fading = True
    Next
    '---

   '페이드 인(Fade In) 효과
   Frm.Enabled = False
   Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
   Call SetLayeredWindowAttributes(Frm.hwnd, 0, 0, LWA_ALPHA)
   Frm.Show
   For i = 1 To 255 Step 7
      lastTimer = Timer
      Do While Timer < lastTimer + (timeFadeIn / 1000)
         DoEvents
      Loop
      Call SetLayeredWindowAttributes(Frm.hwnd, 0, i, LWA_ALPHA)
      DoEvents
   Next
   Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Xor WS_EX_LAYERED)
   Frm.Enabled = True
    
    '---fading원상복귀
    For Each obj In Frm
        obj.fading = False
    Next
End Function

 Function FadeOUT(Frm As Form): On Error Resume Next

   Const timeFadeIn As Long = 3 '점점 불투명해지는 간격(밀리초)
   Dim lastTimer As Single, i As Long

    Dim obj As Object
        
    '---페이드하고컬러링노이즈방지
    For Each obj In Frm
        obj.fading = True
    Next

   '페이드 아웃(Fade Out) 효과
   'frm.Enabled = False
   Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
   Call SetLayeredWindowAttributes(Frm.hwnd, 0, 0, LWA_ALPHA)
   'frm.Show
   For i = 255 To 0 Step -7
      lastTimer = Timer
      Do While Timer < lastTimer + (timeFadeIn / 1000)
         DoEvents
      Loop
      Call SetLayeredWindowAttributes(Frm.hwnd, 0, i, LWA_ALPHA)
      DoEvents
   Next
   Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, GWL_EXSTYLE) Xor WS_EX_LAYERED)
   Frm.Enabled = True
   Frm.Hide
   
   '---fading원상복귀할필요없겟징
   
End Function


Function SetAlpha(ByRef Frm As Form, Optional ByVal Alpha As Double = 0.4)
    Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, GetWindowLong(Frm.hwnd, -20) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Frm.hwnd, 0, 255 * (value_SetAlpha_Alpha), 2)
    
End Function

'Function SetAlphaEX(ByRef Frm As Form, Optional ByVal Alpha As Double = 0.4)
'    Dim ret As Long
'    ret = ret Or WS_EX_LAYERED                  '투명스타일 적용
'    Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, ret)
'    Call SetLayeredWindowAttributes(Frm.hwnd, vbMagenta, 0, 1) '자홍색(vbMagenta)을 투명으로 변경
'End Function
Function SetAlphaEX(ByVal hwnd As Long, Optional ByVal Alpha As Double = 0.4)
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, -20) Or WS_EX_LAYERED)
    If Alpha = 1 Then
        Call SetLayeredWindowAttributes(hwnd, 0, 255, 2)
    Else
        Call SetLayeredWindowAttributes(hwnd, 0, 255 * (value_SetAlpha_Alpha), 2)
    End If
    If value_SetAlpha_AlwaysTop Then
        AlwaysTopEX hwnd, True
    Else
        AlwaysTopEX hwnd, False
    End If
    If Alpha = 1 Then
        AlwaysTopEX hwnd, False
    End If
End Function


'### 폼 최상위로
Function AlwaysTop(Frm As Form, ByVal Use As Boolean)
    If Use Then
        SetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
End Function


Function AlwaysTopEX(ByVal hwnd As Long, ByVal Use As Boolean)
    If Use Then
        SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
    Else
        SetWindowPos hwnd, -2, 0, 0, 0, 0, 3
    End If
End Function
