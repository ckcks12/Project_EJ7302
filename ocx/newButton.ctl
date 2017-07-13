VERSION 5.00
Begin VB.UserControl newButton 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   ScaleHeight     =   4830
   ScaleWidth      =   5160
   Begin VB.Label lb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "newButton"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   5175
   End
End
Attribute VB_Name = "newButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event DblClick()
Event MouseExit()
Event MouseMove(X As Single, Y As Single)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim R As Integer
Dim G As Integer
Dim b As Integer
Dim dostop As Boolean
Dim coloring As Boolean
Public fading As Boolean 'fadein, fadeout할때 컬러링이벤트때문에 작업진행이안됨..
Dim saveX As Single, saveY As Single
 
 
Private Sub lb_Click()
    RaiseEvent Click
End Sub

Private Sub lb_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Color R, G, b
    dostop = False
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    lb.Left = 0
    lb.Top = (UserControl.Height / 10) * 5
    lb.Width = UserControl.Width
    
    Randomize
    R = Int(Rnd * 256 / 10) * 10: G = Int(Rnd * 256 / 10) * 10: b = Int(Rnd * 256 / 10) * 10
    UserControl.BackColor = RGB(R, G, b)
    lb.ForeColor = RGB(255 - R, 255 - G, 255 - b)
    'lb.ForeColor = cc(R, G, b)
    dostop = False
End Sub

Function Color(ByVal RR As Integer, ByVal GG As Integer, ByVal BB As Integer): On Error Resume Next
    
    If coloring Then Exit Function
    coloring = True
    Do
        
        Do Until RR = R And BB = b And GG = G
            If RR < R Then R = R - 10
            If RR > R Then R = R + 10
            If GG < G Then G = G - 10
            If GG > G Then G = G + 10
            If BB < b Then b = b - 10
            If BB > b Then b = b + 10
             
            Dim before As Long
                before = GetTickCount()
                Do Until GetTickCount() >= before + 10
                    DoEvents
                Loop

            If dostop Then dostop = False:  coloring = False: Exit Function
    
            UserControl.BackColor = RGB(R, G, b)
            lb.ForeColor = RGB(255 - R, 255 - G, 255 - b)
            'lb.ForeColor = cc(R, G, b)
            DoEvents
            
            'If (saveX < 100 Or saveY < 100 Or saveX > UserControl.Width - 100 Or saveY > UserControl.Height - 100) Then
                dostop = True
            'End If
        Loop
        
        Randomize
        RR = Int(Rnd * 256 / 10) * 10
        BB = Int(Rnd * 256 / 10) * 10
        GG = Int(Rnd * 256 / 10) * 10
    Loop
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not fading Then
        Color R, G, b
        dostop = False
    End If
End Sub
    

Private Sub UserControl_Resize()
    lb.Left = 0
    lb.Top = (UserControl.Height / 10) * 5
    lb.Width = UserControl.Width
End Sub

Property Get Title() As String
    Title = lb.Caption
End Property
Property Let Title(ByVal new_tmp As String)
    lb.Caption = new_tmp
    PropertyChanged "title"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lb.Caption = PropBag.ReadProperty("title", vbNullString)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("title", lb.Caption, vbNullString)
End Sub
 
 
Function cc(ByVal R%, ByVal G%, ByVal b%) As Long
    Dim max%, min%
    If R > G Then
        If R > b Then
            max = R
            If G > b Then
                min = b
            Else
                min = G
            End If
        Else
            max = b
            If R > G Then
                min = G
            Else
                min = R
            End If
        End If
    Else
        If G > b Then
            max = G
            If R > b Then
                min = b
            Else
                min = R
            End If
        Else
            max = b
            If R > G Then
                min = G
            Else
                min = R
            End If
        End If
    End If
        
    max = max + min
    
    cc = RGB(max - R, max - G, max - b)
End Function
