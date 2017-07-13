VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form2"
   ScaleHeight     =   405
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = Screen.Width
    Me.Left = 0
    Label1.Width = Me.Width
    'AlwaysTopEX Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    AlwaysTopEX Me, False
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub
