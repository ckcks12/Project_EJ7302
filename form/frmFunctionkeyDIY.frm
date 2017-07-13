VERSION 5.00
Begin VB.Form frmFunctionkeyDIY 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmFunctionkeyDIY"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmFunctionkeyDIY"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   3000
      Width           =   3975
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
   Begin VB.Image Image1 
      Height          =   870
      Left            =   3360
      Top             =   1560
      Width           =   3495
   End
End
Attribute VB_Name = "frmFunctionkeyDIY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FadeIN Me
End Sub

