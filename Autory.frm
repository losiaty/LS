VERSION 5.00
Begin VB.Form Autory 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   495
      Left            =   480
      Picture         =   "Autory.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autoryzacja..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "Autory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim X As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long, i As Long

DoEvents
Sleep 800
'Unload Me
Hide

End Sub

