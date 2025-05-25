VERSION 5.00
Begin VB.Form Szukaj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jest czy nie ma ?"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Pytanie 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Zamknij"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Wpisz szukane s這wo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Szukaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
Pytanie.SetFocus
End Sub

Private Sub Pytanie_GotFocus()
Pytanie.SelStart = 0
Pytanie.SelLength = Len(Pytanie.Text)
End Sub

Private Sub Pytanie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Me.Hide
    Exit Sub
End If
If KeyCode = 13 Then
    Szuk
    Command1.SetFocus
    Pytanie.SetFocus
End If
End Sub

Public Sub Szuk()
If Len(Pytanie.Text) < 2 Or Len(Pytanie.Text) > 15 Then
    MsgBox "Nie ma takiego s這wa !", vbCritical
    Exit Sub
End If
On Error GoTo nm
If graslow.SlownikCheck(UCase(Pytanie.Text)) = False Then
    MsgBox "Takie s這wo jest w s這wniku", vbInformation
Else
nm:    MsgBox "Nie ma takiego s這wa !", vbCritical
End If
Pytanie.SetFocus

End Sub
