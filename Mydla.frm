VERSION 5.00
Begin VB.Form Mydla 
   BorderStyle     =   0  'None
   Caption         =   "Dopasowanie ""Myde≥ka"""
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Mydlonie 
      Caption         =   "Anuluj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton mydlotak 
      Caption         =   "Zastosuj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox asciiM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox znakM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Wprowadü nowy kod ASCII ""Myde≥ka"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wprowadü nowy znak ""Myde≥ka"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Mydla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub asciiM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then mydlotak.SetFocus
End Sub

Private Sub asciiM_LostFocus()
On Error Resume Next
znakM.Text = Chr$(CInt(asciiM.Text))
End Sub

Private Sub Mydlonie_Click()
Mydla.Hide
End Sub

Private Sub mydlotak_Click()
Dim bk As Long, i As Long, j As Long

For i = 0 To 97
    If znakM.Text = Worek(i) Then
        MsgBox "Ten znak nie moøe byÊ uøyty jako myd≥o", vbOKOnly
        Exit Sub
    End If
Next i

OldMydlo = Mydlo
Mydlo = znakM.Text
For i = 0 To 224
    If Pole(i).Caption = OldMydlo And Pole(i).Tag = True Then
        bk = Pole(i).NumerObrazka
        Call Pole(i).Po≥Ûø(Mydlo, True)
        Call Pole(i).Zatwierdü(bk)
    End If
Next i

For i = 1 To 4
    For j = 0 To 6
        If Stojak(i, j) = OldMydlo Then Stojak(i, j) = Mydlo
    Next j
Next i

For i = 0 To 6
    If plytka(i).Caption = OldMydlo Then
        bk = plytka(i).NumerObrazka
        Call plytka(i).Po≥Ûø(Mydlo, True)
        Call plytka(i).Zatwierdü(bk)
    End If
Next i
Worek(98) = Mydlo
Worek(99) = Mydlo
Mydla.Hide

End Sub

Private Sub znakM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then mydlotak.SetFocus
End Sub

Private Sub znakM_LostFocus()
asciiM.Text = Asc(znakM)
End Sub
