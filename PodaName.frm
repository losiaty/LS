VERSION 5.00
Begin VB.Form PodaName 
   Caption         =   "Kto gra ?"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Level"
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
      Begin VB.OptionButton LevelGracza 
         Caption         =   "Maniac"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton LevelGracza 
         Caption         =   "Hard"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton LevelGracza 
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton LevelGracza 
         Caption         =   "Easy"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Typ gracza"
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
      Begin VB.OptionButton TypGracza 
         Caption         =   "Komputer"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton TypGracza 
         Caption         =   "Cz³owiek"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton dawaj 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox imie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Komunikat 
      Alignment       =   2  'Center
      Caption         =   "Proszê podaæ Imiê / Nazwisko / Pseudonim zawodnika nr 1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "PodaName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public im As String, kom As Boolean, nrGR As Long
Dim tg As Long, lg As Long
Private Sub Dawaj_Click()
im = imie.Text
If TypGracza(1).Value Then
    kom = True
    If LevelGracza(0).Value Then MaxPsz = 0
    If LevelGracza(1).Value Then MaxPsz = 1
    If LevelGracza(2).Value Then MaxPsz = 2
    If LevelGracza(3).Value Then MaxPsz = 8
Else
    kom = False
End If
SaveSetting "£ów S³ów", "Ustawienia", "TypGracza" & CStr(nrGR), tg
SaveSetting "£ów S³ów", "Ustawienia", "LevelGracza" & CStr(nrGR), lg
PodaName.Hide
End Sub

Private Sub Form_Activate()
Dim i As Long
imie.Font.Size = 14
tg = CLng(GetSetting("£ów S³ów", "Ustawienia", "TypGracza" & CStr(nrGR), "0"))
lg = CLng(GetSetting("£ów S³ów", "Ustawienia", "LevelGracza" & CStr(nrGR), "0"))
LevelGracza(lg).Value = True
TypGracza(tg).Value = True
If tg = 0 Then
   Frame2.Enabled = False
   For i = 0 To 3
      LevelGracza(i).Enabled = False
   Next i
Else
   Frame2.Enabled = True
   For i = 0 To 3
      LevelGracza(i).Enabled = True
   Next i
End If
imie.SetFocus

End Sub

Private Sub imie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Dawaj_Click
End If
End Sub

Private Sub LevelGracza_Click(Index As Integer)
lg = Index
End Sub

Private Sub TypGracza_Click(Index As Integer)
Dim i As Long
If Index = 0 Then
   Frame2.Enabled = False
   For i = 0 To 3
      LevelGracza(i).Enabled = False
   Next i
Else
   Frame2.Enabled = True
   For i = 0 To 3
      LevelGracza(i).Enabled = True
   Next i

End If
tg = Index

End Sub

