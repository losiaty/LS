VERSION 5.00
Begin VB.Form Tytul 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "£ów S³ów"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton won 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6120
      Picture         =   "Tytul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   4095
   End
   Begin VB.CommandButton Dawaj 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   600
      Picture         =   "Tytul.frx":136F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   4605
   End
End
Attribute VB_Name = "Tytul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dawaj_Click()

'On Error Resume Next
'Word.Application.Quit

'WordOff:
'On Error GoTo NieMaWorda
'Set Czek = New Word.Application
'JestWord = True

Me.MousePointer = 11
'film.Stop
'film.Cancel
'film.Enabled = False
Me.Hide
'Zacz.Show
graslow.Show
'Exit Sub

'NieMaWorda:
'DoEvents
'graslow.InfoBox "W twoim systemie nie ma zainstalowanego programu Microsoft Word w wersji 32-bit. Sprawdzanie pisowni ograniczy siê do w³asnego s³ownika programu '£ów S³ów'", False, False
'JestWord = False
'Me.MousePointer = 11
'DoEvents
'film.Stop
'film.Cancel
'film.Enabled = False
'Me.Hide
'Zacz.Show
'graslow.Show

End Sub

Private Sub Form_Load()
'On Error GoTo BladFilm

'tytul.Film.Visible = True
'Tytul.Film.Open App.Path & "\LOWSLOW.avi"
'Film.Play
'BladFilm:
'Exit Sub
End Sub

Private Sub Label1_Click()

End Sub

Private Sub won_Click()

Call KoniecLS
End Sub
