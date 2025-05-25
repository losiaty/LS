VERSION 5.00
Begin VB.Form Brak 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Nie ma w s³owniku !!!"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Dopisz 
      BackColor       =   &H00C0C000&
      Caption         =   "Akceptuj i dopisz do s³ownika"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton Nie 
      BackColor       =   &H000000FF&
      Caption         =   "Nie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton tak 
      BackColor       =   &H0000C000&
      Caption         =   "Tak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Komunikat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Takiego wyrazu nie ma w s³owniku. Czy wszyscy uczestnicy zgadzaj¹ siê na jego u¿ycie ?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Brak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Dopisz_Click()
BrakAnswer = 3
Brak.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Public Sub Nie_Click()
BrakAnswer = 2
Brak.Hide
End Sub

Public Sub tak_Click()
BrakAnswer = 1
Brak.Hide
End Sub
