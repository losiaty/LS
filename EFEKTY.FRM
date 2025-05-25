VERSION 5.00
Begin VB.Form efekty 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Efekty wysi³ków"
   ClientHeight    =   5655
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "efekty.frx":0000
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Anuluj"
      DownPicture     =   "efekty.frx":1A4A1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      MaskColor       =   &H008080FF&
      Picture         =   "efekty.frx":1C0BB
      TabIndex        =   1
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "efekty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public decyzja As Boolean
Public Sub CancelButton_Click()
decyzja = False
Lst.Clear
Unload efekty
End Sub
Private Sub LST_GotFocus()
OKButton.SetFocus
End Sub

Private Sub Form_Activate()
Set Lst.obiekt = Me
Lst.Drukuj 10, 10
Me.MousePointer = 0
End Sub
Private Sub OKButton_Click()
decyzja = True
Lst.Clear
Unload efekty
End Sub
