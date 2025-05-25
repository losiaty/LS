VERSION 5.00
Begin VB.Form PlusMinus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dodaj/Usuñ s³owo"
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Anuluj"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
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
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Usuñ"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Dodaj"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PlusMinus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Dim wr As String, tbl As String
wr = Trim(UCase(Pytanie.Text))
If Len(wr) > 1 And Len(wr) < 16 Then
   MousePointer = 11
   MojaBaza.Remove wr
   MousePointer = 0
End If
Me.Hide
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
Pytanie.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27
    Me.Hide
Case 13
    If OKButton.Enabled Then
        OKButton_Click
    Else
        CancelButton_Click
    End If
End Select
End Sub

Private Sub OKButton_Click()
Dim wr As String, tbl As String
wr = Trim(UCase(Pytanie.Text))
MousePointer = 11
MojaBaza.Add wr
MousePointer = 0
Me.Hide
End Sub

Private Sub Pytanie_GotFocus()
Pytanie.SelStart = 0
Pytanie.SelLength = Len(Pytanie.Text)
End Sub

