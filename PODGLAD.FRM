VERSION 5.00
Begin VB.Form Podglad 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   64.823
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   130.44
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Koniec"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drukuj"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Ost 
      Caption         =   "Ostatnia"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Pier 
      Caption         =   "Pierwsza"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox nrStrony 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton nas 
      Caption         =   ">>"
      Height          =   255
      Left            =   3360
      MaskColor       =   &H00FFFF80&
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton pop 
      Caption         =   "<<"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox tlo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   46.831
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   82.815
      TabIndex        =   2
      Top             =   600
      Width           =   4695
      Begin VB.PictureBox obraz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         MousePointer    =   15  'Size All
         ScaleHeight     =   48.419
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   80.169
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.HScrollBar pPoziomy 
      Height          =   255
      LargeChange     =   30
      Left            =   120
      SmallChange     =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   5055
   End
   Begin VB.VScrollBar pPionowy 
      Height          =   3495
      LargeChange     =   30
      Left            =   5280
      SmallChange     =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lStrony 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "27"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lZ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " z"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   227
   End
   Begin VB.Label lStrona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Strona"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "Podglad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpx As Integer, tmpy As Integer, licz As Boolean
Public PrintSource As Object, strona As Integer, strony As Integer, NazwaWydruku As String

Private Sub Command1_Click()
Form1.DrukujWyniki Printer
Printer.EndDoc
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim skok As Integer, krok As Integer
krok = 20
skok = tlo.Height
If KeyCode = 36 Then
    pPionowy.Value = pPionowy.Min
    pPoziomy.Value = pPoziomy.Min
    Exit Sub
End If

If KeyCode = 35 Then
    pPionowy.Value = pPionowy.Max
    pPoziomy.Value = pPoziomy.Max
    Exit Sub
End If

If KeyCode = 34 Then
    If pPionowy.Max > pPionowy.Value + skok Then
        pPionowy.Value = pPionowy.Value + skok
    Else
        pPionowy.Value = pPionowy.Max
    End If
    Exit Sub
End If

If KeyCode = 33 Then
    If pPionowy.Min < pPionowy.Value - skok Then
        pPionowy.Value = pPionowy.Value - skok
    Else
        pPionowy.Value = pPionowy.Min
    End If
    Exit Sub
End If

If KeyCode = 38 Then
    If pPionowy.Min < pPionowy.Value - krok Then
        pPionowy.Value = pPionowy.Value - krok
    Else
        pPionowy.Value = pPionowy.Min
    End If
    Exit Sub
End If

If KeyCode = 40 Then
    If pPionowy.Max > pPionowy.Value + krok Then
        pPionowy.Value = pPionowy.Value + krok
    Else
        pPionowy.Value = pPionowy.Max
    End If
    Exit Sub
End If

If KeyCode = 39 Then
    If pPoziomy.Max > pPoziomy.Value + krok Then
        pPoziomy.Value = pPoziomy.Value + krok
    Else
        pPoziomy.Value = pPoziomy.Max
    End If
    Exit Sub
End If

If KeyCode = 37 Then
    If pPoziomy.Min < pPoziomy.Value - krok Then
        pPoziomy.Value = pPoziomy.Value - krok
    Else
        pPoziomy.Value = pPoziomy.Min
    End If
    Exit Sub
End If

End Sub

Private Sub Form_Resize()

obraz.Width = 210
obraz.Height = 297
Pier.Top = Podglad.ScaleHeight - Pier.Height
Pier.Left = 0
pop.Left = Pier.Width
pop.Top = Pier.Top
lStrona.Top = Pier.Top
lStrona.Left = pop.Left + pop.Width
lZ.Top = Pier.Top
Command1.Left = 5
Command1.Top = 1
Command2.Top = 1
Command2.Left = Command1.Left + Command1.Width + 2
nrStrony.Top = Pier.Top
nrStrony.Left = lStrona.Left + lStrona.Width
lZ.Left = nrStrony.Left + nrStrony.Width
lStrony.Top = Pier.Top
lStrony.Left = lZ.Left + lZ.Width
nas.Top = Pier.Top
nas.Left = lStrony.Left + lStrony.Width
Ost.Top = Pier.Top
Ost.Left = nas.Left + nas.Width
pPionowy.Left = Podglad.ScaleWidth - pPionowy.Width
pPoziomy.Top = Podglad.ScaleHeight - pPoziomy.Height
pPoziomy.Left = Ost.Left + Ost.Width + 5
pPoziomy.Width = Podglad.ScaleWidth - pPoziomy.Left - pPionowy.Width
pPionowy.Height = Podglad.ScaleHeight - pPoziomy.Height - 10
pPionowy.Top = 10
tlo.Left = 0
tlo.Top = 10
tlo.Width = Podglad.ScaleWidth - pPionowy.Width
tlo.Height = Podglad.ScaleHeight - pPoziomy.Height - 10

pPoziomy.Min = -2
pPoziomy.Max = obraz.Width - tlo.ScaleWidth + 2
pPionowy.Min = -2
pPionowy.Max = obraz.Height - tlo.ScaleHeight + 2
pPionowy.Value = -2
pPoziomy.Value = -2
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub nas_Click()

strona = strona + 1
'If strona = strony Then nas.Enabled = False

PrintSource.tabprint strona

End Sub


Private Sub obraz_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
End Sub

Private Sub obraz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If licz = False Then
        tmpx = X
        tmpy = Y
        licz = True
    End If
    obraz.MousePointer = 99
End If
End Sub

Private Sub obraz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dx As Integer, dy As Integer, newx As Integer, newy As Integer
If Button = 1 Then
    dy = Int(tmpy - Y)
    dx = Int(tmpx - X)
    
    newx = pPoziomy.Value + dx
    newy = pPionowy.Value + dy
    
    Select Case newx
        Case Is > pPoziomy.Max: pPoziomy.Value = pPoziomy.Max
        Case Is < pPoziomy.Min: pPoziomy.Value = pPoziomy.Min
        Case Else: pPoziomy.Value = newx
    End Select
    
    Select Case newy
        Case Is > pPionowy.Max: pPionowy.Value = pPionowy.Max
        Case Is < pPionowy.Min: pPionowy.Value = pPionowy.Min
        Case Else: pPionowy.Value = newy
    End Select
End If
End Sub

Private Sub obraz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
licz = False
obraz.MousePointer = 15
End Sub

Private Sub Pier_Click()
Form1.DrukujWyniki Printer
End Sub

Private Sub pop_Click()
strona = strona - 1
'If strona = 1 Then pop.Enabled = False
PrintSource.tabprint strona
End Sub

Private Sub pPionowy_Change()
obraz.Top = -pPionowy.Value
End Sub

Private Sub pPionowy_Scroll()
obraz.Top = -pPionowy.Value
End Sub

Private Sub pPoziomy_Change()
obraz.Left = -pPoziomy.Value
End Sub

Private Sub pPoziomy_Scroll()
obraz.Left = -pPoziomy.Value
End Sub
