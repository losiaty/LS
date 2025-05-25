VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Zacz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zaczynamy !"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar UD 
      Height          =   255
      Left            =   3960
      Max             =   0
      Min             =   10
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Iloœæ wymian ograniczona"
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CheckBox CCT 
      BackColor       =   &H8000000A&
      Caption         =   "Czas ca³kowity ograniczony"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox CRT 
      Caption         =   "Czas ruchu ograniczony"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Iloœæ uczestników"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   6255
      Begin VB.OptionButton Option1 
         Caption         =   "Klient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   5040
         TabIndex        =   20
         Top             =   200
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   3840
         TabIndex        =   19
         Top             =   200
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   3120
         TabIndex        =   11
         Top             =   200
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Top             =   200
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Top             =   200
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   200
         Width           =   645
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Domyœlne"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   30
      SmallChange     =   5
      Max             =   300
      SelStart        =   120
      TickFrequency   =   30
      Value           =   120
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Max             =   120
      SelStart        =   30
      TickFrequency   =   10
      Value           =   30
   End
   Begin VB.Label IlewCap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Jêzyk:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label CR 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label CC 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ustaw parametry gry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Zacz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IleW As Long

Private Sub CCT_Click()
Slider1.Visible = CCT.Value
CC.Visible = CCT.Value
End Sub

Private Sub Check1_Click()
IlewCap.Visible = Check1.Value
UD.Visible = Check1.Value
End Sub

Private Sub Command1_Click()
Dim i As Long, hj As Long, znak As String * 1, warto As Long, Plitery As String
Dim OldKap As String, Czasek As Single
Me.MousePointer = 11
Dim nIG As Long
Dim bServer As Boolean
Dim bKlient As Boolean

nIG = IG()

If nIG = 5 Then
    Set m_Server = New CServer
    nIG = 0
    IleGraczy = 0
ElseIf nIG = 6 Then

    Set m_Klient = New CKlient
    nIG = 1
    IleGraczy = 1
    graslow.Caption = graslow.Caption & " - Klient"
    graslow.Solve.Visible = True
    graslow.anulujemy.Visible = True

    Zacz.Hide
    Exit Sub
End If

If nIG = 0 Then
    Me.MousePointer = 0
    'Zacz.Hide
    'Exit Sub
End If



ZapiszDom1 nIG, IleW, Combo1.ListIndex, CRT.Value, CCT.Value, Check1.Value, Slider2.Value, Slider1.Value
IleGraczy = nIG
WymianaMydla = False
RMax = CLng(Slider2.Value)
FMax = CLng(Slider1.Value * 60)
CzasCTak = CCT.Value
CzasRTak = CRT.Value
Set Jêzyk = Jêzyki(Combo1.ListIndex + 1)
hj = FreeFile
Open App.Path & "\" & Jêzyk.Plik For Input As hj
Input #hj, Plitery
While Not EOF(hj)
    Input #hj, znak, warto
    Wart(Asc(znak)) = warto
    ReDim Preserve PLit(i + 1)
    PLit(i) = znak
    i = i + 1
Wend
For i = 0 To 97
    Worek(i) = Mid(Plitery, i + 1, 1)
Next i

OldKap = graslow.Caption
graslow.Caption = "Inicjownie bazy..."
If Not MojaBaza Is Nothing Then
   If MojaBaza.Sciezka <> App.Path & "\bazy\" & Left(Jêzyk.Klucz, 2) Then
      Set MojaBaza = New cBazaDanych
      MojaBaza.Inicjuj App.Path & "\bazy\" & Left(Jêzyk.Klucz, 2)
      graslow.DefMyd1.Clear
   
      For i = 0 To UBound(PLit) - 2
         graslow.DefMyd1.AddItem PLit(i)
      Next i
   End If
Else
   Set MojaBaza = New cBazaDanych
   MojaBaza.Inicjuj App.Path & "\bazy\" & Left(Jêzyk.Klucz, 2)
End If

If Not m_Klient Is Nothing Then
    graslow.Caption = OldKap & " - Klient"
ElseIf Not m_Server Is Nothing Then
    graslow.Caption = OldKap & " - Server"
Else
    graslow.Caption = OldKap
End If
Me.MousePointer = 0
Zacz.Hide
End Sub

Private Sub Command2_Click()
Slider1.Value = 20
Slider2.Value = 120
Option1(2).Value = True
CCT.Value = 1
CRT.Value = 1
Me.IlewCap.Visible = True
Me.IlewCap.Caption = "2"
IleWymian = 2
Check1.Value = 1
End Sub

Private Sub CRT_Click()
Slider2.Visible = CRT.Value
CR.Visible = CRT.Value

End Sub

Private Sub Form_Activate()
Dim item As tJezyk, dIleGraczy As Long, dNrJezyka As Long
IleW = CLng(GetSetting("£ów S³ów", "Ustawienia", "IleWymian", 3))
IlewCap.Caption = CStr(IleW)
dIleGraczy = CLng(GetSetting("£ów S³ów", "Ustawienia", "IleGraczy", 2))
If dIleGraczy = 0 Then dIleGraczy = 2

Option1(dIleGraczy).Value = True
'Slider1.Value = 0
'Slider2.Value = 0
Slider1.Value = CLng(GetSetting("£ów S³ów", "Ustawienia", "CzasCa³y", 2))
Slider2.Value = CLng(GetSetting("£ów S³ów", "Ustawienia", "CzasRuchu", 2))
CRT.Value = CLng(GetSetting("£ów S³ów", "Ustawienia", "CzasRuchuO", 0))
CCT.Value = CLng(GetSetting("£ów S³ów", "Ustawienia", "CzasCa³yO", 0))
Check1.Value = CLng(GetSetting("£ów S³ów", "Ustawienia", "WymianyO", 0))
Combo1.Clear
dNrJezyka = CLng(GetSetting("£ów S³ów", "Ustawienia", "Jêzyk", 1))
For Each item In Jêzyki
    Combo1.AddItem item.Nazwa
Next item
Combo1.Text = Combo1.List(dNrJezyka)
End Sub

Private Sub IleWCap_Change()
On Error Resume Next
IleW = CLng(IlewCap.Caption)
   UD.Value = IleW
End Sub

Private Sub Slider1_Change()
CC.Caption = Slider1.Value & " minut."
End Sub

Private Sub Slider1_Scroll()
CC.Caption = Slider1.Value & " minut."
End Sub

Private Sub Slider2_Change()
CR.Caption = graslow.FormaCzasu(Slider2.Value)
End Sub

Private Sub Slider2_Scroll()
CR.Caption = graslow.FormaCzasu(Slider2.Value)
End Sub

Private Function IG() As Long
Dim i As Long

For i = 1 To 4
    If Option1(i).Value Then
        IG = i
        Exit Function
    End If
Next i

If Option1(5).Value = True Then
    IG = 5
    MsgBox ("Server")
ElseIf Option1(6).Value = True Then
    IG = 6
    MsgBox ("Klient")
End If





End Function


Private Sub UD_Change()
IleW = UD.Value
IlewCap.Caption = CStr(UD.Value)
End Sub
