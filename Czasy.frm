VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Czasy 
   Caption         =   "Ustawienia czasów"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider pasek2 
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   60
      Max             =   1200
      TickFrequency   =   30
   End
   Begin MSComctlLib.Slider pasek1 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   120
      Max             =   7200
      TickFrequency   =   120
   End
   Begin VB.CommandButton RDef 
      Caption         =   "Domyœlny"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton FDef 
      Caption         =   "Domyœlny"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CzasAnuluj 
      Caption         =   "Anuluj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton CzasOK 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label CzasRuchu 
      Caption         =   "Czas "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label CzasCa³y 
      Caption         =   "Ca³kowity czas "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Czasy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CzasAnuluj_Click()
Me.Hide
End Sub

Private Sub CzasOK_Click()
FMax = pasek1.Value
RMax = pasek2.Value
Call graslow.UstawCzasy
Me.Hide
End Sub

Private Sub FDef_Click()
pasek1.Value = DFMax
CzasCa³y.Caption = "Ca³kowity czas : " & graslow.FormaCzasu(DFMax)
End Sub

Private Sub Form_Load()
pasek1.Value = FMax
pasek2.Value = RMax
CzasCa³y.Caption = "Ca³kowity czas :" & graslow.FormaCzasu(FMax)
CzasRuchu.Caption = "Czas ruchu :" & graslow.FormaCzasu(RMax)


End Sub

Private Sub pasek1_Change()

CzasCa³y.Caption = "Ca³kowity czas : " & graslow.FormaCzasu(pasek1.Value)

End Sub

Private Sub pasek1_Scroll()
CzasCa³y.Caption = "Ca³kowity czas : " & graslow.FormaCzasu(pasek1.Value)
End Sub

Private Sub pasek2_Change()
CzasRuchu.Caption = "Czas ruchu : " & graslow.FormaCzasu(pasek2.Value)
End Sub

Private Sub pasek2_Scroll()
CzasRuchu.Caption = "Czas ruchu : " & graslow.FormaCzasu(pasek2.Value)
End Sub

Private Sub RDef_Click()
pasek2.Value = DRMax
CzasRuchu.Caption = "Czas ruchu : " & graslow.FormaCzasu(DRMax)

End Sub
