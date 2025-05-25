VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form graslow 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "£ów S³ów"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11880
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "graslow.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "graslow.frx":08CA
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   Begin VB.PictureBox Postep 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2760
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   29
      Top             =   6480
      Width           =   3975
      Begin MSComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   1080
         TabIndex        =   30
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Myœlê..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList Podstawki 
      Left            =   3960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   33
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":1C5C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":1D296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":1DF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":1EC3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":1F912
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":205E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":212BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dodajplik 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "*.txt"
      Filter          =   "*.txt"
      Orientation     =   2
   End
   Begin VB.ComboBox DefMyd1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10800
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox Sloweczka 
      Height          =   3570
      Left            =   10200
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList Klocki 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":21F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":22256
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":22586
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":228AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Obrazki 
      Left            =   4200
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":22C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":22FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":234F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":2394E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":23C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":23FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":242BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":24652
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "graslow.frx":249FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox T³o 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DragIcon        =   "graslow.frx":24E0E
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   240
      ScaleHeight     =   246
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   230
      TabIndex        =   12
      Top             =   1320
      Width           =   3450
   End
   Begin VB.PictureBox Podstawka 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DragIcon        =   "graslow.frx":25118
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
      Begin VB.PictureBox Zaslonka 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   495
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4800
      Top             =   600
   End
   Begin VB.Timer Zegar 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   3720
   End
   Begin VB.CommandButton Kolejka 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Picture         =   "graslow.frx":259E2
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dysk 
      Left            =   5160
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   255
      DefaultExt      =   "gsl"
      FileName        =   "*.gsl"
      Filter          =   "*.gsl"
      Flags           =   2
      Orientation     =   2
   End
   Begin VB.CommandButton wymiana 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      MaskColor       =   &H0080FFFF&
      Picture         =   "graslow.frx":275FC
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton anulujemy 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Picture         =   "graslow.frx":2A212
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Solve 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      MaskColor       =   &H00C0FFFF&
      Picture         =   "graslow.frx":2C0C0
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7920
      Picture         =   "graslow.frx":2D6AA
      Stretch         =   -1  'True
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   10830
      TabIndex        =   27
      Top             =   4080
      Width           =   675
   End
   Begin VB.Label IleLitCap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7320
      TabIndex        =   26
      Top             =   7920
      Width           =   330
   End
   Begin VB.Label EfektyTu 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   7920
      TabIndex        =   25
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Image Kolorek 
      Height          =   255
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Kolorek 
      Height          =   255
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Kolorek 
      Height          =   255
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Kolorek 
      Height          =   255
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label player 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   3
      Left            =   7920
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblRekord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label CR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10800
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label cas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   3
      Left            =   11100
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   2
      Left            =   11100
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Index           =   1
      Left            =   11100
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   0
      Left            =   11100
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label aktGracz 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label WynikGracza 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Wynik     Czas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   10320
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label player 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   2
      Left            =   7800
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label player 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label NazwaGracza 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gracz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   0
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label player 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   0
      Left            =   7800
      MouseIcon       =   "graslow.frx":536DC
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label wynik 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   0
      Left            =   10320
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label wynik 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   1
      Left            =   10320
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label wynik 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   2
      Left            =   10320
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label wynik 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   3
      Left            =   10320
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu MenuGame 
      Caption         =   "Gra"
      Begin VB.Menu MenuNew 
         Caption         =   "&Nowa"
         Shortcut        =   ^N
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuOpen 
         Caption         =   "&Otwórz"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuSave 
         Caption         =   "&Zapisz"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEnd 
         Caption         =   "Koniec"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MenuUstaw 
      Caption         =   "Ustawienia"
      WindowList      =   -1  'True
      Begin VB.Menu MenuFont 
         Caption         =   "Czcionka"
         Shortcut        =   ^F
      End
      Begin VB.Menu MenuMydlo 
         Caption         =   "&Mydlo"
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuShowTime 
         Caption         =   "Pokazuj Czas"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu MenuShowLit 
         Caption         =   "Zas³aniaj &Literki"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu MenuHist 
         Caption         =   "Pokazuj &Historiê"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
      Begin VB.Menu MenuSound 
         Caption         =   "&DŸwiêk"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MenuAkcje 
      Caption         =   "Akcje"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu AkcjeSub 
         Caption         =   "OK"
         Index           =   1
      End
      Begin VB.Menu AkcjeSub 
         Caption         =   "Anuluj"
         Index           =   2
      End
      Begin VB.Menu AkcjeSub 
         Caption         =   "Wymieñ Literki"
         Index           =   3
      End
      Begin VB.Menu AkcjeSub 
         Caption         =   "Omiñ Kolejkê"
         Index           =   4
      End
   End
   Begin VB.Menu MenuSlowa1 
      Caption         =   "S³owa"
      Begin VB.Menu MenuSlowa 
         Caption         =   "Dodaj"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuSlowa 
         Caption         =   "Usuñ"
         Index           =   2
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenuSlowa 
         Caption         =   "ZnajdŸ"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuSlowa 
         Caption         =   "Dodaj z pliku"
         Index           =   4
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuSlowa 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MenuSlowa 
         Caption         =   "Szukaj"
         Index           =   6
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuHelp2 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu MenuBazy 
      Caption         =   "Bazy"
      Visible         =   0   'False
      Begin VB.Menu MenuBazy1 
         Caption         =   "Stwórz"
      End
   End
End
Attribute VB_Name = "graslow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlanszaTAG(0 To 224) As Long, PlanszaRAZYS(0 To 224) As Long, PlanszaRAZYL(0 To 224) As Long, PlanszaCAPTION(0 To 224) As String
Dim PlanszaLICZ(0 To 224) As Long, VirtPlytka(7) As String
Dim Kaption As String
Private Type tDlugosci
   MinL As Long
   MaxL As Long
End Type
Const Horror As Single = 0.2
Dim IlePolLezy As Long, MaxLiter As Long
Public Kolek As Collection
Dim Kierunek As wKierunek, Dlug As Long, KPolaCPP(0 To 6) As tPole
Dim OldX As Single, OldY As Single, VirtLezy(0 To 6) As Long
Dim NumerW As Long, Czyt As Boolean, Razem As Long
Dim CzasKompa As Single, KPolaTAB(7) As tPole, KLiteryTAB(7) As Long
Dim SortLiterki() As Long
Dim CzytP As Boolean, OldXP As Single, OldYP As Single
Dim Omin As Long, Rekord As String, ImieRek As String
Dim MaxPkt As Long, Zakaz(0 To 449) As Long
Dim Konc(50) As String, nwyrazu As Long
Public Sub Dodaj()
Dim FSO As FileSystemObject, fil As TextStream, Plik As String
Dim a As String, i As Long, L As Long, tbl As String, ZAPYT As String
Dim ileDod As Long, IleBylo As Long
On Error GoTo AnulujDodaj
Set FSO = New FileSystemObject
dodajplik.Filter = "*.txt"
dodajplik.ShowOpen
Plik = dodajplik.FileName
Set fil = FSO.OpenTextFile(Plik, ForReading)
While Not fil.AtEndOfStream
    a = UCase(fil.ReadLine)
    a = Replace(a, Chr$(34), "")
   If a <> "" Then
      If MojaBaza.FindB2(a) = 0 Then
         If DodajSlownik(a) = True Then
            ileDod = ileDod + 1
         End If
         IleBylo = IleBylo + 1
         anulujemy.Caption = a
         DoEvents
      End If
   End If
Wend
fil.Close
MojaBaza.AktualizujAll
anulujemy.Caption = ""
Set fil = Nothing
Set FSO = Nothing
MsgBox "Dodawanie s³ów zakoñczone !" & vbCrLf & "Dodano " & ileDod & " z " & IleBylo & " s³ów.", vbInformation
anulujemy.Caption = ""
AnulujDodaj:
End Sub
Public Sub UpdateStojak(strStojak As String)

Dim i As Long
Dim bBlank As Boolean

For i = 0 To 6
    If Mid$(strStojak, i * 2 + 2, 1) = "1" Then
        bBlank = True
    Else
        bBlank = False
    End If
    plytka(i).Po³ó¿ Mid$(strStojak, i * 2 + 1, 1), bBlank
Next i

For i = 0 To 6
    If plytka(i).Caption = "_" Then
        
    Else
        bBlank = False
    End If
    plytka(i).Po³ó¿ Mid$(strStojak, i * 2 + 1, 1), bBlank
Next i



End Sub

Public Function GetStojak() As String

Dim strValue As String
Dim i As Long

For i = 0 To 6
    If plytka(i).Blank Then
        strValue = strValue & plytka(i).Caption & "1"
    Else
        strValue = strValue & plytka(i).Caption & "0"
    End If
Next i

GetStojak = strValue

End Function

Private Sub AkcjeSub_Click(Index As Integer)
Select Case Index
    Case 1: Obliczamy
    Case 2: Anuluj
    Case 3: ZmianaLiter
    Case 4: OminKolejke
End Select
End Sub

Private Sub anulujemy_Click()
Call Anuluj
End Sub
Private Sub DefMyd1_Click()
DefiniujMydlo 1, DefMyd1.List(DefMyd1.ListIndex)
End Sub

Private Sub Form_Activate()
Dim i As Long
For i = 0 To 3
   player(i).Font.Size = 12
   wynik(i).Font.Size = 12
   cas(i).Font.Size = 12
Next i

End Sub

Private Sub Form_DblClick()
Dim znaki As String, mczas As Single
If Not Czasy Is Nothing Then
   mczas = Czasy.Count
Else
   mczas = 0
End If
znaki = "Ca³kowity czas: " & FormaCzasu(CLng(SumaCzas)) & vbCrLf & "œredni czas ruchu komputera: " & FormaCzasu(CLng(SredniCzas)) & vbCrLf & "Maksymalny czas ruchu komputera: " & FormaCzasu(CLng(MaxCzas)) & vbCrLf & "Iloœæ ruchów komputera: " & mczas
MsgBox znaki, vbInformation

End Sub

Private Sub Form_Load()
Dim i As Long
Sloweczka.Visible = False
IlePolLezy = 0
For i = 1 To 4
   Kolorek(i - 1).Left = 520
   player(i - 1).Left = Kolorek(i - 1).Left + Kolorek(i - 1).Width + 2
   player(i - 1).Top = 20 + ((player(i - 1).Height + 2) * (i))
   Kolorek(i - 1).Top = player(i - 1).Top
   cas(i - 1).Top = player(i - 1).Top
   wynik(i - 1).Top = player(i - 1).Top
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call KoniecLS
End Sub

Private Sub Form_Resize()
On Error Resume Next
Postep.Top = graslow.Top + graslow.Height - Postep.Height - 20
Postep.Left = graslow.Left + graslow.Width - Postep.Width - 20
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call KoniecLS
End Sub

Private Sub IleLitCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
IleLitCap.ToolTipText = "Pozosta³o jeszcze " & IleLitCap.Caption & " liter."
End Sub

Private Sub Kolejka_Click()
OminKolejke
End Sub

Private Sub MenuBazy1_Click()
Dim Nazwa As String

Nazwa = InputBox("Podaj nazwê bazy")
If Nazwa <> "" Then
   CreateBaza App.Path & "\bazy\" & Nazwa
End If
End Sub

Private Sub MenuEnd_Click()
Call konczymy
End Sub

Private Sub MenuFont_Click()
Call Fontuj
End Sub

Private Sub MenuHelp2_Click()
Dim Tekst As String, Pliki As FileSystemObject, Plik As TextStream
Set Pliki = New FileSystemObject
Set Plik = Pliki.OpenTextFile(App.Path & "\help.txt")
Tekst = Plik.ReadAll
Plik.Close
Load Dialog
Dialog.Text1.Text = LTrim(Tekst)
Dialog.Show vbModal
Set Plik = Nothing
Set Pliki = Nothing
Unload Dialog
End Sub

Private Sub MenuHist_Click()
If MenuHist.Checked Then
   MenuHist.Checked = False
Else
   MenuHist.Checked = True
End If
End Sub
Private Sub MenuMydlo_Click()
Call Mydelko
End Sub

Private Sub MenuNew_Click()
Call NowaGra
End Sub

Private Sub MenuOpen_Click()
Call Importuj
End Sub

Private Sub MenuSave_Click()
Call Zapisz
End Sub
Public Sub Punktuj()
Dim i As Long
For i = 0 To 6
    Call plytka(i).Po³ó¿(Left(Stojak(Ktory, i), 1), CBool(Right(Stojak(Ktory, i), 1)))
    Call plytka(i).ZatwierdŸ(3)
Next i

End Sub
Public Function Losuj() As String

Dim mY As Long, LS As String, Ziarno As Double
Ziarno = Timer + (Timer * 8) ^ 2 + (7 * GetTickCount())
10 Randomize Ziarno
'Sleep 20
Ziarno = Int(Rnd(Ziarno) * 45135) + 1
Randomize Ziarno
mY = Int((Rnd(Ziarno) * 100))
If Wolne(mY) = False Then
    If mY > 97 Then
        LS = "1"
    Else
        LS = "0"
    End If
    Losuj = Worek(mY) & LS
    Wolne(mY) = True
    IleLiter = IleLiter - 1
    'ileL.Caption = IleLiter
    IleLitCap.Caption = IleLiter
    DoEvents
    Exit Function
Else
    GoTo 10
End If

End Function

Private Sub MenuShowLit_Click()
If MenuShowLit.Checked Then
    MenuShowLit.Checked = False
Else
    MenuShowLit.Checked = True
End If

End Sub

Private Sub MenuShowTime_Click()
If MenuShowTime.Checked Then
    MenuShowTime.Checked = False
Else
    MenuShowTime.Checked = True
End If
End Sub

Private Sub MenuSlowa_Click(Index As Integer)
Dim t As Single
Select Case Index
    Case 1:
        PlusMinus.OKButton.Enabled = True
        PlusMinus.CancelButton.Enabled = False
        PlusMinus.Show vbModal
    Case 2
        PlusMinus.OKButton.Enabled = False
        PlusMinus.CancelButton.Enabled = True
        PlusMinus.Show vbModal
    Case 3
        Szukaj.Show vbModal, graslow
   Case 4
      Dodaj
   Case 6
   
      VirtUklada
      KLiterujTAB
      t = Timer
      SzukajWszystkie2 False
      Sloweczka.AddItem "Czas szukania: " & CStr(Timer - t)
      Sloweczka.Clear

   End Select

End Sub

Private Sub MenuSound_Click()
If MenuSound.Checked Then
   MenuSound.Checked = False
Else
   MenuSound.Checked = True
End If
End Sub

Private Sub player_Click(Index As Integer)
Historia CLng(Index + 1)
End Sub

Private Sub Podstawka_DblClick()

If JuzGramy = False Then Exit Sub
    
        Anuluj

End Sub

Private Sub Podstawka_DragDrop(Source As Control, X As Single, Y As Single)

Dim xxp As Single, yyP As Single, OldXXP As Single, OldYYP As Single, OldIndex As Integer, Indeks As Integer
Dim tmpCP As String, StCo As Long, OldXX As Long, OldYY As Long, i As Long
Dim tmpp As String, OldM As Boolean
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
    If PLR(Ktory).Komp Then Exit Sub
End If
xxp = Int((X) / (bok + Prz)) + 1
yyP = 1
If xxp > 7 Or xxp < 1 Then Exit Sub

Indeks = ciagiem(xxp, yyP)

If Source Is Podstawka Then
    OldXXP = Int((OldXP) / (bok + Prz)) + 1
    OldYYP = Int((OldYP) / (bok + Prz)) + 1
    If OldXXP > 7 Or OldYYP > 7 Or OldXXP < 1 Or OldYYP < 1 Then
        CzytP = False
        Exit Sub
    End If
    OldIndex = ciagiem(OldXXP, OldYYP)
    If plytka(OldIndex).Caption = "_" Then
        GoTo koniecDD
    End If
    
    tmpCP = plytka(Indeks).Litera
    StCo = plytka(Indeks).NumerObrazka
    OldM = plytka(Indeks).Blank
    Call plytka(Indeks).Po³ó¿("_", False)
    Call plytka(Indeks).ZatwierdŸ(3)
    Call plytka(Indeks).Po³ó¿(plytka(OldIndex).Litera, plytka(OldIndex).Blank)
    Call plytka(Indeks).ZatwierdŸ(plytka(OldIndex).NumerObrazka)
    Call plytka(OldIndex).Po³ó¿("_", False)
    Call plytka(OldIndex).ZatwierdŸ(3)
    Call plytka(OldIndex).Po³ó¿(tmpCP, OldM)
    Call plytka(OldIndex).ZatwierdŸ(StCo)
    EfektujTu
End If

If Source Is T³o Then
    OldXX = Int((OldX - px) / (bok + Prz)) + 1
    OldYY = Int((OldY - py) / (bok + Prz)) + 1
    If OldXX > 15 Or OldYY > 15 Or OldXX < 1 Or OldYY < 1 Then
        Czyt = False
        Exit Sub
    End If
    OldIndex = ciagiem(OldXX, OldYY)
    
    If plytka(Indeks).Caption = "_" Then
        If Pole(OldIndex).NumerObrazka <> 1 Then Exit Sub
        Dlug = Dlug - 1
        If Dlug = 1 Then Kierunek = wSama
        Call plytka(Indeks).Po³ó¿("_", False)
        Call plytka(Indeks).ZatwierdŸ(3)
        Call plytka(Indeks).Po³ó¿(Pole(OldIndex).Caption, Pole(OldIndex).Blank)
        Le¿¹ce.Remove "K" & CStr(OldIndex)
        Call plytka(Indeks).ZatwierdŸ(3)
        Call Pole(OldIndex).PustePole
        Call EfektujTu
        GoTo koniecDD
    Else
        If Pole(OldIndex).NumerObrazka <> 1 Then Exit Sub
        tmpp = plytka(Indeks).Caption
        OldM = plytka(Indeks).Blank
        Call plytka(Indeks).Po³ó¿("_", OldM)
        Call plytka(Indeks).ZatwierdŸ(3)
        plytka(Indeks).Po³ó¿ Pole(OldIndex).Caption, Pole(OldIndex).Blank
        Call plytka(Indeks).ZatwierdŸ(3)
        Call Pole(OldIndex).PustePole
        Call Pole(OldIndex).Po³ó¿(tmpp, OldM)
        Call EfektujTu
    End If
End If
koniecDD:
CzytP = False
Czyt = False

End Sub

Private Sub Podstawka_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Dim xxp As Long, kapp As String
If Not Source Is Podstawka Then Exit Sub
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
    If PLR(Ktory).Komp Then Exit Sub
End If

If State = 0 And X > 20 And Y > 20 And X < Podstawka.Width - 20 And Y < Podstawka.Height - 20 Then CzytP = False
xxp = X \ (bok + Prz)
If xxp > 6 Then Exit Sub
If State = 0 Then
   If plytka(xxp).Caption = "_" Then
      Podstawka.DragIcon = LoadPicture(App.Path & "\nodrop01.cur")
      Exit Sub
   End If
   If plytka(xxp).Blank Then
      If plytka(xxp).Litera = "" Then
         kapp = "MYDLO"
      Else
         kapp = "blank_" & Asc(plytka(xxp).Litera)
      End If
   Else
      kapp = "Ikona_" & Asc(plytka(xxp).Caption)
   End If
    If Not Jêzyk Is Nothing Then
        Podstawka.DragIcon = LoadPicture(App.Path & "\ikony\" & Jêzyk.Klucz & "\" & kapp & ".ico")
    End If
    plytka(xxp).Zaslon
End If

If CzytP = False Then
    OldXP = X
    OldYP = Y
    CzytP = True
End If

End Sub

Private Sub Podstawka_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim xxp As Single, yyP As Single, Index As Integer
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
End If

If Button = 2 Then
    xxp = Int((X) / (bok + Prz)) + 1
    yyP = Int((Y) / (bok + Prz)) + 1
    If xxp > 7 Or yyP > 7 Or xxp < 1 Or yyP < 1 Then Exit Sub
    Index = ciagiem(xxp, yyP)
    If plytka(Index).Caption <> "_" Then
        If plytka(Index).NumerObrazka = 3 Then
            Call plytka(Index).ZatwierdŸ(4)
        Else
            Call plytka(Index).ZatwierdŸ(3)
        End If
    End If
End If

If PLR(Ktory).IloscWymian >= IleWymian Then
   wymiana.Enabled = False
   AkcjeSub(3).Enabled = False
Else
   wymiana.Enabled = CzyDoWymiany
   AkcjeSub(3).Enabled = wymiana.Enabled
End If

End Sub

Private Sub Podstawka_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xxp As Single, yyP As Single, Index As Integer
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
End If
    
    xxp = Int((X) / (bok + Prz)) + 1
    yyP = Int((Y) / (bok + Prz)) + 1
    If xxp > 7 Or yyP > 7 Or xxp < 1 Or yyP < 1 Then Exit Sub
    Index = ciagiem(xxp, yyP)
    If plytka(Index).Caption <> "_" Then
        Anuluj
    End If
End Sub

Public Function svrFindWords() As String

Dim strWords As String
Dim i As Long
Dim nIndex As Long

If Le¿¹ce.Count = 0 Then Exit Function

If Le¿¹ce.Count = 1 Then
    nIndex = CLng(Le¿¹ce(1))
    strWords = Pole(nIndex).Caption
    Exit Function
End If
 



For i = 1 To 15
    strWords = svrFindLineWords(i, wPionowy)
    If strWords <> "" Then
        svrFindWords = strWords
        Exit Function
    End If
Next i

For i = 1 To 15
    strWords = svrFindLineWords(i, wPionowy)
    If strWords <> "" Then
        svrFindWords = strWords
        Exit Function
    End If
Next i

svrFindWords = strWords

End Function

Public Sub svrSetPlayerPresent(nNumer As Long, bPresent As Boolean)

If bPresent = True Then

    Set PLR(nNumer + 1) = New GR
    PLR(nNumer + 1).ClearWyrazy
    PLR(nNumer + 1).imie = "Klient " & CStr(nNumer)
    player(nNumer).Visible = True
    wynik(nNumer).Visible = True
    player(nNumer).Caption = PLR(nNumer + 1).imie
    wynik(nNumer).Caption = "0"

Else

    Set PLR(nNumer) = Nothing

End If

End Sub

Public Function svrFindLineWords(nNumer As Long, wKier As wKierunek) As String

Dim strWords As String
Dim nIndex As Long
Dim i As Long
Dim bFound As Boolean

If wKier = wPionowy Then
    For i = 1 To 15
        nIndex = ciagiem(nNumer, i)
        If Pole(nIndex).Tag = True Then
            
        End If
    
    Next i
ElseIf wKier = wPoziomy Then


End If

svrFindLineWords = strWords

End Function
Private Sub Solve_Click()

If Not m_Klient Is Nothing Then
    
'    m_Klient.SendTurn "Kladzie", "WYRAZIK", wPoziomy, 108, "W0Y0R0A0Z0I0K0"

    Exit Sub
End If

Call Obliczamy

End Sub

Public Sub Zatwierdz()
Dim co³nt As Integer, i As Long, item As Variant
'If m_Server Is Nothing Then
    PLR(Ktory).DodajWyrazy Kolek, Razem
'End If

If MenuSound.Checked = True Then
    If Le¿¹ce.Count = 7 Then
        gramy "flinston.wav"
    End If
End If
For Each item In Le¿¹ce
    Pole(CLng(item)).licz = False
    Pole(CLng(item)).ZatwierdŸ 2
    DoEvents
    Kolkon(CLng(item)) = Ktory - 1
Next item
co³nt = Le¿¹ce.Count
For i = co³nt To 1 Step -1
    Le¿¹ce.Remove i
Next i
Set Kolek = Nothing
'If m_Server Is Nothing Then
    If PLR(Ktory).Komp = False Then
        If MenuHist.Checked Then Historia (Ktory)
    End If
'End If

NumerW = NumerW + 1
Omin = 0

End Sub

Public Sub Anuluj()
Dim a As Boolean, co³nt As Integer, i As Long, j As Long
Dim item As Variant
For Each item In Le¿¹ce
    For j = 0 To 6
        If plytka(j).Caption = "_" Then
            Call plytka(j).Po³ó¿(Pole(CLng(item)).Caption, Pole(CLng(item)).Blank)
            Call plytka(j).ZatwierdŸ(3)
            Exit For
        End If
    Next j
    Call Pole(CLng(item)).PustePole
Next item

For i = 0 To 224
    If Pole(i).Tag = True Then
        a = True
        Exit For
    End If
Next i

If a = False Then NumerW = 0
co³nt = Le¿¹ce.Count
For i = 1 To co³nt
    Le¿¹ce.Remove (1)
Next i
Dlug = 0
Kierunek = wBrak
EfektujTu
End Sub


Private Sub Timer1_Timer()

If IleLitCap.BackColor = vbGreen Then
    
    IleLitCap.BackColor = vbRed
Else
    
    IleLitCap.BackColor = vbGreen
End If

End Sub

Private Sub T³o_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

Dim xx As Long, yy As Long, Indeks As Long, Plik As String
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
    If PLR(Ktory).Komp Then Exit Sub
End If
If Not Source Is T³o Then Exit Sub
If State = 0 And X > 5 And Y > 5 And X < T³o.Width - 5 And Y < T³o.Height - 5 Then Czyt = False
If Czyt = False Then
    OldX = X
    OldY = Y
    Czyt = True
    xx = Int((X - px) / (bok + Prz)) + 1
    yy = Int((Y - py) / (bok + Prz)) + 1
    If xx > 15 Or yy > 15 Or xx < 1 Or yy < 1 Then Exit Sub
    Indeks = ciagiem(xx, yy)
    If Pole(Indeks).licz = False Then
        T³o.DragIcon = LoadPicture(App.Path & "\nodrop01.cur")
    Else
    
      If Pole(Indeks).Blank Then
         If Pole(Indeks).Litera = "" Then
            Plik = "MYDLO"
         Else
            Plik = "blank_" & Asc(Pole(Indeks).Litera)
         End If
      Else
         Plik = "Ikona_" & Asc(Pole(Indeks).Caption)
      End If
        If Not Jêzyk Is Nothing Then
            T³o.DragIcon = LoadPicture(App.Path & "\ikony\" & Jêzyk.Klucz & "\" & Plik & ".ico")
        End If
    End If
End If

End Sub

Private Sub T³o_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Long, yy As Long, Indeks As Long
If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
End If
If Button = 2 Then
    graslow.PopupMenu MenuAkcje, , X, Y
End If
End Sub

Private Sub T³o_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Long, yy As Long, Indeks As Long
If JuzGramy = False Then Exit Sub
xx = Int((X - px) / (bok + Prz)) + 1
yy = Int((Y - py) / (bok + Prz)) + 1
If xx > 15 Or yy > 15 Or xx < 1 Or yy < 1 Then Exit Sub
Indeks = ciagiem(xx, yy)

If Pole(Indeks).Tag Then
    T³o.ToolTipText = ""
    Exit Sub
End If
    
Select Case Pole(Indeks).RazyL
Case 2: T³o.ToolTipText = "Podwójna premia literowa ": Exit Sub
Case 3: T³o.ToolTipText = "Potrójna premia literowa ": Exit Sub
Case Else
    Select Case Pole(Indeks).RazyS
    Case 2: T³o.ToolTipText = "Podwójna premia s³owna ": Exit Sub
    Case 3: T³o.ToolTipText = "Potrójna premia s³owna ": Exit Sub
    Case Else:  T³o.ToolTipText = ""
    End Select
End Select

End Sub

Private Sub wymiana_Click()
Call ZmianaLiter
End Sub

Public Function Nastepny(Kt As Long) As Long

Dim i As Long

If m_Server Is Nothing Then
    If Kt = IleGraczy Then
        Nastepny = 1
    Else
        Nastepny = Kt + 1
    End If
Else
    For i = Kt + 1 To 4
        If Not PLR(i) Is Nothing Then
            Nastepny = i
            Exit Function
        End If
    Next i

    For i = 0 To Kt - 1
        If Not PLR(i) Is Nothing Then
            Nastepny = i + 1
            Exit Function
        End If
    Next i
End If

End Function

Public Function Poprzedni(Kt As Long) As Long
If Kt = 1 Then
    Poprzedni = IleGraczy
Else
    Poprzedni = Kt - 1
End If
End Function
Private Sub Zapisz()
Dim Plik As String, i As Long, lanc As String, j As Long, t As Long
Dim item As Ruchy, IlePol As Long, hh As Long, hj As Long
hh = FreeFile
Open App.Path & "\save\tmp.tmp" For Binary As hh
Close hh
Kill (App.Path & "\save\tmp.tmp")

For i = 1 To IleGraczy
    lanc = lanc & PLR(i).imie & "_"
Next i
lanc = lanc & Date

dysk.Filter = "Pliki '£ów S³ów' (*.lsw)|*.lsw"
dysk.FilterIndex = 1
dysk.FileName = lanc
dysk.InitDir = App.Path & "\save"
hj = FreeFile
On Error GoTo KoniecZapisz
dysk.ShowSave
On Error GoTo 0
Plik = dysk.FileName

Open App.Path & "\save\tmp.tmp" For Output As hj

IlePol = 0

Write #hj, "LW2"
For i = 0 To 224
    If Pole(i).Tag Then
        IlePol = IlePol + 1
    End If
Next i
Write #hj, IlePol, IleGraczy, IleLiter, Ktory - 1, Omin, FMax, RMax, Zero, NumerW, CInt(KoniecGry), Jêzyk.Klucz, Jêzyk.Nazwa, Jêzyk.Plik, Jêzyk.S³ownik, CInt(CzasRTak), CInt(CzasCTak), IleWymian, Demo

For i = 1 To IleGraczy
    Write #hj, PLR(i).IleRuchów, PLR(i).iw
    For Each item In PLR(i).MyKolek
        Write #hj, item.Punkty, item.S³owa
    Next item
Next i

For i = 0 To 224
    If Pole(i).Tag Then
        Write #hj, i, Pole(i).Caption & CStr(Abs(CLng(Pole(i).Blank))), Kolkon(i)
    End If
Next i
For j = 1 To IleGraczy
    Write #hj, PLR(j).imie, PLR(j).wynik, PLR(j).CzasCa³kowity, CInt(PLR(j).Komp), PLR(j).MaxPsz, PLR(j).IloscWymian
    For t = 0 To 6
        If Right(Stojak(j, t), 1) = "1" Then Stojak(j, t) = Chr$(32) & "1"
        Write #hj, Stojak(j, t);
    Next t
    Write #hj, "X"
Next j
For i = 0 To 99
    If Wolne(i) Then
        Write #hj, i;
    End If
Next i
Close hj
Call Szyfruj(Plik)
Kill (App.Path & "\save\tmp.tmp")
Exit Sub
KoniecZapisz:

Close hj
End Sub

Public Sub Mydelko()

Mydla.asciiM.Text = CStr(Asc(Mydlo))
Mydla.znakM.Text = CStr(Mydlo)
Mydla.Show

End Sub

Public Sub konczymy()
Dim a As Long, i As Long
a = InfoBox("Na pewno koñczymy ?", True)
If a = 2 Then Exit Sub
JuzGramy = False
For i = 0 To 224
   Set Pole(i) = Nothing
Next i
For i = 0 To 6
   Set plytka(i) = Nothing
Next i
For i = 0 To 4
   Set PLR(i) = Nothing
Next i
Me.MousePointer = 11
KoniecLS
End Sub

Public Sub Fontuj()

Dim ff As String, bk As ColorConstants, i As Long
On Error GoTo NieWybranoFontu
dysk.ShowFont
If dysk.FontName = "" Or dysk.FontName = " " Then Exit Sub
ff = dysk.FontName
T³o.Font = ff
Podstawka.Font = ff
Podstawka.Font.Charset = 238
T³o.Font.Charset = 238
For i = 0 To 224
    Pole(i).Font = ff
    If Pole(i).Tag Then
        bk = Pole(i).NumerObrazka
        Call Pole(i).ZatwierdŸ(bk)
    End If
Next i

For i = 0 To 6
    bk = plytka(i).NumerObrazka
    plytka(i).Font = ff
    Call plytka(i).ZatwierdŸ(bk)
Next i
NieWybranoFontu:

End Sub
Public Sub gramy(Plik As String)
GrajWave App.Path & "\" & Plik, &H1
End Sub
Public Function SlownikCheck(Wyraz As String) As Boolean
If MojaBaza.FindB2(Wyraz) > 0 Then
   SlownikCheck = False
Else
   SlownikCheck = True
End If
End Function

Public Function DodajSlownik(Wyraz As String, Optional Nie23 As Boolean = True) As Boolean

If Len(Wyraz) > 15 Or Len(Wyraz) < 2 Then Exit Function
If InStr(1, Wyraz, Mydlo) > 0 Then Exit Function
If Nie23 = True Then
   If Len(Wyraz) = 2 Or Len(Wyraz) = 3 Then Exit Function
End If
Me.MousePointer = 11
If MojaBaza.FindB2(Wyraz) = 0 Then
   MojaBaza.Add Wyraz
   DodajSlownik = True
Else
   DodajSlownik = False
End If
Me.MousePointer = 0
End Function
Public Sub enduj()
Dim j As Long, i As Long, Y As Long, AllMinus As Long, Minus(4) As Long
Dim HiScore As Long, HiName As String, Suma As Long, Reszty(4) As String
Dim Tek As String
On Error GoTo 0
Timer1.Enabled = False

    For j = 1 To IleGraczy
        For i = 0 To 6
            If Not Left(Stojak(j, i), 1) = "_" Then
                If Right(Stojak(j, i), 1) = 1 Then
                    Minus(j) = Minus(j) + 50
                Else
                    Minus(j) = Minus(j) + Wart(Asc(Left(Stojak(j, i), 1)))
                End If
                Reszty(j) = Reszty(j) & "  " & BezMydla(Left(Stojak(j, i), 1))
            End If
        Next i
        AllMinus = AllMinus + Minus(j)
    Next j
    
    If KoniecGry = False Then
        For j = 1 To IleGraczy
            If Minus(j) = 0 Then
                PLR(j).wynik = PLR(j).wynik + AllMinus
                wynik(j - 1).Caption = PLR(j).wynik
            Else
                PLR(j).wynik = PLR(j).wynik - Minus(j)
                wynik(j - 1).Caption = PLR(j).wynik
            End If
        Next j
    End If
    KoniecGry = True
For i = 0 To 224
    If Pole(i).Tag Then Pole(i).ZatwierdŸ Kolkon(i) + 4
Next i
wymiana.Enabled = False
Solve.Enabled = False
anulujemy.Enabled = False
Kolejka.Enabled = False
Zegar.Enabled = False
For i = 1 To 4
   AkcjeSub(i).Enabled = False
Next i

HiScore = 0

For i = 1 To IleGraczy
    If PLR(i).wynik > HiScore Then
        HiScore = PLR(i).wynik
        HiName = PLR(i).imie
    End If
Next i

If IsNumeric(Rekord) = False Then
    Rekord = 0
End If

If HiScore > CLng(Rekord) Then
   SaveSetting "£ów S³ów", "Rekordy", "Rekord_" & Jêzyk.Klucz & "_" & IleGraczy, CStr(HiScore)
   SaveSetting "£ów S³ów", "Rekordy", "Imie_" & Jêzyk.Klucz & "_" & IleGraczy, HiName
End If

Call Sortuj

For j = 1 To IleGraczy
    Koniec.plajer(j - 1).Caption = PLR(max(j)).imie
    Koniec.wynikk(j - 1).Caption = PLR(max(j)).wynik
    Koniec.Kara(j - 1).Caption = Minus(max(j))
    Koniec.Reszta(j - 1).Caption = Reszty(max(j))
    Suma = Suma + PLR(max(j)).wynik
Next j

Koniec.Sumapkt.Caption = Suma

For Y = IleGraczy + 1 To 4
    Koniec.plajer(Y - 1).Visible = False
    Koniec.wynikk(Y - 1).Visible = False
    Koniec.Reszta(Y - 1).Visible = False
    Koniec.Kara(Y - 1).Visible = False
Next Y

Zegar.Enabled = False

Koniec.Show vbModal

For i = 0 To 3
    Minus(i) = 0
    Reszty(i) = ""
Next i
Timer1.Enabled = False

End Sub

Private Sub Zegar_Timer()

If JuzGramy = False Then Exit Sub
Dim kl As Collection

    cas(Ktory - 1).Caption = FormaCzasu(FMax - PLR(Ktory).CzasCa³kowity)
    CR.Caption = FormaCzasu(RMax - Zero)
    If PLR(Ktory).Komp Then Exit Sub
   If CzasCTak = True Then
      PLR(Ktory).CzasCa³kowity = PLR(Ktory).CzasCa³kowity + 1
   End If
   If CzasRTak = True Then
      Zero = Zero + 1
   End If
   If (RMax - Zero) < 10 Then
      If MenuSound.Checked = True Then gramy "10sekund.wav"
   End If
    If PLR(Ktory).CzasCa³kowity > FMax Then
       
        Call InfoBox("Przekroczy³eœ ustalony czas. Sorry ... ")
        Omin = Omin + 1
        Set kl = New Collection
        kl.Add "*** koniec czasu gry ***"
        PLR(Ktory).DodajWyrazy kl, 0
        Set kl = Nothing
        Call Anuluj
        Call NextPlayer
    End If
        
    If Zero > RMax Then
            Anuluj
            efekty.CancelButton_Click
            Anuluj
        Call InfoBox("Przekroczy³eœ ustalony limit czasu na wykonanie ruchu. Sorry ... ")
        Set kl = New Collection
        kl.Add "*** przekroczony czas ruchu ***"
        PLR(Ktory).DodajWyrazy kl, 0
        Set kl = Nothing
        Call Anuluj
        Call NextPlayer
    End If
DoEvents
End Sub

Public Function FormaCzasu(sek As Long) As String
Dim Zeroo As String
    If sek - ((sek \ 60) * 60) < 10 Then
        Zeroo = "0"
    Else
        Zeroo = ""
    End If
    FormaCzasu = (sek \ 60) & ":" & Zeroo & sek - (sek \ 60) * 60
End Function


Public Sub NextPlayer()
Dim i As Long, Pustej As Long, Kapon As String, Mdl As Boolean

If Not m_Klient Is Nothing Then Exit Sub

If Not m_Server Is Nothing Then

    Ktory = Nastepny(Ktory)
    For i = 0 To 3
       player(i).FontBold = False
       wynik(i).FontBold = False
       cas(i).FontBold = False
    Next i
    
    player(Ktory - 1).FontBold = True
    wynik(Ktory - 1).FontBold = True
    wynik(Ktory - 1).Caption = PLR(Ktory).wynik
    cas(Ktory - 1).FontBold = True
    aktGracz.Caption = "Uk³ada " & PLR(Ktory).imie
    Exit Sub

End If


Sloweczka.Visible = False
Zegar.Enabled = False
If JuzGramy = False Then Exit Sub
Podstawka.Visible = False
DoEvents
'Demo = Demo + 1
'If IleLiter < 20 Then
'If Demo >= 87 Then
   'WersjaDemo
  ' enduj
 '  Exit Sub
'End If
On Error Resume Next
Kill App.Path & "\save\autoold.lsw"
FileCopy App.Path & "\save\autosave.lsw", App.Path & "\save\autoold.lsw"
Kill App.Path & "\save\autosave.lsw"
AutoZapis

On Error GoTo 0

For i = 0 To 6
    Stojak(Ktory, i) = plytka(i).Caption & CStr(Abs(CLng(plytka(i).Blank)))
Next i

If Omin = 2 * IleGraczy And IleLiter = 0 Then
    Call enduj
    Exit Sub
End If
On Error GoTo 0
For i = 0 To 6
    If IleLiter = 0 Then
        
        Timer1.Enabled = True
        Exit For
    Else
        If Left(Stojak(Ktory, i), 1) = "_" Then
            Stojak(Ktory, i) = Losuj()
        End If
    End If
Next i

For i = 0 To 6
    If Left(Stojak(Ktory, i), 1) = "_" Then
        Pustej = Pustej + 1
    End If
Next i

If Pustej = 7 And IleLiter = 0 Then
    Call enduj
    Exit Sub
End If
Pustej = 0
Ktory = Nastepny(Ktory)
For i = 0 To 3
   player(i).FontBold = False
   wynik(i).FontBold = False
   cas(i).FontBold = False
Next i

player(Ktory - 1).FontBold = True
wynik(Ktory - 1).FontBold = True
cas(Ktory - 1).FontBold = True
aktGracz.Caption = "Uk³ada " & PLR(Ktory).imie
Zero = 0
IleLitCap.Caption = CStr(IleLiter)
For i = 0 To 6
    Call plytka(i).PustyStojak
    Kapon = Left(Stojak(Ktory, i), 1)
    Mdl = CBool(Right(Stojak(Ktory, i), 1))
    Call plytka(i).Po³ó¿(Kapon, Mdl)
    Call plytka(i).ZatwierdŸ(3)
Next i

Call AutoZapis
Call EfektujTu
MojaBaza.ClearWlasne
Call Nazywaj

DefMyd1.Visible = CzyBlank And (Not PLR(Ktory).Komp)
Label1.Visible = DefMyd1.Visible
ButOnOff (PLR(Ktory).Komp)

DoEvents

If PLR(Ktory).Komp = True Then
    If MenuShowLit.Checked = True Then
        Podstawka.Visible = False
        DoEvents
    Else
        Podstawka.Visible = True
        DoEvents
    End If
    KompKladzie
Else
     Podstawka.Visible = True
     DoEvents
End If
Solve.Enabled = False
AkcjeSub(1).Enabled = False
wymiana.Enabled = False
AkcjeSub(3).Enabled = False

Zegar.Enabled = True
End Sub

Public Sub Szyfruj(Plik As String)
Dim a As Byte, i As Long, hh As Long, hj As Long
hj = FreeFile
Open App.Path & "\save\tmp.tmp" For Binary As hj
hh = FreeFile
Open Plik For Binary As hh
While Not EOF(hj)
    Get hj, i + 1, a
    Put hh, i + 1, (a Xor 45) Xor 211
    i = i + 1
Wend
Close hj
Close hh
End Sub

Private Sub DeSzyfruj(Plik As String)
Dim hh As Long, hj As Long
Dim a As Byte, i As Long
hh = FreeFile
Open App.Path & "\save\tmp.tmp" For Binary As hh
hj = FreeFile
Open Plik For Binary As hj
While Not EOF(hj)
    Get hj, i + 1, a
    Put hh, i + 1, (a Xor 45) Xor 211
    i = i + 1
Wend
Close hj
Close hh
End Sub

Public Sub UstawCzasy()
Dim i As Long
For i = 1 To IleGraczy
    cas(i).Caption = FormaCzasu(FMax - PLR(i).CzasCa³kowity)
Next i
End Sub
Private Sub AutoZapis()

If Not m_Klient Is Nothing Then Exit Sub

Dim Plik As String, i As Long, j As Long, t As Long, item As Ruchy
Dim IlePol As Long, hh As Long, hj As Long, Myda As String
Plik = App.Path & "\save\autosave.lsw"
hh = FreeFile
Open App.Path & "\save\tmp.tmp" For Binary As hh
Close hh
Kill (App.Path & "\save\tmp.tmp")
hj = FreeFile
Open App.Path & "\save\tmp.tmp" For Output As hj

IlePol = 0
Write #hj, "LW2"
For i = 0 To 224
    If Pole(i).Tag Then
        IlePol = IlePol + 1
    End If
Next i
Write #hj, IlePol, IleGraczy, IleLiter, Ktory - 1, Omin, FMax, RMax, Zero, NumerW, CInt(KoniecGry), Jêzyk.Klucz, Jêzyk.Nazwa, Jêzyk.Plik, Jêzyk.S³ownik, CInt(CzasRTak), CInt(CzasCTak), IleWymian, Demo

For i = 1 To IleGraczy
    Write #hj, PLR(i).IleRuchów, PLR(i).iw
    For Each item In PLR(i).MyKolek
        Write #hj, item.Punkty, item.S³owa
    Next item
Next i

For i = 0 To 224
    If Pole(i).Tag Then
        Write #hj, i, Pole(i).Caption & CStr(Abs(CLng(Pole(i).Blank))), Kolkon(i)
    End If
Next i
For j = 1 To IleGraczy
    Write #hj, PLR(j).imie, PLR(j).wynik, PLR(j).CzasCa³kowity, CInt(PLR(j).Komp), PLR(j).MaxPsz, PLR(j).IloscWymian
    For t = 0 To 6
        If Right(Stojak(j, t), 1) = "1" Then Stojak(j, t) = Chr$(32) & "1"
        Write #hj, Stojak(j, t);
    Next t
    Write #hj, "X"
Next j
For i = 0 To 99
    If Wolne(i) Then
        Write #hj, i;
    End If
Next i
Close hj
Call Szyfruj(Plik)
Kill (App.Path & "\save\tmp.tmp")
End Sub
Public Function CzyPionowo(Indeks As Long, Kolekcja As Collection, Optional OldIndex As Long = -1) As Boolean
Dim item As Variant
If OldIndex = -1 Then
    For Each item In Kolekcja
        If wx(Indeks) <> wx(item) Then
            CzyPionowo = False
            Exit Function
        End If
    Next item
    CzyPionowo = True
Else
    For Each item In Kolekcja
        If (wx(Indeks) <> wx(item)) And (item <> OldIndex) Then
            CzyPionowo = False
            Exit Function
        End If
    Next item
    CzyPionowo = True
End If
End Function
Public Function CzyPoziomo(Indeks As Long, Kolekcja As Collection, Optional OldIndex As Long = -1) As Boolean
Dim item As Variant
If OldIndex = -1 Then
    For Each item In Kolekcja
        If wy(Indeks) <> wy(item) Then
            CzyPoziomo = False
            Exit Function
        End If
    Next item
    CzyPoziomo = True
Else
    For Each item In Kolekcja
        If (wy(Indeks) <> wy(item)) And (item <> OldIndex) Then
            CzyPoziomo = False
            Exit Function
        End If
    Next item
    CzyPoziomo = True
End If
End Function
Public Function BezMydla(Wyraz As String) As String
BezMydla = Replace(Wyraz, Mydlo, "_")
End Function
Public Sub Sortuj()
Dim Byl(4) As Boolean, i As Long, j As Long, Maxim As Long
Dim ChMax As Long
Maxim = -500
For i = 1 To 4
    max(i) = -500
Next i
For j = 1 To IleGraczy
    For i = 1 To IleGraczy
        If PLR(i).wynik >= Maxim And Byl(i) = False Then
            Maxim = PLR(i).wynik
            ChMax = i
        End If
    Next i
    max(j) = ChMax
    Byl(ChMax) = True
    Maxim = -500
Next j
End Sub

Public Sub Historia(NumerGracza As Long)
Dim item As Variant, Klr As ColorConstants, gruby As Boolean
Load Historie
Historie.Height = (PLR(NumerGracza).MyKolek.Count + 3) * 330 + 600
Historie.Width = 7600
For Each item In PLR(NumerGracza).MyKolek
    Set w = Historie.lista.AddNewWiersz(330)
    Select Case CLng(item.Punkty)
        Case 0: Klr = vbBlack: gruby = False
        Case Is > 49: Klr = vbRed: gruby = True
        Case Is > 29: Klr = vbBlack: gruby = True
        Case Else: Klr = vbBlack: gruby = False
    End Select
        
    Set e = w.AddNewElement((CStr(item.S³owa)), 6000, , , , , Klr, vbAlignNone, vbAlignNone, "Arial CE", 12, , gruby)
    If item.Punkty Then
        Set e = w.AddNewElement((CStr(item.Punkty)), 1500, , , , , Klr, vbAlignNone, vbAlignNone, "Arial CE", 12, , gruby)
    Else
        Set e = w.AddNewElement("-", 1500, , , , , , vbAlignNone, vbAlignNone, "Arial CE", 12)
    End If
Next item

Set w = Historie.lista.AddNewWiersz(330)
Set w = Historie.lista.AddNewWiersz(330)
Set e = w.AddNewElement("Œrednia s³ów: ", 6000, , , , , , vbAlignNone, , , 14)
Set e = w.AddNewElement(Format(PLR(NumerGracza).Œrednia, "0.0#"), 1500, , , , , , vbAlignNone, , , 14)
Set w = Historie.lista.AddNewWiersz(330)

Set e = w.AddNewElement("Œrednia punktów: ", 6000, , , , , , vbAlignNone, , , 14)
If PLR(NumerGracza).wynik Then
    Set e = w.AddNewElement(Format(CSng((PLR(NumerGracza).wynik / PLR(NumerGracza).IleRuchów)), "#.0"), 1500, , , , , , vbAlignNone, , , 14)
Else
    Set e = w.AddNewElement("-", 1500, , , , , , vbAlignNone, , , 14)
End If

Historie.Caption = "Dotychczasowe osi¹gniêcia - " & PLR(NumerGracza).imie
Historie.lista.Drukuj 30, 30

Historie.Show vbModal

End Sub
Public Sub ZmianaLiter(Optional Kmp As Boolean = False)
Dim czer As Long, a As Long, i As Long, j As Long, kl As Collection
Dim LS1 As String, Ls2 As String, LS As String

If IleLiter = 0 Then
    Call InfoBox("Przykro mi bardzo, ale woreczek œwieci pustkami, zatem wymiana jest niemozliwa...")
    Exit Sub
End If

If PLR(Ktory).IloscWymian >= IleWymian Then
    Call InfoBox("Przykro mi bardzo, ale wykorzysta³eœ ju¿ limit wymian...")
    Exit Sub
End If

czer = 0
For i = 0 To 6
    If plytka(i).NumerObrazka = 4 Then
        czer = czer + 1
    End If
Next i

If czer = 0 Then
    Call InfoBox("W celu wymiany literki zaznacz j¹ prawym przyciskiem myszy.")
    Exit Sub
End If

If Kmp = False Then
    a = InfoBox("Czy na pewno chcesz dokonaæ wymiany zaznaczonych literek ?", True)
    If a = 2 Then Exit Sub
End If
Call Anuluj
Omin = 0
Podstawka.Visible = False
For j = 0 To 6
    If plytka(j).NumerObrazka = 4 Then
        For i = 0 To 99
            If Worek(i) = plytka(j).Caption And Wolne(i) Then
                Wolne(i) = False
                IleLiter = IleLiter + 1
                LS = Losuj
                LS1 = Left(LS, 1)
                Ls2 = Right(LS, 1)
                plytka(j).Po³ó¿ LS1, CBool(Ls2)
                plytka(j).ZatwierdŸ 3
                DoEvents
                Exit For
            End If
        Next i
    End If
Next j
Set kl = New Collection
kl.Add " *** wymiana literek ( " & czer & " ) ***"
PLR(Ktory).DodajWyrazy kl, 0
Set kl = Nothing
PLR(Ktory).IloscWymian = PLR(Ktory).IloscWymian + 1
Zero = 0
If PLR(Ktory).Komp = False Then Call NextPlayer

End Sub
Public Sub OminKolejke(Optional Komp As Boolean = False)

Dim a As Long, kl As Collection
If Komp = False Then
    a = InfoBox("Czy na pewno chcesz opuœciæ kolejkê ?", True, False)
    If a = 2 Then Exit Sub
End If
Omin = Omin + 1
Set kl = New Collection
kl.Add " *** opuœci³ kolejkê ***"
PLR(Ktory).DodajWyrazy kl, 0
Set kl = Nothing
Call Anuluj
If PLR(Ktory).Komp = False Then Call NextPlayer

End Sub
Public Function WyrazyOK() As Boolean
Dim item As Variant, i As Long
Dim MaxX As Long, MaxY As Long, MinX As Long, MinY As Long

MinX = 16
MinY = 16

If Le¿¹ce.Count = 0 Then
    WyrazyOK = False
    Exit Function
End If

i = Le¿¹ce(1)
If CzyPionowo(i, Le¿¹ce) = True Then
    Kierunek = wPionowy
Else
    If CzyPoziomo(i, Le¿¹ce) = True Then
        Kierunek = wPoziomy
    Else
        WyrazyOK = False
        Exit Function
    End If
End If

For Each item In Le¿¹ce
    i = CLng(item)
    If wx(i) > MaxX Then MaxX = wx(i)
    If wy(i) > MaxY Then MaxY = wy(i)
    If wx(i) < MinX Then MinX = wx(i)
    If wy(i) < MinY Then MinY = wy(i)
Next item

If Kierunek = wPionowy Then
    For i = MinY To MaxY
        If Not Pole(ciagiem(MinX, i)).Tag = True Then
            WyrazyOK = False
            Exit Function
        End If
    Next i
End If

If Kierunek = wPoziomy Then
    For i = MinX To MaxX
        If Pole(ciagiem(i, MinY)).Tag = False Then
            WyrazyOK = False
            Exit Function
        End If
    Next i
End If

If NumerW = 0 Then
    If Pole(112).Tag = False Then
        WyrazyOK = False
        Exit Function
    End If
Else
    For Each item In Le¿¹ce
        i = CLng(item)
            If wx(i) < 15 Then
                If Pole(ciagiem(wx(i) + 1, wy(i))).NumerObrazka = 2 Then GoTo WOK
            End If
            If wx(i) > 1 Then
                If Pole(ciagiem(wx(i) - 1, wy(i))).NumerObrazka = 2 Then GoTo WOK
            End If
            If wy(i) < 15 Then
                If Pole(ciagiem(wx(i), wy(i) + 1)).NumerObrazka = 2 Then GoTo WOK
            End If
            If wy(i) > 1 Then
                If Pole(ciagiem(wx(i), wy(i) - 1)).NumerObrazka = 2 Then GoTo WOK
            End If
    Next item
    WyrazyOK = False
    Exit Function
End If

WOK:
WyrazyOK = True
End Function

Private Function ZbierajOK() As Boolean

Dim YChMin As Long, XChMin As Long, YChMax As Long, XChMax As Long
Dim ind As Long, ba As Integer, kZB As Long, i As Long
Dim j As Long, MinY As Long, MinX As Long, MaxX As Long, MaxY As Long
Dim Slowo(9) As String, Mno As Long, Pkt(9) As Long
Dim item As Variant, YMinW(9) As Long, YMaxW(9) As Long, XMinW(9) As Long, XMaxW(9) As Long
Dim NPola As Long, pt As Tpunktacja
Dim lezy(0 To 6) As Long, IleV As Long

Set Kolek = New Collection

Razem = 0
Mno = 1
MinX = 16
MinY = 16

For Each item In Le¿¹ce
   i = CLng(item)
   If wx(i) > MaxX Then MaxX = wx(i)
   If wx(i) < MinX Then MinX = wx(i)
   If wy(i) > MaxY Then MaxY = wy(i)
   If wy(i) < MinY Then MinY = wy(i)
Next item

'*************** WYRAZY PIONOWE ************8

If Kierunek = wPionowy Or Kierunek = wSama Then

122 If MinY > 1 And MinY < 16 Then
    If Pole(ciagiem(MinX, MinY - 1)).Tag Then
       MinY = MinY - 1
       GoTo 122
    End If
End If
        
124  If MaxY < 15 Then
     If Pole(ciagiem(MaxX, MaxY + 1)).Tag Then
            MaxY = MaxY + 1
            GoTo 124
        End If
    End If
    
    If MinY < MaxY Then
        For i = MinY To MaxY
            NPola = ciagiem(MinX, i)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
   
    For i = MinY To MaxY
        If Pole(ciagiem(MinX, i)).NumerObrazka = 1 Then
            XChMin = MinX
            XChMax = MaxX
20          If XChMin > 1 Then
                If Pole(ciagiem(XChMin - 1, i)).Tag Then
                    XChMin = XChMin - 1
                    GoTo 20
                End If
            End If
30          If XChMax < 15 Then
                If Pole(ciagiem(XChMax + 1, i)).Tag Then
                    XChMax = XChMax + 1
                    GoTo 30
                End If
            End If
            XMinW(kZB) = XChMin
            XMaxW(kZB) = XChMax
            YMinW(kZB) = i
            YMaxW(kZB) = i
            kZB = kZB + 1
        End If
    Next i
        
    For i = 1 To kZB
        If XMinW(i) < XMaxW(i) Then
            For j = XMinW(i) To XMaxW(i)
                NPola = ciagiem(j, YMinW(i))
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If

'*************** WYRAZY POZIOME ************8

If Kierunek = wPoziomy Then

222 If MinX > 1 Then
        If Pole(ciagiem(MinX - 1, MinY)).Tag Then
            MinX = MinX - 1
            GoTo 222
        End If
    End If
        
224 If MaxX < 15 Then
        If Pole(ciagiem(MaxX + 1, MinY)).Tag Then
            MaxX = MaxX + 1
            GoTo 224
        End If
    End If
    
    If MinX < MaxX Then
        For i = MinX To MaxX
            NPola = ciagiem(i, MinY)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
    For i = MinX To MaxX
        If Pole(ciagiem(i, MinY)).NumerObrazka = 1 Then
            YChMin = MinY
            YChMax = MaxY
202         If YChMin > 1 Then
                If Pole(ciagiem(i, YChMin - 1)).Tag Then
                     YChMin = YChMin - 1
                     GoTo 202
                End If
            End If
302         If YChMax < 15 Then
                If Pole(ciagiem(i, YChMax + 1)).Tag Then
                    YChMax = YChMax + 1
                    GoTo 302
                End If
            End If
            XMinW(kZB) = i
            XMaxW(kZB) = i
            YMinW(kZB) = YChMin
            YMaxW(kZB) = YChMax
            kZB = kZB + 1
        End If
    Next i
    For i = 1 To kZB
        If YMinW(i) < YMaxW(i) Then
            For j = YMinW(i) To YMaxW(i)
                NPola = ciagiem(XMinW(i), j)
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If
    
For i = 0 To kZB
    If Slowo(i) <> "" And InStr(1, Slowo(i), Mydlo) = 0 Then
        Me.MousePointer = 0
           If SlownikCheck(Slowo(i)) = True Then
                ba = InfoBox("Wyrazu " & Slowo(i) & " nie ma w s³owniku. Czy wszyscy uczestnicy zgadzaj¹ siê na jego u¿ycie ?", True, True)
                If ba = 2 Then
                    ZbierajOK = False
                    Exit Function
                End If
                If ba = 3 Then Call DodajSlownik(Slowo(i))
            End If
        Me.MousePointer = 11
    End If
    Razem = Razem + Pkt(i)
Next i

If PLR(Ktory).Komp = True Then
    PLR(Ktory).wynik = PLR(Ktory).wynik + Razem
    wynik(Ktory - 1).Caption = PLR(Ktory).wynik
    Kierunek = wBrak
    ZbierajOK = True
End If
Lst.Clear

Set w = Lst.AddNewWiersz(20)
Set e = w.AddNewElement("S³owo:", 200, , , , vbYellow, , , , , 15)
Set e = w.AddNewElement("Punkty:", 90, , , , vbYellow, , , , , 15)
Set w = Lst.AddNewWiersz(25)
Set e = w.AddNewElement(eBackColor:=vbWhite, szerokoœæ:=200)
Set e = w.AddNewElement(eBackColor:=vbWhite, szerokoœæ:=90)

For i = 0 To kZB
    If Not Slowo(i) = "" Then
        Kolek.Add BezMydla(Slowo(i))
        Set w = Lst.AddWiersz(25)
        Set e = w.AddNewElement(BezMydla(Slowo(i)), 200, , , , , , , , , 15)
        Set e = w.AddNewElement(CStr(Pkt(i)), 90, , , , , , vbAlignNone, , , 15)
    End If
Next i

Set w = Lst.AddNewWiersz(6)
Set e = w.AddNewElement(, 200, , , , vbWhite)
Set e = w.AddNewElement(, 90, , , , vbWhite)

Dim nPremia As Long

Select Case Le¿¹ce.Count

    Case 1
        nPremia = 0
    Case 2
        nPremia = 0
    Case 3
        nPremia = 0
    Case 4
        nPremia = 0
    Case 5
        nPremia = 0
    Case 6
        nPremia = 0
    Case 7
        nPremia = 50
    Case Else
        nPremia = 0

End Select

Razem = Razem + nPremia

Set w = Lst.AddNewWiersz(25)
Set e = w.AddNewElement("PREMIA:", 200, , , , , vbRed, , , , 15)
Set e = w.AddNewElement(CStr(nPremia), 90, , , , , vbRed, vbAlignNone, , , 15)

'If Le¿¹ce.Count = 7 Then
'    Set w = Lst.AddNewWiersz(25)
'    Set e = w.AddNewElement("PREMIA:", 200, , , , , vbRed, , , , 15)
'    Set e = w.AddNewElement("50", 90, , , , , vbRed, vbAlignNone, , , 15)
'    Razem = Razem + 50
'End If

Set w = Lst.AddNewWiersz(15)
Set e = w.AddNewElement(, 200, , , , vbWhite)
Set e = w.AddNewElement(, 90, , , , vbWhite)
Set w = Lst.AddNewWiersz(25)
Set e = w.AddNewElement("Razem punktów:", 200, , , , vbWhite, vbBlue, , , , 16)
Set e = w.AddNewElement(CStr(Razem), 90, , , , vbWhite, vbMagenta, vbAlignNone, , , 16)

If Zalicz = False Then
    Set Kolek = Nothing
    ZbierajOK = False
Else
    PLR(Ktory).wynik = PLR(Ktory).wynik + Razem
    wynik(Ktory - 1).Caption = PLR(Ktory).wynik
    Kierunek = wBrak
    ZbierajOK = True
End If
Me.MousePointer = 0
End Function

Private Function Zalicz() As Boolean
efekty.Show vbModal
Zalicz = efekty.decyzja
End Function
Public Sub Obliczamy()
Me.MousePointer = 11
Me.MousePointer = 0
If WyrazyOK Then
    If ZbierajOK Then
        Zegar.Enabled = False
        Zatwierdz
        NextPlayer
        Zegar.Enabled = True
    End If
Else
    InfoBox "Po³o¿y³eœ litery nieprawid³owo.", False, False
End If

End Sub

Public Sub SzukajWszystkie2(Komp As Boolean)
Dim ZAPYTANIE(7) As String, dc As Single
Dim i As Long, Ile As Long, Lit As String, z As String * 1, IleRek As Long
Dim r As cMojRecordset, w() As String
Static oldlitery As String
dc = Timer
    For i = 0 To 6
        z = plytka(i).Caption
        If z <> "_" Then
            Ile = Ile + 1
            If z = Mydlo Then
                Lit = Lit & "_"
            Else
                Lit = Lit & z
            End If
        End If
    Next i
    
If Komp = False Then
    Me.MousePointer = 11
    Szukam.Show
    OnTop Szukam.hwnd, True
    DoEvents
    Znalaz.List1.Clear
End If


If Komp = False Then
   If oldlitery <> Lit Then
      oldlitery = Lit
         MojaBaza.ClearWlasne
         For i = Ile To 2 Step -1
            w = KtoSzukaMB(i, Lit)
            MojaBaza.InsertToWlasne i, w
         Next i
      End If
   ReDim w(0)
   For i = Ile To 2 Step -1
      Set r = MojaBaza.OpenWlasneWyraz(i, 0, w)
      If r.NoMatch = False Then
         r.MoveFirst
         While Not r.EOF
            Znalaz.List1.AddItem CStr(r.Element())
            r.MoveNext
         Wend
      End If
      IleRek = IleRek + r.RecordCount
      Set r = Nothing
   Next i
   Me.MousePointer = 0
   Set r = Nothing
   Znalaz.List1.AddItem ""
   Znalaz.List1.AddItem "Znalaz³em " & IleRek & " wyrazów."
   OnTop Szukam.hwnd, False
   Szukam.Hide
   Znalaz.Show vbModal, graslow
Else
   MojaBaza.ClearWlasne
   For i = Ile To 2 Step -1
      MojaBaza.InsertToWlasne i, KtoSzukaMB(i, Lit)
   Next i
End If

End Sub
Public Sub OnTop(hwnd As Long, OnTop As Boolean)
Dim flaga As Long
Dim fl As Long
fl = Swp_Nomove Or Swp_Nosize Or Swp_ShowWindow Or Swp_NoActivate
If OnTop = False Then
    flaga = Hwnd_NoTopMost
Else
    flaga = Hwnd_TopMost
End If
SetWindowPos hwnd, flaga, 0, 0, 0, 0, fl

End Sub

Public Sub ZbierajTu()
Dim YChMin As Long, XChMin As Long, YChMax As Long, XChMax As Long
Dim ind As Long, ba As Integer, kZB As Long, i As Long
Dim j As Long, MinY As Long, MinX As Long, MaxX As Long, MaxY As Long
Dim Slowo(9) As String, Mno As Long, Pkt(9) As Long
Dim item As Variant, YMinW(9) As Long, YMaxW(9) As Long, XMinW(9) As Long, XMaxW(9) As Long
Dim NPola As Long, RazemTu As Long
Set Kolek = New Collection
RazemTu = 0
Mno = 1
MinX = 16
MinY = 16

For Each item In Le¿¹ce
    i = CLng(item)
    If wx(i) > MaxX Then MaxX = wx(i)
    If wy(i) > MaxY Then MaxY = wy(i)
    If wx(i) < MinX Then MinX = wx(i)
    If wy(i) < MinY Then MinY = wy(i)
Next item

'*************** WYRAZY PIONOWE ************8

If Kierunek = wPionowy Or Kierunek = wSama Then

122 If MinY > 1 And MinY < 16 Then
    If Pole(ciagiem(MinX, MinY - 1)).Tag Then
       MinY = MinY - 1
       GoTo 122
    End If
End If
        
124  If MaxY < 15 Then
     If Pole(ciagiem(MaxX, MaxY + 1)).Tag Then
            MaxY = MaxY + 1
            GoTo 124
        End If
    End If
    
    If MinY < MaxY Then
        For i = MinY To MaxY
            NPola = ciagiem(MinX, i)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
   
    For i = MinY To MaxY
        If Pole(ciagiem(MinX, i)).NumerObrazka = 1 Then
            XChMin = MinX
            XChMax = MaxX
20          If XChMin > 1 Then
                If Pole(ciagiem(XChMin - 1, i)).Tag Then
                    XChMin = XChMin - 1
                    GoTo 20
                End If
            End If
30          If XChMax < 15 Then
                If Pole(ciagiem(XChMax + 1, i)).Tag Then
                    XChMax = XChMax + 1
                    GoTo 30
                End If
            End If
            XMinW(kZB) = XChMin
            XMaxW(kZB) = XChMax
            YMinW(kZB) = i
            YMaxW(kZB) = i
            kZB = kZB + 1
        End If
    Next i
        
    For i = 1 To kZB
        If XMinW(i) < XMaxW(i) Then
            For j = XMinW(i) To XMaxW(i)
                NPola = ciagiem(j, YMinW(i))
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If

'*************** WYRAZY POZIOME ************8

If Kierunek = wPoziomy Then

222 If MinX > 1 Then
        If Pole(ciagiem(MinX - 1, MinY)).Tag Then
            MinX = MinX - 1
            GoTo 222
        End If
    End If
        
224 If MaxX < 15 Then
        If Pole(ciagiem(MaxX + 1, MinY)).Tag Then
            MaxX = MaxX + 1
            GoTo 224
        End If
    End If
    
    If MinX < MaxX Then
        For i = MinX To MaxX
            NPola = ciagiem(i, MinY)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
    For i = MinX To MaxX
        If Pole(ciagiem(i, MinY)).NumerObrazka = 1 Then
            YChMin = MinY
            YChMax = MaxY
202         If YChMin > 1 Then
                If Pole(ciagiem(i, YChMin - 1)).Tag Then
                     YChMin = YChMin - 1
                     GoTo 202
                End If
            End If
302         If YChMax < 15 Then
                If Pole(ciagiem(i, YChMax + 1)).Tag Then
                    YChMax = YChMax + 1
                    GoTo 302
                End If
            End If
            XMinW(kZB) = i
            XMaxW(kZB) = i
            YMinW(kZB) = YChMin
            YMaxW(kZB) = YChMax
            kZB = kZB + 1
        End If
    Next i
    For i = 1 To kZB
        If YMinW(i) < YMaxW(i) Then
            For j = YMinW(i) To YMaxW(i)
                NPola = ciagiem(XMinW(i), j)
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If
For i = 0 To kZB
    RazemTu = RazemTu + Pkt(i)
Next i

EfektyTu.Caption = ""
For i = 0 To kZB
    If Slowo(i) <> "" Then EfektyTu.Caption = EfektyTu.Caption & " " & BezMydla(Slowo(i)) & " - " & Pkt(i) & vbCrLf & " "
Next i

Dim nPremia As Long

Select Case Le¿¹ce.Count

    Case 1
        nPremia = 0
    Case 2
        nPremia = 0
    Case 3
        nPremia = 0
    Case 4
        nPremia = 0
    Case 5
        nPremia = 0
    Case 6
        nPremia = 0
    Case 7
        nPremia = 50
    Case Else
        nPremia = 0

End Select

Razem = Razem + nPremia
'If Le¿¹ce.Count = 7 Then
EfektyTu.Caption = EfektyTu.Caption & vbCrLf & " "
EfektyTu.Caption = EfektyTu.Caption & " PREMIA - " & CStr(nPremia)
EfektyTu.Caption = EfektyTu.Caption & vbCrLf
RazemTu = RazemTu + nPremia
'End If
EfektyTu.Caption = EfektyTu.Caption & vbCrLf & " "
EfektyTu.Caption = EfektyTu.Caption & " RAZEM - " & CStr(RazemTu)
End Sub

Public Sub EfektujTu()

If WyrazyOK = True Then
    ZbierajTu
    Solve.Enabled = True
    AkcjeSub(1).Enabled = True
Else
    EfektyTu.Caption = "- 0 -"
    Solve.Enabled = False
    AkcjeSub(1).Enabled = False
End If

End Sub

Public Sub Nazywaj()
Dim i As Long
For i = 1 To IleGraczy
    If IleLiter = 0 Then
        player(i - 1).Caption = PLR(i).imie & " (" & CStr(Policz(i)) & ")"
    Else
        player(i - 1).Caption = PLR(i).imie
    End If
Next i

End Sub

Public Function Policz(Kt As Long) As Long
Dim i As Long, n As Long
   
For i = 0 To 6
    If Left(Stojak(Kt, i), 1) <> "_" Then n = n + 1
Next i
Policz = n

End Function

Public Function svrKladzie(Wyraz As String, Indeks As Long, KirUnek As wKierunek, strStojak As String) As Boolean
Dim bBlank As Boolean
Dim i As Long
Dim strNormWord As String

Podstawka.Visible = False
DoEvents

For i = 1 To Len(Wyraz) Step 2
    strNormWord = strNormWord & Mid$(Wyraz, i, 1)
Next i

For i = 0 To 6
    If Mid$(strStojak, i * 2 + 2, 1) = "1" Then
        bBlank = True
    Else
        bBlank = False
    End If
    plytka(i).Po³ó¿ Mid$(strStojak, i * 2 + 1, 1), bBlank
Next i

svrKladzie = Kladzie(strNormWord, Indeks, KirUnek, 0)

SprawdzKladzie
DoEvents
'Call ZbierajOK
Call Zatwierdz

NextPlayer

End Function

Public Function Kladzie(Wyraz As String, Indeks As Long, KirUnek As wKierunek, Punkty As Long) As Boolean
Dim i As Long, xx As Long, yy As Long
Dim n As Long, Uzyte(7) As Boolean, znalazlem As Boolean
Dim j As Long

For i = 0 To 6
    If plytka(i).Caption = "_" Then Uzyte(i) = True
Next i

Kierunek = KirUnek

xx = wx(Indeks)
yy = wy(Indeks)

If KirUnek = wPoziomy Then
    For i = 1 To Len(Wyraz)
        If (xx + i - 1) > 15 Then Exit For
        If Pole(ciagiem(xx + i - 1, yy)).Tag = False Then
            For j = 0 To 6
                znalazlem = False
                If Uzyte(j) = False And plytka(j).Caption = Mid(Wyraz, i, 1) And plytka(j).Blank = False Then
                    Indeks = ciagiem(xx + i - 1, yy)
                    Pole(Indeks).Po³ó¿ Mid(Wyraz, i, 1), False
                    Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
                    Call plytka(j).Po³ó¿("_", False)
                    Call plytka(j).ZatwierdŸ(3)
                    Dlug = Dlug + 1
                    DoEvents
                    Sleep Horror * 1000
                    Uzyte(j) = True
                    znalazlem = True
                    Exit For
                End If
            Next j
            If znalazlem = False Then
                For j = 0 To 6
                    If Uzyte(j) = False And plytka(j).Blank Then
                        Indeks = ciagiem(xx + i - 1, yy)
                        Pole(Indeks).Po³ó¿ Mid(Wyraz, i, 1), True
                        Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
                        Call plytka(j).Po³ó¿("_", False)
                        Call plytka(j).ZatwierdŸ(3)
                        Dlug = Dlug + 1
                        DoEvents
                        Sleep Horror * 1000
                        Uzyte(j) = True
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
Else
    For i = 1 To Len(Wyraz)
        If (yy + i - 1) > 15 Then Exit For
        If Pole(ciagiem(xx, yy + i - 1)).Tag = False Then
            For j = 0 To 6
                znalazlem = False
                If Uzyte(j) = False And plytka(j).Caption = Mid(Wyraz, i, 1) And plytka(j).Blank = False Then
                    Indeks = ciagiem(xx, yy + i - 1)
                    Pole(Indeks).Po³ó¿ Mid(Wyraz, i, 1), False
                    Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
                    Call plytka(j).Po³ó¿("_", False)
                    Call plytka(j).ZatwierdŸ(3)
                    Dlug = Dlug + 1
                    DoEvents
                    Sleep Horror * 1000
                    Uzyte(j) = True
                    znalazlem = True
                    Exit For
                End If
            Next j
            If znalazlem = False Then
                For j = 0 To 6
                    If Uzyte(j) = False And plytka(j).Blank Then
                        Indeks = ciagiem(xx, yy + i - 1)
                        Call Pole(Indeks).Po³ó¿(Mid(Wyraz, i, 1), True)
                        Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
                        Call plytka(j).Po³ó¿("_", False)
                        Call plytka(j).ZatwierdŸ(3)
                        Dlug = Dlug + 1
                        DoEvents
                        Sleep Horror * 1000
                        Uzyte(j) = True
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
End If
    
1234 If WyrazyOK = True Then
    Sloweczka.AddItem Punkty & " - " & Wyraz & " - " & Indeks
    Kladzie = True
Else
    Kladzie = False
End If

End Function
Public Function CzyDoWymiany() As Boolean
Dim i As Long

For i = 0 To 6
   If plytka(i).NumerObrazka = 4 Then
      CzyDoWymiany = True
      Exit Function
   End If
Next i
CzyDoWymiany = False

End Function

Public Function SprawdzKladzie() As Boolean
Dim YChMin As Long, XChMin As Long, YChMax As Long, XChMax As Long
Dim ind As Long, ba As Integer, kZB As Long, i As Long
Dim j As Long, MinY As Long, MinX As Long, MaxX As Long, MaxY As Long
Dim Slowo(9) As String, Mno As Long, Pkt(9) As Long
Dim item As Variant, YMinW(9) As Long, YMaxW(9) As Long, XMinW(9) As Long, XMaxW(9) As Long
Dim NPola As Long, JestM As Boolean, RazemTu As Long
Set Kolek = New Collection
Razem = 0
Mno = 1
MinX = 16
MinY = 16
JestM = False
If Le¿¹ce.Count = 0 Then
    SprawdzKladzie = False
    Exit Function
End If
For Each item In Le¿¹ce
    i = CLng(item)
    If wx(i) > MaxX Then MaxX = wx(i)
    If wy(i) > MaxY Then MaxY = wy(i)
    If wx(i) < MinX Then MinX = wx(i)
    If wy(i) < MinY Then MinY = wy(i)
Next item

'*************** WYRAZY PIONOWE ************8

If Kierunek = wPionowy Or Kierunek = wSama Then

122 If MinY > 1 And MinY < 16 Then
    If Pole(ciagiem(MinX, MinY - 1)).Tag Then
       MinY = MinY - 1
       GoTo 122
    End If
End If
        
124  If MaxY < 15 Then
     If Pole(ciagiem(MaxX, MaxY + 1)).Tag Then
            MaxY = MaxY + 1
            GoTo 124
        End If
    End If
    
    If MinY < MaxY Then
        For i = MinY To MaxY
            NPola = ciagiem(MinX, i)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
   
    For i = MinY To MaxY
        If Pole(ciagiem(MinX, i)).NumerObrazka = 1 Then
            XChMin = MinX
            XChMax = MaxX
20          If XChMin > 1 Then
                If Pole(ciagiem(XChMin - 1, i)).Tag Then
                    XChMin = XChMin - 1
                    GoTo 20
                End If
            End If
30          If XChMax < 15 Then
                If Pole(ciagiem(XChMax + 1, i)).Tag Then
                    XChMax = XChMax + 1
                    GoTo 30
                End If
            End If
            XMinW(kZB) = XChMin
            XMaxW(kZB) = XChMax
            YMinW(kZB) = i
            YMaxW(kZB) = i
            kZB = kZB + 1
        End If
    Next i
        
    For i = 1 To kZB
        If XMinW(i) < XMaxW(i) Then
            For j = XMinW(i) To XMaxW(i)
                NPola = ciagiem(j, YMinW(i))
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + (Pole(NPola).RazyL * Pole(NPola).Wartoœæ)
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If

'*************** WYRAZY POZIOME ************8

If Kierunek = wPoziomy Then

222 If MinX > 1 Then
        If Pole(ciagiem(MinX - 1, MinY)).Tag Then
            MinX = MinX - 1
            GoTo 222
        End If
    End If
        
224 If MaxX < 15 Then
        If Pole(ciagiem(MaxX + 1, MinY)).Tag Then
            MaxX = MaxX + 1
            GoTo 224
        End If
    End If
    
    If MinX < MaxX Then
        For i = MinX To MaxX
            NPola = ciagiem(i, MinY)
            Slowo(0) = Slowo(0) & Pole(NPola).Caption
            If Pole(NPola).NumerObrazka = 1 Then
                Pkt(0) = Pkt(0) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                Mno = Mno * Pole(NPola).RazyS
            Else
                Pkt(0) = Pkt(0) + Pole(NPola).Wartoœæ
            End If
        Next i
        Pkt(0) = Pkt(0) * Mno
        Mno = 1
    End If
    kZB = 1
    For i = MinX To MaxX
        If Pole(ciagiem(i, MinY)).NumerObrazka = 1 Then
            YChMin = MinY
            YChMax = MaxY
202         If YChMin > 1 Then
                If Pole(ciagiem(i, YChMin - 1)).Tag Then
                     YChMin = YChMin - 1
                     GoTo 202
                End If
            End If
302         If YChMax < 15 Then
                If Pole(ciagiem(i, YChMax + 1)).Tag Then
                    YChMax = YChMax + 1
                    GoTo 302
                End If
            End If
            XMinW(kZB) = i
            XMaxW(kZB) = i
            YMinW(kZB) = YChMin
            YMaxW(kZB) = YChMax
            kZB = kZB + 1
        End If
    Next i
    For i = 1 To kZB
        If YMinW(i) < YMaxW(i) Then
            For j = YMinW(i) To YMaxW(i)
                NPola = ciagiem(XMinW(i), j)
                Slowo(i) = Slowo(i) & Pole(NPola).Caption
                If Pole(NPola).NumerObrazka = 1 Then
                    Pkt(i) = Pkt(i) + Pole(NPola).RazyL * Pole(NPola).Wartoœæ
                    Mno = Mno * Pole(NPola).RazyS
                Else
                    Pkt(i) = Pkt(i) + Pole(NPola).Wartoœæ
                End If
            Next j
            Pkt(i) = Pkt(i) * Mno
            Mno = 1
        End If
    Next i
End If

For i = 0 To kZB
    Razem = Razem + Pkt(i)
Next i
'If Le¿¹ce.Count = 7 Then Razem = Razem + 50
Dim nPremia As Long

Select Case Le¿¹ce.Count

    Case 1
        nPremia = 0
    Case 2
        nPremia = 0
    Case 3
        nPremia = 0
    Case 4
        nPremia = 0
    Case 5
        nPremia = 0
    Case 6
        nPremia = 0
    Case 7
        nPremia = 50
    Case Else
        nPremia = 0

End Select

Razem = Razem + nPremia


For i = 0 To kZB
    If Slowo(i) <> "" And InStr(1, Slowo(i), Mydlo) = 0 Then
        Me.MousePointer = 0
        'If Sprawdzaj(Slowo(i)) = False Then
            If SlownikCheck(Slowo(i)) Then
                SprawdzKladzie = False
                Exit Function
            End If
        'End If
        Me.MousePointer = 11
    End If
Next i

Lst.Clear

For i = 0 To kZB
    If Not Slowo(i) = "" Then
        Kolek.Add BezMydla(Slowo(i))
    End If
Next i
    PLR(Ktory).wynik = PLR(Ktory).wynik + Razem
    'If m_Server Is Nothing Then
        wynik(Ktory - 1).Caption = PLR(Ktory).wynik
    'End If
    Kierunek = wBrak
    SprawdzKladzie = True
    Exit Function

If Zalicz = False Then
    Set Kolek = Nothing
    SprawdzKladzie = False
Else
    PLR(Ktory).wynik = PLR(Ktory).wynik + Razem
    'If m_Server Is Nothing Then
        wynik(Ktory - 1).Caption = PLR(Ktory).wynik
    'End If
    Kierunek = wBrak
    SprawdzKladzie = True
End If
End Function
Public Sub KompKladzie()
Dim dtp As tDlugosci
CzasKompa = Timer
Postep.Top = graslow.ScaleHeight - Postep.Height - 2
Postep.Left = graslow.ScaleWidth - Postep.Width - 2
Bar.Value = 0
Postep.Visible = True
Postep.ZOrder
DoEvents
'OnTop Postep.hwnd, True
graslow.SetFocus
DoEvents
Sloweczka.Clear
MaxPkt = 0
MaxLiter = 0
Me.MousePointer = 11
On Error Resume Next
Call VirtUklada

If Pole(112).Tag Then
    SortLiterki = SortLiteryTAB()
    Call SzukajPlus(1)
    If PLR(Ktory).MaxPsz > 0 Then
        dtp = ZnajdzSlowa2()
        Call DodajSlowa3(dtp.MinL, dtp.MaxL)
    End If
    Call VirtLinieAll
Else
    Call KompZaczyna(False)
End If
OnTop Postep.hwnd, True
Postep.Visible = False
Postep.ZOrder 1
Call Kladz

Me.MousePointer = 0

End Sub

Public Sub Kladz()
Dim kr As wKierunek, i As Long
Dim RC As Long, strata As Single
Dim punkty1 As Long, los As Long, mpkt As Long
Dim Cz As cCzas, Tek As String, Polozyl As Boolean
Dim Wyraz As String, St As Long, st2 As Long

MojaBaza.SortujAW MaxPkt

If MojaBaza.AllWyrazySort.RecordCount > 0 Then
   RC = MojaBaza.AllWyrazySort.RecordCount
   Randomize Timer
   los = Int(Rnd() * RC) + 1
   DoEvents
   Wyraz = MojaBaza.AllWyrazySort.Element(los).Wyraz
   kr = MojaBaza.AllWyrazySort.Element(los).Kierunek
   punkty1 = MojaBaza.AllWyrazySort.Element(los).Punkty
   St = MojaBaza.AllWyrazySort.Element(los).Start
   
   If IleLiter > MinLiter Then
      If punkty1 > PLR(Ktory).MinPKT Then
         Call Kladzie(Wyraz, St, kr, punkty1)
         Call SprawdzKladzie
         'status.Caption = CStr(punkty1) & ": " & Wyraz & ":" & St
         strata = Le¿¹ce.Count * Horror
         DoEvents
         Call Zatwierdz
      Else
         If PLR(Ktory).IloscWymian < IleWymian Then
            Call Zaznacz
            Call ZmianaLiter(True)
         Else
            Call Kladzie(Wyraz, St, kr, punkty1)
            SprawdzKladzie
            'status.Caption = CStr(punkty1) & ": " & Wyraz & ":" & St
            strata = Le¿¹ce.Count * Horror
            DoEvents
            Call Zatwierdz
         End If
      End If
   Else
      If punkty1 > 0 Then
         Call Kladzie(Wyraz, St, kr, punkty1)
         SprawdzKladzie
         'status.Caption = CStr(punkty1) & ": " & Wyraz & ":" & St
         strata = Le¿¹ce.Count * Horror
         DoEvents
         Call Zatwierdz
      Else
         Call OminKolejke(True)
      End If
   End If
Else
   If PLR(Ktory).IloscWymian < IleWymian And IleLiter > MinLiter Then
      Call Zaznacz
      Call ZmianaLiter(True)
   Else
      Call OminKolejke(True)
   End If
End If
      
Call VirtAnuluj

Set Cz = New cCzas
Cz.Czas = CSng(Timer - CzasKompa - strata)
Czasy.Add Cz
Set Cz = Nothing
Me.MousePointer = 0
If MenuShowTime.Checked Then
   Tek = "Ca³kowity czas: " & Format(SumaCzas, "0.00") & vbCrLf & "Œredni czas ruchu komputera: " & FormaCzasu(CLng(SredniCzas)) & vbCrLf & "Maksymalny czas ruchu komputera: " & FormaCzasu(CLng(MaxCzas)) & vbCrLf & "Iloœæ ruchów komputera: " & Czasy.Count & vbCrLf & "Czas ostatniego ruchu: " & Format(Timer - CzasKompa - strata, "00.00")
   MsgBox Tek, vbInformation
End If

Call NextPlayer

End Sub

Private Function Punkty(Wyraz As String, ByVal Start As Long, ByVal Kier As wKierunek) As Tpunktacja
Dim Pkt As Tpunktacja

Pkt.Punkty = -1

If VirtKladzie(Wyraz, Kier, Start) Then
   Pkt = VirtPunkty(IleLiter, VirtLezy(0), IlePolLezy, Kier, PlanszaWART(0), PlanszaRAZYL(0), PlanszaRAZYS(0), PlanszaTAG(0), PlanszaLICZ(0))
End If
Punkty = Pkt

End Function

Public Function svrPunkty(Wyraz As String, ByVal Start As Long, ByVal Kier As wKierunek) As Long

svrPunkty = Punkty(Wyraz, Kier, Start).Wartosc

End Function

Public Function svrGetNewStojak(strStojak As String) As String

Dim i As Long
Dim strNew As String
Dim strNewStojak As String

For i = 0 To 6
    If Mid$(strStojak, i * 2 + 1, 1) = "_" Then
        strNew = Losuj()
        strNewStojak = strNewStojak & strNew
        
    Else
        strNewStojak = strNewStojak & Mid$(strStojak, i * 2 + 1, 2)
    End If

Next i

svrGetNewStojak = strNewStojak

End Function

Public Sub CzytajHS()

If m_Klient Is Nothing Then

    Rekord = GetSetting("£ów S³ów", "Rekordy", "Rekord_" & Jêzyk.Klucz & "_" & IleGraczy)
    ImieRek = GetSetting("£ów S³ów", "Rekordy", "Imie_" & Jêzyk.Klucz & "_" & IleGraczy)

    lblRekord.Caption = "Rekord : " & ImieRek & " - " & Rekord
End If
End Sub

Public Sub KompZaczyna(Optional CzyWlasne As Boolean = False)
Dim dc As Single
Dim i As Long, Ile As Long, Lit As String, z As String * 1, Wyraz As String, Kier As Long, pnkt As Long
Dim Start As Long, Pkt3 As Tpunktacja, Kierun As wKierunek
Dim ileL As Long, Warki() As String, czaswl As Long
MaxPkt = 0
For i = 0 To 6
   z = plytka(i).Caption
   If z <> "_" Then
      ileL = ileL + 1
      If z = Mydlo Then
         Lit = Lit & "_"
      Else
         Lit = Lit & z
      End If
   End If
Next i
Me.MousePointer = 11

If CzyWlasne = False Then
    MojaBaza.ClearAllWyrazy
End If

If ileL < 2 Then Exit Sub

If CzyWlasne = True Then
   MojaBaza.ClearWlasne
   For i = ileL To 2 Step -1
      Warki = KtoSzukaMB(i, Lit)
      MojaBaza.InsertToWlasne i, Warki
   Next i
   Me.MousePointer = 0
   Exit Sub
End If

On Error Resume Next

For i = ileL To 2 Step -1
   Warki = KtoSzukaMB(i, Lit)
   MojaBaza.InsertToAllWyrazy2 i, Warki
Next i

On Error Resume Next
Dim FW As cFullWyraz
For i = 1 To MojaBaza.AllWyrazy.RecordCount
    Kierun = wPoziomy
    Set FW = MojaBaza.AllWyrazy.Element(i)
    Wyraz = FW.Wyraz
    Start = StartZacz(Wyraz)
    Pkt3 = Punkty(Wyraz, Start, wPoziomy)
    If Pkt3.Wsp >= MaxPkt Then
        MaxPkt = Pkt3.Wsp
        MojaBaza.InsertToAllWyrazy Wyraz, Start, Pkt3.Wsp, Kierun, Pkt3.Wartosc, Len(Wyraz)
    End If
    VirtAnuluj
Next i
On Error GoTo 0
NieZaczynam:

End Sub

Private Function UstalWzorzec(Numer As Long, Kierunek As wKierunek, nLiter As Long) As String

Dim X As Long, Y As Long, Wz As String, Pelne As Boolean, xx As Long, i As Long
Dim yy As Long, PoczX As Long, PoczY As Long, KonX As Long, KonY As Long
Dim psz As Long
X = wx(Numer)
Y = wy(Numer)

If X < 1 Or Y < 1 Or X > 15 Or Y > 15 Then
    UstalWzorzec = "0"
    Exit Function
End If

If Kierunek = wPoziomy Then
    Pelne = True
    xx = X
    While xx > 1 And Pelne = True
        If xx < 1 Then GoTo Zero
        If Pole(ciagiem(xx - 1, Y)).Tag Then
            xx = xx - 1
        Else
            Pelne = False
        End If
    Wend
Zero:
    
    If xx > 1 Then
        PoczX = xx
    Else
        PoczX = 1
    End If
    
    While i < nLiter And xx < 16
        If Pole(ciagiem(xx, Y)).Tag Then
            Wz = Wz & Pole(ciagiem(xx, Y)).Caption
        Else
            Wz = Wz & "_"
            i = i + 1
        End If
        xx = xx + 1
    Wend
    
    Pelne = True
    
    While xx < 16 And Pelne = True
        If Pole(ciagiem(xx, Y)).Tag Then
            Wz = Wz & Pole(ciagiem(xx, Y)).Caption
            xx = xx + 1
        Else
            Pelne = False
        End If
    Wend
    If xx <= 16 Then
        KonX = xx - 1
    Else
        KonX = 15
    End If
Else

    Pelne = True
    yy = Y
    While yy > 1 And Pelne = True
        If Pole(ciagiem(X, yy - 1)).Tag Then
            yy = yy - 1
        Else
            Pelne = False
        End If
    Wend

Zero1:
    If yy > 1 Then
        PoczX = yy
    Else
        PoczX = 1
    End If
    
    While i < nLiter And yy < 16
        If Pole(ciagiem(X, yy)).Tag Then
            Wz = Wz & Pole(ciagiem(X, yy)).Caption
        Else
            Wz = Wz & "_"
            i = i + 1
        End If
        yy = yy + 1
    Wend
    
    Pelne = True
    
    While yy < 16 And Pelne = True
        If Pole(ciagiem(X, yy)).Tag Then
            Wz = Wz & Pole(ciagiem(X, yy)).Caption
            yy = yy + 1
        Else
            Pelne = False
        End If
    Wend
    If yy <= 16 Then
        KonX = yy - 1
    Else
        KonX = 15
    End If
End If


psz = 0
For i = PoczX To KonX
    If Kierunek = wPoziomy Then
        If Y > 1 Then
            If Pole(ciagiem(i, Y - 1)).Tag And Pole(ciagiem(i, Y)).Tag = False Then
                psz = psz + 1
            End If
        End If
        If Y < 15 Then
            If Pole(ciagiem(i, Y + 1)).Tag And Pole(ciagiem(i, Y)).Tag = False Then
                psz = psz + 1
            End If
        End If
    Else
        If X > 1 Then
            If Pole(ciagiem(X - 1, i)).Tag And Pole(ciagiem(X, i)).Tag = False Then
                psz = psz + 1
            End If
        End If
        If X < 15 Then
            If Pole(ciagiem(X + 1, i)).Tag And Pole(ciagiem(X, i)).Tag = False Then
                psz = psz + 1
            End If
        End If
    End If
Next i

If i < nLiter Or psz > 1 Then
    UstalWzorzec = "0"
Else
    UstalWzorzec = Wz
End If
    
End Function




Public Function VirtKladzie(Wyraz As String, ByVal Kierunek As wKierunek, Indeks As Long) As Boolean

Dim i As Long, xx As Long, yy As Long
Dim n As Long, Uzyte(7) As Boolean, znalazlem As Boolean
Dim j As Long, zz As String * 1

If Indeks < 0 Then
    VirtKladzie = False
    Exit Function
End If

For i = 0 To 6
    If Left(VirtPlytka(i), 1) = "_" Then Uzyte(i) = True
Next i

xx = wx(Indeks)
yy = wy(Indeks)

If Kierunek = wPoziomy Then
    If xx + Len(Wyraz) > 16 Then
        VirtKladzie = False
        Exit Function
    End If
Else
    If yy + Len(Wyraz) > 16 Then
        VirtKladzie = False
        Exit Function
    End If
End If

If Kierunek = wPoziomy Then
    For i = 1 To Len(Wyraz)
        If (xx + i - 1) > 15 Then Exit For
        If PlanszaTAG(ciagiem(xx + i - 1, yy)) = 0 Then
            For j = 0 To 6
                znalazlem = False
                zz = Mid(Wyraz, i, 1)
                If Uzyte(j) = False And Left(VirtPlytka(j), 1) = zz And Right(VirtPlytka(j), 1) = "0" Then
                    Indeks = ciagiem(xx + i - 1, yy)
                    VirtPo³ó¿ Indeks, zz
                    Uzyte(j) = True
                    znalazlem = True
                    Exit For
                End If
            Next j
            If znalazlem = False Then
                For j = 0 To 6
                    If Uzyte(j) = False And Right(VirtPlytka(j), 1) = "1" Then
                        Indeks = ciagiem(xx + i - 1, yy)
                        VirtPo³ó¿ Indeks, Mid(Wyraz, i, 1), True
                        Uzyte(j) = True
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
Else
    For i = 1 To Len(Wyraz)
        If (yy + i - 1) > 15 Then Exit For
        If PlanszaTAG(ciagiem(xx, yy + i - 1)) = 0 Then
            For j = 0 To 6
                znalazlem = False
                zz = Mid(Wyraz, i, 1)
                If Uzyte(j) = False And Left(VirtPlytka(j), 1) = zz And Right(VirtPlytka(j), 1) = "0" Then
                    Indeks = ciagiem(xx, yy + i - 1)
                    VirtPo³ó¿ Indeks, zz
                    Uzyte(j) = True
                    znalazlem = True
                    Exit For
                End If
            Next j
            If znalazlem = False Then
                For j = 0 To 6
                    If Uzyte(j) = False And Right(VirtPlytka(j), 1) = "1" Then
                        Indeks = ciagiem(xx, yy + i - 1)
                        VirtPo³ó¿ Indeks, Mid(Wyraz, i, 1), True
                        Uzyte(j) = True
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
End If
    
VirtKladzie = True

End Function
Public Sub FillVirtPlytka(strWord As String)
Dim i As Long

For i = 0 To 6
    VirtPlytka(i) = Mid$(strWord, i, 1) & "0"
Next i

For i = 0 To 224
   PlanszaTAG(i) = Abs(CLng(Pole(i).Tag))
   PlanszaLICZ(i) = Abs(CLng(Pole(i).licz))
   If Pole(i).Caption = "" Then
      PlanszaCAPTION(i) = "_"
      Kaption = Kaption & "_"
   Else
      PlanszaCAPTION(i) = Pole(i).Caption
      Kaption = Kaption & Pole(i).Caption
   End If
   PlanszaWART(i) = Pole(i).Wartoœæ
   PlanszaRAZYL(i) = Pole(i).RazyL
   PlanszaRAZYS(i) = Pole(i).RazyS
   PlanszaBLANK(i) = Pole(i).Blank
Next i

End Sub

Private Sub VirtUklada()

Dim i As Long
Kaption = ""
For i = 0 To 6
    VirtPlytka(i) = plytka(i).Caption & CStr(Abs(CLng(plytka(i).Blank)))
Next i

For i = 0 To 224
   PlanszaTAG(i) = Abs(CLng(Pole(i).Tag))
   PlanszaLICZ(i) = Abs(CLng(Pole(i).licz))
   If Pole(i).Caption = "" Then
      PlanszaCAPTION(i) = "_"
      Kaption = Kaption & "_"
   Else
      PlanszaCAPTION(i) = Pole(i).Caption
      Kaption = Kaption & Pole(i).Caption
   End If
   PlanszaWART(i) = Pole(i).Wartoœæ
   PlanszaRAZYL(i) = Pole(i).RazyL
   PlanszaRAZYS(i) = Pole(i).RazyS
   PlanszaBLANK(i) = Pole(i).Blank
Next i

End Sub

Private Function VirtPo³ó¿(Numer As Long, znak As String, Optional Mydl As Boolean = False) As Boolean
Dim i As Long, j As Long
If Mydl = False Then
    For i = 0 To 6
        If Left(VirtPlytka(i), 1) = znak And Right(VirtPlytka(i), 1) = "0" Then
            VirtPlytka(i) = "_0"
            PlanszaTAG(Numer) = 1
            PlanszaLICZ(Numer) = 1
            PlanszaCAPTION(Numer) = znak
            PlanszaBLANK(Numer) = Mydl
            PlanszaWART(Numer) = Wart(Asc(znak))
            VirtLezy(IlePolLezy) = Numer
            IlePolLezy = IlePolLezy + 1
            VirtPo³ó¿ = True
            Exit Function
        End If
    Next i
Else
    For i = 0 To 6
        If Right(VirtPlytka(i), 1) = "1" Then
            VirtPlytka(i) = "_0"
            PlanszaTAG(Numer) = 1
            PlanszaLICZ(Numer) = 1
            PlanszaCAPTION(Numer) = znak
            PlanszaBLANK(Numer) = True
            PlanszaWART(Numer) = 0
            VirtLezy(IlePolLezy) = Numer
            IlePolLezy = IlePolLezy + 1
            VirtPo³ó¿ = True
            Exit Function
        End If
    Next i
End If

VirtPo³ó¿ = False

End Function
Private Function VirtAnuluj()

Dim a As Boolean, co³nt As Integer, i As Long, j As Long
Dim item As Variant, nmr As Long

For i = 0 To IlePolLezy - 1
    nmr = VirtLezy(i)
    For j = 0 To 6
        If Left(VirtPlytka(j), 1) = "_" Then
            VirtPlytka(j) = PlanszaCAPTION(nmr) & CStr(Abs(CLng(PlanszaBLANK(nmr))))
            Exit For
        End If
    Next j
    PlanszaTAG(nmr) = 0
    PlanszaLICZ(nmr) = 0
    PlanszaBLANK(nmr) = False
Next i

IlePolLezy = 0

End Function
Private Sub VirtLinieAll()
Dim i As Long, Litery As String, j As Long
Dim nn As Long, tc1 As String, czasik As Single
Dim plrmax As Long, Numlit As Long

plrmax = PLR(Ktory).MaxPsz
For i = 0 To 6
    If Left(VirtPlytka(i), 1) <> "_" Then
        If Right(VirtPlytka(i), 1) <> "1" Then
            Litery = Litery & Left(VirtPlytka(i), 1)
        Else
            Litery = Litery & Mydlo
        End If
    End If
Next i
Numlit = Len(Litery)
If PLR(Ktory).MaxPsz = 0 Then
    MojaBaza.ClearAllWyrazy
End If

For i = 1 To 15 Step 7
   czasik = Timer
   Call VirtObszukajLinie2(i, wPionowy, Litery, Numlit, plrmax)
   'tc1 = Format(CStr(Timer - czasik), "#0.#0")
   j = j + 1
   Bar.Value = (j * 100) \ 30
   DoEvents
   'czasik = Timer
   Call VirtObszukajLinie2(i, wPoziomy, Litery, Numlit, plrmax)
   'Sloweczka.AddItem CStr(i) & " : " & tc1 & "/" & Format(CStr(Timer - czasik), "#0.#0")
   j = j + 1
   Bar.Value = (j * 100) \ 30
   DoEvents
Next i

For nn = 1 To 8 Step 7
   For i = 1 + nn To 6 + nn
      czasik = Timer
      Call VirtObszukajLinie2(i, wPionowy, Litery, Numlit, plrmax)
      j = j + 1
      'tc1 = Format(CStr(Timer - czasik), "#0.#0")
      Bar.Value = (j * 100) \ 30
      DoEvents
      Call VirtObszukajLinie2(i, wPoziomy, Litery, Numlit, plrmax)
      'Sloweczka.AddItem CStr(i) & " : " & tc1 & "/" & Format(CStr(Timer - czasik), "#0.#0")
      j = j + 1
      Bar.Value = (j * 100) \ 30
      DoEvents
   Next i
Next nn
Me.MousePointer = 0

End Sub

Private Function VirtSprawdz(Kierunek As wKierunek) As Boolean

Dim YChMin As Long, XChMin As Long, YChMax As Long, XChMax As Long
Dim kZB As Long, i As Long
Dim j As Long, MinY As Long, MinX As Long, MaxX As Long, MaxY As Long
Dim Slowo(9) As String
Dim item As Variant, YMinW(9) As Long, YMaxW(9) As Long, XMinW(9) As Long, XMaxW(9) As Long
Dim NPola As Long, Kon As Boolean, pr As Long
Dim vFirst As Long, vLast As Long, vCount As Long

MinX = 16
MinY = 16
vCount = IlePolLezy - 1
vLast = VirtLezy(vCount)
vFirst = VirtLezy(0)

MaxX = wx(vLast)
MaxY = wy(vLast)
MinX = wx(vFirst)
MinY = wy(vFirst)


'*************** WYRAZY PIONOWE ************8

If Kierunek = wPionowy Or Kierunek = wSama Then
Kon = False
While MinY > 1 And MinY < 16 And Kon = False
    If PlanszaTAG(ciagiem(MinX, MinY - 1)) Then
        MinY = MinY - 1
    Else
        Kon = True
    End If
Wend
Kon = False
While MaxY < 15 And Kon = False
    If PlanszaTAG(ciagiem(MaxX, MaxY + 1)) Then
        MaxY = MaxY + 1
    Else
        Kon = True
    End If
Wend
    
    kZB = 1
   
    For i = MinY To MaxY
        If PlanszaLICZ(ciagiem(MinX, i)) Then
            XChMin = MinX
            XChMax = MaxX
2031        If XChMin > 1 Then
                If PlanszaTAG(ciagiem(XChMin - 1, i)) Then
                    XChMin = XChMin - 1
                    GoTo 2031
                End If
            End If
3031        If XChMax < 15 Then
                If PlanszaTAG(ciagiem(XChMax + 1, i)) Then
                    XChMax = XChMax + 1
                    GoTo 3031
                End If
            End If
            XMinW(kZB) = XChMin
            XMaxW(kZB) = XChMax
            YMinW(kZB) = i
            YMaxW(kZB) = i
            kZB = kZB + 1
        End If
    Next i
        
    For i = 1 To kZB
        If XMinW(i) < XMaxW(i) Then
            For j = XMinW(i) To XMaxW(i)
                NPola = ciagiem(j, YMinW(i))
                Slowo(i) = Slowo(i) & PlanszaCAPTION(NPola)
            Next j
        End If
        
        If Slowo(i) <> "" Then
         If SlownikCheck(Slowo(i)) Then
            VirtSprawdz = False
            
            Exit Function
         End If
      End If
      
    Next i
End If

'*************** WYRAZY POZIOME ************8

If Kierunek = wPoziomy Then

1222 If MinX > 1 Then
        If PlanszaTAG(ciagiem(MinX - 1, MinY)) Then
            MinX = MinX - 1
            GoTo 1222
        End If
    End If
        
1224 If MaxX < 15 Then
        If PlanszaTAG(ciagiem(MaxX + 1, MinY)) Then
            MaxX = MaxX + 1
            GoTo 1224
        End If
    End If
    
    kZB = 1
    For i = MinX To MaxX
        If PlanszaLICZ(ciagiem(i, MinY)) Then
            YChMin = MinY
            YChMax = MaxY
1202        If YChMin > 1 Then
                If PlanszaTAG(ciagiem(i, YChMin - 1)) Then
                     YChMin = YChMin - 1
                     GoTo 1202
                End If
            End If
1302        If YChMax < 15 Then
                If PlanszaTAG(ciagiem(i, YChMax + 1)) Then
                    YChMax = YChMax + 1
                    GoTo 1302
                End If
            End If
            XMinW(kZB) = i
            XMaxW(kZB) = i
            YMinW(kZB) = YChMin
            YMaxW(kZB) = YChMax
            kZB = kZB + 1
        End If
    Next i
    For i = 1 To kZB
        If YMinW(i) < YMaxW(i) Then
            For j = YMinW(i) To YMaxW(i)
                NPola = ciagiem(XMinW(i), j)
                Slowo(i) = Slowo(i) & PlanszaCAPTION(NPola)
            Next j
        End If
        
      If Slowo(i) <> "" Then
         If SlownikCheck(Slowo(i)) Then
            VirtSprawdz = False
            Exit Function
         End If
      End If

    Next i
End If
    
VirtSprawdz = True

End Function
Private Sub DefiniujMydlo(Numer As Long, znak As String)
Dim i As Long, nmr As Long
If JuzGramy = False Then Exit Sub
nmr = 1
For i = 0 To 6
    If plytka(i).Blank Then
        If nmr = Numer Then
            plytka(i).Po³ó¿ "_", True
            plytka(i).ZatwierdŸ 3
            plytka(i).Po³ó¿ znak, True
            plytka(i).ZatwierdŸ 3
            Exit For
        Else
            nmr = nmr + 1
        End If
    End If
Next i

End Sub
Private Sub ObszukajLinie2(Numer As Long, Warunek As String, Kierunek As wKierunek, MyWar As String)
Dim i As Long, Y As Long, X As Long, Wzor As String, WarLew As String, WarPra As String
Dim NPola As Long, ZAPYT As String
Dim Pos As Long, tbl As String, WLCPP As String, WPCPP As String
Dim Warki() As String, ttx As String, mR As cMojRecordset
On Error Resume Next
If MyWar = "" Then
   ReDim Warki(0)
Else
   Warki = Split(MyWar, ",")
End If

If Kierunek = wPionowy Then
    X = Numer
    For Y = 1 To 15
        Wzor = ""
        NPola = ciagiem(X, Y)
        If PlanszaTAG(NPola) = 0 Then
            Wzor = "_"
            WarLew = WzorkujCPP(NPola, wPionowy, 0, PlanszaTAG(0), Kaption)
            WarPra = WzorkujCPP(NPola, wPionowy, 1, PlanszaTAG(0), Kaption)
            If WarLew <> "0" Then Wzor = WarLew & Wzor
            If WarPra <> "0" Then Wzor = Wzor & WarPra
            If Wzor <> "_" Then
                Pos = InStr(1, Wzor, "_")
                tbl = "Sl" & CStr(Len(Wzor))
                If Warunek = "" Then
                    Set mR = MojaBaza.OpenMidTabela(tbl, Wzor, Pos, Warki)
                Else
                    Set mR = MojaBaza.OpenMidTabela(tbl, Wzor, Pos, Warki)
                End If
                If mR.NoMatch = True Then
                    Zakaz(NPola + 225) = 1
                Else
                  mR.MoveFirst
                     While mR.EOF = False
                        ttx = mR.Element()
                        MojaBaza.AddDoDo ttx, NPola, Nie(Kierunek)
                        mR.MoveNext
                    Wend
                End If
            End If
        End If
    Next Y
Else
    Y = Numer
    For X = 1 To 15
        Wzor = ""
        NPola = ciagiem(X, Y)
        If PlanszaTAG(NPola) = 0 Then
            Wzor = "_"
            WarLew = WzorkujCPP(NPola, wPoziomy, 0, PlanszaTAG(0), Kaption)
            WarPra = WzorkujCPP(NPola, wPoziomy, 1, PlanszaTAG(0), Kaption)
            If WarLew <> "0" Then Wzor = WarLew & Wzor
            If WarPra <> "0" Then Wzor = Wzor & WarPra
            If Wzor <> "_" Then
                Pos = InStr(1, Wzor, "_")
                tbl = "Sl" & CStr(Len(Wzor))
                
                If Warunek = "" Then
                     Set mR = MojaBaza.OpenMidTabela(tbl, Wzor, Pos, Warki)
                Else
                     Set mR = MojaBaza.OpenMidTabela(tbl, Wzor, Pos, Warki)
                End If
                
                If mR.NoMatch = True Then
                    Zakaz(NPola) = 1
                Else
                  mR.MoveFirst
                    While mR.EOF = False
                         ttx = mR.Element()
                         MojaBaza.AddDoDo ttx, NPola, Nie(Kierunek)
                        mR.MoveNext
                    Wend
                End If
            End If
        End If
    Next X
End If

Set mR = Nothing
End Sub


Private Function Nie(Kierunek As wKierunek) As wKierunek

If Kierunek = wPionowy Then
    Nie = wPoziomy
Else
    Nie = wPionowy
End If

End Function

Private Function ZnajdzSlowa2() As tDlugosci
Dim i As Long, j As Long, X As Long, Y As Long
Dim Sl As String, Pos As Long, ds2 As tDlugosci, MyWar As String
Dim SX As Long, SY As Long, Pionowe As String, Poziome As String, tbl As String
Dim Warunek As String, zpt As String, Literka As String, starcik As Long
Dim Slowo As String, z As String
Dim nmr As Long, Pkt As Long, NewStart As Long, Kier As wKierunek, ll As Long, LitPlr As Long
Dim m As Long, ZAPYT As String, MinL As Long, MaxL As Long, AllPkt As Tpunktacja
Dim czasik As Single

czasik = Timer
MojaBaza.ClearDoDo
MojaBaza.ClearAllWyrazy

For i = 0 To 449
   Zakaz(i) = 0
Next i

Call KompZaczyna(True)

ds2.MaxL = MojaBaza.ExtDlugosc(False)
ds2.MinL = MojaBaza.ExtDlugosc(True)

If ds2.MaxL = 0 Then ds2.MaxL = 2
If ds2.MinL = 0 Then ds2.MinL = 2
ZnajdzSlowa2 = ds2

For i = 0 To 6
    If Left(VirtPlytka(i), 1) <> "_" Then LitPlr = LitPlr + 1
Next i

For i = 0 To 6
    If Left(VirtPlytka(i), 1) <> "_" Then
        If Right(VirtPlytka(i), 1) = "1" Then
            Warunek = ""
            MyWar = ""
            Exit For
        Else
            Warunek = Warunek & Left(VirtPlytka(i), 1) & "','"
        End If
    End If
Next i

If Warunek <> "" Then
   MyWar = Replace(Left(Warunek, Len(Warunek) - 2), "'", "")
End If


For i = 1 To 15
    Call ObszukajLinie2(i, Warunek, wPionowy, MyWar)
    'status.Caption = i & " <> " & "Pion"
    DoEvents
    Call ObszukajLinie2(i, Warunek, wPoziomy, MyWar)
    'status.Caption = i & " <> " & "Poziom"
    DoEvents
Next i

End Function
Private Function CzySzukac(Numer As Long, Kierunek As wKierunek, Ile As Long) As Boolean

Dim X As Long, Y As Long, i As Long, NPola As Long

X = wx(Numer)
Y = wy(Numer)
If X < 1 Or Y < 1 Or X > 15 Or Y > 15 Then
     CzySzukac = False
     Exit Function
End If

If Kierunek = wPoziomy Then
    If X + Ile > 16 Then
        CzySzukac = False
        Exit Function
    End If
Else
    If Y + Ile > 16 Then
        CzySzukac = False
        Exit Function
    End If
End If

If Kierunek = wPionowy Then
    For i = Y To Y + Ile - 1
        If i > 0 And i < 16 Then
            NPola = ciagiem(X, i)
            If PlanszaTAG(NPola) > 0 Or Zakaz(NPola) > 0 Then
                CzySzukac = False
                Exit Function
            End If
        Else
            CzySzukac = False
            Exit Function
        End If
    Next i
    If Y > 1 Then
        If PlanszaTAG(ciagiem(X, Y - 1)) > 0 Then
            CzySzukac = False
            Exit Function
        End If
    End If
    If Y + Ile < 16 Then
        If PlanszaTAG(ciagiem(X, Y + Ile)) > 0 Then
            CzySzukac = False
            Exit Function
        End If
    End If
Else
    For i = X To X + Ile - 1
        If i > 0 And i < 16 Then
            NPola = ciagiem(i, Y)
            If PlanszaTAG(NPola) > 0 Or Zakaz(NPola + 225) > 0 Then
                CzySzukac = False
                Exit Function
            End If
        Else
            CzySzukac = False
            Exit Function
        End If
    Next i
    If X > 1 Then
        If PlanszaTAG(ciagiem(X - 1, Y)) > 0 Then
            CzySzukac = False
            Exit Function
        End If
    End If
    If X + Ile < 16 Then
        If PlanszaTAG(ciagiem(X + Ile, Y)) > 0 Then
            CzySzukac = False
            Exit Function
        End If
    End If

End If

CzySzukac = True

End Function


Private Sub T³o_DragDrop(Source As Control, X As Single, Y As Single)

Dim ss As Boolean, a As Integer, OldXX As Single, OldYY As Single, OldIndex As Long
Dim xx As Single, yy As Single, Indeks As Long, stary As Boolean
Dim tmpC As String, StCo As ColorConstants, kap As String, StCo1 As ColorConstants
Dim ba As Integer, co³nt As Integer, OldXXP As Long, OldYYP As Long, i As Long, j As Long
Dim ttp As String, nrob As Long, OldM As Boolean

If m_Klient Is Nothing Then
    If JuzGramy = False Then Exit Sub
    If PLR(Ktory).Komp Then Exit Sub
End If
'Zaslonka.Visible = False
xx = Int((X - px) / (bok + Prz)) + 1
yy = Int((Y - py) / (bok + Prz)) + 1
If xx > 15 Or yy > 15 Or xx < 1 Or yy < 1 Then Exit Sub

Indeks = ciagiem(xx, yy)

If Source Is T³o Then
    
    If Pole(Indeks).Tag = True Then
        
        OldXX = Int((OldX - px) / (bok + Prz)) + 1
        OldYY = Int((OldY - py) / (bok + Prz)) + 1
        If OldXX > 15 Or OldYY > 15 Or OldXX < 1 Or OldYY < 1 Then Exit Sub
        OldIndex = ciagiem(OldXX, OldYY)
        If Indeks <> OldIndex Then
            If Pole(Indeks).NumerObrazka = 1 And Pole(OldIndex).Tag And Pole(OldIndex).NumerObrazka = 1 Then
                tmpC = Pole(Indeks).Litera
                OldM = Pole(Indeks).Blank
            
                Call Pole(Indeks).PustePole
                Call Pole(Indeks).Po³ó¿(Pole(OldIndex).Litera, Pole(OldIndex).Blank)
            
                Call Pole(OldIndex).PustePole
                Call Pole(OldIndex).Po³ó¿(tmpC, OldM)
            End If
            Czyt = False
            EfektujTu
            
        End If
        Exit Sub
    Else
        OldXX = Int((OldX - px) / (bok + Prz)) + 1
        OldYY = Int((OldY - py) / (bok + Prz)) + 1
        If OldXX > 15 Or OldYY > 15 Or OldXX < 1 Or OldYY < 1 Then Exit Sub
        OldIndex = ciagiem(OldXX, OldYY)
        If Pole(OldIndex).NumerObrazka = 1 And Pole(OldIndex).Tag Then
            If CzyPionowo(Indeks, Le¿¹ce, OldIndex) Or CzyPoziomo(Indeks, Le¿¹ce, OldIndex) Then
                Call Pole(Indeks).PustePole
                Pole(Indeks).Po³ó¿ Pole(OldIndex).Litera, Pole(OldIndex).Blank
                Pole(OldIndex).PustePole
                Le¿¹ce.Remove "K" & CStr(OldIndex)
                Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
                EfektujTu
                Exit Sub
            End If
        End If
    End If
End If

If Not Source Is Podstawka Then Exit Sub

'******* CZÊŒÆ ZASADNICZA ( K£ADZENIE KLOCKA NA PLANSZY ) *******

If Dlug < 0 Then Dlug = 0

If Le¿¹ce.Count Then
    If CzyPoziomo(Indeks, Le¿¹ce) = False Then
        If CzyPionowo(Indeks, Le¿¹ce) = False Then
            Exit Sub
        Else
            Kierunek = wPionowy
        End If
    Else
        Kierunek = wPoziomy
    End If
Else
    Kierunek = wSama
End If

If Source Is Podstawka Then
    OldXXP = Int((OldXP) / (bok + Prz)) + 1
    OldYYP = Int((OldYP) / (bok + Prz)) + 1
    If OldXXP > 7 Or OldYYP > 7 Or OldXXP < 1 Or OldYYP < 1 Then
        CzytP = False
        Exit Sub
    End If
    
    OldIndex = ciagiem(OldXXP, OldYYP)
    If plytka(OldIndex).Litera = Mydlo Then
        InfoBox "Musisz zdefiniowaæ blanka !", False, False
        Exit Sub
    End If
    If Pole(Indeks).Tag = False And plytka(OldIndex).Caption <> "_" Then
        Call Pole(Indeks).Po³ó¿(plytka(OldIndex).Litera, plytka(OldIndex).Blank)
        Le¿¹ce.Add Indeks, "K" & CStr(Indeks)
        Call plytka(OldIndex).Po³ó¿("_", False)
        Call plytka(OldIndex).ZatwierdŸ(3)
        Dlug = Dlug + 1
        EfektujTu
    End If
    
    If Pole(Indeks).Tag And Pole(Indeks).licz And plytka(OldIndex).Caption <> "_" Then
        ttp = Pole(Indeks).Caption
        nrob = plytka(OldIndex).NumerObrazka
        OldM = Pole(Indeks).Blank
        
        Pole(Indeks).PustePole
        Pole(Indeks).Po³ó¿ plytka(OldIndex).Litera, plytka(OldIndex).Blank
        
        plytka(OldIndex).PustePole
        plytka(OldIndex).Po³ó¿ "_", False
        plytka(OldIndex).ZatwierdŸ 3
        plytka(OldIndex).Po³ó¿ ttp, OldM
        plytka(OldIndex).ZatwierdŸ nrob
    End If
    CzytP = False
    EfektujTu
End If
If Dlug = 1 Then
    Kierunek = wSama
End If

End Sub


Private Function StartZacz(Wyraz As String) As Long
Dim LW As Long, i As Long, maxP As Long, MaxL As Long

LW = Len(Wyraz)

Select Case LW
    
Case 5
    If Wart(Asc(Left(Wyraz, 1))) >= Wart(Asc(Right(Wyraz, 1))) Then
        StartZacz = 108
    Else
        StartZacz = 112
    End If
    Exit Function
Case 2, 3
    StartZacz = 111
    Exit Function
Case 4
    StartZacz = 110
    Exit Function
Case 6
    maxP = 0
    For i = 1 To 2
        If Wart(Asc(Mid(Wyraz, i, 1))) > maxP Then
            maxP = Wart(Asc(Mid(Wyraz, i, 1)))
            MaxL = i
        End If
    Next i
    For i = 5 To 6
        If Wart(Asc(Mid(Wyraz, i, 1))) > maxP Then
            maxP = Wart(Asc(Mid(Wyraz, i, 1)))
            MaxL = i
        End If
    Next i
    Select Case MaxL
        Case 1: StartZacz = 108
        Case 2: StartZacz = 107
        Case 5: StartZacz = 112
        Case 6: StartZacz = 111
    End Select
Case 7
    maxP = 0
    For i = 1 To 3
        If Wart(Asc(Mid(Wyraz, i, 1))) > maxP Then
            maxP = Wart(Asc(Mid(Wyraz, i, 1)))
            MaxL = i
        End If
    Next i
    For i = 5 To 7
        If Wart(Asc(Mid(Wyraz, i, 1))) > maxP Then
            maxP = Wart(Asc(Mid(Wyraz, i, 1)))
            MaxL = i
        End If
    Next i
    Select Case MaxL
        Case 1: StartZacz = 108
        Case 2: StartZacz = 107
        Case 3: StartZacz = 106
        Case 5: StartZacz = 112
        Case 6: StartZacz = 111
        Case 7: StartZacz = 110
    End Select
Case Else: StartZacz = 112

End Select

End Function

Private Sub DodajSlowa3(MinL As Long, MaxL As Long)
Dim i As Long, czasik As Single
'dim r As New ADODB.Recordset
Dim k As Long, ZAPYT As String, Warunek As String, Z3 As String
Dim n As Long, j As Long, Ns As Long, Mmm As Long, Z2 As String, Pkt As Tpunktacja
Dim w As String, Wyraz As String, mmm2 As Long
'Dim R0 As New ADODB.Recordset, R3 As New ADODB.Recordset
Dim LokZakaz(0 To 224, 1 To 2, 1 To 7) As Boolean
Dim mR As cMojRecordset, mR2 As cMojRecordset, mR3 As cMojRecordset, MojWarunek() As String, MWCount As Long
czasik = Timer
'r.ActiveConnection = ADObaza
'R0.ActiveConnection = ADObaza
'R3.ActiveConnection = ADObaza
'R0.CursorLocation = adUseClient

For k = 1 To 2
   Set mR = MojaBaza.OpenDoDoNPola(k, True)
   If mR.NoMatch = False Then
      mR.MoveFirst
      While mR.EOF = False
         i = CLng(mR.Element())
         'status.Caption = CStr(i)
         DoEvents
         If Zakaz(i + (225 * (k - 1))) = 0 And PlanszaTAG(i) = 0 Then
            Set mR2 = MojaBaza.OpenDoDoLitera(i, k)
            If mR2.NoMatch = False Then
               ReDim MojWarunek(0)
               MWCount = 0
               mR2.MoveFirst
               While Not mR2.EOF
                  w = mR2.Element
                  MWCount = MWCount + 1
                  ReDim Preserve MojWarunek(MWCount)
                  MojWarunek(MWCount) = w
                  mR2.MoveNext
               Wend
                  For n = MaxL To MinL Step -1
                     For j = 1 To n
                        Ns = JakiStart3(i, k, j, n, PlanszaTAG(0), Zakaz(0))
                        If Ns > -1 Then
                           If CzySzukac2(Ns, k, n, PlanszaTAG(0), Zakaz(0)) = 1 And LokZakaz(Ns, k, n) = False Then
                              LokZakaz(Ns, k, n) = True
                              If IleLiter > 0 Then
                                 Mmm = MaxWynik(KPolaCPP(0), SortLiterki(0), n, Ns, n, k, PlanszaRAZYS(0), PlanszaTAG(0), PlanszaWART(0), PlanszaRAZYL(0))
                              Else
                                 Mmm = -1
                              End If
                                 
                              If Mmm >= MaxPkt Or Mmm = -1 Then
                                 Set mR3 = MojaBaza.OpenWlasneWyraz(n, j, MojWarunek)
                                 If mR3.EOF = False Then
                                    mR3.MoveFirst
                                    While Not mR3.EOF
                                       Wyraz = mR3.Element
                                       Pkt = Punkty(Wyraz, Ns, k)
                                       If IleLiter > 0 Then
                                          If Pkt.Wsp >= MaxPkt And Pkt.Wsp > 0 Then
                                             If VirtSprawdz(k) = True Then
                                                MaxPkt = Pkt.Wsp
                                                MojaBaza.InsertToAllWyrazy Wyraz, Ns, MaxPkt, k, Pkt.Wartosc, Pkt.IleLiter
                                             End If
                                          End If
                                       Else
                                          If n >= Policz(Ktory) Then
                                             If Pkt.Punkty >= MinPktToWiN(Ktory) Then
                                                If VirtSprawdz(k) = True Then
                                                   MaxPkt = Pkt.Punkty + 500
                                                   MojaBaza.InsertToAllWyrazy Wyraz, Ns, MaxPkt, k, Pkt.Wartosc, Pkt.IleLiter
                                                End If
                                             Else
                                                If Pkt.Wsp >= MaxPkt And Pkt.Wsp > 0 Then
                                                   If VirtSprawdz(k) = True Then
                                                      MaxPkt = Pkt.Wsp
                                                      MojaBaza.InsertToAllWyrazy Wyraz, Ns, MaxPkt, k, Pkt.Wartosc, Pkt.IleLiter
                                                   End If
                                                End If
                                             End If
                                          Else
                                             If Pkt.Wsp >= MaxPkt And Pkt.Wsp > 0 Then
                                                If VirtSprawdz(k) = True Then
                                                   MaxPkt = Pkt.Wsp
                                                   MojaBaza.InsertToAllWyrazy Wyraz, Ns, MaxPkt, k, Pkt.Wartosc, Pkt.IleLiter
                                                End If
                                             End If
                                          End If
                                       End If
                                       VirtAnuluj
                                       mR3.MoveNext
                                    Wend
                                 End If
                              End If
                           End If
                        End If
                     Next j
                  Next n
               End If
            End If
           mR.MoveNext
        Wend
    End If
Next k

Set mR = Nothing
Set mR2 = Nothing
Sloweczka.AddItem "Dopisz: " & Format(CStr(Timer - czasik), "##0.#0")

End Sub

Private Sub ButOnOff(Stan As Boolean)
Dim i As Long
anulujemy.Enabled = Not Stan
Solve.Enabled = Not Stan
wymiana.Enabled = Not Stan

Kolejka.Enabled = Not Stan
For i = 1 To 4
   AkcjeSub(i).Enabled = Not Stan
Next i

End Sub

Private Sub VirtObszukajLinie2(Numer As Long, Kierunek As wKierunek, Litery As String, Numlit As Long, plrmax As Long)

Dim pxx As Long, X As Long, Y As Long, Pl As Long, kl As Long, i As Long
Dim j As Long, NPola As Long, Wzorzec As String, LW As Long, Tabela As String, Warunio As String
Dim Star As Long, Mmm As Long, ZAPYT As String
Dim mWyraz As cFullWyraz
Dim ll As Long, AllPkt As Tpunktacja
Dim Kierunek2 As wKierunek, mmm2 As Long, PlusJeden As Boolean
Dim mR As cMojRecordset
'If Numer = 13 And Kierunek = wPionowy Then Stop
Me.MousePointer = 11

'*****POZIOME*****

If Kierunek = wPoziomy Then
    
    Y = Numer
    Pl = 0
    kl = 0
    
    For i = 1 To 15
        If PlanszaTAG(ciagiem(i, Y)) > 0 Then
            Pl = i
            Exit For
        End If
    Next i
    If Pl = 0 Then Exit Sub
    For i = 15 To 1 Step -1
        If PlanszaTAG(ciagiem(i, Y)) > 0 Then
            kl = i
            Exit For
        End If
    Next i
    
   MojaBaza.ClearWyrazy
   
   For j = Numlit To 1 Step -1
      For i = kl + 1 To Pl - j Step -1
         If i > 0 And i < 16 Then
            NPola = ciagiem(i, Y)
            If PlanszaTAG(NPola) = 0 Then
               Wzorzec = VirtUstalWzorzec2(NPola, wPoziomy, j, plrmax, PlanszaTAG(0), Kaption, Zakaz(0))
               If Wzorzec <> "0" Then
                  LW = Len(Wzorzec)
                  If LW > j Then
                     If Len(Replace(Wzorzec, "_", "")) = 1 Then
                        Tabela = "Plus1"
                        PlusJeden = True
                     Else
                        Tabela = "Sl" & CStr(LW)
                        PlusJeden = False
                     End If
                     Star = ciagiem(i - VirtMinusST2(NPola, wPoziomy, PlanszaTAG(0)), Y)
                     If IleLiter > 0 Then
                        Mmm = MaxWynik(KPolaCPP(0), SortLiterki(0), LW, Star, j, Kierunek, PlanszaRAZYS(0), PlanszaTAG(0), PlanszaWART(0), PlanszaRAZYL(0))
                     Else
                        Mmm = -1
                     End If
                           
                     If Mmm >= MaxPkt Or Mmm = -1 Then
                        MojaBaza.InsertToWyrazy2 LW, Wzorzec, Kierunek, Star, PlusJeden, KtoSzukaMB(j, Litery, Replace(Wzorzec, "_", ""))
                        DoEvents
                     End If
                  End If
               End If
            End If
         End If
      Next i
   Next j
Else

'*****Pionowe*****
    
    X = Numer
    Pl = 0
    kl = 0
    
    For i = 1 To 15
        If PlanszaTAG(ciagiem(X, i)) > 0 Then
            Pl = i
            Exit For
        End If
    Next i
    
    If Pl = 0 Then Exit Sub
    
    For i = 15 To 1 Step -1
        If PlanszaTAG(ciagiem(X, i)) > 0 Then
            kl = i
            Exit For
        End If
    Next i
    
   MojaBaza.ClearWyrazy
   
    For j = Numlit To 1 Step -1
        For i = kl + 1 To Pl - j Step -1
            If i > 0 And i < 16 Then
                NPola = ciagiem(X, i)
                If PlanszaTAG(NPola) = 0 Then
                    Wzorzec = VirtUstalWzorzec2(NPola, wPionowy, j, plrmax, PlanszaTAG(0), Kaption, Zakaz(0))
                    If Wzorzec <> "0" Then
                        LW = Len(Wzorzec)
                        If LW > j Then
                           If Len(Replace(Wzorzec, "_", "")) = 1 Then
                              PlusJeden = True
                           Else
                              PlusJeden = False
                           End If
                           Star = ciagiem(X, i - VirtMinusST2(NPola, wPionowy, PlanszaTAG(0)))
                           If IleLiter > 0 Then
                              Mmm = MaxWynik(KPolaCPP(0), SortLiterki(0), LW, Star, j, Kierunek, PlanszaRAZYS(0), PlanszaTAG(0), PlanszaWART(0), PlanszaRAZYL(0))
                           Else
                              Mmm = -1
                           End If
                           
                           If Mmm >= MaxPkt Or Mmm = -1 Then
                              MojaBaza.InsertToWyrazy2 LW, Wzorzec, Kierunek, Star, PlusJeden, KtoSzukaMB(j, Litery, Replace(Wzorzec, "_", ""))
                              DoEvents
                           End If
                        End If
                    End If
                End If
            End If
        Next i
    Next j
End If

If MojaBaza.Wyrazy.RecordCount > 0 Then
   For i = 1 To MojaBaza.Wyrazy.RecordCount
      Set mWyraz = MojaBaza.Wyrazy.Element(i)
      Kierunek2 = mWyraz.Kierunek
      Wzorzec = mWyraz.Wyraz
      NPola = mWyraz.Start
      ll = mWyraz.IleLiter
      
      AllPkt = Punkty(Wzorzec, NPola, Kierunek2)
        
      If IleLiter > 0 Then
         If AllPkt.Wsp >= MaxPkt And AllPkt.Wsp > 0 Then
            If VirtSprawdz(Kierunek) = True Then
               MaxPkt = AllPkt.Wsp
               MojaBaza.InsertToAllWyrazy Wzorzec, NPola, MaxPkt, Kierunek, AllPkt.Wartosc, AllPkt.IleLiter
            End If
         End If
      Else
         If AllPkt.IleLiter >= Policz(Ktory) Then
            If AllPkt.Punkty >= MinPktToWiN(Ktory) Then
               If VirtSprawdz(Kierunek) = True Then
                  If MaxPkt <= AllPkt.Punkty + 500 Then
                     MaxPkt = AllPkt.Punkty + 500
                     MojaBaza.InsertToAllWyrazy Wzorzec, NPola, MaxPkt, Kierunek, AllPkt.Wartosc, AllPkt.IleLiter
                  End If
               End If
            Else
               If AllPkt.Wsp + CLng(IleMaReszta(Ktory) * 1.3) >= MaxPkt And AllPkt.Wsp > 0 Then
                  If VirtSprawdz(Kierunek) = True Then
                     MaxPkt = AllPkt.Wsp + CLng(IleMaReszta(Ktory) * 1.6)
                     MojaBaza.InsertToAllWyrazy Wzorzec, NPola, MaxPkt, Kierunek, AllPkt.Wartosc, AllPkt.IleLiter
                  End If
               End If
            End If
         Else
            If AllPkt.Wsp >= MaxPkt And AllPkt.Wsp > 0 Then
               If VirtSprawdz(Kierunek) = True Then
                  MaxPkt = AllPkt.Wsp
                  MojaBaza.InsertToAllWyrazy Wzorzec, NPola, MaxPkt, Kierunek, AllPkt.Wartosc, AllPkt.IleLiter
               End If
            End If
         End If
      End If
      VirtAnuluj
   Next i
End If

Set mR = Nothing
Me.MousePointer = 0

End Sub

Private Function IleMaReszta(Kto As Long) As Long
Dim i As Long, il As Long

For i = 1 To IleGraczy
    If i <> Kto Then
        il = il + Policz(i)
    End If
Next i
IleMaReszta = il

End Function

Private Function MaxCzas() As Single
Dim item As cCzas, maxC As Single
If Not Czasy Is Nothing Then
For Each item In Czasy
      If item.Czas > maxC Then
         maxC = item.Czas
      End If
   Next item
End If
MaxCzas = maxC

End Function

Private Function SredniCzas() As Single
Dim item As cCzas, maxC As Single
If Not Czasy Is Nothing And Czasy.Count > 0 Then
   For Each item In Czasy
      maxC = item.Czas + maxC
   Next item
   SredniCzas = CSng(maxC / Czasy.Count)
Else
   SredniCzas = 0
End If

End Function

Private Function SumaCzas() As Single

Dim item As cCzas, maxC As Single
If Not Czasy Is Nothing Then
   For Each item In Czasy
      maxC = item.Czas + maxC
   Next item
End If

SumaCzas = maxC

End Function
Public Sub SzukajPlus(IlePlus As Long)
Dim dc As Single
Dim i As Long, Ile As Long, Lit As String, z As String * 1
Dim Czas As Single
dc = Timer

   For i = 0 To 6
      z = plytka(i).Caption
      If z <> "_" Then
         Ile = Ile + 1
         If plytka(i).Blank Then
            Lit = Lit & "_"
         Else
            Lit = Lit & z
         End If
      End If
   Next i

MojaBaza.ClearPlus1
For i = 1 To Ile
   MojaBaza.InsertToPlus1 i + IlePlus, KtoSzukaMB(i, Lit)
Next i

Sloweczka.AddItem "" & "Plus1" & ": " & CStr(MojaBaza.Plus1Count) & " / " & Format(CStr(Timer - dc), "#0.#0")
DoEvents


End Sub
Public Function CzyBlank() As Boolean

Dim i  As Long
If PLR(Ktory).Komp Then
   CzyBlank = False
   Exit Function
End If
For i = 0 To 6
   If plytka(i).Blank Then
      CzyBlank = True
      Exit Function
   End If
Next i
CzyBlank = False

End Function


Public Sub KLiterujTAB()

Dim i As Long, n As Long

For i = 0 To 6
   If plytka(i).Caption <> "_" Then
      KLiteryTAB(n) = plytka(i).Wartoœæ
      n = n + 1
   End If
Next i
QuickSort KLiteryTAB, 0, n - 1, True

End Sub

Public Sub QuickSort(tbl() As Long, X As Long, Y As Long, Optional Descending As Boolean = False)

Dim i As Long, j As Long, v As Long, Temp As Long

If Descending = False Then

   i = X
   j = Y
   v = tbl((X + Y) \ 2)
   While i <= j
      While tbl(i) < v
         i = i + 1
      Wend
      While v < tbl(j)
         j = j - 1
      Wend
      If i <= j Then
         Temp = tbl(i)
         tbl(i) = tbl(j)
         tbl(j) = Temp
         i = i + 1
         j = j - 1
      End If
   Wend

   If X < j Then QuickSort tbl, X, j
   If i < Y Then QuickSort tbl, i, Y
Else

   i = X
   j = Y
   v = tbl((X + Y) \ 2)
   While i <= j
      While tbl(i) > v
         i = i + 1
      Wend
      While v > tbl(j)
         j = j - 1
      Wend
      If i <= j Then
         Temp = tbl(i)
         tbl(i) = tbl(j)
         tbl(j) = Temp
         i = i + 1
         j = j - 1
      End If
   Wend

   If X < j Then QuickSort tbl, X, j, True
   If i < Y Then QuickSort tbl, i, Y, True
End If

End Sub

Private Function SortLiteryTAB() As Long()
Dim i As Long, Bylo(7) As Boolean, max As Long, w As Long, Ile As Long, w2 As String
Dim nmr As Long, IleM As Long, SD(0 To 6) As Long

For i = 0 To 6
   If Left(VirtPlytka(i), 1) <> "_" Then
      If Right(VirtPlytka(i), 1) = "1" Then
         SD(w) = 0
      Else
         SD(w) = Wart(Asc(Left(VirtPlytka(i), 1)))
      End If
      w = w + 1
   End If
Next i
   
QuickSort SD, 0, w - 1, True

SortLiteryTAB = SD()

End Function

Public Sub Zaznacz()
Dim i As Long
For i = 0 To 6
   If plytka(i).Blank = False Then
      Call plytka(i).ZatwierdŸ(4)
      DoEvents
   End If
Next i
End Sub

Public Sub Importuj()
Dim znak As String, Plik As String, i As Long, t As Long, e1 As String, e2 As String, e3 As String
Dim pustka As String, j As Long, ee As String, item As Ruchy, SaveIR As Long
Dim SaveŒrednia As Single, IR As Long, Slowo As String, SaveIW As Long, hh As Long
Dim IlePol As Long, j1 As String, j2 As String, j3 As String, j4 As String
Dim warto As Long, znakk As String * 1, hj1 As Long, komput As String, e4 As String
Dim KG As String, Mpsz As Long, Plitery As String

Zegar.Enabled = False
Timer1.Enabled = False

dysk.FileName = "*.lsw"
dysk.InitDir = App.Path & "\save"
On Error GoTo KoniecOdczytaj
dysk.Filter = "Pliki '£ów S³ów' (*.lsw)|*.lsw"
dysk.FilterIndex = 1
dysk.ShowOpen
On Error GoTo 0

MenuFont.Enabled = True
Podstawka.Visible = True

For i = 1 To 4
    Set PLR(i) = New GR
Next i

52 hh = FreeFile
Open App.Path & "\save\tmp.tmp" For Binary As hh
Close hh
Kill (App.Path & "\save\tmp.tmp")

Dim e(7) As String, kkon As Long, hj As Long

On Error GoTo 0

For i = 0 To 99
    Wolne(i) = False
Next i
j = Le¿¹ce.Count
If j > 0 Then
    For i = j To 1 Step -1
        Le¿¹ce.Remove i
    Next i
End If
Plik = dysk.FileName
Call DeSzyfruj(Plik)

For i = 0 To 224
    Pole(i).PustePole
Next i
For i = 0 To 6
    plytka(i).PustePole
    plytka(i).Po³ó¿ "_", False
    plytka(i).ZatwierdŸ 3
Next i

hj = FreeFile

Open App.Path & "\save\tmp.tmp" For Input As hj
Dim wer As String
Input #hj, wer
Close #hj

If wer = "LW2" Then
   Odczytaj2
   Exit Sub
End If

hj = FreeFile
Open App.Path & "\save\tmp.tmp" For Input As hj
Input #hj, IlePol, IleGraczy, IleLiter, Ktory, Omin, FMax, RMax, Zero, NumerW, KG, j1, j2, j3, j4
Ktory = Ktory + 1

If KG = "0" Then
    KoniecGry = False
Else
    KoniecGry = True
End If

'*******************************8
IleWymian = 3

Set Jêzyk = New tJezyk
Jêzyk.Klucz = j1
Jêzyk.Nazwa = j2
Jêzyk.Plik = j3
Jêzyk.S³ownik = j4
Me.Caption = "£ów S³ów - " & Jêzyk.Nazwa

hj1 = FreeFile
Open App.Path & "\" & Jêzyk.Plik For Input As hj1
Input #hj1, Plitery
i = 0

While Not EOF(hj1)
    Input #hj1, znakk, warto
    Wart(Asc(znakk)) = warto
    ReDim Preserve PLit(i + 1)
    PLit(i) = znakk
    i = i + 1
Wend
For i = 0 To 97
    Worek(i) = Mid(Plitery, i + 1, 1)
Next i
Close hj1

'On Error GoTo BladImportuj

For i = 1 To IleGraczy
    Input #hj, SaveIR, SaveIW
    PLR(i).IleRuchów = SaveIR
    PLR(i).iw = SaveIW
    For j = 1 To SaveIR
        Input #hj, IR, Slowo
        Set item = New Ruchy
        item.Punkty = IR
        item.S³owa = Slowo
        PLR(i).MyKolek.Add item
    Next j
Next i
If IlePol Then
    For t = 1 To IlePol
        Input #hj, i, znak, kkon
        Kolkon(i) = kkon
        Call Pole(i).Po³ó¿(Left(znak, 1), CBool(Abs(CLng(Right(znak, 1)))))
        Call Pole(i).ZatwierdŸ
   Next t
End If

For t = 1 To IleGraczy
    Input #hj, e1, e2, e3, komput, Mpsz, e(0), e(1), e(2), e(3), e(4), e(5), e(6), pustka
        PLR(t).imie = e1
        If komput = "0" Then
            PLR(t).Komp = False
        Else
            PLR(t).Komp = True
        End If
        PLR(t).MaxPsz = Mpsz
        PLR(t).wynik = CLng(e2)
        PLR(t).CzasCa³kowity = e3
        For j = 0 To 6
            Stojak(t, j) = e(j)
        Next j
Next t
For i = 1 To 100 - IleLiter
    Input #hj, ee
    Wolne(ee) = True
Next i
For i = 0 To 3
    player(i).Visible = False
    wynik(i).Visible = False
    Kolorek(i).Visible = False
    cas(i).Visible = False
    player(i).FontBold = False
    cas(i).FontBold = False
    wynik(i).FontBold = False
Next i
For i = 1 To IleGraczy
    player(i - 1).Visible = True
    Kolorek(i - 1).Left = 520
    player(i - 1).Left = Kolorek(i - 1).Left + Kolorek(i - 1).Width + 2
    player(i - 1).Top = 20 + ((player(i - 1).Height + 2) * (i))
    Kolorek(i - 1).Top = player(i - 1).Top
    cas(i - 1).Top = player(i - 1).Top
    wynik(i - 1).Top = player(i - 1).Top
    player(i - 1).Caption = PLR(i).imie
    wynik(i - 1).Caption = PLR(i).wynik
    wynik(i - 1).Visible = True
    Kolorek(i - 1).Visible = True
    cas(i - 1).Visible = True
    player(i - 1).FontBold = False
    Kolorek(i - 1).Picture = Obrazki.ListImages.item(i + 3).Picture
    wynik(i - 1).FontBold = False
    cas(i - 1).FontBold = False
    cas(i - 1).Caption = FormaCzasu(FMax - PLR(i).CzasCa³kowity)
Next i
CR.Visible = True
player(Ktory - 1).FontBold = True
wynik(Ktory - 1).FontBold = True
cas(Ktory - 1).FontBold = True
IleLitCap.Caption = IleLiter
aktGracz.Caption = "Uk³ada " & PLR(Ktory).imie
CR.Caption = RMax - Zero
JuzGramy = True

For i = 0 To 6
    If Right(Stojak(Ktory, i), 1) = "1" Then
        Call plytka(i).Po³ó¿(Left(Stojak(Ktory, i), 1), True)
        Call plytka(i).ZatwierdŸ(3)
    Else
        Call plytka(i).Po³ó¿(Left(Stojak(Ktory, i), 1), False)
        Call plytka(i).ZatwierdŸ(3)
    End If
Next i
Close hj

Solve.Enabled = True
anulujemy.Enabled = True
wymiana.Enabled = True
AkcjeSub(3).Enabled = True
AkcjeSub(1).Enabled = True
Kolejka.Enabled = True

Kill (App.Path & "\save\tmp.tmp")
If KoniecGry Then
    enduj
Else
    If IleLiter = 0 Then
        Timer1.Enabled = True
    End If
End If


On Error GoTo 0
Call Nazywaj

Call CzytajHS

Pokazuj True
If PLR(Ktory).Komp Then
   Podstawka.Visible = Not (MenuShowLit.Checked)
End If

Set Czasy = Nothing
Set Czasy = New Collection
DefMyd1.Visible = CzyBlank And (Not PLR(Ktory).Komp)
Label1.Visible = DefMyd1.Visible
ButOnOff (PLR(Ktory).Komp)
DoEvents
   
graslow.Caption = "£ów S³ów - " & Jêzyk.Nazwa
graslow.DefMyd1.Clear
   
For i = 0 To UBound(PLit) - 2
   graslow.DefMyd1.AddItem PLit(i)
Next i
   
graslow.DefMyd1.ListIndex = 0

DoEvents

If PLR(Ktory).Komp Then
   KompKladzie
End If
Zegar.Enabled = True

Close hj

Exit Sub

MkDir (App.Path & "\save")
GoTo 52
KoniecOdczytaj:
Exit Sub

BladImportuj:
Close hj
InfoBox "Wskazany plik jest nieprawid³owy.", False, False

End Sub

Public Sub Odczytaj2()
Dim hj As Long, wersja As String, ccc As Long, crr As Long, j5 As Long
Dim znak As String, Plik As String, i As Long, t As Long, e1 As String, e2 As String, e3 As String
Dim pustka As String, j As Long, ee As String, item As Ruchy, SaveIR As Long
Dim SaveŒrednia As Single, IR As Long, Slowo As String, SaveIW As Long, hh As Long
Dim IlePol As Long, j1 As String, j2 As String, j3 As String, j4 As String
Dim warto As Long, znakk As String * 1, hj1 As Long, komput As String, e4 As String
Dim KG As String, Mpsz As Long, Plitery As String, kkon As Long
Dim e(6) As String, j6 As Long
hj = FreeFile
522 Open App.Path & "\save\tmp.tmp" For Input As hj
Input #hj, wersja

Input #hj, IlePol, IleGraczy, IleLiter, Ktory, Omin, FMax, RMax, Zero, NumerW, KG, j1, j2, j3, j4, crr, ccc, j5, j6
Ktory = Ktory + 1

If KG = "0" Then
    KoniecGry = False
Else
    KoniecGry = True
End If

If ccc = 0 Then
   CzasCTak = False
Else
   CzasCTak = True
End If

If crr = 0 Then
   CzasRTak = False
Else
   CzasRTak = True
End If

'*******************************8
IleWymian = j5
Demo = j6
Set Jêzyk = New tJezyk
Jêzyk.Klucz = j1
Jêzyk.Nazwa = j2
Jêzyk.Plik = j3
Jêzyk.S³ownik = j4
Me.Caption = "£ów S³ów - " & Jêzyk.Nazwa

hj1 = FreeFile
Open App.Path & "\" & Jêzyk.Plik For Input As hj1
Input #hj1, Plitery
i = 0

While Not EOF(hj1)
    Input #hj1, znakk, warto
    Wart(Asc(znakk)) = warto
    ReDim Preserve PLit(i + 1)
    PLit(i) = znakk
    i = i + 1
Wend
For i = 0 To 97
    Worek(i) = Mid(Plitery, i + 1, 1)
Next i
Close hj1
Me.MousePointer = 11
Set MojaBaza = New cBazaDanych
MojaBaza.Inicjuj App.Path & "\bazy\" & Left(Jêzyk.Klucz, 2)
Me.MousePointer = 0
For i = 1 To IleGraczy
    Input #hj, SaveIR, SaveIW
    PLR(i).IleRuchów = SaveIR
    PLR(i).iw = SaveIW
    For j = 1 To SaveIR
        Input #hj, IR, Slowo
        Set item = New Ruchy
        item.Punkty = IR
        item.S³owa = Slowo
        PLR(i).MyKolek.Add item
    Next j
Next i
If IlePol Then
    For t = 1 To IlePol
        Input #hj, i, znak, kkon
        Kolkon(i) = kkon
        Call Pole(i).Po³ó¿(Left(znak, 1), CBool(Abs(CLng(Right(znak, 1)))))
        Call Pole(i).ZatwierdŸ
   Next t
End If

For t = 1 To IleGraczy
    Input #hj, e1, e2, e3, komput, Mpsz, j5, e(0), e(1), e(2), e(3), e(4), e(5), e(6), pustka
        PLR(t).imie = e1
        If komput = "0" Then
            PLR(t).Komp = False
        Else
            PLR(t).Komp = True
        End If
        PLR(t).IloscWymian = j5
        PLR(t).MaxPsz = Mpsz
        PLR(t).wynik = CLng(e2)
        PLR(t).CzasCa³kowity = e3
        For j = 0 To 6
            Stojak(t, j) = e(j)
        Next j
Next t
For i = 1 To 100 - IleLiter
    Input #hj, ee
    Wolne(ee) = True
Next i
For i = 0 To 3
    player(i).Visible = False
    wynik(i).Visible = False
    Kolorek(i).Visible = False
    cas(i).Visible = False
    player(i).FontBold = False
    cas(i).FontBold = False
    wynik(i).FontBold = False
Next i
For i = 1 To IleGraczy
    player(i - 1).Visible = True
    Kolorek(i - 1).Left = 520
    player(i - 1).Left = Kolorek(i - 1).Left + Kolorek(i - 1).Width + 2
    player(i - 1).Top = 20 + ((player(i - 1).Height + 2) * (i))
    Kolorek(i - 1).Top = player(i - 1).Top
    cas(i - 1).Top = player(i - 1).Top
    wynik(i - 1).Top = player(i - 1).Top
    player(i - 1).Caption = PLR(i).imie
    wynik(i - 1).Caption = PLR(i).wynik
    wynik(i - 1).Visible = True
    Kolorek(i - 1).Visible = True
    cas(i - 1).Visible = True
    player(i - 1).FontBold = False
    Kolorek(i - 1).Picture = Obrazki.ListImages.item(i + 3).Picture
    wynik(i - 1).FontBold = False
    cas(i - 1).FontBold = False
    cas(i - 1).Caption = FormaCzasu(FMax - PLR(i).CzasCa³kowity)
Next i
CR.Visible = True
player(Ktory - 1).FontBold = True
wynik(Ktory - 1).FontBold = True
cas(Ktory - 1).FontBold = True
IleLitCap.Caption = IleLiter
aktGracz.Caption = "Uk³ada " & PLR(Ktory).imie
CR.Caption = RMax - Zero
JuzGramy = True

For i = 0 To 6
    If Right(Stojak(Ktory, i), 1) = "1" Then
        Call plytka(i).Po³ó¿(Left(Stojak(Ktory, i), 1), True)
        Call plytka(i).ZatwierdŸ(3)
    Else
        Call plytka(i).Po³ó¿(Left(Stojak(Ktory, i), 1), False)
        Call plytka(i).ZatwierdŸ(3)
    End If
Next i
Close hj

Solve.Enabled = True
anulujemy.Enabled = True
wymiana.Enabled = True
AkcjeSub(3).Enabled = True
AkcjeSub(1).Enabled = True
Kolejka.Enabled = True

Kill (App.Path & "\save\tmp.tmp")
Timer1.Enabled = False
If KoniecGry Then
    enduj
Else
    If IleLiter = 0 Then
        Timer1.Enabled = True
    End If
End If


On Error GoTo 0
Call Nazywaj
Call CzytajHS

Pokazuj True
If PLR(Ktory).Komp Then
   Podstawka.Visible = Not (MenuShowLit.Checked)
End If

Set Czasy = Nothing
Set Czasy = New Collection
DefMyd1.Visible = CzyBlank And (Not PLR(Ktory).Komp)
Label1.Visible = DefMyd1.Visible
ButOnOff (PLR(Ktory).Komp)
DoEvents
   
graslow.Caption = "£ów S³ów - " & Jêzyk.Nazwa
graslow.DefMyd1.Clear
   
For i = 0 To UBound(PLit) - 2
   graslow.DefMyd1.AddItem PLit(i)
Next i
   
graslow.DefMyd1.ListIndex = 0

DoEvents
If PLR(Ktory).Komp Then
   KompKladzie
End If
Zegar.Enabled = True
Close hj

Exit Sub

NieMaKatalogu2:
MkDir (App.Path & "\save")
GoTo 522
KoniecOdczytaj2:
Exit Sub

BladOdczytaj2:
Close hj
InfoBox "Wskazany plik jest nieprawid³owy.", False, False

End Sub

Public Function MinPktToWiN(Kto As Long) As Long
Dim MaxWyn As Long, i As Long, Best As Long, Roznica As Long
Dim IleR As Long

For i = 1 To IleGraczy
   If i <> Kto Then
      If PLR(i).wynik > MaxWyn Then
         MaxWyn = PLR(i).wynik
      End If
   End If
Next i

IleR = IleMaReszta(Kto)

MinPktToWiN = (MaxWyn - PLR(Kto).wynik) - IleR

End Function

Public Function MaxKount() As Long

Dim i As Long, n As Long

For i = 1 To IleGraczy
   If PLR(i).MyKolek.Count > n Then n = PLR(i).MyKolek.Count
Next i

MaxKount = n

End Function
Public Function KtoSzukaMB(Ile As Long, Litery As String, Optional znaki As String = "") As String()
Dim ks As String, dlg As Long, i As Long, j As Long
Dim k As Long, L As Long, m As Long, n As Long, o As Long
Dim znak(15) As String * 1, Wyrazik As String, Warunek As String
Dim PopZnak As String, NewZnaki() As String, NZankow As Long

ReDim NewZnaki(0)
znaki = Replace(znaki, Mydlo, "")

dlg = Len(Litery)
If dlg < Ile Then
   KtoSzukaMB = NewZnaki
   Exit Function
End If

For i = 1 To dlg
   znak(i) = Mid(Litery, i, 1)
Next i
NZankow = -1
Select Case Ile
    
Case 1
   
   For i = 1 To dlg
      Wyrazik = znak(i)
      If PopZnak <> znak(i) Then
         NZankow = NZankow + 1
         ReDim Preserve NewZnaki(NZankow)
         NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
         PopZnak = znak(i)
      End If
   Next i
   
Case 2

   For i = 1 To dlg - 1
      For j = i + 1 To dlg
         Wyrazik = znak(i) & znak(j)
         If Wyrazik <> PopZnak Then
            NZankow = NZankow + 1
            ReDim Preserve NewZnaki(NZankow)
            NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
            PopZnak = Wyrazik
         End If
      Next j
   Next i

Case 3

   For j = 1 To dlg - 2
       For k = j + 1 To dlg - 1
           For L = k + 1 To dlg
               Wyrazik = znak(j) & znak(k) & znak(L)
               If Wyrazik <> PopZnak Then
                  NZankow = NZankow + 1
                  ReDim Preserve NewZnaki(NZankow)
                  NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
                  PopZnak = Wyrazik
               End If
           Next L
      Next k
   Next j

Case 4
    
   For i = 1 To dlg - 3
      For j = i + 1 To dlg - 2
         For k = j + 1 To dlg - 1
            For L = k + 1 To dlg
               Wyrazik = znak(i) & znak(j) & znak(k) & znak(L)
               If Wyrazik <> PopZnak Then
                  NZankow = NZankow + 1
                  ReDim Preserve NewZnaki(NZankow)
                  NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
                  PopZnak = Wyrazik
               End If
            Next L
         Next k
      Next j
   Next i
   
Case 5
   
   For m = 1 To dlg - 4
      For i = m + 1 To dlg - 3
         For j = i + 1 To dlg - 2
            For k = j + 1 To dlg - 1
               For L = k + 1 To dlg
                  Wyrazik = znak(i) & znak(j) & znak(k) & znak(L) & znak(m)
                  If Wyrazik <> PopZnak Then
                     NZankow = NZankow + 1
                     ReDim Preserve NewZnaki(NZankow)
                     NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
                     PopZnak = Wyrazik
                  End If
               Next L
            Next k
         Next j
      Next i
   Next m

Case 6

   For n = 1 To dlg - 5
      For m = n + 1 To dlg - 4
         For i = m + 1 To dlg - 3
            For j = i + 1 To dlg - 2
               For k = j + 1 To dlg - 1
                  For L = k + 1 To dlg
                     Wyrazik = znak(i) & znak(j) & znak(k) & znak(L) & znak(m) & znak(n)
                     If Wyrazik <> PopZnak Then
                        NZankow = NZankow + 1
                        ReDim Preserve NewZnaki(NZankow)
                        NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
                        PopZnak = Wyrazik
                     End If
                  Next L
               Next k
            Next j
         Next i
      Next m
   Next n

Case 7

   For o = 1 To dlg - 6
      For n = o + 1 To dlg - 5
         For m = n + 1 To dlg - 4
            For i = m + 1 To dlg - 3
               For j = i + 1 To dlg - 2
                  For k = j + 1 To dlg - 1
                     For L = k + 1 To dlg
                        Wyrazik = znak(i) & znak(j) & znak(k) & znak(L) & znak(m) & znak(n) & znak(o)
                        NZankow = NZankow + 1
                        ReDim Preserve NewZnaki(NZankow)
                        NewZnaki(NZankow) = Replace(Replace(Wyrazik, "_", ""), Mydlo, "") & znaki
                     Next L
                  Next k
               Next j
            Next i
         Next m
      Next n
   Next o

End Select
    
KtoSzukaMB = NewZnaki

End Function

Public Sub CreateBaza(Katalog As String)
Dim i As Long, FSOM As New FileSystemObject, Plik As TextStream

FSOM.CreateFolder Katalog
For i = 2 To 15
   FSOM.CreateTextFile Katalog & "\Sl" & CStr(i) & ".bmt"
   FSOM.CreateTextFile Katalog & "\Sl" & CStr(i) & ".imt"
Next i

Set FSOM = Nothing
End Sub
Public Sub WersjaDemo()
InfoBox "To jest wersja demo. Pe³n¹ wersjê programu mo¿na zamówiæ pod adresem: www.lowslow.prv.pl", False, False
End Sub

Public Sub ZapiszUstaw()

SaveSetting "£ów S³ów", "Ustawienia", "ShowTime", CStr(CLng(graslow.MenuShowTime.Checked))
SaveSetting "£ów S³ów", "Ustawienia", "ShowLit", CStr(CLng(graslow.MenuShowLit.Checked))
SaveSetting "£ów S³ów", "Ustawienia", "ShowHist", CStr(CLng(graslow.MenuHist.Checked))
SaveSetting "£ów S³ów", "Ustawienia", "Sound", CStr(CLng(graslow.MenuSound.Checked))

End Sub

Public Sub OdczytUstaw()

graslow.MenuShowTime.Checked = CBool(GetSetting("£ów S³ów", "Ustawienia", "ShowTime", graslow.MenuShowTime.Checked))
graslow.MenuShowLit.Checked = CBool(GetSetting("£ów S³ów", "Ustawienia", "ShowLit", graslow.MenuShowLit.Checked))
graslow.MenuHist.Checked = CBool(GetSetting("£ów S³ów", "Ustawienia", "ShowHist", graslow.MenuHist.Checked))
graslow.MenuSound.Checked = CBool(GetSetting("£ów S³ów", "Ustawienia", "Sound", graslow.MenuSound.Checked))

End Sub

