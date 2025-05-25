VERSION 5.00
Begin VB.Form ttt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   139
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Zegar 
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
End
Attribute VB_Name = "ttt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.Print ("ToolTipText")

End Sub

Private Sub Zegar_Timer()
Me.Hide
Zegar.Enabled = False
End Sub
