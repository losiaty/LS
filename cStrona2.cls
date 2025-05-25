VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrona2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Tabele As Collection
Public Miejsce As Single, Numer As Long, IleStron As Long
Public Cel As Object, MarginesH As Single, MarginesV As Single
Private Stopka As cTabela2, Naglowek As cTabela2

Public Sub Drukuj()
Dim T As cTabela2, Y As Single, X As Single
Y = MarginesV
If Not Naglowek Is Nothing Then
   Naglowek.Drukuj MarginesH, Y
   Y = Y + Naglowek.Height
End If
For Each T In Tabele
   T.Drukuj MarginesH, Y
   Y = Y + T.Height
Next T
If Not Stopka Is Nothing Then
   Set Stopka.MojaStrona = Me
   Stopka.Drukuj MarginesH, Cel.ScaleHeight - MarginesV - Stopka.Height
End If
If Cel Is Printer Then Printer.EndDoc
End Sub

Public Function CzyZmiescisz(T As cTabela2) As Boolean

If T.Height < Miejsce Then
   CzyZmiescisz = True
Else
   CzyZmiescisz = False
End If

End Function

Public Sub AddTabele(ByVal T As cTabela2)

Miejsce = Miejsce - T.Height
Tabele.Add T

End Sub

Private Sub Class_Initialize()
Set Tabele = New Collection

End Sub

Public Sub AddStopka(T As cTabela2)
Set Stopka = T
Set Stopka.MojaStrona = Me
Miejsce = Miejsce - Stopka.Height
End Sub

Public Sub AddNaglowek(T As cTabela2)
Set Naglowek = T
Set Naglowek.MojaStrona = Me
Miejsce = Miejsce - Naglowek.Height
End Sub

Public Sub AddNumer()

End Sub
