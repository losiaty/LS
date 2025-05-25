VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWydruk2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Cel As Object, MarH As Long, MarV As Long
Private Strony As Collection, Tabele As Collection
Private Naglowek As cTabela2, Stopka As cTabela2

Public Sub Drukuj(nrStrony As Long)
   If Not Cel Is Printer Then
      Cel.Cls
   Else
      Cel.ScaleMode = 6
   End If
   Strony(nrStrony).IleStron = Strony.Count
   Strony(nrStrony).Drukuj
End Sub

Public Sub DrukujAll()
Dim i As Long
For i = 1 To Strony.Count
   Drukuj i
Next i
End Sub

Public Function AddTabela(Optional Szerokosc As Single = 0, Optional horizontalCellSpacing As Single = 1, Optional verticalCellSpacing As Single = 0, Optional Nazwa As String, Optional eBackColor As ColorConstants = vbWhite, Optional eForeColor As ColorConstants = vbBlack, Optional eAlignH As AlignConstants = vbAlignLeft, Optional eAlignV As AlignConstants = vbAlignTop, Optional eFontName As String = "Arial CE", Optional eFontSize As Integer = 12, Optional eFontItalic As Boolean = False, Optional eFontBold As Boolean, Optional eFontUnderline As Boolean = False) As cTabela2
Dim NL As New cTabela2

Set NL.Cel = Cel
Tabele.Add NL
Set AddTabela = NL

End Function

Public Sub Stronicuj()
Dim S As cStrona2, T As cTabela2, i As Long
i = 1
Set S = New cStrona2
Set S.Cel = Cel
S.MarginesH = MarH
S.MarginesV = MarV
S.Miejsce = Cel.ScaleHeight - 2 * S.MarginesV
If Not Stopka Is Nothing Then S.AddStopka Stopka
If Not Naglowek Is Nothing Then S.AddNaglowek Naglowek
S.Numer = i
Strony.Add S

For Each T In Tabele
   If S.CzyZmiescisz(T) Then
      Set T.MojaStrona = S
      S.AddTabele T
   Else
      i = i + 1
      Set S = New cStrona2
      Set S.Cel = Cel
      S.Numer = i
      S.Miejsce = Cel.ScaleHeight
      S.MarginesH = MarH
      S.MarginesV = MarV
      S.Miejsce = Cel.ScaleHeight - 2 * S.MarginesV
      If Not Stopka Is Nothing Then S.AddStopka Stopka
      If Not Naglowek Is Nothing Then S.AddNaglowek Naglowek
      Set T.MojaStrona = S
      S.AddTabele T
      Strony.Add S
   End If
Next T

End Sub

Private Sub Class_Initialize()
Set Tabele = New Collection
Set Strony = New Collection
End Sub

Public Function AddNaglowek() As cTabela2

Dim n As New cTabela2
Set Naglowek = n
Set n.Cel = Cel
Set AddNaglowek = n

End Function

Public Function AddStopka() As cTabela2

Dim S As New cTabela2
Set Stopka = S
Set S.Cel = Cel
Set AddStopka = S

End Function
Public Property Get IleStron() As Long
If Not Strony Is Nothing Then
   IleStron = Strony.Count
Else
   IleStron = -1
End If
End Property

