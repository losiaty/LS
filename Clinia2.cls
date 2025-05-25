VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clinia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Pola As Collection, Cel As Object, Podkresl As Boolean
Public Szer As Single, MinHeight As Single, MojaTabela As cTabela2, Nadkresl As Boolean

Public Sub Drukuj(ByVal X As Single, ByVal Y As Single)
Dim e As cPole2, px As Single
px = X
For Each e In Pola
   e.Drukuj px, Y
   px = px + e.Szer
Next e

If Nadkresl = True Then
   Cel.Line (MojaTabela.MojaStrona.MarginesV, Y)-(Cel.ScaleWidth - MojaTabela.MojaStrona.MarginesV, Y)
End If

Y = Y + Height
If Podkresl = True Then
   Cel.Line (MojaTabela.MojaStrona.MarginesV, Y)-(Cel.ScaleWidth - MojaTabela.MojaStrona.MarginesV, Y)
End If

End Sub

Public Property Get Height() As Single
Dim e As cPole2, H As Single

For Each e In Pola
   If e.Height > H Then H = e.Height
Next e
If H > MinHeight Then
   Height = H
Else
   Height = MinHeight
End If

End Property

Public Function AddPole(Optional Tekst As String = "", Optional szerokoœæ As Single = 0, Optional horizontalCellSpacing As Single = 1, Optional verticalCellSpacing As Single = 0, Optional Nazwa As String, Optional eBackColor As ColorConstants = vbWhite, Optional eForeColor As ColorConstants = vbBlack, Optional eAlignH As AlignConstants = vbAlignLeft, Optional eAlignV As AlignConstants = vbAlignTop, Optional eFontName As String = "Arial CE", Optional eFontSize As Integer = 12, Optional eFontItalic As Boolean = False, Optional eFontBold As Boolean, Optional eFontUnderline As Boolean = False) As cPole2
Dim e As cPole2
Set e = New cPole2
Set e.Cel = Cel
e.eAlignH = eAlignH
e.eFontName = eFontName
e.eAlignV = eAlignV
e.vCs = verticalCellSpacing
e.hCs = horizontalCellSpacing
e.eBackColor = eBackColor
e.eFontBold = eFontBold
e.eFontItalic = eFontItalic
e.eFontSize = eFontSize
e.eFontUnderline = eFontUnderline
e.eForeColor = eForeColor
e.Szer = szerokoœæ
e.Tekst = Tekst
Set e.MojaLinia = Me
If Nazwa = "" Then
    Pola.Add e
Else
    Pola.Add e, Nazwa
End If
Set AddPole = e

End Function

Private Sub Class_Initialize()
Set Pola = New Collection
End Sub
