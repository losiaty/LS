VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPole2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Tekst As String
Public eBackColor As ColorConstants, eForeColor As ColorConstants, eAlignH As AlignConstants
Public eFontSize As Integer, eFontItalic As Boolean, eFontBold As Boolean, eFontUnderline As Boolean
Public vCs As Single, hCs As Single, eAlignV As AlignConstants, eFontName As String
Public Szer As Single, Cel As Object, X As Single, Y As Single
Public MojaLinia As Clinia2
Private Const NrStr As String = "%%N"
Private Const SumStr As String = "%%S"

Public Property Get Height() As Single
   Cel.Font = eFontName
   Cel.FontBold = eFontBold
   Cel.FontSize = eFontSize
   Cel.FontItalic = eFontItalic
   Cel.FontUnderline = eFontUnderline
   Cel.Font.Charset = 238
   Height = LineCount2(Tekst) * Cel.TextHeight(Tekst)
End Property

Public Sub Drukuj(ByVal X As Single, ByVal Y As Single, Optional Align As AlignConstants = vbAlignLeft)
Dim NTekst() As String, i As Long, Tek As String
   
   Cel.Font = eFontName
   Cel.FontBold = eFontBold
   Cel.FontSize = eFontSize
   Cel.FontItalic = eFontItalic
   Cel.FontUnderline = eFontUnderline
   Cel.ForeColor = eForeColor
   Cel.Font.Charset = 238
   If eBackColor > -1 Then Cel.Line (X, Y)-(X + Szer, Y + Height), eBackColor, BF

NTekst = Dziel2(Tekst)
For i = 1 To UBound(NTekst)
   Select Case eAlignH
      Case vbAlignLeft
         Cel.CurrentX = X
      Case vbAlignRight
         Cel.CurrentX = X + Szer - Cel.TextWidth(NTekst(i))
      Case vbAlignNone
         Cel.CurrentX = X + Szer / 2 - Cel.TextWidth(NTekst(i)) / 2
   End Select
   
   Select Case eAlignV
      Case vbAlignTop
         Cel.CurrentY = Y
      Case vbAlignNone
         Cel.CurrentY = Y + (Me.MojaLinia.Height - Me.Height) / 2
      Case vbAlignBottom
         Cel.CurrentY = Y + (Me.MojaLinia.Height - Me.Height)
   End Select
   
   Tek = Replace(NTekst(i), NrStr, CStr(MojaLinia.MojaTabela.MojaStrona.Numer))
   Tek = Replace(Tek, SumStr, CStr(MojaLinia.MojaTabela.MojaStrona.IleStron))
   Cel.Print Trim(Tek)
   Y = Y + Cel.TextHeight(NTekst(i))
Next i

End Sub

Public Function Dziel2(Tekst As String) As String()

Dim NTekst() As String, Temp() As String, Tmp As String
Dim i As Long, n As Long
   
   Cel.Font = eFontName
   Cel.FontBold = eFontBold
   Cel.FontSize = eFontSize
   Cel.FontItalic = eFontItalic
   Cel.FontUnderline = eFontUnderline
   Cel.Font.Charset = 238

Temp = Split(Tekst, " ")
ReDim NTekst(0)
For i = LBound(Temp) To UBound(Temp)
   If Cel.TextWidth(Tmp) + Cel.TextWidth(" ") + Cel.TextWidth(Temp(i)) < Szer Then
      Tmp = Tmp & " " & Temp(i)
   Else
      ReDim Preserve NTekst(UBound(NTekst) + 1)
      NTekst(UBound(NTekst)) = Trim(Tmp)
      Tmp = Temp(i)
   End If
Next i

If Tmp <> "" Then
   ReDim Preserve NTekst(UBound(NTekst) + 1)
   NTekst(UBound(NTekst)) = Tmp
End If

Dziel2 = NTekst

End Function
Public Function LineCount2(Tekst As String) As Long

Dim NTekst() As String, Temp() As String, Tmp As String
Dim i As Long, n As Long, Ile As Long
   
   Cel.Font = eFontName
   Cel.FontBold = eFontBold
   Cel.FontSize = eFontSize
   Cel.FontItalic = eFontItalic
   Cel.FontUnderline = eFontUnderline
   Cel.Font.Charset = 238
   
Temp = Split(Tekst, " ")
Tmp = ""
For i = LBound(Temp) To UBound(Temp)
   
   If Cel.TextWidth(Tmp) + Cel.TextWidth(" ") + Cel.TextWidth(Temp(i)) < Szer Then
      Tmp = Tmp & " " & Temp(i)
   Else
      Tmp = Temp(i)
      Ile = Ile + 1
   End If
Next i

If Tmp <> "" Then
   Ile = Ile + 1
End If

LineCount2 = Ile

End Function

