VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIndeks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type tIndeks
   First As Long
   Last As Long
End Type
Public FSO As FileSystemObject, Znaczki As Collection
Public Nazwa As String

Public Sub Inicjuj(iNazwa As String, Katalog As String)
Dim Plik42 As TextStream, a As String, i As Long, n As Long, spl() As String
Dim z As cIndeksik
Set Plik42 = FSO.OpenTextFile(Katalog & "\" & iNazwa & ".imt", ForReading, True)
Set Znaczki = New Collection
While Plik42.AtEndOfStream = False
   a = Plik42.ReadLine
   spl = Split(a, ",")
   Set z = New cIndeksik
   z.First = spl(1)
   z.Last = spl(2)
   Znaczki.Add z, spl(0)
   Set z = Nothing
Wend
Plik42.Close

Set Plik42 = Nothing

End Sub
