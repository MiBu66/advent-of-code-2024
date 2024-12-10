Attribute VB_Name = "Day01"
Option Explicit
Dim Daten As String
Dim DatArr

Sub AdventDay01a()
Dim i As Integer, Summe As Long
Dim CollLeft As New Collection, CollRight As New Collection, Zeile

  Daten = HoleDaten(1, True)  'Lese die Vorgabedaten
  DatArr = Split(Daten, vbCrLf)  'Zerlege die Daten zeilenweise
  
  For i = 0 To UBound(DatArr)
    Zeile = Split(DatArr(i), "   ")   'Trenne zwischen linkem und rechtem Wert
    CollLeft.Add CLng(Zeile(0))          'und füge die linke Seite der Collection 1 hinzu
    CollRight.Add CLng(Zeile(1))          'und füge die rechte Seite der Collection 2 hinzu
  Next i
  Set CollLeft = SortCollection(CollLeft, False)   'sortiere die Collections
  Set CollRight = SortCollection(CollRight, False)
  For i = 1 To CollLeft.Count
    Summe = Summe + Abs(CollLeft(i) - CollRight(i))  'bilde die Summe der Differenzen li/re
  Next i
  MsgBox Summe
End Sub

Sub AdventDay01b()
Dim i As Integer, Summe As Long, Wert As Long, Ndx As String
Dim CollLeft As New Collection, CollRight As New Collection, Zeile

  Daten = HoleDaten(1, True)  'Lese die Vorgabedaten
  DatArr = Split(Daten, vbCrLf)  'Zerlege die Daten zeilenweise
  
  For i = 0 To UBound(DatArr)
    Zeile = Split(DatArr(i), "   ")   'Trenne zwischen linkem und rechtem Wert
    CollLeft.Add CLng(Zeile(0))          'und füge die linke Seite der Collection 1 hinzu
    On Error Resume Next
      Ndx = Zeile(1)
      CollRight.Add 1, Ndx                'versuche neuen Index anzulegen
      If Err.Number <> 0 Then         'wenn ein Fehler auftritt
        Wert = CollRight(Ndx)             'erhöhe den Wert des existierenden Index
        CollRight.Remove Ndx
        CollRight.Add Wert + 1, Ndx
      End If
    On Error GoTo 0
  Next i
  For i = 1 To CollLeft.Count
    On Error Resume Next
      Wert = CollRight(Trim(Str(CollLeft(i))))    'suche den Wert von CollLeft in der CollRight
      If Err.Number = 0 Then Summe = Summe + CollLeft(i) * Wert       'wenn gefunden, dann addiere das Produkt hinzu
    On Error GoTo 0
  Next i
  MsgBox Summe  'zeige Ergebnis
End Sub


