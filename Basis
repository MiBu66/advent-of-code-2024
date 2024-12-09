Option Explicit

Function HoleDaten(Tag As Integer, Full As Boolean, Optional AddOn As String = "") As String
Dim Pfad As String, Datei As String, Zeile As String, aktWB As Workbook
Dim Erg As String, i As Integer
  
  For Each aktWB In Application.Workbooks
    If InStr(1, aktWB.Path, "Advent") <> 0 Then
      Pfad = aktWB.Path
    End If
  Next
  Datei = Pfad & "\Day" & Format(Tag, "00")
  If Full Then Datei = Datei & ".txt" Else Datei = Datei & "_Test" & AddOn & ".txt"
  Open Datei For Input As #1
    Line Input #1, Erg
    Do While Not EOF(1)
      Line Input #1, Zeile
      Erg = Erg & vbCrLf & Zeile
    Loop
  Close #1
  HoleDaten = Erg
End Function

Public Function SortCollection(colInput As Collection, Optional KillDubl As Boolean = False) As Collection
Dim iCounter As Long, Loesch As Long
Dim iCounter2 As Long
Dim Temp As Variant
 
  Set SortCollection = New Collection
  For iCounter = 1 To colInput.Count - 1
    For iCounter2 = iCounter + 1 To colInput.Count
      If colInput(iCounter) > colInput(iCounter2) Then
        Temp = colInput(iCounter2)
        colInput.Remove iCounter2
        colInput.Add Temp, , iCounter
      End If
    Next iCounter2
  Next iCounter
  If KillDubl Then
  Loesch = 1
  For iCounter = colInput.Count - 1 To 2 Step -1
    If colInput(iCounter) = colInput(iCounter + 1) Then
      colInput.Remove iCounter
      'iCounter = iCounter - 1: Loesch = Loesch + 1
    End If
  Next iCounter
  End If
  Set SortCollection = colInput
End Function

