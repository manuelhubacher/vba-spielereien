Attribute VB_Name = "AddSheets"
Option Explicit

Sub neue_tabellenblaeter()

' Einfaches Makro, um neue Tabellenblätter zu erstellen
' Manuel S. Hubacher
' 2016-03-07
' Erstellt mit/für Excel für Mac, Version 15.19.1

Dim intAnzBlatt As Integer
Dim i As Integer

' Userinput: Anzahl Tabellenblätter
intAnzBlatt = InputBox("Bitte geben Sie die Anzahl der zu erstellenden Tabellenblätter an.", "Anzahl eingeben", vbOKCancel)

' Testen, ob zu viele (mehr als 20) neue Tabellenblätter erstellt werden
If intAnzBlatt > 20 Then
    MsgBox "Sie möchten zu viele neue Tabellenblätter erstellen. Dieses Makro erlaubt höchsten 20 neue Tabellenblätter zu erstellen. Sie haben aber " & intAnzBlatt & " eingegeben."
    Exit Sub
End If

' Tabellenblätter erstellen
For i = 1 To intAnzBlatt
    Sheets.Add After:=Worksheets(Worksheets.Count)
Next i

End Sub