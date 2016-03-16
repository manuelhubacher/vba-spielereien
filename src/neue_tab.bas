Attribute VB_Name = "AddSheets"
Option Explicit

Sub neue_tabellenblaeter()

' Einfaches Makro, um neue Tabellenbl�tter zu erstellen
' Manuel S. Hubacher
' 2016-03-07
' Erstellt mit/f�r Excel f�r Mac, Version 15.19.1

Dim intAnzBlatt As Integer
Dim i As Integer

' Userinput: Anzahl Tabellenbl�tter
intAnzBlatt = InputBox("Bitte geben Sie die Anzahl der zu erstellenden Tabellenbl�tter an.", "Anzahl eingeben", vbOKCancel)

' Testen, ob zu viele (mehr als 20) neue Tabellenbl�tter erstellt werden
If intAnzBlatt > 20 Then
    MsgBox "Sie m�chten zu viele neue Tabellenbl�tter erstellen. Dieses Makro erlaubt h�chsten 20 neue Tabellenbl�tter zu erstellen. Sie haben aber " & intAnzBlatt & " eingegeben."
    Exit Sub
End If

' Tabellenbl�tter erstellen
For i = 1 To intAnzBlatt
    Sheets.Add After:=Worksheets(Worksheets.Count)
Next i

End Sub