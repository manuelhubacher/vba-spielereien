Attribute VB_Name = "AddSheets"

Sub neue_tabellenblaetter()

' Einfaches Makro, um neue Tabellenblätter zu erstellen
' Manuel S. Hubacher
' 2016-03-09
' Erstellt mit/für Excel für Mac, Version 15.19.1

' Userinput: Anzahl Tabellenblätter
n = InputBox("Bitte geben Sie die Anzahl der zu erstellenden Tabellenblätter an.", "Anzahl eingeben", vbOKCancel)

' Testen, ob etwas eingegeben wurde
If n = "" Then
    MsgBox "Nichts eingegeben. Makro wird beendet.", vbInformation
    Exit Sub
End If

' Testen, ob eine Zahl eingegeben wurde
If Not IsNumeric(n) Then
    MsgBox "Keine Zahl eingegeben. Makro wird beendet.", vbInformation
    Exit Sub
End If

' Testen, ob zu viele (mehr als 20) neue Tabellenblätter erstellt werden
If n > 20 Then
    MsgBox "Sie möchten zu viele neue Tabellenblätter erstellen. Dieses Makro erlaubt höchsten 20 neue Tabellenblätter zu erstellen. Sie haben aber " & n & " eingegeben."
    Exit Sub
End If

' Tabellenblätter erstellen
For i = 1 To n
    Sheets.Add After:=Worksheets(Worksheets.Count)
Next i

End Sub