Attribute VB_Name = "DeleteSheets"

Sub tabellenblaeter_loeschen()

' Einfaches Makro, um leere Tabellenblätter zu löschen
' Manuel S. Hubacher
' 2016-03-09
' Erstellt mit/für Excel für Mac, Version 15.19.1

' Variablen deklarieren
Dim wks As Worksheet

' Veränderungen und Meldungen nicht mehr anzeigen
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' Leere Tabellenblätter löschen
'------------------------------

For Each wks In ActiveWorkbook.Worksheets

    ' 1. Bedingung: WorksheetFunction.CountA zählt die nicht leeren Zellen eines Tabellenblattes
    ' 2. Bedingung: Mindestens ein Tabellenblatt übrig lassen, auch wenn dieses leer ist
    If WorksheetFunction.CountA(wks.Cells) = 0 And _
    ActiveWorkbook.Sheets.Count > 1 Then wks.Delete

Next wks

' Veränderungen und Meldungen wieder anzeigen
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub