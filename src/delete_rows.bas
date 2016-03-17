Attribute VB_Name = "DeleteStuff"
Option Explicit

Sub DeleteRows ()
	' Makro zum Löschen von nicht benötigten Spalten
	' Manuel S. Hubacher
	' 2016-03-17
	' Erstellt mit/für Excel für Mac, Version 15.20

	'-----------------------------------------------
	'STATUS: LÖSCHT ERST EINE ZEILE
	'ZIEL:   ALLE FUNDSTELLEN WERDEN GELÖSCHT 
	'-----------------------------------------------

	Dim strSuchen As String

	strSuchen = InputBox("Alle Zeilen, welche den eingegebenen Suchbegriff enthalten, werden gelöscht", "Zeileninhalt")

	If strSuchen = "" then
		MsgBox "Sie haben keine Eingabe getätigt. Das Makro wird beendet", vbOKOnly + vbInformation
		Exit Sub
	End If

	ActiveSheet.UsedRange.Find (strSuchen).EntireRow.Delete
	
End Sub