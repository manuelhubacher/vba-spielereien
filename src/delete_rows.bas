Attribute VB_Name = "DeleteStuff"
Option Explicit

Sub DeleteRows ()
	' Makro zum L�schen von nicht ben�tigten Spalten
	' Manuel S. Hubacher
	' 2016-03-17
	' Erstellt mit/f�r Excel f�r Mac, Version 15.20

	'-----------------------------------------------
	'STATUS: L�SCHT ERST EINE ZEILE
	'ZIEL:   ALLE FUNDSTELLEN WERDEN GEL�SCHT 
	'-----------------------------------------------

	Dim strSuchen As String

	strSuchen = InputBox("Alle Zeilen, welche den eingegebenen Suchbegriff enthalten, werden gel�scht", "Zeileninhalt")

	If strSuchen = "" then
		MsgBox "Sie haben keine Eingabe get�tigt. Das Makro wird beendet", vbOKOnly + vbInformation
		Exit Sub
	End If

	ActiveSheet.UsedRange.Find (strSuchen).EntireRow.Delete
	
End Sub