Attribute VB_Name = "DeleteStuff"
Option Explicit

Sub DeleteRows()
    ' Makro zum L�schen von nicht ben�tigten Spalten
    ' Manuel S. Hubacher
    ' 2016-04-08
    ' Erstellt mit/f�r Excel f�r Mac, Version 15.20

    ' To do:
    ' - Suche auf die erste Spalte beschr�nken, in welcher Gesuchte Variable vorkommt
    ' - Fehlerhandhabung einbauen


    Dim strSuchen As String
    Dim strGefunden As String ' Erste Fundstelle
    Dim rngGefunden As Range  ' Fundstellen
    Dim rngKombo As Range

    strSuchen = InputBox("Alle Zeilen, welche den eingegebenen Suchbegriff enthalten, werden gel�scht", "Zeileninhalt")

    If strSuchen = "" Then
        MsgBox "Sie haben keine Eingabe get�tigt. Das Makro wird beendet", vbOKOnly + vbInformation
        Exit Sub
    End If

    With ActiveSheet.UsedRange
        ' Erste Fundstelle in Variable speichern
        Set rngGefunden = .Find(strSuchen, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns)
        ' Position der ersten Fundstelle, um Ende der Schleife zu finden
        strGefunden = rngGefunden.Address
        
        ' Fundstellen (Zeilen) als Variable
        ' Im ersten Durchlauf: Erste Fundstelle als rngKombo definieren
        ' In weiteren Durchl�ufen: Weitere Fundstellen mit rngKombo fusionieren
        Do
            If rngKombo Is Nothing Then
                Set rngKombo = rngGefunden.EntireRow
            Else
                Set rngKombo = Union(rngKombo, rngGefunden.EntireRow)
            End If
            
            ' N�chste Fundstelle
            Set rngGefunden = .FindNext(rngGefunden)
        Loop While Not rngGefunden Is Nothing And rngGefunden.Address <> strGefunden
        
    End With

    rngKombo.EntireRow.Hidden = True
    
    'ActiveSheet.UsedRange.Find(strSuchen).EntireRow.Delete
    
End Sub