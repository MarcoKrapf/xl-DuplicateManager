Attribute VB_Name = "modArraySortieren"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Die Funktion nimmt ein einspaltiges Array an, sortiert die Werte
'mit dem Bubble Sort Algorithmus aufsteigend und gibt es zurück.
'Der Code ist entnommen von https://msdn.microsoft.com/de-de/library/bb979305.aspx
'und wurde angepasst.
'------------------------------------------------------------------------------------------

'Variablen für dieses Modul definieren
Dim i As Long, j As Long, vTemp As Long


Public Function BubbleSort(ByRef ArrayToSort As Variant) As Variant 'Array aufsteigend sortieren

    'Sanduhr neu starten
        g_strSanduhrAktion = "Duplikate löschen"
        g_strSanduhrNummer = "[2/6]"
        g_strSanduhrSchritt = "Aktuellen Zustand sichern"
        'Fortschrittsbalken zurücksetzen
        Call frmDuplikatManager.FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / (UBound(ArrayToSort))
    
    For j = UBound(ArrayToSort) - 1 To LBound(ArrayToSort) Step -1
        ' Alle links davon liegenden Zeichen auf richtige Sortierung
        ' der jeweiligen Nachfolger überprüfen:
        For i = LBound(ArrayToSort) To j
            ' Ist das aktuelle Element seinem Nachfolger gegenüber korrekt sortiert?
            If ArrayToSort(i) > ArrayToSort(i + 1) Then
                ' Element und seinen Nachfolger vertauschen.
                vTemp = ArrayToSort(i)
                ArrayToSort(i) = ArrayToSort(i + 1)
                ArrayToSort(i + 1) = vTemp
            End If
        Next i
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call frmDuplikatManager.FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next j
    
    BubbleSort = ArrayToSort 'Sortiertes Array zurückgeben
    
End Function
