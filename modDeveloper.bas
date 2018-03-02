Attribute VB_Name = "modDeveloper"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Prozeduren zum Auslesen der UserForm-Elemente w�hrend der Entwicklung

Private Sub auslesen_frmDuplikatManager()
    
    'Variablen
    Dim steuerelement As control
    Dim zeile As Integer
    
    On Error Resume Next
    
    'Alle Steuerelemente auf dem UserForm durchlaufen und die
    'Namen (also die ID) in Spalte A schreiben
    zeile = 1 'Z�hler f�r die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmDuplikatManager.Controls
        ThisWorkbook.Worksheets("Controls_GUI").Cells(zeile, 1).Value = steuerelement.Name 'bzw. ("Tooltips_GUI")
        zeile = zeile + 1
    Next steuerelement
    
End Sub
