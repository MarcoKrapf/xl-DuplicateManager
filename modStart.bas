Attribute VB_Name = "modStart"
Option Explicit

'Modulbeschreibung:
'Das ist das Startmodul wenn das Tool per Mausklick auf das Icon im Ribbon geklickt wird
'------------------------------------------------------------------------------------------

'Versionsnummer des Tools
Public Const cstrVersion = "2.0" '(02. März 2018) 'ToDo
    
'Globale Variablen (gültig in allen Modulen und Formularen) definieren
Public g_strSaveDateipfad As String 'Dateipfad, in dem das Tool und die Textdatei mit den Tool-Einstellungen gespeichert werden
Public g_strSaveOptionen As String 'Dateiname der Textdatei, in der die Tool-Einstellungen gespeichert werden
Public g_intAktiviertesTabellenblatt As Integer 'Index des aktuell aktivierten Tabellenblatts
Public g_strSanduhrAktion As String 'Aktion, bei der die Sanduhr aufgerufen wird
Public g_strSanduhrNummer As String 'Nummer des Einzelschritts, bei der die Sanduhr aufgerufen wird
Public g_strSanduhrSchritt As String 'Einzelschritt, bei der die Sanduhr aufgerufen wird
Public g_dblBalkenAnteil As Double 'Breite des Fortschrittsbalkens der Sanduhr pro Schleifendurchlauf
Public g_dblBalkenAktuell As Double 'Aktuelle Breite des Fortschrittsbalkens der Sanduhr

'GUI aufrufen wenn das Icon im Ribbon angeklickt wird
Sub UserformAufrufen(control As IRibbonControl)
    Load frmDuplikatManager 'GUI laden
    frmDuplikatManager.Show 'GUI starten
End Sub

'Während der Entwicklung Einkommentieren
'Sub ToolStartenTest() 'Diese Prozedur manuell starten zum Testen der Entwicklung
'    Load frmDuplikatManager 'GUI laden
'    frmDuplikatManager.Show 'GUI starten
'End Sub
