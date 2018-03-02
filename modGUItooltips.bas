Attribute VB_Name = "modGUItooltips"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Anpassung der Tooltips je nach gewählter Sprache bzw. Ausblenden der Tooltips
'------------------------------------------------------------------------------------------

Dim steuerelement As control
Dim zeile As Integer
    
Public Sub TooltipsON(ByRef spalte As Integer) ' 2 = Deutsch / 3 = Englisch

    On Error Resume Next
    
    'Steuerelemente beschriften
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmDuplikatManager.Controls
        If ThisWorkbook.Worksheets("Tooltips_GUI").Cells(zeile, spalte).Value <> "" Then
            steuerelement.ControlTipText = ThisWorkbook.Worksheets("Tooltips_GUI").Cells(zeile, spalte).Value
        End If
        zeile = zeile + 1
    Next steuerelement
    
End Sub

Public Sub TooltipsOFF()

    On Error Resume Next
    
    'Steuerelemente beschriften
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmDuplikatManager.Controls
        steuerelement.ControlTipText = ""
        zeile = zeile + 1
    Next steuerelement
    
End Sub
