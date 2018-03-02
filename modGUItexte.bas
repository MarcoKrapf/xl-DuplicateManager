Attribute VB_Name = "modGUItexte"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Anpassung der statischen und dynamischen GUI-Beschriftungen je nach gewählter Sprache
'------------------------------------------------------------------------------------------

Public Sub spracheAendern(spalte As Integer) ' 2 = Deutsch / 3 = Englisch
    
    Dim steuerelement As control
    Dim zeile As Integer

    On Error Resume Next
    
    'Steuerelemente beschriften
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For Each steuerelement In frmDuplikatManager.Controls
        If ThisWorkbook.Worksheets("Controls_GUI").Cells(zeile, spalte).Value <> "" Then
            steuerelement.Caption = ThisWorkbook.Worksheets("Controls_GUI").Cells(zeile, spalte).Value
        End If
        zeile = zeile + 1
    Next steuerelement
    
    'MultiPage-Beschriftungen
    zeile = 1 'Zähler für die Zeile auf dem Tabellenblatt auf 1 setzen
    For zeile = 1 To 7
        frmDuplikatManager.MultiPage.Pages(zeile - 1).Caption _
            = ThisWorkbook.Worksheets("Other_GUI").Cells(zeile, spalte).Value
    Next zeile

    With frmDuplikatManager
        .SuchModus 'Dynamische Beschriftung aufrufen
        .MarkierModus 'Dynamische Beschriftung aufrufen
        .AusgabeModus 'Dynamische Beschriftung aufrufen
        .LoeschModusZeilen 'Dynamische Beschriftung aufrufen
        .LoeschModusKomprimieren 'Dynamische Beschriftung aufrufen
    End With
    
    Call frmDuplikatManager.AktuelleSelektionAnzeigen 'Daten der aktuellen Selektion anzeigen
    
    'Überschrift GUI
    frmDuplikatManager.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(9, spalte).Value & " (" & _
        ThisWorkbook.Worksheets("Other_GUI").Cells(10, spalte).Value & " " & cstrVersion & ")"
    
    'Überschrift UserForm QR-Code
    frmQRcode.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(12, spalte).Value

    'UserForm Versionshistorie
    frmVersionshinweise.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(33, spalte).Value
    frmVersionshinweise.lblVersionsInfo10a.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(34, spalte).Value
    frmVersionshinweise.lblVersionsInfo10b.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(35, spalte).Value
    frmVersionshinweise.lblVersionsInfo11a.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(36, spalte).Value
    frmVersionshinweise.lblVersionsInfo11b.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(37, spalte).Value
    frmVersionshinweise.lblVersionsInfo12a.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(38, spalte).Value
    frmVersionshinweise.lblVersionsInfo12b.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(39, spalte).Value
    frmVersionshinweise.lblVersionsInfo20a.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(40, spalte).Value
    frmVersionshinweise.lblVersionsInfo20b.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(41, spalte).Value
    
    'UserForm Nutzungsbedingungen
    frmNutzungsbedingungen.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(14, spalte).Value
    frmNutzungsbedingungen.lblNutzungsbedingungen = ThisWorkbook.Worksheets("Other_GUI").Cells(15, spalte).Value
    
    'UserForm Anleitung
        'Allgemeine Elemente
            frmAnleitung.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(47, spalte).Value
            frmAnleitung.lblAnl1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(48, spalte).Value

        'MultiPage-Tabs
            For zeile = 17 To 23
                frmAnleitung.MultiPageAnleitung.Pages(zeile - 17).Caption _
                    = ThisWorkbook.Worksheets("Other_GUI").Cells(zeile, spalte).Value
            Next zeile

        'Registerkarte "Allgemein"
            frmAnleitung.lblAnlAllgemein1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(105, spalte).Value
    
        'Registerkarte "Suchen"
            frmAnleitung.lblAnlSuchen1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(49, spalte).Value
            frmAnleitung.lblAnlSuchen2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(50, spalte).Value
            frmAnleitung.lblAnlSuchen3.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(51, spalte).Value
            frmAnleitung.lblAnlSuchen4.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(52, spalte).Value
            frmAnleitung.lblAnlSuchen5.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(53, spalte).Value
            frmAnleitung.lblAnlSuchen6.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(54, spalte).Value
            frmAnleitung.lblAnlSuchen7.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(55, spalte).Value
            frmAnleitung.lblAnlSuchen8.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(56, spalte).Value
            frmAnleitung.lblAnlSuchen9.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(57, spalte).Value
            frmAnleitung.lblAnlSuchen10.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(58, spalte).Value

        'Registerkarte "Hervorheben"
            frmAnleitung.lblAnlHervorheben1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(65, spalte).Value
            frmAnleitung.lblAnlHervorheben2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(66, spalte).Value
            frmAnleitung.lblAnlHervorheben3.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(67, spalte).Value
            frmAnleitung.lblAnlHervorheben4.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(68, spalte).Value
            frmAnleitung.lblAnlHervorheben5.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(69, spalte).Value
            frmAnleitung.lblAnlHervorheben6.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(70, spalte).Value
            frmAnleitung.lblAnlHervorheben7.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(71, spalte).Value
            frmAnleitung.lblAnlHervorheben8.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(72, spalte).Value
            frmAnleitung.lblAnlHervorheben9.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(73, spalte).Value
            frmAnleitung.lblAnlHervorheben10.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(74, spalte).Value
            frmAnleitung.lblAnlHervorheben11.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(75, spalte).Value
            frmAnleitung.lblAnlHervorheben12.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(76, spalte).Value
            frmAnleitung.lblAnlHervorheben13.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(77, spalte).Value
            frmAnleitung.lblAnlHervorheben14.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(78, spalte).Value
            
        'Registerkarte "Ausgeben"
            frmAnleitung.lblAnlAusgeben1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(85, spalte).Value
            frmAnleitung.lblAnlAusgeben2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(86, spalte).Value
            frmAnleitung.lblAnlAusgeben3.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(87, spalte).Value
            frmAnleitung.lblAnlAusgeben4.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(88, spalte).Value
            frmAnleitung.lblAnlAusgeben5.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(89, spalte).Value
            frmAnleitung.lblAnlAusgeben6.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(90, spalte).Value
            frmAnleitung.lblAnlAusgeben7.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(91, spalte).Value
            frmAnleitung.lblAnlAusgeben8.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(92, spalte).Value
            frmAnleitung.lblAnlAusgeben9.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(93, spalte).Value
            frmAnleitung.lblAnlAusgeben10.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(94, spalte).Value
            frmAnleitung.lblAnlAusgeben11.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(95, spalte).Value
            frmAnleitung.lblAnlAusgeben12.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(96, spalte).Value
            frmAnleitung.lblAnlAusgeben13.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(97, spalte).Value
            frmAnleitung.lblAnlAusgeben14.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(98, spalte).Value
            frmAnleitung.lblAnlAusgeben15.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(99, spalte).Value
            
        'Registerkarte "Löschen"
            frmAnleitung.lblAnlLoeschen1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(110, spalte).Value
            frmAnleitung.lblAnlLoeschen2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(111, spalte).Value
            frmAnleitung.lblAnlLoeschen3.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(112, spalte).Value
            frmAnleitung.lblAnlLoeschen4.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(113, spalte).Value
            frmAnleitung.lblAnlLoeschen5.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(114, spalte).Value
            frmAnleitung.lblAnlLoeschen6.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(115, spalte).Value
            
        'Registerkarte "Bereich Duplikate"
            frmAnleitung.lblAnlDuplikatfenster1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(120, spalte).Value
            frmAnleitung.lblAnlDuplikatfenster2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(121, spalte).Value
            frmAnleitung.lblAnlDuplikatfenster3.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(122, spalte).Value
        
        'Registerkarte "Bugs"
            frmAnleitung.lblAnlBugs1.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(130, spalte).Value
            frmAnleitung.lblAnlBugs2.Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(131, spalte).Value
            
End Sub
