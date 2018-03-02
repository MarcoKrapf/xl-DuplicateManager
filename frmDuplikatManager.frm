VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuplikatManager 
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   OleObjectBlob   =   "frmDuplikatManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmDuplikatManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Hauptcode des Tools, steuert die GUI
'------------------------------------------------------------------------------------------

'Dateityp: Excel-Add-In (.xlam)
'Name: xl DuplikatMananger
'Autor: Marco Krapf
'Version: 2.0 (02. März 2018)


'Variablen für dieses Modul definieren
'-------------------------------------
    Dim opt1 As String, opt2 As String, opt3 As String, opt4 As String, opt5 As String, _
        opt6 As String, opt7 As String, opt8 As String, opt9 As String, opt10 As String, _
        opt11 As String, opt12 As String, opt13 As String, opt14 As String, opt15 As String, _
        opt17 As String, opt18 As String, opt19 As String, opt20 As String, _
        opt21 As String 'Variablen mit den Werten der Tool-Einstellungen
    Dim intSprache As Integer 'Sprache (2 = Deutsch, 3 = Englisch)
    Dim blnStop As Boolean 'Kennzeichen ob die Verarbeitung abgebrochen werden soll (z.B. wenn Selektion ungültig)
    Dim strTabellenblatt As String 'Variable für den Namen des Tabellenblatts, das selektiert ist
    Dim rngSelection As Range 'Variable für selektierten Bereich auf dem Tabellenblatt
    Dim rngZelle As Range 'Variable für eine einzelne Zelle in For Each Schleifen
    Dim varSaveAll2() As Variant 'Array um alle Zellen eines selektierten Bereichs zu speichern (Ganze Zeile löschen)
    Dim varListeZeilen() As Variant '1-dimensionales Array mit allen Zeilen, in denen ein Duplikat vorkommt
    Dim varVergleichsListe() As Variant 'Array mit den auf Duplikate zu vergleichenden Werten
    Dim intAnzahlAreas As Integer 'Anzahl der selektierten Bereiche auf dem Tabellenblatt
    Dim lngAnzahlZeilen As Long 'Anzahl der insgesamt selektierten Zeilen
    Dim intAnzahlSpalten As Integer 'Anzahl der Spalten des selektierten Bereichs
    Dim lngMaxAnzahlZeilenBereich As Long 'Anzahl der Zeilen der längsten Spalte in einem Bereich
    Dim strSelektierterBereich As String 'Adresse des selektierten Bereichs
    Dim intAnzahlSpaltenDuplikatfenster As Integer 'Anzahl der Spalten des Duplikatfensters (max. 6)
    Dim blnUnikat As Boolean 'Kennzeichen, ob der Wert nur 1x vorkommt (Unikat)
    Dim strAktuellerWert As String 'Aktueller Wert, der auf Duplikate abgeglichen wird
    Dim strLetztesDuplikat As String 'Variable, die sich den verketteten String des letzten gefundenen Duplikats merkt
    Dim lngDuplikatGruppenNummer As Long 'Variable zum Gruppieren der gefundenen Duplikate
    Dim lngDuplikatCounter As Long 'Anzahl der gefundenen Duplikate
    Dim lngOriginalCounter As Long 'Anzahl der gefundenen 'Originale'
    Dim lngUnikatCounter As Long 'Anzahl der gefundenen Unikate
    Dim lngEindeutigeWerte As Long 'Summe aus 'Originalen' und Unikaten
    Dim blnLoeschmodus As Boolean 'Kennzeichen, ob die ganze Zeile gelöscht werden soll
    Dim lngAnzahlGeloeschteZeilen As Long 'Variable die zählt, wie viele komplette Zeilen gelöscht wurden
    Dim arrAusgabetyp(1 To 2) As String 'Array für die Typen der Ausgabe in einem neuen Tabellenblatt
    Dim intAktuellesTabellenblatt As Integer 'Index des Tabellenblatts
    Dim strAusgabetyp As String 'Duplikat oder Original
    Dim lngAktuelleZeile As Long 'Aktuelle Zeile
    Dim lngArrayZeile As Long, intArraySpalte As Integer 'Variablen für die Zuweisung des selektierten Bereichs in das Array
    Dim blnAlleFarbigMarkieren As Boolean 'Kennzeichen, ob auch die Originale farbig markiert werden sollen
    Dim blnVerdichtenBeimLoeschen As Boolean 'Kennzeichen, ob die Originale beim Löschen zusammengeschoben werden
    Dim lngAnzahlGeloeschteDuplikate As Long 'Anzahl der aktuell gelöschten Duplikate
    Dim varListeGeloeschteZeilen() As Variant 'Liste mit der Position der gelöschten Zeilen
    Dim varFarbListe As Variant 'Liste für die Farben zur Markierung von Duplikaten
    Dim strVorschauSpaltenBreite As String 'Breite der Spalten im Duplikatfenster
    Dim byteDuplikatfensterStatus As Byte 'Kennzeichen, ob ein Duplikat gelöscht oder wiederhergestellt werden kann
    Dim wksTabellenblatt As Worksheet 'Variable für Worksheets in For Each Schleifen
    Dim objMail As Object 'Shell-Objekt für E-Mail
    Dim strFehlerberichtProzedur As String 'Fehlerbericht: Name der Prozedur
    Dim strFehlerberichtErrNumber As String 'Fehlerbericht: Fehlernummer
    Dim strFehlerberichtErrDescription As String 'Fehlerbericht: Fehlerbeschreibung
    Dim i As Long, j As Long, k As Long, m As Long 'Zählvariablen für Schleifen
    

'PROZEDUREN
'----------

Private Sub UserForm_Initialize() 'Initialisierung beim Aufruf der GUI

    'Standardordner für Office-Add-Ins und Dateinamen für Systemdateien setzen
    g_strSaveDateipfad = Application.UserLibraryPath
    g_strSaveOptionen = "xlDuMa.settings"

    On Error Resume Next 'Falls ein Fehler auftritt überspringen
    
    'Startwerte setzen wenn die GUI aufgerufen wird
        'GUI maximieren
        Call MinimierenMaximieren
        'Button in Duplikatfenster ausblenden
        lblDuplikatEinzelnLoeschen.Visible = False
        'Standardsprache Englisch
        intSprache = 3
        'Sanduhr initialisieren
        With frmSanduhr
            .Caption = ThisWorkbook.Worksheets("Other_GUI").Cells(9, intSprache).Value
            .lblFortschrittSchritt = ""
            .lblFortschrittProzent = ""
            .lblFortschrittBalken.Width = 0
        End With
        'Suchmodus, Markiermodus und Löschmodus setzen
        Call SuchModus
        Call MarkierModus
        Call AusgabeModus
        Call LoeschModusZeilen
        Call LoeschModusKomprimieren
        'Aktives Tabellenblatt setzen
        g_intAktiviertesTabellenblatt = ActiveWorkbook.Worksheets(ActiveSheet.Index).Index
        'Speicherort setzen
        lblSpeicherortOptionen.Caption = g_strSaveDateipfad

        'Erste Registerkarte aktivieren
        MultiPage.Value = 0
    
    'Gespeicherte Tool-Einstellungen laden wenn Datei vorhanden
    If Dir(g_strSaveDateipfad & g_strSaveOptionen) <> "" Then
        'Kanal für den Output
        Dim intFileNr As Integer
        intFileNr = FreeFile 'Nächste freie Nummer zuweisen
        
        Open g_strSaveDateipfad & g_strSaveOptionen For Input As #intFileNr 'Eingangskanal öffnen
            Line Input #intFileNr, opt1 'Tooltips
            Line Input #intFileNr, opt2 'Warnungen
            Line Input #intFileNr, opt3 'Hinweise
            Line Input #intFileNr, opt4 'Sprache (2 = Deutsch, 3 = Englisch)
'            Line Input #intFileNr, opt5 '(unbenutzt)
            Line Input #intFileNr, opt6 'Farbige Markierungen entfernen
            Line Input #intFileNr, opt7 'Farbige Markierungen erhalten
            Line Input #intFileNr, opt8 'Suchkriterien: Groß-/Kleinschreibung
            Line Input #intFileNr, opt9 'Suchkriterien: Leerzeichen
            Line Input #intFileNr, opt10 'Suchmodus
            Line Input #intFileNr, opt11 'Titelzeile auf Tabellenblatt für Ausgabe
            Line Input #intFileNr, opt12 'Auto-Markieren: Duplikate
            Line Input #intFileNr, opt13 'Auto-Markieren: Alle mehrfachen
            Line Input #intFileNr, opt14 'Auto-Markieren: Unikate
            Line Input #intFileNr, opt15 'Auto-Markieren: nichts
            Line Input #intFileNr, opt17 'Markiermodus (einfarbig oder bunt)
            Line Input #intFileNr, opt18 'Ausgabemodus (nur Zellen der Selektion oder komplette Zeilen)
            Line Input #intFileNr, opt19 'Löschmodus (ganze Zeilen oder nur Zellen mit Duplikaten)
            Line Input #intFileNr, opt20 'Löschmodus (nur löschen oder auch komprimieren)
            Line Input #intFileNr, opt21 'Aktualisierung des Tabellenblatts während Berechnung
        Close #intFileNr 'Eingangskanal schließen
        
        'Ausgelesene Werte in UserForm setzen
        checkboxTooltip.Value = CBool(opt1)
        checkboxWarnung.Value = CBool(opt2)
        checkboxHinweis.Value = CBool(opt3)
        intSprache = CInt(opt4)
'        xxxxxxxxxxxxx = xxxxxx(opt5)
        optBtnFarbeEntfernen.Value = CBool(opt6)
        optBtnFarbeErhalten.Value = CBool(opt7)
        checkboxGrossKleinBuchstaben.Value = CBool(opt8)
        checkboxLeerzeichen.Value = CBool(opt9)
        toggleSuchModus.Value = CBool(opt10)
        checkboxAusgabeTitelzeile.Value = CBool(opt11)
        optBtnMarkieren1.Value = CBool(opt12)
        optBtnMarkieren2.Value = CBool(opt13)
        optBtnMarkieren3.Value = CBool(opt14)
        optBtnMarkieren4.Value = CBool(opt15)
        toggleMarkierModus.Value = CBool(opt17)
        toggleAusgabeModus.Value = CBool(opt18)
        toggleLoeschModusZeilen.Value = CBool(opt19)
        toggleLoeschModusKomprimieren.Value = CBool(opt20)
        checkboxAktualisierung.Value = CBool(opt21)
    End If
    
    'Haken für Speichern der Tool-Einstellungen setzen
    checkboxToolSettingsSpeichern.Value = True
    
    'Eventuelle Markierungen löschen
    Selection.Interior.Pattern = xlNone
    
    'Sprache
    Call Sprache(intSprache)
    'Tooltips an- bzw. ausschalten
    Call Tooltips
    'Selektierten Bereich auf dem Tabellenblatt einlesen
    Call Selektion
    If blnStop = True Then 'Wenn Selektion ungültig
        blnStop = False
        Exit Sub
    End If
    'Suche nach Duplikaten starten
    Call DuplikateFinden
    
End Sub


Private Sub UserForm_Terminate() 'Aktionen beim Schließen des UserForms

    On Error Resume Next 'falls ein Fehler auftritt: Anweisung überspringen
    
    'Tool-Einstellungen speichern wenn angehakt
    If checkboxToolSettingsSpeichern.Value = True Then
        'Kanal für den Output
        Dim intFileNr As Integer
        intFileNr = FreeFile 'Nächste freie Nummer zuweisen
        
        Open g_strSaveDateipfad & g_strSaveOptionen For Output As #intFileNr 'Ausgangskanal öffnen
            'Werte in Textdatei schreiben (-1 für Haken gesetzt, 0 für nicht gesetzt)
            Print #intFileNr, CInt(checkboxTooltip.Value) 'Tooltips
            Print #intFileNr, CInt(checkboxWarnung.Value) 'Warnungen
            Print #intFileNr, CInt(checkboxHinweis.Value) 'Hinweise
            Print #intFileNr, intSprache 'Sprache (2 = Deutsch, 3 = Englisch)
'            Print #intFileNr, CInt(xxxxxxxxx) '(unbenutzt)
            Print #intFileNr, CInt(optBtnFarbeEntfernen.Value) 'Farbige Markierungen entfernen
            Print #intFileNr, CInt(optBtnFarbeErhalten.Value) 'Farbige Markierungen erhalten
            Print #intFileNr, CInt(checkboxGrossKleinBuchstaben.Value) 'Suchkriterien: Groß-/Kleinschreibung
            Print #intFileNr, CInt(checkboxLeerzeichen.Value) 'Suchkriterien: Leerzeichen
            Print #intFileNr, CInt(toggleSuchModus.Value) 'Suchmodus
            Print #intFileNr, CInt(checkboxAusgabeTitelzeile.Value) 'Titelzeile auf Tabellenblatt für Ausgabe
            Print #intFileNr, CInt(optBtnMarkieren1.Value) 'Auto-Markieren: Duplikate
            Print #intFileNr, CInt(optBtnMarkieren2.Value) 'Auto-Markieren: Alle mehrfachen
            Print #intFileNr, CInt(optBtnMarkieren3.Value) 'Auto-Markieren: Unikatge
            Print #intFileNr, CInt(optBtnMarkieren4.Value) 'Auto-Markieren: nichts
            Print #intFileNr, CInt(toggleMarkierModus.Value) 'Markiermodus (einfarbig oder bunt)
            Print #intFileNr, CInt(toggleAusgabeModus.Value) 'Ausgabemodus (nur Zellen der Selektion oder komplette Zeilen)
            Print #intFileNr, CInt(toggleLoeschModusZeilen.Value) 'Löschmodus (ganze Zeilen oder nur Zellen mit Duplikaten)
            Print #intFileNr, CInt(toggleLoeschModusKomprimieren.Value) 'Löschmodus (nur löschen oder auch komprimieren)
            Print #intFileNr, CInt(checkboxAktualisierung.Value) 'Aktualisierung des Tabellenblatts während Berechnung
        Close #intFileNr 'Ausgangskanal schließen
    End If
    
    'Popups schließen
    Unload frmAnleitung
    Unload frmNutzungsbedingungen
    Unload frmQRcode
    Unload frmSanduhr
    Unload frmVersionshinweise
    
End Sub

'Neue Selektion übernehmen
Private Sub Selektion()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Tabellenblatt, Areas, Adressen und Abmessungen der Selektion in Userform eintragen
    lblAktuelleMappeWert.Caption = ActiveWorkbook.Name
    lblAktuellesBlattWert.Caption = ActiveSheet.Name

    'Duplikatfenster leeren und ausblenden
    boxVorschau.Clear
    boxVorschau.Visible = False
    
    'Buttons deaktivieren
    btnHervorhebenAlle.Enabled = False
    btnHervorhebenDuplikate.Enabled = False
    btnHervorhebenUnikate.Enabled = False
    btnFarbigeMarkierungLoeschen.Enabled = False
    toggleMarkierModus.Enabled = False
    btnDuplikateAusgeben.Enabled = False
    btnOriginaleAusgeben.Enabled = False
    btnUnikateAusgeben.Enabled = False
    btnDuplikateLoeschen.Enabled = False
    btnLoeschenRueckgaengig.Enabled = False
    
    'Duplikatfenster aktualisieren
    byteDuplikatfensterStatus = 0
    lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
    lblDuplikatEinzelnLoeschen.Enabled = False
        
    'Variable für die Adresse des selektierten Bereichs zurücksetzen
    strSelektierterBereich = ""
    
    Set rngSelection = Selection 'Selektierten Bereich in Variable einlesen
    strTabellenblatt = ActiveSheet.Name 'Name des Tabellenblatts auslesen

    'Auslesen, wie viele Bereiche die Selektion hat
    intAnzahlAreas = rngSelection.Areas.Count
    
    'Zähler zurücksetzen
    lngAnzahlZeilen = 0
    intAnzahlSpalten = 0
    lngMaxAnzahlZeilenBereich = 0
    
    'Check ob die Selektion gültig ist
        'Wenn selektierte Bereiche einzeln behandelt werden: Prüfung ob die selektierten Bereiche gleich viele Spalten haben
        If toggleSuchModus.Value = False Then
            'Anzahl Spalten
            intAnzahlSpalten = rngSelection.Columns.Count 'nimmt die Spaltenanzahl des zuerst selektierten Bereichs
            
            For i = 1 To intAnzahlAreas
                If rngSelection.Areas(i).Columns.Count <> intAnzahlSpalten Then
                    MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(65, intSprache).Value), _
                            vbExclamation + vbOKOnly, ThisWorkbook.Worksheets("Messages_GUI").Cells(64, intSprache).Value
                    Call SelektionUngueltig
                    Exit Sub
                End If
                
                'Zeilen der selektierten Bereiche aufaddieren
                lngAnzahlZeilen = lngAnzahlZeilen + rngSelection.Areas(i).Rows.Count
            Next i
        End If
        
        'Wenn selektierte Spalten zu 1 Bereich vereinigt werden: Prüfung ob die selektierten Bereiche in den gleichen Zeilen liegen
        If toggleSuchModus.Value = True Then
            'Oberste Zeile und Anzahl der Zeilen
            lngAktuelleZeile = rngSelection.Rows.Row 'nimmt die oberste Zeile des zuerst selektierten Bereichs
            lngAnzahlZeilen = rngSelection.Rows.Count 'nimmt die Anzahl der Zeilen des zuerst selektierten Bereichs

            For i = 1 To intAnzahlAreas
                If rngSelection.Areas(i).Row <> lngAktuelleZeile Or _
                    rngSelection.Areas(i).Rows.Count <> lngAnzahlZeilen Then
                    MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(66, intSprache).Value), _
                            vbExclamation + vbOKOnly, ThisWorkbook.Worksheets("Messages_GUI").Cells(64, intSprache).Value
                    Call SelektionUngueltig
                    Exit Sub
                End If

                'Spalten der selektierten Bereiche aufaddieren
                intAnzahlSpalten = intAnzahlSpalten + rngSelection.Areas(i).Columns.Count
            Next i
        End If
    
    'Selektierten Bereich neu einlesen (nötig für den Fall, dass ganze Spalten markiert sind)
        'Wenn selektierte Bereiche einzeln behandelt werden
        If toggleSuchModus.Value = False Then
            'Zähler für die Anzahl der insgesamt selektierten Zeilen wieder zurücksetzen
            lngAnzahlZeilen = 0
            
            For i = 1 To rngSelection.Areas.Count
                'Zähler für den neuen Bereich zurücksetzen
                lngMaxAnzahlZeilenBereich = 0
            
                'Bereich, bei dem die ganze Spalte selektiert ist
                If rngSelection.Areas(i).Rows.Count = 1048576 Then
                
                    'Längste Spalte des Bereichs ermitteln
                    For j = 1 To intAnzahlSpalten
                        If Cells(rngSelection.Areas(i).Rows.Count, rngSelection.Areas(i).Column + j - 1).End(xlUp).Row > lngMaxAnzahlZeilenBereich Then
                            If IsEmpty(Cells(1048576, rngSelection.Areas(i).Column + j - 1)) Then
                                lngMaxAnzahlZeilenBereich = Cells(rngSelection.Areas(i).Rows.Count, rngSelection.Areas(i).Column + j - 1).End(xlUp).Row
                            Else
                                lngMaxAnzahlZeilenBereich = 1048576
                            End If
                        End If
                    Next j
                    
                    'Zeilen des Bereichs aufaddieren
                    lngAnzahlZeilen = lngAnzahlZeilen + lngMaxAnzahlZeilenBereich
                
                    'Bereich anpassen, kürzen bis zur letzten Zeile mit Daten
                    strSelektierterBereich = strSelektierterBereich & _
                        Range( _
                        Cells(1, rngSelection.Areas(i).Column), _
                        Cells(lngMaxAnzahlZeilenBereich, rngSelection.Areas(i).Column + intAnzahlSpalten - 1)) _
                        .AddressLocal & ","
                        
                    'Hinweis anzeigen, wenn Checkbox aktiv
                    If checkboxHinweis.Value = True Then
                        MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(5, intSprache).Value & i & " (" & _
                            ThisWorkbook.Worksheets("Messages_GUI").Cells(6, intSprache).Value & rngSelection.Areas(i).Columns.Column & _
                            "-" & rngSelection.Areas(i).Columns.Column + rngSelection.Areas(i).Columns.Count - 1 & _
                            ") " & ThisWorkbook.Worksheets("Messages_GUI").Cells(7, intSprache).Value & "." & vbNewLine & vbNewLine & _
                            ThisWorkbook.Worksheets("Messages_GUI").Cells(8, intSprache).Value & vbNewLine & _
                            ThisWorkbook.Worksheets("Messages_GUI").Cells(9, intSprache).Value & lngMaxAnzahlZeilenBereich & "."), _
                            vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(4, intSprache).Value
                    End If
                    
                'Bereich, bei dem keine ganze Spalte selektiert ist
                Else
                    'Zeilen des Bereichs aufaddieren
                    lngAnzahlZeilen = lngAnzahlZeilen + rngSelection.Areas(i).Rows.Count
                    'Bereich unverändert lassen
                    strSelektierterBereich = strSelektierterBereich & rngSelection.Areas(i).AddressLocal & ","
                End If
            Next i
            
            'Komma am Ende des neu zusammengesetzten Strings mit dem selektierten Bereich entfernen
            strSelektierterBereich = Left(strSelektierterBereich, Len(strSelektierterBereich) - 1)
            
            'Selektion neu einlesen
            Range(strSelektierterBereich).Select
            Set rngSelection = Selection
        End If
        
        'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
        If toggleSuchModus.Value = True And rngSelection.Rows.Count = 1048576 Then
            'Zähler für die Anzahl der insgesamt selektierten Zeilen wieder zurücksetzen
            lngAnzahlZeilen = 0
            'Maximale Zeilenanzahl mit Daten ermitteln
            For i = 1 To rngSelection.Areas.Count
                'Längste Spalte des Bereichs ermitteln
                For j = 1 To rngSelection.Areas(i).Columns.Count
                    If Cells(rngSelection.Areas(i).Rows.Count, rngSelection.Areas(i).Column + j - 1).End(xlUp).Row > lngMaxAnzahlZeilenBereich Then
                        If IsEmpty(Cells(1048576, rngSelection.Areas(i).Column + j - 1)) Then
                            lngMaxAnzahlZeilenBereich = Cells(rngSelection.Areas(i).Rows.Count, rngSelection.Areas(i).Column + j - 1).End(xlUp).Row
                        Else
                            lngMaxAnzahlZeilenBereich = 1048576
                        End If
                    End If
                Next j
            Next i
            lngAnzahlZeilen = lngMaxAnzahlZeilenBereich
            
            'Bereich anpassen, kürzen bis zur letzten Zeile mit Daten
            For i = 1 To rngSelection.Areas.Count
                strSelektierterBereich = strSelektierterBereich & _
                    Range( _
                    Cells(1, rngSelection.Areas(i).Column), _
                    Cells(lngAnzahlZeilen, rngSelection.Areas(i).Column + rngSelection.Areas(i).Columns.Count - 1)) _
                    .AddressLocal & ","
            Next i
                
            'Hinweis anzeigen, wenn Checkbox aktiv
            If checkboxHinweis.Value = True And lngAnzahlZeilen < 1048576 Then
                MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(12, intSprache).Value & "." & vbNewLine & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(13, intSprache).Value & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(14, intSprache).Value & lngAnzahlZeilen & "."), _
                    vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(11, intSprache).Value
            End If
            
            'Komma am Ende des neu zusammengesetzten Strings mit dem selektierten Bereich entfernen
            strSelektierterBereich = Left(strSelektierterBereich, Len(strSelektierterBereich) - 1)
            
            'Selektion neu einlesen
            Range(strSelektierterBereich).Select
            Set rngSelection = Selection
        End If
    
    Call AktuelleSelektionAnzeigen 'Daten der aktuellen Selektion anzeigen
    
    'Selektierte Einzelbereiche anzeigen
    lblAktuellerBereichWert.Caption = rngSelection.AddressLocal(False, False) 'Anzeige ohne $
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "Selektion"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub

Public Sub AktuelleSelektionAnzeigen() 'Daten der aktuellen Selektion anzeigen

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'GUI-Beschriftungen anpassen
    If intAnzahlAreas > 1 Then
        lblAktuelleAnzahlBereicheWert.Caption = intAnzahlAreas & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(23, intSprache).Value
    Else
        lblAktuelleAnzahlBereicheWert.Caption = intAnzahlAreas & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(24, intSprache).Value
    End If
    
    If toggleSuchModus.Value = False Then 'Wenn selektierte Bereiche einzeln behandelt werden
        If lngAnzahlZeilen > 1 Then
            lblSuchbereichGroesse.Caption = lngAnzahlZeilen & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(26, intSprache).Value & _
                                            " (" & intAnzahlSpalten & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(28, intSprache).Value
        Else
            lblSuchbereichGroesse.Caption = lngAnzahlZeilen & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(27, intSprache).Value & _
                                            " (" & intAnzahlSpalten & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(28, intSprache).Value
        End If
    Else 'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
        If lngAnzahlZeilen > 1 Then
            lblSuchbereichGroesse.Caption = lngAnzahlZeilen & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(30, intSprache).Value
        Else
            lblSuchbereichGroesse.Caption = lngAnzahlZeilen & ThisWorkbook.Worksheets("Dynamic_GUI").Cells(31, intSprache).Value
        End If
    End If
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "SelektionsBereichAnzeigen"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub

Private Sub SelektionUngueltig()
    
    'Anzeige des selektierten Bereichs löschen
    lblAktuellerBereichWert.Caption = ""
    lblAktuelleAnzahlBereicheWert.Caption = ""
    lblSuchbereichGroesse.Caption = ""
    
    'Anzeige der Funde zurücksetzen
    lblAnzahlDuplikate.Caption = ""
    lblAnzahlUnikate.Caption = ""
    lblAnzahlVerschiedene.Caption = ""
    
    'Abbruchkennzeichen setzen
    blnStop = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
End Sub


'Duplikate finden und anzeigen
'-----------------------------

Private Sub DuplikateFinden() 'Suche nach Duplikaten starten

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'VergleichsListe zurücksetzen und Größe anpassen:
        'Spalte 0 wird für den verketteten String und den Vergleich benötigt
        'Die am Ende angefügten 5 Spalten werden für die Markierung als Original, Duplikat bzw. Unikat,
        'die Position der Zelle (Zeile und Spalte) auf dem Tabellenblatt, die DuplikatGruppenNummer
        'und die Markierung, dass das Duplikat schon berücksichtigt wurde, benötigt
    ReDim varVergleichsListe(1 To lngAnzahlZeilen, 0 To intAnzahlSpalten + 5)

    'Farbige Markierung im selektierten Bereich des Tabellenblatts entfernen
    rngSelection.Interior.ColorIndex = xlNone
    
    Call VergleichsListeErstellen
    Call SpaltenVerketten
    Call DuplikateSuchen
    
    If lngDuplikatCounter > 0 Then
        Call Farben 'Legt die Farben für die Markierung der Duplikate auf dem Tabellenblatt fest
        Call VorschauSpaltenbreiten 'Ermittelt die Spaltenbreiten für das Vorschaufenster
        boxVorschau.Visible = True 'Duplikatfenster einblenden
        Call AusgabeVorschau 'Duplikate in Fenster im User-Form ausgeben
    End If
        
    'Buttons aktivieren, wenn Duplikate gefunden wurden
    If lngDuplikatCounter > 0 Then
        btnHervorhebenAlle.Enabled = True '"Alle mehrfachen Werte markieren"
        btnHervorhebenDuplikate.Enabled = True 'Button "Nur Duplikate markieren"
        btnDuplikateAusgeben.Enabled = True 'Button "Duplikate ausgeben"
        btnOriginaleAusgeben.Enabled = True 'Button "Alle Werte 1x ausgeben"
        btnDuplikateLoeschen.Enabled = True 'Alle Duplikate löschen und verdichten"
    Else:
        btnHervorhebenAlle.Enabled = False
        btnHervorhebenDuplikate.Enabled = False
        btnHervorhebenUnikate.Enabled = False
        btnDuplikateAusgeben.Enabled = False
        btnOriginaleAusgeben.Enabled = False
        btnDuplikateLoeschen.Enabled = False
    End If
    
    'Buttons aktivieren, wenn Unikate gefunden wurden
    If lngUnikatCounter > 0 Then
        btnHervorhebenUnikate.Enabled = True 'Button "Nur Unikate markieren"
        btnUnikateAusgeben.Enabled = True 'Button "Unikate ausgeben"
    Else:
        btnHervorhebenUnikate.Enabled = False
        btnUnikateAusgeben.Enabled = False
    End If
    
    'Markieren-Buttons: Farbe zurücksetzen
    btnHervorhebenDuplikate.BackColor = &HFFC0C0
    btnHervorhebenAlle.BackColor = &HFFC0C0
    btnHervorhebenUnikate.BackColor = &HFFC0C0
    
    'Tabellenblatt, in dem nach Duplikaten gesucht wurde, merken
    intAktuellesTabellenblatt = ActiveSheet.Index 'ToDo ist das okay so???
    
    'Auto-Markierung ausführen wenn ausgewählt
    Call AutoMarkierung
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

    Exit Sub
    
Fehlerbehandlung:
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    If Err.Number = 7 Then
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(16, intSprache).Value & Err.Description & "." & _
            vbCrLf & ThisWorkbook.Worksheets("Messages_GUI").Cells(17, intSprache).Value & ThisWorkbook.Worksheets("Messages_GUI").Cells(18, intSprache).Value
    ElseIf Err.Number = 6 Then
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(16, intSprache).Value & Err.Description & "." & _
            vbCrLf & ThisWorkbook.Worksheets("Messages_GUI").Cells(17, intSprache).Value & ThisWorkbook.Worksheets("Messages_GUI").Cells(18, intSprache).Value
    Else:
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(19, intSprache).Value
        strFehlerberichtProzedur = "DuplikateFinden"
        strFehlerberichtErrNumber = Err.Number
        strFehlerberichtErrDescription = Err.Description
        Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    End If
    
    Call SelektionUngueltig

End Sub

Private Sub AutoMarkierung() 'Farbige Markierung auf dem Tabellenblatt direkt nach der Duplikatsuche
    If optBtnMarkieren1 = True And lngDuplikatCounter > 0 Then 'Nur Duplikate
        Call HervorhebenDuplikate
        Call DuplikateHervorheben
    ElseIf optBtnMarkieren2 = True And lngDuplikatCounter > 0 Then 'Alle mehrfachen
        Call HervorhebenAlle
        Call DuplikateHervorheben
    ElseIf optBtnMarkieren3 = True And lngUnikatCounter > 0 Then 'Nur Unikate
        Call HervorhebenUnikate
        Call UnikateHervorheben
    End If
End Sub

Private Sub VergleichsListeErstellen() 'Selektierte Bereiche in VergleichsListe einlesen

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(1, intSprache).Value
        g_strSanduhrNummer = "[1/4]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(2, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
        
    lngArrayZeile = 0 'Zeiger der Zeile der VergleichsListe auf 0 setzen
    
    'Wenn selektierte Bereiche einzeln behandelt werden
    If toggleSuchModus.Value = False Then
        
        For i = 1 To intAnzahlAreas 'Alle selektierten Bereiche nacheinander durchlaufen
            For j = rngSelection.Areas(i).Row To rngSelection.Areas(i).Row + rngSelection.Areas(i).Rows.Count - 1
                lngArrayZeile = lngArrayZeile + 1 'Zeiger der Zeile in VergleichsListe hochzählen
                intArraySpalte = 0 'Zeiger der Spalte in VergleichsListe auf 0 setzen
                For k = rngSelection.Areas(i).Column To rngSelection.Areas(i).Column + rngSelection.Areas(i).Columns.Count - 1
                    intArraySpalte = intArraySpalte + 1 'Zeiger der Spalte in VergleichsListe hochzählen
                    varVergleichsListe(lngArrayZeile, intArraySpalte) = ActiveSheet.Cells(j, k).Value 'Inhalt der Zelle in die VergleichsListe schreiben
                Next k
                varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 1) = "---leer---" 'Initial alle Zeilen als Original markieren
                varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 2) = ActiveSheet.Cells(j, k - intAnzahlSpalten).Row 'Zeile auf dem Tabellenblatt
                varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 3) = ActiveSheet.Cells(j, k - intAnzahlSpalten).Column 'Erste Spalte auf dem Tabellenblatt
                
                'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
                
             Next j
        Next i
    End If
    
    'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
    If toggleSuchModus.Value = True Then
        For i = rngSelection.Rows.Row To rngSelection.Rows.Row + rngSelection.Rows.Count - 1 'Alle Zeilen nacheinander durchlaufen
            lngArrayZeile = lngArrayZeile + 1 'Zeiger der Zeile in VergleichsListe hochzählen
            intArraySpalte = 0 'Zeiger der Spalte in VergleichsListe auf 0 setzen
            For j = 1 To intAnzahlAreas 'Alle selektierten Bereiche nacheinander durchlaufen
                For k = rngSelection.Areas(j).Columns.Column To rngSelection.Areas(j).Columns.Column + rngSelection.Areas(j).Columns.Count - 1 'Alle Spalten des Bereichs nacheinander durchlaufen
                    intArraySpalte = intArraySpalte + 1 'Zeiger der Spalte in VergleichsListe hochzählen
                    varVergleichsListe(lngArrayZeile, intArraySpalte) = ActiveSheet.Cells(i, k).Value 'Inhalt der Zelle in die VergleichsListe schreiben
                Next k
            Next j
            varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 1) = "---leer---" 'Initial alle Zeilen als Original markieren
            varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 2) = ActiveSheet.Cells(i, rngSelection.Areas(1).Columns.Column).Row 'Zeile auf dem Tabellenblatt
            varVergleichsListe(lngArrayZeile, intAnzahlSpalten + 3) = ActiveSheet.Cells(i, rngSelection.Areas(1).Columns.Column).Column 'Erste Spalte des ersten Bereichs auf dem Tabellenblatt
            
            'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
        Next i
    End If

    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "VergleichsListeErstellen"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub SpaltenVerketten() 'Werte der Spalten des selektierten Bereichs verketten, um Duplikate finden zu können

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(1, intSprache).Value
        g_strSanduhrNummer = "[2/4]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(3, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
    For i = 1 To lngAnzahlZeilen 'Zeilen des Arrays durchlaufen
        For j = 1 To intAnzahlSpalten 'Spalten des Arrays durchlaufen
            
            'Für jede selektierte Zeile die Inhalte der Spalten verbinden
            If checkboxGrossKleinBuchstaben.Value = True And checkboxLeerzeichen.Value = True Then 'wenn beide Checkboxen aktiv
                varVergleichsListe(i, 0) = varVergleichsListe(i, 0) & Replace(LCase(varVergleichsListe(i, j)), " ", "") 'alles zu Kleinbuchstaben konvertieren und Leerzeichen entfernen
            ElseIf checkboxGrossKleinBuchstaben.Value = True And checkboxLeerzeichen.Value = False Then 'nur Checkbox "Groß-/Kleinbuchstaben ignorieren" aktiv
                varVergleichsListe(i, 0) = varVergleichsListe(i, 0) & LCase(varVergleichsListe(i, j)) 'alles zu Kleinbuchstaben konvertieren, Leerzeichen beibehalten
            ElseIf checkboxGrossKleinBuchstaben.Value = False And checkboxLeerzeichen.Value = True Then 'nur Checkbox "Leerzeichen ignorieren" aktiv
                varVergleichsListe(i, 0) = varVergleichsListe(i, 0) & Replace(varVergleichsListe(i, j), " ", "") 'alle Leerzeichen löschen, Groß-/Kleinschreibung beibehalten
            Else 'wenn keine Checkboxen aktiv
                varVergleichsListe(i, 0) = varVergleichsListe(i, 0) & varVergleichsListe(i, j) 'Groß-/Kleinschreibung und Leerzeichen beibehalten
            End If
        
        Next j
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            
    Next i
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "SpaltenVerketten"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub DuplikateSuchen() 'VergleichListe nach Duplikaten durchsuchen

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(1, intSprache).Value
        g_strSanduhrNummer = "[3/4]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(4, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
    'Werte zurücksetzen
    strLetztesDuplikat = ""
    lngDuplikatGruppenNummer = 0
    lngDuplikatCounter = 0
    lngOriginalCounter = 0
    lngUnikatCounter = 0
    
    For i = 1 To lngAnzahlZeilen - 1
        
        If varVergleichsListe(i, 0) <> "" Then
        
            'Aktuellen VergleichsWert merken
            strAktuellerWert = varVergleichsListe(i, 0)
            
            blnUnikat = False
        
            If varVergleichsListe(i, intAnzahlSpalten + 1) = "---leer---" Then
                varVergleichsListe(i, intAnzahlSpalten + 1) = "Original" 'vorläufig als Original eintragen
                lngOriginalCounter = lngOriginalCounter + 1
                blnUnikat = True 'vorläufig als Unikat kennzeichnen
            End If
        
            For j = i + 1 To lngAnzahlZeilen
            
                If varVergleichsListe(i, 0) = varVergleichsListe(j, 0) _
                    And varVergleichsListe(j, intAnzahlSpalten + 5) <> "X" Then
                
                    'Wenn neues Duplikat gefunden wird: Duplikatgruppennummer hochzählen
                    If varVergleichsListe(i, 0) <> strLetztesDuplikat Then
                        lngDuplikatGruppenNummer = lngDuplikatGruppenNummer + 1
                    End If
                    
                    'Duplikatgruppennummer in VergleichsListe schreiben
                    varVergleichsListe(j, intAnzahlSpalten + 4) = lngDuplikatGruppenNummer 'Duplikat
                    varVergleichsListe(i, intAnzahlSpalten + 4) = lngDuplikatGruppenNummer 'Original
                    
                    'In VergleichsListe als Duplikat markieren
                    varVergleichsListe(j, intAnzahlSpalten + 1) = "Duplikat"
                                    
                    'Duplikat in der VergleichsListe markieren, damit es nicht nochmal gefunden wird
                    varVergleichsListe(j, intAnzahlSpalten + 5) = "X"
                    
                    'Letztes Duplikat merken
                    strLetztesDuplikat = CStr(varVergleichsListe(i, 0))
                
                    'Duplikatzähler erhöhen (Anzahl der gefundenen Duplikate)
                    lngDuplikatCounter = lngDuplikatCounter + 1
                    
                    'Markierung als Unikat entfernen
                    blnUnikat = False
                
                End If
            Next j
            
            'Wenn der VergleichsWert nur ein einziges mal vorkommt,
            'dann als Unikat in VergleichsListe eintragen und Zähler hochzählen
            If blnUnikat = True Then
                varVergleichsListe(i, intAnzahlSpalten + 1) = "Unikat"
                'Zähler aktualisieren
                lngOriginalCounter = lngOriginalCounter - 1 'Korrektur
                lngUnikatCounter = lngUnikatCounter + 1
            End If
            
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i
    
    'Wenn die letzte markierte Zeile kein Duplikat ist: Prüfung ob Original oder leer
    If varVergleichsListe(lngAnzahlZeilen, intAnzahlSpalten + 1) = "---leer---" _
        And varVergleichsListe(lngAnzahlZeilen, 0) <> "" Then
            varVergleichsListe(lngAnzahlZeilen, intAnzahlSpalten + 1) = "Unikat"
            'Zähler aktualisieren
            lngUnikatCounter = lngUnikatCounter + 1
    End If
    
    'Zählerstatus anpassen
    lblAnzahlDuplikate.Caption = lngDuplikatCounter
    lblAnzahlUnikate.Caption = lngUnikatCounter
    lblAnzahlVerschiedene.Caption = lngOriginalCounter + lngUnikatCounter
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "DuplikateSuchen"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub

Private Sub Farben() 'Legt die Farben für die Markierung der Duplikate auf dem Tabellenblatt fest
    
    'Farbliste anpassen an Anzahl der DuplikatGruppen
    ReDim varFarbListe(1 To lngDuplikatGruppenNummer, 1 To 3)
    
    For i = 1 To lngDuplikatGruppenNummer
        For j = 1 To 3
            
            Randomize 'Zufallszahlengenerator initialisieren
            varFarbListe(i, j) = Int((240 - 110 + 1) * Rnd + 110) 'Zufallsfarbe im RGB-Raum (zwischen 110 und 240)
        
        Next j
    Next i
End Sub

Private Sub VorschauSpaltenbreiten()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Im Duplikatfenster werden maximal die ersten 6 Spalten angezeigt
    If intAnzahlSpalten > 6 Then
        intAnzahlSpaltenDuplikatfenster = 6
    Else
        intAnzahlSpaltenDuplikatfenster = intAnzahlSpalten
    End If

    'String für die Spaltenbreite zurücksetzen
    strVorschauSpaltenBreite = ""

    'String zusammenfügen für die Spaltenbreite des Vorschaufensters
        'Erste Spalte: Spaltenbreite für (x)
        strVorschauSpaltenBreite = strVorschauSpaltenBreite & "0,6cm;" 'ToDo ändern auf pt (wieviel?)
        'ToDo passt noch nicht... Nächste 1-6 Spalten: Spaltenbreiten berechnet für die Spalten der Duplikate
        For j = 1 To intAnzahlSpaltenDuplikatfenster
            strVorschauSpaltenBreite = strVorschauSpaltenBreite & rngSelection.Areas(1).Columns(j).ColumnWidth & "cm;"
        Next j
        'Letzte 3 Spalten minimieren, damit sie nicht zu sehen sind
        'Reihenfolge: Zeile auf dem Tabellenblatt;Spalte auf dem Tabellenblatt;Elementnummer in der VergleichListe
        strVorschauSpaltenBreite = strVorschauSpaltenBreite & "0cm;0cm;0cm" 'Während Entwicklung: "1cm;1cm;1cm" 'ToDo ändern auf "0pt;0pt;0pt"
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "VorschauSpaltenbreiten"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub AusgabeVorschau() 'Liste der Duplikate in Fenster auf dem User-Form ausgeben

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(1, intSprache).Value
        g_strSanduhrNummer = "[4/4]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(5, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
    'Duplikat-Fenster auf dem Userform leeren, Spaltenanzahl und -breiten zuweisen
    boxVorschau.Clear
    boxVorschau.ColumnCount = intAnzahlSpaltenDuplikatfenster + 4 'Die erste Spalte wird für die Markierung
            'ob ein Duplikat gelöscht wurde verwendet (Position 0), dann kommen die Werte der Spalten
            '(Position 1 bis n), und in die letzten 3 Spalten werden die Position des Duplikats auf dem
            'Tabellenblatt (Zeile und Spalte) und die Nummer des Elements in der VergleichsListe geschrieben
    
    boxVorschau.ColumnWidths = strVorschauSpaltenBreite
    
    'VergleichsListe durchlaufen bis keine Einträge mehr drin sind
    For i = 1 To lngAnzahlZeilen
    
        If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
            'Zeile an die ListBox des Duplikat-Fensters anfügen
            boxVorschau.AddItem
            
            For j = 1 To intAnzahlSpaltenDuplikatfenster
                'Duplikate in Duplikat-Fenster auf dem Userform ausgeben
                boxVorschau.List(boxVorschau.ListCount - 1, j) = varVergleichsListe(i, j)
            Next j
            
            'Position der Zelle auf dem Tabellenblatt in die ListBox schreiben
            boxVorschau.List(boxVorschau.ListCount - 1, intAnzahlSpaltenDuplikatfenster + 1) = varVergleichsListe(i, intAnzahlSpalten + 2) 'Zeile
            boxVorschau.List(boxVorschau.ListCount - 1, intAnzahlSpaltenDuplikatfenster + 2) = varVergleichsListe(i, intAnzahlSpalten + 3) 'Spalte
            'Position des Elements in der VergleichsListe in die ListBox schreiben
            boxVorschau.List(boxVorschau.ListCount - 1, intAnzahlSpaltenDuplikatfenster + 3) = i 'Nummer des Eintrags
            
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "AusgabeVorschau"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub


'Farbig markieren auf Tabellenblatt
'----------------------------------

Private Sub DuplikateHervorheben() 'Duplikate und evtl. Originale auf Tabellenblatt farbig hervorheben
                
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Farbe aller markierten Zellen im selektierten Bereich entfernen
    rngSelection.Interior.ColorIndex = xlColorIndexNone
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(6, intSprache).Value
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(7, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
    For i = 1 To lngAnzahlZeilen
    
        If varVergleichsListe(i, 0) <> "" Then
            'Wenn selektierte Bereiche einzeln behandelt werden
            If toggleSuchModus.Value = False Then
                'Original farbig markieren
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Original" _
                    And blnAlleFarbigMarkieren = True Then
                    For k = 1 To intAnzahlSpalten
                        If toggleMarkierModus.Value = False Then 'Zufallsfarben aus der FarbListe
                            ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + k - 1).Interior.Color _
                                = RGB(varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 1), _
                                    varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 2), _
                                    varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 3))
                        Else 'einfarbig (hellblau)
                            ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + k - 1).Interior.Color _
                                = RGB(255, 100, 100)
                        End If
                        
                    Next k
                End If
            
                'Duplikat farbig markieren
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                    For k = 1 To intAnzahlSpalten
                        If toggleMarkierModus.Value = False Then 'Zufallsfarben aus der FarbListe
                            ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + k - 1).Interior.Color _
                                    = RGB(varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 1), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 2), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 3))
                        Else 'einfarbig (hellblau)
                            ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + k - 1).Interior.Color _
                                = RGB(255, 100, 100)
                        End If
                    Next k
                End If
            End If
            
            'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
            If toggleSuchModus.Value = True Then
                'Original farbig markieren
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Original" _
                    And blnAlleFarbigMarkieren = True Then
                    For j = 1 To intAnzahlAreas
                        For k = 1 To rngSelection.Areas(j).Columns.Count
                            If toggleMarkierModus.Value = False Then 'Zufallsfarben aus der FarbListe
                                ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), rngSelection.Areas(j).Columns.Column + k - 1).Interior.Color _
                                    = RGB(varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 1), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 2), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 3))
                            Else 'einfarbig (hellblau)
                                ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), rngSelection.Areas(j).Columns.Column + k - 1).Interior.Color _
                                    = RGB(255, 100, 100)
                            End If
                        Next k
                    Next j
                End If
            
                'Duplikat farbig markieren
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                    For j = 1 To intAnzahlAreas
                        For k = 1 To rngSelection.Areas(j).Columns.Count
                            If toggleMarkierModus.Value = False Then 'Zufallsfarben aus der FarbListe
                                ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), rngSelection.Areas(j).Columns.Column + k - 1).Interior.Color _
                                    = RGB(varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 1), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 2), _
                                        varFarbListe(varVergleichsListe(i, intAnzahlSpalten + 4), 3))
                            Else 'einfarbig (hellblau)
                                ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), rngSelection.Areas(j).Columns.Column + k - 1).Interior.Color _
                                    = RGB(255, 100, 100)
                            End If
                        Next k
                    Next j
                End If
            End If
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    
    Next i
    
    'Blauen Balken in ListBox entfernen
    boxVorschau.ListIndex = -1
    
    'Radiergummi aktivieren ("Farbige Markierungen löschen")
    btnFarbigeMarkierungLoeschen.Enabled = True
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

    Exit Sub

Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

    strFehlerberichtProzedur = "DuplikateHervorheben1"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub UnikateHervorheben() 'Unikate auf Tabellenblatt farbig hervorheben

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Farbe aller markierten Zellen im selektierten Bereich entfernen
    rngSelection.Interior.ColorIndex = xlColorIndexNone
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(6, intSprache).Value
        g_strSanduhrNummer = ""
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(8, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen

    For i = 1 To lngAnzahlZeilen
        If varVergleichsListe(i, 0) <> "" _
            And varVergleichsListe(i, intAnzahlSpalten + 1) = "Unikat" Then
        
            'Wenn selektierte Bereiche einzeln behandelt werden
            If toggleSuchModus.Value = False Then
                For k = 1 To intAnzahlSpalten
                    'Unikate farbig markieren (hellblau)
                    ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + k - 1).Interior.Color _
                            = RGB(100, 150, 255)
                Next k
            End If
            
            'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
            If toggleSuchModus.Value = True Then
                For j = 1 To intAnzahlAreas
                    For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                        'Unikate farbig markieren (hellblau)
                        ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), k).Interior.Color _
                            = RGB(100, 150, 255)
                    Next k
                Next j
            End If
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i
    
    'Blauen Balken in ListBox entfernen
    boxVorschau.ListIndex = -1
    
    'Radiergummi aktivieren ("Farbige Markierungen löschen")
    btnFarbigeMarkierungLoeschen.Enabled = True
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

    Exit Sub
    
Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "UnikateHervorheben"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub


'Ausgabe in neuem Tabellenblatt
'------------------------------

Private Sub AusgabeTabellenblatt(ByRef strAusgabetyp As String, ByRef arrAusgabetyp() As String, ByRef strSheet As String) 'Liste in neuem Tabellenblatt ausgeben
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Leeres Tabellenblatt anfügen und aktivieren
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name _
        = Left(strAusgabetyp & ThisWorkbook.Worksheets("Other_GUI").Cells(28, intSprache).Value & strSheet, 25) & " _" & Worksheets.Count 'Name des Tabellenblattes zusammenbauen (max. 31 Zeichen)
    Worksheets(Worksheets.Count).Activate
    
    'Prüfung ob eine Titelzeile generiert werden soll
    If checkboxAusgabeTitelzeile.Value = True Then
        Cells(1, 1).Value = strAusgabetyp & ThisWorkbook.Worksheets("Other_GUI").Cells(29, intSprache).Value & strSheet & _
                            ThisWorkbook.Worksheets("Other_GUI").Cells(30, intSprache).Value & lblAktuelleMappeWert.Caption & ")"
        lngAktuelleZeile = 3
    Else
        'Zähler für die nächste freie Zeile auf dem Ausgabeblatt setzen
        lngAktuelleZeile = 1
    End If
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(9, intSprache).Value
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(10, intSprache).Value
        g_strSanduhrNummer = ""
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen

    'VergleichsListe durchlaufen bis keine Einträge mehr drin sind
    For i = 1 To lngAnzahlZeilen
        If varVergleichsListe(i, 0) <> "" _
            And (varVergleichsListe(i, intAnzahlSpalten + 1) = arrAusgabetyp(1) Or _
                varVergleichsListe(i, intAnzahlSpalten + 1) = arrAusgabetyp(2)) Then

            For j = 1 To intAnzahlSpalten
                ActiveSheet.Cells(lngAktuelleZeile, j).Value = varVergleichsListe(i, j)
            Next j
            
            'Zähler für die Ausgabe in der nächsten Zeile hochzählen
            lngAktuelleZeile = lngAktuelleZeile + 1
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Popup mit Info anzeigen
    Call AusgabeInfo(strAusgabetyp, intAktuellesTabellenblatt)
    
    Exit Sub
    
Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "AusgabeTabellenblatt"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub

Private Sub AusgabeTabellenblattGanzeZeile(ByRef strAusgabetyp As String, ByRef arrAusgabetyp() As String, ByRef strSheet As String) 'Liste in neuem Tabellenblatt ausgeben (ganze Zeile)
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Leeres Tabellenblatt anfügen und aktivieren
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name _
        = Left(strAusgabetyp & ThisWorkbook.Worksheets("Other_GUI").Cells(28, intSprache).Value & strSheet, 25) & " _" & Worksheets.Count 'Name des Tabellenblattes zusammenbauen (max. 31 Zeichen)
    Worksheets(Worksheets.Count).Activate
    
    'Prüfung ob eine Titelzeile generiert werden soll
    If checkboxAusgabeTitelzeile.Value = True Then
        Cells(1, 1).Value = strAusgabetyp & ThisWorkbook.Worksheets("Other_GUI").Cells(29, intSprache).Value & strSheet & _
                            ThisWorkbook.Worksheets("Other_GUI").Cells(30, intSprache).Value & lblAktuelleMappeWert.Caption & ")"
        lngAktuelleZeile = 3
    Else
        'Zähler für die nächste freie Zeile auf dem Ausgabeblatt setzen
        lngAktuelleZeile = 1
    End If
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(9, intSprache).Value
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(10, intSprache).Value
        g_strSanduhrNummer = ""
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen

    'VergleichsListe durchlaufen bis keine Einträge mehr drin sind
    For i = 1 To lngAnzahlZeilen
        If (varVergleichsListe(i, intAnzahlSpalten + 1) = arrAusgabetyp(1) Or _
            varVergleichsListe(i, intAnzahlSpalten + 1) = arrAusgabetyp(2)) Then
            For j = 1 To Worksheets(strTabellenblatt).Cells.SpecialCells(xlCellTypeLastCell).Column
                Worksheets(Worksheets.Count).Cells(lngAktuelleZeile, j).Value _
                    = Worksheets(strTabellenblatt).Cells(varVergleichsListe(i, intAnzahlSpalten + 2), j)
            Next j

            'Zähler für die Ausgabe in der nächsten Zeile hochzählen
            lngAktuelleZeile = lngAktuelleZeile + 1
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Popup mit Info anzeigen
    Call AusgabeInfo(strAusgabetyp, intAktuellesTabellenblatt)
    
    Exit Sub
    
Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "AusgabeTabellenblattGanzeZeile"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub


'Info über die Ausgabe von Duplikaten/Originalen auf neuem Tabellenblatt
'-------------------------------------------------------------------

Private Sub AusgabeInfo(ByRef strAusgabetyp As String, ByRef intAktuellesTabellenblatt As Integer)
        
    'Pop-up wenn Hinweise aktiviert
    If checkboxHinweis.Value = True Then
        MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(25, intSprache).Value & strAusgabetyp & _
            ThisWorkbook.Worksheets("Messages_GUI").Cells(26, intSprache).Value & vbNewLine & _
            ThisWorkbook.Worksheets("Messages_GUI").Cells(27, intSprache).Value & ActiveSheet.Name & _
            ThisWorkbook.Worksheets("Messages_GUI").Cells(28, intSprache).Value), vbInformation, _
            ThisWorkbook.Worksheets("Messages_GUI").Cells(24, intSprache).Value
    End If
    
    'Zu Worksheet zurückspringen
    Worksheets(intAktuellesTabellenblatt).Activate

End Sub


'Duplikate auf Tabellenblatt löschen und wiederherstellen
'--------------------------------------------------------

Private Sub DuplikateLoeschen() 'Löschen der Zellen mit den gefundenen Duplikaten
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Farbe aller markierten Zellen im selektierten Bereich entfernen
    rngSelection.Interior.ColorIndex = xlColorIndexNone
    
    'Prüfen, ob Button "verdichten" gedrückt wurde
    If toggleLoeschModusKomprimieren.Value = False Then 'Nur Duplikate löschen, Originale nicht verdichten
        
        'Sanduhr neu starten
            g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
            g_strSanduhrNummer = "[1/2]"
            g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(12, intSprache).Value
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen
        
        'Alle Einträge der VergleichsListe durchlaufen
        For i = 1 To lngAnzahlZeilen
            'Wenn selektierte Bereiche einzeln behandelt werden
            If toggleSuchModus.Value = False Then
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                    For j = 1 To intAnzahlSpalten
                        'Zellen mit Duplikaten: Inhalte löschen und Farbe entfernen
                        ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), _
                            varVergleichsListe(i, intAnzahlSpalten + 3) + j - 1).ClearContents
                    Next j
                End If
            End If
            'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
            If toggleSuchModus.Value = True Then
                    If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                        For j = 1 To intAnzahlAreas
                            For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                                'Zellen mit Duplikaten: Inhalte löschen und Farbe entfernen
                                ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), k).ClearContents
                            Next k
                        Next j
                    End If
            End If
            
            'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
        
        'Alle Einträge im Duplikatfenster als gelöscht markieren
            'Sanduhr neu starten
            g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
            g_strSanduhrNummer = "[2/2]"
            g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(13, intSprache).Value
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / boxVorschau.ListCount
            
            For i = 0 To boxVorschau.ListCount - 1
                boxVorschau.List(i, 0) = "(x)"
                'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
        
        'Kennzeichen, dass Zeilen nicht komprimiert wurden
        blnVerdichtenBeimLoeschen = False
        
        'Duplikatfenster aktualisieren
        byteDuplikatfensterStatus = 2 'wiederherstellen
        lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
        lblDuplikatEinzelnLoeschen.Enabled = True
    
    Else: 'Duplikate löschen und Originale verdichten
        
        'Alle Zellen der selektierten Bereiche leeren
        rngSelection.ClearContents
    
        'Zähler für die nächste freie Zeile auf dem Ausgabeblatt zurücksetzen
        lngAktuelleZeile = 1
        
        'Sanduhr neu starten
            g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
            g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(14, intSprache).Value
            g_strSanduhrNummer = ""
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen
        
        'Alle Einträge der VergleichsListe durchlaufen
        For i = 1 To lngAnzahlZeilen
            If (varVergleichsListe(i, intAnzahlSpalten + 1) = "Original" Or _
                    varVergleichsListe(i, intAnzahlSpalten + 1) = "Unikat") Then
                
                'Wenn selektierte Bereiche einzeln behandelt werden
                If toggleSuchModus.Value = False Then
                    For j = 1 To intAnzahlSpalten
                        'Originale aus der VergleichsListe in Tabellenblatt schreiben
                        ActiveSheet.Cells(varVergleichsListe(lngAktuelleZeile, intAnzahlSpalten + 2), _
                            varVergleichsListe(lngAktuelleZeile, intAnzahlSpalten + 3) + j - 1).Value _
                            = varVergleichsListe(i, j)
                    Next j
                End If
                
                'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
                If toggleSuchModus.Value = True Then
                    m = 1 'Zähler zurücksetzen
                    For j = 1 To intAnzahlAreas
                        For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                            'Originale aus der VergleichsListe in Tabellenblatt schreiben
                            ActiveSheet.Cells(varVergleichsListe(lngAktuelleZeile, intAnzahlSpalten + 2), k).Value _
                                = varVergleichsListe(i, m)
                            m = m + 1
                        Next k
                    Next j
                End If
                
                'Zähler in der VergleichsListe hochzählen
                lngAktuelleZeile = lngAktuelleZeile + 1
                        
            End If
            
            'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
        
        'Kennzeichen, dass Zeilen komprimiert wurden
        blnVerdichtenBeimLoeschen = False
        
        'Duplikatfenster aktualisieren
        byteDuplikatfensterStatus = 0 'kein Text
        lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
        lblDuplikatEinzelnLoeschen.Enabled = False
        boxVorschau.Visible = False 'Fenster ausblenden
        
    End If
    
    'Anzahl der gelöschten Duplikate aktualisieren
    lngAnzahlGeloeschteDuplikate = lngDuplikatCounter
    lblAnzahlDuplikate.Caption = lngDuplikatCounter - lngAnzahlGeloeschteDuplikate
    lblAnzahlGeloescht.Caption = lngAnzahlGeloeschteDuplikate
    
    'Löschmodus festlegen (wird beim Wiederherstellen benötigt)
    blnLoeschmodus = False 'nur Duplikate löschen
    
    'Buttons deaktivieren
    btnDuplikateLoeschen.Enabled = False '"Duplikate löschen"
    btnHervorhebenAlle.Enabled = False '"Alle mehrfachen Werte markieren"
    btnHervorhebenDuplikate.Enabled = False '"Nur Duplikate markieren"
    btnHervorhebenUnikate.Enabled = False '"Nur Unikate markieren"
    
    'Button aktivieren
    btnLoeschenRueckgaengig.Enabled = True ' Pfeil zurück ("Alle Duplikate wiederherstellen")
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    lngEindeutigeWerte = lngOriginalCounter + lngUnikatCounter
    
    'Hinweis anzeigen, wenn Checkbox aktiv
    If checkboxHinweis.Value = True Then
        If lngDuplikatCounter = 1 And lngEindeutigeWerte = 1 Then
            MsgBox (lngDuplikatCounter & ThisWorkbook.Worksheets("Messages_GUI").Cells(31, intSprache).Value & vbCrLf & _
                lngEindeutigeWerte & ThisWorkbook.Worksheets("Messages_GUI").Cells(32, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
        ElseIf lngDuplikatCounter <> 1 And lngEindeutigeWerte = 1 Then
            MsgBox (lngDuplikatCounter & ThisWorkbook.Worksheets("Messages_GUI").Cells(33, intSprache).Value & vbCrLf & _
                lngEindeutigeWerte & ThisWorkbook.Worksheets("Messages_GUI").Cells(32, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
        ElseIf lngDuplikatCounter = 1 And lngEindeutigeWerte <> 1 Then
            MsgBox (lngDuplikatCounter & ThisWorkbook.Worksheets("Messages_GUI").Cells(31, intSprache).Value & vbCrLf & _
                lngEindeutigeWerte & ThisWorkbook.Worksheets("Messages_GUI").Cells(34, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
        Else:
            MsgBox (lngDuplikatCounter & ThisWorkbook.Worksheets("Messages_GUI").Cells(33, intSprache).Value & vbCrLf & _
                lngEindeutigeWerte & ThisWorkbook.Worksheets("Messages_GUI").Cells(34, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
        End If
    End If
        
    Exit Sub
        
Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    If Err.Number = 1004 Then
        'Bereits gelöschte Teile wiederherstellen
        Call DuplikateLoeschenUnDo
        'Info
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(20, intSprache).Value & Err.Description
    Else:
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(19, intSprache).Value
        strFehlerberichtProzedur = "DuplikateLoeschen"
        strFehlerberichtErrNumber = Err.Number
        strFehlerberichtErrDescription = Err.Description
        Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    End If

End Sub

Private Sub DuplikateLoeschenGanzeZeile() 'Löschen der kompletten Zeilen mit den gefundenen Duplikaten
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
    'Farbe aller markierten Zellen im selektierten Bereich entfernen
    rngSelection.Interior.ColorIndex = xlColorIndexNone
    
    'Prüfen, ob Button "verdichten" gedrückt wurde
    If toggleLoeschModusKomprimieren.Value = False Then 'Nur Duplikate löschen, Originale nicht verdichten
        
        'Wenn selektierte Bereiche einzeln behandelt werden
        If toggleSuchModus.Value = False Then
        
            'Alle Werte der Zeilen, in denen ein Duplikat steht, in Array sichern
            Call SelektionSpeichern3
            
            'Sanduhr neu starten
                g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
                g_strSanduhrNummer = "[5/6]"
                g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(15, intSprache).Value
                'Fortschrittsbalken zurücksetzen
                Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varListeZeilen)
            
            'Alle Zeilen mit Duplikaten durchlaufen und komplett leeren
                k = 0 'Wert der aktuell höchsten Zeile zurücksetzen
                'Liste mit den Zeilen, in denen Duplikate stehen, durchlaufen
                For i = 1 To UBound(varListeZeilen)
                    If varListeZeilen(i) > k Then
                        'Zeilen mit Duplikaten: Komplette Zeile leeren
                        Worksheets(strTabellenblatt).Rows(varListeZeilen(i)).ClearContents
                        k = varListeZeilen(i) 'Aktuell höchste Zeile neu setzen
                    End If
                    
                    'Sanduhr aktualisieren
                        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
                Next i
        End If
        
        'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
        If toggleSuchModus.Value = True Then
            'Alle Werte der Zeilen, in denen ein Duplikat steht, in Array sichern
            Call SelektionSpeichern2
            
            'Sanduhr neu starten
                g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
                g_strSanduhrNummer = "[3/6]"
                g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
                'Fortschrittsbalken zurücksetzen
                Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen
            
            'Alle Einträge der VergleichsListe durchlaufen und Zeilen mit Duplikaten komplett leeren
            For i = 1 To lngAnzahlZeilen
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                    'Zeilen mit Duplikaten: Komplette Zeile leeren
                    Worksheets(strTabellenblatt).Rows(varVergleichsListe(i, intAnzahlSpalten + 2)).ClearContents
                End If
                
                'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
            Next i
        End If
        
        'Sanduhr neu starten
            g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
            g_strSanduhrNummer = "[6/6]"
            g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(13, intSprache).Value
        'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)

    'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / boxVorschau.ListCount

        
        'Alle Einträge im Duplikatfenster als gelöscht markieren
        For i = 0 To boxVorschau.ListCount - 1
            boxVorschau.List(i, 0) = "(x)"
            'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        Next i
        
        'Kennzeichen, dass Zeilen nicht komprimiert wurden
        blnVerdichtenBeimLoeschen = False
        
        'Duplikatfenster aktualisieren
        byteDuplikatfensterStatus = 2 'wiederherstellen
        lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
        lblDuplikatEinzelnLoeschen.Visible = False

    
    Else: 'Duplikate löschen und Originale verdichten

        'Zähler für die Anzahl gelöschter Zeilen zurücksetzen
        lngAnzahlGeloeschteZeilen = 0
        
        'Array neu dimensionieren: Für jede gelöschte Zeile eine Zeile
        ReDim varListeGeloeschteZeilen(1 To lngDuplikatCounter, 1 To 1)
        
        'Wenn selektierte Bereiche einzeln behandelt werden
        If toggleSuchModus.Value = False Then
        
            'Alle Werte der Zeilen, in denen ein Duplikat steht, in Array sichern
            Call SelektionSpeichern3
            
            'Sanduhr neu starten
                g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
                g_strSanduhrNummer = "[5/5]"
                g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(16, intSprache).Value
                'Fortschrittsbalken zurücksetzen
                Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varListeZeilen)
            
            'Alle Zeilen mit Duplikaten durchlaufen und entfernen
                k = 0 'Wert der aktuell höchsten Zeile zurücksetzen
                'Liste mit den Zeilen, in denen Duplikate stehen, durchlaufen
                For i = 1 To UBound(varListeZeilen)
                    If varListeZeilen(i) > k Then
                    
                        'Nummer der gelöschten Zeile in Liste eintragen
                        varListeGeloeschteZeilen(lngAnzahlGeloeschteZeilen + 1, 1) = varListeZeilen(i)
                        
                        'Komplette Zeile löschen
                        Worksheets(strTabellenblatt).Rows(varListeZeilen(i) - lngAnzahlGeloeschteZeilen).Delete
                        lngAnzahlGeloeschteZeilen = lngAnzahlGeloeschteZeilen + 1 'Zähler hochsetzen
                        
                        k = varListeZeilen(i) 'Aktuell höchste Zeile neu setzen
                    End If
                    
                    'Sanduhr aktualisieren
                        g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                        Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
                Next i
                
                Set rngSelection = Selection 'Selektierten Bereich in Variable einlesen
                
        End If
        
        'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
        If toggleSuchModus.Value = True Then
            'Alle Werte der Zeilen, in denen ein Duplikat steht, in Array sichern
            Call SelektionSpeichern2
            
            'Sanduhr neu starten
                g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
                g_strSanduhrNummer = "[3/5]"
                g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(12, intSprache).Value
                'Fortschrittsbalken zurücksetzen
                Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
            'Alle Einträge der VergleichsListe durchlaufen
            For i = 1 To lngAnzahlZeilen
                If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
                    
                    'Nummer der gelöschten Zeile in Liste eintragen
                    varListeGeloeschteZeilen(lngAnzahlGeloeschteZeilen + 1, 1) = varVergleichsListe(i, intAnzahlSpalten + 2)
                    
                    'Komplette Zeile löschen
                    Worksheets(strTabellenblatt).Rows(varVergleichsListe(i, intAnzahlSpalten + 2) - lngAnzahlGeloeschteZeilen).Delete
                    lngAnzahlGeloeschteZeilen = lngAnzahlGeloeschteZeilen + 1 'Zähler hochsetzen
                End If
                
                'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    
            Next i
        End If
        
        'Kennzeichen, dass Zeilen komprimiert wurden
        blnVerdichtenBeimLoeschen = True
        
        'Duplikatfenster aktualisieren
        byteDuplikatfensterStatus = 0 'kein Text
        lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
        lblDuplikatEinzelnLoeschen.Enabled = False
        boxVorschau.Visible = False 'Fenster ausblenden
        
    End If
    
    'Anzahl der gelöschten Duplikate aktualisieren
    lngAnzahlGeloeschteDuplikate = lngDuplikatCounter
    lblAnzahlDuplikate.Caption = lngDuplikatCounter - lngAnzahlGeloeschteDuplikate
    lblAnzahlGeloescht.Caption = lngAnzahlGeloeschteDuplikate
    
    'Löschmodus festlegen (wird beim Wiederherstellen benötigt)
    blnLoeschmodus = True 'ganze Zeilen löschen
    
    'Buttons deaktivieren
    btnDuplikateLoeschen.Enabled = False '"Duplikate löschen"
    btnHervorhebenAlle.Enabled = False '"Alle mehrfachen Werte markieren"
    btnHervorhebenDuplikate.Enabled = False '"Nur Duplikate markieren"
    btnHervorhebenUnikate.Enabled = False '"Nur Unikate markieren"
    
    'Button aktivieren
    btnLoeschenRueckgaengig.Enabled = True 'Pfeil zurück ("Alle Duplikate wiederherstellen")
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden

    lngEindeutigeWerte = lngOriginalCounter + lngUnikatCounter
    
    'Hinweis anzeigen, wenn Checkbox aktiv
    If checkboxHinweis.Value = True Then
        MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(35, intSprache).Value), _
            vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
    End If
        
    Exit Sub
        
Fehlerbehandlung:
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    If Err.Number = 1004 Then
        'Bereits gelöschte Teile wiederherstellen
        Call DuplikateLoeschenUnDo
        'Info
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(20, intSprache).Value & Err.Description
    Else:
        strFehlerberichtProzedur = "DuplikateLoeschenGanzeZeile"
        strFehlerberichtErrNumber = Err.Number
        strFehlerberichtErrDescription = Err.Description
        Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    End If

End Sub

Private Sub SelektionSpeichern2() 'Alle Zellen einer Zeile, die ein Duplikat enthält, in Array speichern, um die Werte wiederherstellen zu können

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
        g_strSanduhrNummer = "[1/5]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(17, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen

    'Anzahl der benötigten Spalten des Arrays bestimmen: Pro Zelle eine Zeile
    m = 0 'Zähler für die Anzahl der Zellen zurücksetzen
    'Alle Einträge der VergleichsListe durchlaufen
    For i = 1 To lngAnzahlZeilen
        'Anzahl der gefüllten Zellen in jeder Spalte hochzählen
        If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
            m = m + Cells(varVergleichsListe(i, intAnzahlSpalten + 2), Columns.Count).End(xlToLeft).Column
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
        
    'Array neu dimensionieren: Pro Zelle eine Zeile, 3 Spalten für Zeile, Spalte und Inhalt der Zelle
    ReDim varSaveAll2(1 To m, 1 To 3)
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
        g_strSanduhrNummer = "[2/5]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(17, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
        
    'Array füllen
    m = 0 'Zeilenzähler zurücksetzen
    For i = 1 To lngAnzahlZeilen
        If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
            'Alle Spalten eine Zeile mit Duplikat durchlaufen
            For j = 1 To Cells(varVergleichsListe(i, intAnzahlSpalten + 2), Columns.Count).End(xlToLeft).Column
                m = m + 1 'Zeilenzähler hochzählen
                varSaveAll2(m, 1) = varVergleichsListe(i, intAnzahlSpalten + 2) 'Zeile
                varSaveAll2(m, 2) = j 'Spalte
                varSaveAll2(m, 3) = Cells(varVergleichsListe(i, intAnzahlSpalten + 2), j).Value 'Inhalt
            Next j
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
    Exit Sub
    
Fehlerbehandlung:

    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "SelektionSpeichern2"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub SelektionSpeichern3() 'Alle Zellen einer Zeile, die ein Duplikat enthält, in Array speichern, um die Werte wiederherstellen zu können

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
        g_strSanduhrNummer = "[1/6]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(17, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / lngAnzahlZeilen
    
    'Liste mit allen Zeilen, in denen ein Duplikat steht, neu dimensionieren und aufsteigend sortieren
    ReDim varListeZeilen(1 To lngDuplikatCounter)
    m = 1 'Zähler zurücksetzen
    'Alle Einträge der VergleichsListe durchlaufen
    For i = 1 To lngAnzahlZeilen
        'Zeilennummer mit den Duplikaten eintragen
        If varVergleichsListe(i, intAnzahlSpalten + 1) = "Duplikat" Then
            varListeZeilen(m) = varVergleichsListe(i, intAnzahlSpalten + 2)
            m = m + 1 'Zähler hochzählen
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i
    
    'Aufsteigend sortieren
    varListeZeilen = modArraySortieren.BubbleSort(varListeZeilen)
    
    'Anzahl der benötigten Spalten des Arrays bestimmen: Pro Zelle eine Zeile
        m = 0 'Zähler für die Anzahl der benötigten Zellen zurücksetzen
        k = 0 'Wert der aktuell höchsten Zeile zurücksetzen
        
        'Sanduhr neu starten
            g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
            g_strSanduhrNummer = "[3/6]"
            g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(17, intSprache).Value
            'Fortschrittsbalken zurücksetzen
            Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / UBound(varListeZeilen)
        
        'Liste mit den Zeilen, in denen Duplikate stehen, durchlaufen
        For i = 1 To UBound(varListeZeilen)
            If varListeZeilen(i) > k Then
                m = m + Cells(varListeZeilen(i), Columns.Count).End(xlToLeft).Column 'benötigte Zellen aufaddieren
                k = varListeZeilen(i) 'Aktuell höchste Zeile neu setzen
            End If
              
            'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
              
        Next i
        
    'Array neu dimensionieren: Pro Zelle eine Zeile, 3 Spalten für Zeile, Spalte und Inhalt der Zelle
    ReDim varSaveAll2(1 To m, 1 To 3)

    'Array füllen
    m = 0 'Zeilenzähler zurücksetzen
    k = 0 'Wert der aktuell höchsten Zeile zurücksetzen
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(11, intSprache).Value
        g_strSanduhrNummer = "[4/6]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(17, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
        'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / UBound(varListeZeilen)
    
    'Liste mit den Zeilen, in denen Duplikate stehen, durchlaufen
    For i = 1 To UBound(varListeZeilen)
        If varListeZeilen(i) > k Then
            For j = 1 To Cells(varListeZeilen(i), Columns.Count).End(xlToLeft).Column
                m = m + 1 'Zeilenzähler hochzählen
                varSaveAll2(m, 1) = varListeZeilen(i) 'Zeile
                varSaveAll2(m, 2) = j 'Spalte
                varSaveAll2(m, 3) = Cells(varListeZeilen(i), j).Value 'Inhalt
            Next j
            k = varListeZeilen(i) 'Aktuell höchste Zeile neu setzen
        End If
        
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
        
    Next i
    
    Exit Sub
    
Fehlerbehandlung:

    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "SelektionSpeichern3"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub DuplikateLoeschenUnDo() 'Duplikate löschen zurücknehmen

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Sanduhr einblenden
    Call SanduhrEinblenden
    
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(18, intSprache).Value
        g_strSanduhrNummer = "[1/2]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(19, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
    
    'Bildschirmaktualisierung ausschalten (bessere Performance)
    If checkboxAktualisierung.Value = False Then
        Application.ScreenUpdating = False
    End If
    
        rngSelection.Interior.Pattern = xlSolid 'Schraffierung auf dem Tabellenblatt entfernen
        boxVorschau.ListIndex = -1 'Blauen Balken in ListBox entfernen
    
    'Wenn selektierte Bereiche einzeln behandelt werden
    If toggleSuchModus.Value = False Then
    
        'Wenn nur die Duplikate gelöscht wurden
        If blnLoeschmodus = False Then
            
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen
        
            'Alle Einträge der VergleichsListe durchlaufen
            For i = 1 To lngAnzahlZeilen
                For j = 1 To intAnzahlSpalten
                    'Inhalte aus der VergleichsListe wiederherstellen
                    Worksheets(strTabellenblatt).Cells(varVergleichsListe(i, intAnzahlSpalten + 2), varVergleichsListe(i, intAnzahlSpalten + 3) + j - 1) _
                        .Value = varVergleichsListe(i, j)
                Next j
                
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

            Next
        End If
        
        'Wenn ganze Zeilen gelöscht wurden
        If blnLoeschmodus = True Then
        
            'Wenn Zeilen nur geleert, aber nicht komprimiert wurden
            If blnVerdichtenBeimLoeschen = False Then
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varSaveAll2, 1)
            
                m = 1 'Zähler zurücksetzen
                For j = 1 To UBound(varSaveAll2, 1) 'Liste durchlaufen und Werte auf Tabellenblatt zurückschreiben
                    Worksheets(strTabellenblatt).Cells(varSaveAll2(m, 1), varSaveAll2(m, 2)).Value = varSaveAll2(m, 3)
                    m = m + 1
                    
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
            End If
            
            'Wenn Zeilen komprimiert wurden: Zeilen wieder einfügen
            If blnVerdichtenBeimLoeschen = True Then
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varListeGeloeschteZeilen, 1)
                
                For j = 1 To UBound(varListeGeloeschteZeilen, 1) 'ganze Liste durchlaufen
                    If varListeGeloeschteZeilen(j, 1) <> "" Then
                        Worksheets(strTabellenblatt).Rows(varListeGeloeschteZeilen(j, 1)).Insert
                    End If
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
                
                'Fortschrittsbalken zurücksetzen
                Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)
                
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varSaveAll2, 1)

                m = 1 'Zähler zurücksetzen
                For j = 1 To UBound(varSaveAll2, 1) 'Liste durchlaufen und Werte auf Tabellenblatt zurückschreiben
                    Worksheets(strTabellenblatt).Cells(varSaveAll2(m, 1), varSaveAll2(m, 2)).Value = varSaveAll2(m, 3)
                    m = m + 1
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
            End If
        End If
    End If
        
    'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
    If toggleSuchModus.Value = True Then
        'Wenn nur die Duplikate gelöscht wurden
        If blnLoeschmodus = False Then
            'Stückelung des Balkens berechnen
            g_dblBalkenAnteil = 100 / lngAnzahlZeilen

            'Alle Einträge der VergleichsListe durchlaufen
            For i = 1 To lngAnzahlZeilen
                m = 1 'Zähler zurücksetzen
                For j = 1 To intAnzahlAreas
                    For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                        'Inhalte aus der VergleichsListe wiederherstellen
                        ActiveSheet.Cells(varVergleichsListe(i, intAnzahlSpalten + 2), k).Value = varVergleichsListe(i, m)
                    m = m + 1
                    Next k
                Next j
                'Sanduhr aktualisieren
                g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

            Next i
        End If
        
        'Wenn ganze Zeilen gelöscht wurden
        If blnLoeschmodus = True Then
            'Wenn Zeilen nur geleert, aber nicht komprimiert wurden
            If blnVerdichtenBeimLoeschen = False Then
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varSaveAll2, 1)

                m = 1 'Zähler zurücksetzen
                For j = 1 To UBound(varSaveAll2, 1) 'Liste durchlaufen und Werte auf Tabellenblatt zurückschreiben
                    Worksheets(strTabellenblatt).Cells(varSaveAll2(m, 1), varSaveAll2(m, 2)).Value = varSaveAll2(m, 3)
                    m = m + 1
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
            End If
            
            'Wenn Zeilen komprimiert wurden: Zeilen wieder einfügen
            If blnVerdichtenBeimLoeschen = True Then
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varListeGeloeschteZeilen, 1)
                
                For j = 1 To UBound(varListeGeloeschteZeilen, 1) 'ganze Liste durchlaufen
                    Worksheets(strTabellenblatt).Rows(varListeGeloeschteZeilen(j, 1)).Insert
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
                
                'Stückelung des Balkens berechnen
                g_dblBalkenAnteil = 100 / UBound(varSaveAll2, 1)
                
                m = 1 'Zähler zurücksetzen
                For j = 1 To UBound(varSaveAll2, 1) 'Liste durchlaufen und Werte auf Tabellenblatt zurückschreiben
                    Worksheets(strTabellenblatt).Cells(varSaveAll2(m, 1), varSaveAll2(m, 2)).Value = varSaveAll2(m, 3)
                    m = m + 1
                    'Sanduhr aktualisieren
                    g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
                    Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren

                Next j
            End If
        End If
    End If
    
    'Alle Einträge im Duplikatfenster als nicht gelöscht markieren
    'Sanduhr neu starten
        g_strSanduhrAktion = ThisWorkbook.Worksheets("Balken_GUI").Cells(18, intSprache).Value
        g_strSanduhrNummer = "[2/2]"
        g_strSanduhrSchritt = ThisWorkbook.Worksheets("Balken_GUI").Cells(13, intSprache).Value
        'Fortschrittsbalken zurücksetzen
        Call FortschrittsbalkenReset(g_strSanduhrAktion, g_strSanduhrNummer, g_strSanduhrSchritt)

    'Stückelung des Balkens berechnen
        g_dblBalkenAnteil = 100 / boxVorschau.ListCount

    For i = 0 To boxVorschau.ListCount - 1
        boxVorschau.List(i, 0) = ""
        'Sanduhr aktualisieren
            g_dblBalkenAktuell = g_dblBalkenAktuell + g_dblBalkenAnteil 'Aktuelle Balkenlänge berechnen
            Call FortschrittsbalkenAktualisieren(g_dblBalkenAktuell) 'Sanduhr aktualisieren
    Next i
    
    'Duplikatfenster einblenden
    boxVorschau.Visible = True
    
    'Anzahl der gelöschten Duplikate aktualisieren
    lngAnzahlGeloeschteDuplikate = 0
    lblAnzahlDuplikate.Caption = lngDuplikatCounter - lngAnzahlGeloeschteDuplikate
    lblAnzahlGeloescht.Caption = lngAnzahlGeloeschteDuplikate
    
    'Button deaktivieren
    btnLoeschenRueckgaengig.Enabled = False
    
    'Buttons aktivieren
    btnDuplikateLoeschen.Enabled = True '"Duplikate löschen"
    btnHervorhebenAlle.Enabled = True '"Alle mehrfachen Werte markieren"
    btnHervorhebenDuplikate.Enabled = True '"Nur Duplikate markieren"
    btnHervorhebenUnikate.Enabled = True '"Nur Unikate markieren"
    
    'Duplikatfenster aktualisieren
    lblDuplikatEinzelnLoeschen.Visible = False
    byteDuplikatfensterStatus = 1 'löschen
    lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
    
    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    'Hinweis anzeigen, wenn Checkbox aktiv
    If checkboxHinweis.Value = True Then
        If blnLoeschmodus = False Then
            MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(41, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(40, intSprache).Value
        Else
            MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(42, intSprache).Value), _
                vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(40, intSprache).Value
        End If
    End If
    
    Exit Sub
    
Fehlerbehandlung:

    'Bildschirmaktualisierung einschalten
    Application.ScreenUpdating = True
    
    'Sanduhr ausblenden
    Call SanduhrAusblenden
    
    strFehlerberichtProzedur = "DuplikateLoeschenUnDo"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub


'Aktionen im Duplikatfenster
'---------------------------

Private Sub DuplikatfensterAnklickenDuplikat() 'Anklicken eines einzelnen Duplikats in Duplikatfenster

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    'Letzte Markierung löschen
    rngSelection.Interior.Pattern = xlSolid
    
    For i = 0 To boxVorschau.ListCount - 1
        If boxVorschau.Selected(i) Then
            'Angeklicktes Duplikat auf dem Tabellenblatt mit Schraffierung markieren
                'Wenn selektierte Bereiche einzeln behandelt werden
                If toggleSuchModus.Value = False Then
                    For j = 0 To intAnzahlSpalten - 1
                        ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), _
                            j + CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2))).Interior.Pattern = xlGray25
                    Next j
                End If
                'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
                If toggleSuchModus.Value = True Then
                    For j = 1 To intAnzahlAreas
                        For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                            ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), _
                                k).Interior.Pattern = xlGray25
                        Next k
                    Next j
                End If
            
            'Wenn das angeklickte Duplikat schon gelöscht wurde
            If boxVorschau.List(i, 0) = "(x)" Then
                  byteDuplikatfensterStatus = 2 'wiederherstellen
                  lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
                  Exit Sub
            End If
            
            'Scrollen zum angeklickten Duplikat wenn nötig
              'wenn Duplikat unterhalb oder oberhalb des sichtbaren Bereichs liegt
              If CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)) _
                    > ActiveWindow.VisibleRange.Rows(ActiveWindow.VisibleRange.Rows.Count).Row _
                 Or CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)) _
                    < ActiveWindow.VisibleRange.Rows.Row Then
                    
                    ActiveWindow.ScrollRow = CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1))
              End If
                            
              'wenn Duplikat links oder rechts des sichtbaren Bereichs liegt
              If CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2)) _
                    > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column _
                 Or CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2)) _
                    < ActiveWindow.VisibleRange.Columns.Column Then
                    
                    ActiveWindow.ScrollColumn = CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2))
              End If
        End If

        'Button "löschen" im Vorschaufenster aktivieren
        lblDuplikatEinzelnLoeschen.Enabled = True
        byteDuplikatfensterStatus = 1 'löschen
        lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(intSprache, byteDuplikatfensterStatus)
    Next i
    
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "DuplikatfensterAnklickenDuplikat"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    
End Sub

Private Sub DuplikatEinzelnLoeschen()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung

        For i = 0 To boxVorschau.ListCount - 1
            If boxVorschau.Selected(i) Then
                'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
                If toggleSuchModus.Value = True Then
                    'Duplikat auf dem Tabellenblatt löschen
                    For j = 1 To intAnzahlAreas
                        For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                            'Zellen mit Duplikaten: Inhalte löschen und Farbe entfernen
                            ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), _
                                k).ClearContents
                        Next k
                    Next j
                Else 'Wenn selektierte Bereiche einzeln behandelt werden
                    'Duplikat auf dem Tabellenblatt löschen
                    For j = 0 To intAnzahlSpalten - 1
                        ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), _
                          j + CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2))).ClearContents
                    Next j
                End If
                  
                'Duplikat als gelöscht markieren
                boxVorschau.List(i, 0) = "(x)"
                'Anzahl der gelöschten Duplikate aktualisieren
                lngAnzahlGeloeschteDuplikate = lngAnzahlGeloeschteDuplikate + 1
                lblAnzahlDuplikate.Caption = lngDuplikatCounter - lngAnzahlGeloeschteDuplikate
                lblAnzahlGeloescht.Caption = lngAnzahlGeloeschteDuplikate
                
                'Buttons anpassen
                btnLoeschenRueckgaengig.Enabled = True
                If lngDuplikatCounter > lngAnzahlGeloeschteDuplikate Then
                    btnDuplikateLoeschen.Enabled = True
                Else
                    btnDuplikateLoeschen.Enabled = False
                End If
                'Hinweis anzeigen, wenn Checkbox aktiv
                If checkboxHinweis.Value = True Then
                    MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(36, intSprache).Value), _
                        vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(30, intSprache).Value
                End If
                
                Exit Sub
            End If
        Next i
    Exit Sub
    
Fehlerbehandlung:
    If Err.Number = 1004 Then
        'Info
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(20, intSprache).Value & Err.Description
        'Bereits gelöschte Teile wiederherstellen
        Call DuplikatEinzelnWiederherstellen
        'Button anpassen
        lblDuplikatEinzelnLoeschen.Caption = ""
    Else:
        MsgBox ThisWorkbook.Worksheets("Messages_GUI").Cells(19, intSprache).Value
        strFehlerberichtProzedur = "DuplikatEinzelnLoeschen"
        strFehlerberichtErrNumber = Err.Number
        strFehlerberichtErrDescription = Err.Description
        Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
    End If
End Sub

Private Sub DuplikatEinzelnWiederherstellen()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    For i = 0 To boxVorschau.ListCount - 1
            If boxVorschau.Selected(i) Then
                'Wenn selektierte Spalten zu 1 Bereich vereinigt werden
                If toggleSuchModus.Value = True Then
                    'Gelöschtes Duplikat auf dem Tabellenblatt wiederherstellen
                    m = 1 'Zähler zurücksetzen
                    For j = 1 To intAnzahlAreas
                        For k = rngSelection.Areas(j).Column To rngSelection.Areas(j).Column + rngSelection.Areas(j).Columns.Count - 1
                            ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), k) _
                                = boxVorschau.List(i, m)
                            m = m + 1
                        Next k
                    Next j
                Else 'Wenn selektierte Bereiche einzeln behandelt werden
                    'Gelöschtes Duplikat auf dem Tabellenblatt wiederherstellen
                    For j = 0 To intAnzahlSpalten - 1
                        ActiveSheet.Cells(CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 1)), _
                          j + CLng(boxVorschau.List(i, intAnzahlSpaltenDuplikatfenster + 2))) _
                          = boxVorschau.List(i, j + 1)
                    Next j
                End If
                  
                'Markierung als gelöschtes Duplikat entfernen
                boxVorschau.List(i, 0) = ""
                'Anzahl der gelöschten Duplikate aktualisieren
                lngAnzahlGeloeschteDuplikate = lngAnzahlGeloeschteDuplikate - 1
                lblAnzahlDuplikate.Caption = lngDuplikatCounter - lngAnzahlGeloeschteDuplikate
                lblAnzahlGeloescht.Caption = lngAnzahlGeloeschteDuplikate
                
                'Buttons anpassen
                btnDuplikateLoeschen.Enabled = True 'Button "Duplikate löschen"
                If lngAnzahlGeloeschteDuplikate > 0 Then
                    btnLoeschenRueckgaengig.Enabled = True ' Pfeil zurück ("Alle Duplikate wiederherstellen")
                Else
                    btnLoeschenRueckgaengig.Enabled = False
                End If
                'Hinweis anzeigen, wenn Checkbox aktiv
                If checkboxHinweis.Value = True Then
                    MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(43, intSprache).Value), _
                        vbInformation, ThisWorkbook.Worksheets("Messages_GUI").Cells(40, intSprache).Value
                End If
                
                Exit Sub
            End If
        Next i
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "DuplikatEinzelnWiederherstellen"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)

End Sub

Private Sub Tooltips()
    If checkboxTooltip.Value = True Then
            Call modGUItooltips.TooltipsON(intSprache)
        Else
            Call modGUItooltips.TooltipsOFF
        End If
End Sub

Private Sub Sprache(s As Integer)
    Call modGUItexte.spracheAendern(s) 'GUI-Beschriftungen
    lblDuplikatEinzelnLoeschen.Caption = modGUIbuttons.buttonDuplikatfenster(s, byteDuplikatfensterStatus)
    Call Tooltips
End Sub

Private Sub SpeicherortOptionenOeffnen()
    On Error Resume Next 'z.B. wenn anderes Betriebssystem als Windows
    Shell "explorer.exe " & g_strSaveDateipfad, vbNormalFocus
End Sub

Public Sub SuchModus() 'GUI-Beschriftung anpassen
    toggleSuchModus.Caption = modGUIbuttons.buttonSuchModus(intSprache, toggleSuchModus)
End Sub

Public Sub LoeschModusZeilen() 'GUI-Beschriftung anpassen
    toggleLoeschModusZeilen.Caption = modGUIbuttons.buttonLoeschModusZeilen(intSprache, toggleLoeschModusZeilen)
End Sub

Public Sub LoeschModusKomprimieren() 'GUI-Beschriftung anpassen
    toggleLoeschModusKomprimieren.Caption = modGUIbuttons.buttonLoeschModusKomprimieren(intSprache, toggleLoeschModusKomprimieren)
End Sub

Private Sub MinimierenMaximieren() 'Größe und Position der GUI anpassen
    Select Case True
        Case Me.Height = 398 'wenn maximiert
            Me.Width = 270
            Me.Height = 76
        Case Else 'wenn minimiert
            If toggleDuplikatfensterOnOff.Value = False Then 'wenn Duplikatfenster ausgeklappt
                Me.Width = 504
            Else 'wenn Duplikatfenster zugeklappt
                Me.Width = 335
            End If
            Me.Height = 398
    End Select
End Sub

Private Sub DuplikatfensterOnOff() 'Größe der GUI anpassen
    If toggleDuplikatfensterOnOff.Value = True Then
        With Me
            .Width = 335
            .Height = 398
        End With
        toggleDuplikatfensterOnOff.Caption = modGUIbuttons.buttonDuplikatfensterOnOff(intSprache, toggleDuplikatfensterOnOff)
    Else
        With Me
            .Width = 504
            .Height = 398
        End With
        toggleDuplikatfensterOnOff.Caption = modGUIbuttons.buttonDuplikatfensterOnOff(intSprache, toggleDuplikatfensterOnOff)
    End If
End Sub

Public Sub AusgabeModus() 'GUI-Beschriftung anpassen
    toggleAusgabeModus.Caption = modGUIbuttons.buttonAusgabeModus(intSprache, toggleAusgabeModus)
End Sub

Public Sub MarkierModus()
    toggleMarkierModus.Caption = modGUIbuttons.buttonMarkierModus(intSprache, toggleMarkierModus)
End Sub


'E-Mails
'-------

Private Sub eMail() 'Feedback E-Mail

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    Set objMail = CreateObject("Shell.Application")
    objMail.ShellExecute "mailto:" & "excel@marco-krapf.de" _
        & "&subject=" & "Feedback: xl Duplikat-Manager Version " & cstrVersion & " / " _
        & Application.OperatingSystem & " / Excel-Version " & Application.Version
        
    Exit Sub
    
Fehlerbehandlung:
    strFehlerberichtProzedur = "eMail"
    strFehlerberichtErrNumber = Err.Number
    strFehlerberichtErrDescription = Err.Description
    Call mailFehlerbericht(strFehlerberichtProzedur, strFehlerberichtErrNumber, strFehlerberichtErrDescription)
        
End Sub

Public Sub mailFehlerbericht(ByRef errProz As String, ByRef errnum As String, ByRef errDesc As String)
'E-Mail mit Fehlerbericht senden wenn Programm abstürzt
    
    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
        
    If MsgBox(ThisWorkbook.Worksheets("Messages_GUI").Cells(48, intSprache).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Messages_GUI").Cells(49, intSprache).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Messages_GUI").Cells(50, intSprache).Value, _
                vbYesNo, ThisWorkbook.Worksheets("Messages_GUI").Cells(47, intSprache).Value) = vbYes Then
    
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & "excel@marco-krapf.de" _
            & "&subject=" & "Fehlerbericht: xl Duplikat-Manager" _
            & "&body=" _
            & "xl DuplikatManager Version: " & cstrVersion & "%0A" _
            & "Betriebssystem: " & Application.OperatingSystem & "%0A" _
            & "Excel-Version: " & Application.Version & "%0A" _
            & "Prozedur: " & errProz & "%0A" _
            & "Fehlernummer: " & errnum & "%0A" _
            & "Fehlerbeschreibung: " & errDesc & "%0A"
    End If
    
    Exit Sub
    
Fehlerbehandlung:
    MsgBox (ThisWorkbook.Worksheets("Messages_GUI").Cells(52, intSprache).Value), _
            vbCritical, "xl DuplikatManager (Version " & cstrVersion & ")"
    
End Sub


'Sanduhr
'-------

Private Sub SanduhrEinblenden() 'Sanduhr einblenden
    Load frmSanduhr
    Call modUserformPlatzieren.UserFormPlatzieren(frmSanduhr)
    frmSanduhr.Show
End Sub

Private Sub SanduhrAusblenden() 'Sanduhr ausblenden
    Unload frmSanduhr
End Sub

Public Sub FortschrittsbalkenReset(strAktion As String, strNr As String, strSchritt As String) 'Fortschrittsbalken der Sanduhr zurücksetzen
    With frmSanduhr
        .Caption = strAktion
        .lblFortschrittBalken.Width = 0 'Breite des Balkens
        .lblFortschrittProzent.Caption = "" 'Anzeige des Prozentanteils
        .lblFortschrittNr.Caption = strNr 'Nummer des Einzelschritts
        .lblFortschrittSchritt.Caption = strSchritt 'Einzelschritt
    End With
    g_dblBalkenAktuell = 0 'Länge des Balkens zurücksetzen
    DoEvents 'neu zeichnen
End Sub

Public Sub FortschrittsbalkenAktualisieren(dblProzent As Double) 'Fortschrittsbalken der Sanduhr aktualisieren
    With frmSanduhr
        .lblFortschrittBalken.Width = CInt(dblProzent) 'Breite des Balkens
        .lblFortschrittProzent.Caption = CInt(dblProzent) 'Anzeige des Prozentanteils
    End With
    DoEvents 'neu zeichnen
End Sub


'Tool-Einstellungen zurücksetzen
'-------------------------------

Private Sub EinstellungenReset()

    If checkboxWarnung.Value = True Then
        If MsgBox(ThisWorkbook.Worksheets("Messages_GUI").Cells(55, intSprache).Value & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(56, intSprache).Value & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(57, intSprache).Value, vbExclamation + _
                    vbOKCancel, ThisWorkbook.Worksheets("Messages_GUI").Cells(54, intSprache).Value) = vbCancel Then
            Exit Sub
        End If
    End If

    checkboxTooltip.Value = True
    checkboxWarnung.Value = True
    checkboxHinweis.Value = True
    checkboxToolSettingsSpeichern.Value = True
    optBtnFarbeEntfernen.Value = True
    checkboxGrossKleinBuchstaben.Value = False
    checkboxLeerzeichen.Value = False
    toggleSuchModus.Value = False
    checkboxAusgabeTitelzeile.Value = True
    optBtnMarkieren4.Value = True
    toggleMarkierModus.Value = False
    toggleAusgabeModus.Value = False
    toggleLoeschModusZeilen.Value = False
    toggleLoeschModusKomprimieren.Value = False
    checkboxAktualisierung.Value = False
End Sub


'Hyperlinks
'----------

Private Sub SourceCodeURL()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/MarcoKrapf/xlDuplikatManager"
        
    Exit Sub
    
Fehlerbehandlung:
    MsgBox (Err.Description), _
            vbCritical, "xl DuplikatManager (Version " & cstrVersion & ")"
End Sub

Private Sub SpendenLinkURLaufrufen()

    'Wenn ein Fehler auftritt...
    On Error GoTo Fehlerbehandlung
    
    ActiveWorkbook.FollowHyperlink Address:="http://www.ghfkh.de/"
        
    Exit Sub
    
Fehlerbehandlung:
    MsgBox (Err.Description), _
            vbCritical, "xl DuplikatManager (Version " & cstrVersion & ")"
End Sub

Private Sub QRcodeAnzeigen() 'QR-Code einblenden
    If frmQRcode.Visible = False Then
        Load frmQRcode
        Call modUserformPlatzieren.UserFormPlatzieren(frmQRcode)
        frmQRcode.Show
    Else
        Unload frmQRcode
    End If
End Sub


'Klicks
'------

Private Sub toggleAusgabeModus_Click() 'Klick auf den Ausgabemodus-Button
 Call AusgabeModus
End Sub

Private Sub toggleLoeschModusKomprimieren_Click() 'Klick auf den Button, um Komprimierung ein/auszuschalten
    Call LoeschModusKomprimieren
End Sub

Private Sub toggleLoeschModusZeilen_Click() 'Klick auf den Button um das Löschen kompletter Zeilen ein/auszuschalten
    Call LoeschModusZeilen
End Sub

Private Sub toggleMarkierModus_Click() 'Klick auf den Button, um zwischen einfarbigem und buntem Markieren umzuschalten
    Call MarkierModus 'Beschriftung des Buttons anpassen
    If lngDuplikatCounter > 0 Then
        Call DuplikateHervorheben 'neu einfärben
    End If
End Sub


Private Sub imgToolLogo_Click() 'Klick auf das Tool-Logo
    Call MinimierenMaximieren
End Sub

Private Sub lblSpeicherortOptionen_Click() 'Klick auf den Speicherort der Optionen
    Call SpeicherortOptionenOeffnen
End Sub

Private Sub frameSpeicherort_Click() 'Klick auf den Rahmen
    Call SpeicherortOptionenOeffnen
End Sub

Private Sub lblTitel_Click() 'Klick auf den Titel
    Call MinimierenMaximieren
End Sub

Private Sub btnDuplikateLoeschen_Click() 'Klick auf "Duplikate löschen"
    If toggleLoeschModusZeilen.Value = False Then
        Call DuplikateLoeschen
    Else
        Call DuplikateLoeschenGanzeZeile
    End If
End Sub

Private Sub btnSourceCode_Click() 'Klick auf "Quellcode auf GitHub"
    Call SourceCodeURL
End Sub

Private Sub imgFlaggeDE_Click() 'Klick auf die deutsche Flagge
    intSprache = 2
    Call Sprache(intSprache)
End Sub

Private Sub imgFlaggeEN_Click() 'Klick auf die englische Flagge
    intSprache = 3
    Call Sprache(intSprache)
End Sub

Private Sub toggleMinMax_Click() 'Klick auf den Min/Max-Button links oben
    Call MinimierenMaximieren
End Sub

Private Sub toggleDuplikatfensterOnOff_Click() 'Klick auf den Button für das Duplikatfenster
    Call DuplikatfensterOnOff
End Sub

Private Sub checkboxTooltip_Click() 'Klick auf die Checkbox "Tooltips"
    Call Tooltips
End Sub

Private Sub btnFinden_Click() 'Klick auf "Duplikate finden"
    Call SucheStarten
End Sub

Private Sub SucheStarten() 'Suche wird gestartet
'    Call SelektionsVorbereitung
    
    If checkboxWarnung.Value = True And lngAnzahlGeloeschteDuplikate > 0 Then
        If MsgBox(ThisWorkbook.Worksheets("Messages_GUI").Cells(2, intSprache).Value _
            & lngAnzahlGeloeschteDuplikate, vbExclamation + vbOKCancel, _
            ThisWorkbook.Worksheets("Messages_GUI").Cells(1, intSprache).Value) = vbCancel Then
            Exit Sub
        End If
    End If
    
    'Anzahl der gelöschten Duplikate aktualisieren
    lngAnzahlGeloeschteDuplikate = 0
    lblAnzahlGeloescht.Caption = ""

    'Eventuelle Markierungen löschen
    rngSelection.Interior.Pattern = xlNone
    
    'Button in Duplikatfenster ausblenden
    lblDuplikatEinzelnLoeschen.Visible = False
    
    Call Selektion
    If blnStop = True Then 'Wenn Selektion ungültig
        blnStop = False
        Exit Sub
    End If
    Call DuplikateFinden
End Sub

Private Sub toggleSuchModus_Click() 'Klick auf den Button, um den Suchmodus umzuschalten
    Call SuchModus
End Sub

Private Sub btnHervorhebenDuplikate_Click() 'Klick auf "Nur Duplikate markieren"
    Call HervorhebenDuplikate
    Call DuplikateHervorheben
End Sub

Private Sub btnHervorhebenAlle_Click() 'Klick auf "Alle mehrfachen Werte markieren"
    Call HervorhebenAlle
    Call DuplikateHervorheben
End Sub

Private Sub btnHervorhebenUnikate_Click() 'Klick auf "Nur Unikate markieren"
    Call HervorhebenUnikate
    Call UnikateHervorheben
End Sub

Private Sub HervorhebenDuplikate()
    'Nur Duplikate markieren
    blnAlleFarbigMarkieren = False
    'Farben der Buttons anpassen
    btnHervorhebenDuplikate.BackColor = &HC0FFC0
    btnHervorhebenAlle.BackColor = &HFFC0C0
    btnHervorhebenUnikate.BackColor = &HFFC0C0
    'Markiermodus-Button aktivieren
    toggleMarkierModus.Enabled = True
End Sub

Private Sub HervorhebenAlle()
    'Alle mehrfachen markieren
    blnAlleFarbigMarkieren = True
    'Farben der Buttons anpassen
    btnHervorhebenDuplikate.BackColor = &HFFC0C0
    btnHervorhebenAlle.BackColor = &HC0FFC0
    btnHervorhebenUnikate.BackColor = &HFFC0C0
    'Markiermodus-Button aktivieren
    toggleMarkierModus.Enabled = True
End Sub

Private Sub HervorhebenUnikate()
    'Farben der Buttons anpassen
    btnHervorhebenDuplikate.BackColor = &HFFC0C0
    btnHervorhebenAlle.BackColor = &HFFC0C0
    btnHervorhebenUnikate.BackColor = &HC0FFC0
    'Markiermodus-Button deaktivieren
    toggleMarkierModus.Enabled = False
End Sub

Private Sub btnFarbigeMarkierungLoeschen_Click() 'Klick auf Radiergummi ("Farbige Markierungen löschen")
    Call FarbigeMarkierungLoeschen
End Sub

Private Sub FarbigeMarkierungLoeschen()
    'Markierungen entfernen
    rngSelection.Interior.ColorIndex = xlColorIndexNone
    'Farben der Buttons zurücksetzen
    btnHervorhebenDuplikate.BackColor = &HFFC0C0
    btnHervorhebenAlle.BackColor = &HFFC0C0
    btnHervorhebenUnikate.BackColor = &HFFC0C0
    'Button deaktivieren
    toggleMarkierModus.Enabled = False
    'Radiergummi deaktivieren ("Farbige Markierungen löschen")
    btnFarbigeMarkierungLoeschen.Enabled = False
End Sub

Private Sub btnDuplikateAusgeben_Click() 'Klick auf "Nur Duplikate ausgeben"
    strAusgabetyp = "Duplikat"
    arrAusgabetyp(1) = "Duplikat"
    arrAusgabetyp(2) = "Duplikat"
    If toggleAusgabeModus.Value = False Then
        Call AusgabeTabellenblatt(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'nur Zellen der Selektion
    Else
        Call AusgabeTabellenblattGanzeZeile(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'komplette Zeilen
    End If
End Sub

Private Sub btnUnikateAusgeben_Click() 'Klick auf "Nur Unikate ausgeben"
    strAusgabetyp = "Unikat"
    arrAusgabetyp(1) = "Unikat"
    arrAusgabetyp(2) = "Unikat"
    If toggleAusgabeModus.Value = False Then
        Call AusgabeTabellenblatt(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'nur Zellen der Selektion
    Else
        Call AusgabeTabellenblattGanzeZeile(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'komplette Zeilen
    End If
End Sub

Private Sub btnOriginaleAusgeben_Click() 'Klick auf "Alle Werte 1x ausgeben"
    strAusgabetyp = "Original"
    arrAusgabetyp(1) = "Original"
    arrAusgabetyp(2) = "Unikat"
    If toggleAusgabeModus.Value = False Then
        Call AusgabeTabellenblatt(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'nur Zellen der Selektion
    Else
        Call AusgabeTabellenblattGanzeZeile(strAusgabetyp, arrAusgabetyp, strTabellenblatt) 'komplette Zeilen
    End If
End Sub

Private Sub btnLoeschenRueckgaengig_Click() 'Klick auf Pfeil zurück ("Alle Duplikate wiederherstellen")
    Call DuplikateLoeschenUnDo
End Sub

Private Sub boxVorschau_Click() 'Klick auf einen Eintrag im Fenster "Duplikate" auf dem UserForm
    Call DuplikatfensterAnklickenDuplikat
    lblDuplikatEinzelnLoeschen.Visible = True 'Schaltfläche "löschen/wiederherstellen" einblenden
End Sub

Private Sub lblDuplikatEinzelnLoeschen_Click() 'Klick auf "löschen" bzw. "wiederherstellen" im Duplikatfenster
    If byteDuplikatfensterStatus = 1 Then
        Call DuplikatEinzelnLoeschen
        
    ElseIf byteDuplikatfensterStatus = 2 Then
        Call DuplikatEinzelnWiederherstellen
        
    End If
End Sub

Private Sub lblVorschau_Click() 'Klick auf "Duplikate" im Duplikatfenster
    rngSelection.Interior.Pattern = xlSolid 'Schraffierung auf dem Tabellenblatt entfernen
    boxVorschau.ListIndex = -1 'Blauen Balken in ListBox entfernen
    lblDuplikatEinzelnLoeschen.Visible = False  'Schaltfläche "löschen/wiederherstellen" ausblenden
End Sub

Private Sub btnOptionenReset_Click() 'Klick auf "Alle Werte auf Standard zurücksetzen" der Optionen
    Call EinstellungenReset
End Sub

Private Sub btnAnleitung_Click() 'Klick auf "Anleitung"
    Call Sprache(intSprache)
    Call AnleitungAnzeigen
End Sub

Private Sub btnFeatures_Click() 'Klick auf "Features"
    Call Sprache(intSprache)
    Call FeaturesAnzeigen
End Sub

Private Sub btnNutzungsbedingungen_Click() 'Klick auf "Nutzungsbedingungen"
    Call Sprache(intSprache)
    Call NutzungsbedingungenAnzeigen
End Sub

Private Sub AnleitungAnzeigen() 'Öffnen bzw. schließen des Popups
    If frmAnleitung.Visible = False Then
        Load frmAnleitung
        Call modUserformPlatzieren.UserFormPlatzieren(frmAnleitung)
        frmAnleitung.Show
    Else
        Unload frmAnleitung
    End If
End Sub

Private Sub FeaturesAnzeigen() 'Öffnen bzw. schließen des Popups
    If frmVersionshinweise.Visible = False Then
        Load frmVersionshinweise
        Call modUserformPlatzieren.UserFormPlatzieren(frmVersionshinweise)
        frmVersionshinweise.Show
    Else
        Unload frmVersionshinweise
    End If
End Sub

Private Sub NutzungsbedingungenAnzeigen() 'Öffnen bzw. schließen des Popups
    If frmNutzungsbedingungen.Visible = False Then
        Load frmNutzungsbedingungen
        Call modUserformPlatzieren.UserFormPlatzieren(frmNutzungsbedingungen)
        frmNutzungsbedingungen.Show
    Else
        Unload frmNutzungsbedingungen
    End If
End Sub

Private Sub btnFeedback_Click() 'Klick auf "Feedback"
    Call eMail
End Sub

Private Sub imgPrinz_Click() 'Klick auf das Logo (Kleiner Held)
    Call SpendenLinkURLaufrufen
End Sub

Private Sub lblSpendenLink_Click() 'Klick auf den Spendenlink
    Call SpendenLinkURLaufrufen
End Sub

Private Sub lblSpendeQRcode_Click() 'Klick auf Link zu QR-Code
    Call Sprache(intSprache)
    Call QRcodeAnzeigen
End Sub

Private Sub btnBeenden_Click() 'Klick auf den roten Button X ("Schließen")
    Unload frmDuplikatManager
    '---> automatischer Absprung in Sub UserForm_QueryClose
End Sub

Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer) 'Klick auf das rote X des UserForms ("Schließen")

    If checkboxWarnung.Value = True And lngAnzahlGeloeschteDuplikate > 0 Then
        If MsgBox(ThisWorkbook.Worksheets("Messages_GUI").Cells(60, intSprache).Value & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(61, intSprache).Value & vbNewLine & vbNewLine & _
                    ThisWorkbook.Worksheets("Messages_GUI").Cells(62, intSprache).Value & lngAnzahlGeloeschteDuplikate, vbExclamation + _
                    vbOKCancel, ThisWorkbook.Worksheets("Messages_GUI").Cells(59, intSprache).Value) = vbCancel Then
            Cancel = True 'Schließen verhindern
            Exit Sub
        End If
    End If

    On Error Resume Next 'falls ein Fehler auftritt: Anweisung überspringen

    'Eventuelle farbige Markierungen im selektierten Bereich löschen
    If optBtnFarbeEntfernen.Value = True Then
        rngSelection.Interior.Pattern = xlNone
    End If
    
    '---> automatischerAbsprung in Sub UserForm_Terminate
End Sub

'Klick auf eines der Objekte im Bereich "Aktuelle Selektion"
'selektiert auf dem Tabellenblatt wieder den aktuellen Bereich,
'falls inzwischen ein anderer Bereich auf dem Tabellenblatt
'ausgewählt, aber noch nicht übernommen wurde
    Private Sub frameSelektion_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuelleMappe_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuelleMappeWert_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuellesBlatt_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuellesBlattWert_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuellerBereich_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuellerBereichWert_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub lblAktuelleAnzahlBereicheWert_Click()
        Call AktuelleSelektionHolen
    End Sub
    
    Private Sub AktuelleSelektionHolen()
        On Error Resume Next 'wenn z.B. Tabellenblatt mit aktueller Selektion gelöscht wurde
        Workbooks(CStr(lblAktuelleMappeWert.Caption)).Sheets(CStr(lblAktuellesBlattWert.Caption)).Activate 'Tabellenblatt aktivieren
        rngSelection.Select
    End Sub
'----------------------------------------------------------

