Attribute VB_Name = "Pruefplatz_Auswertung"
'===============================================================================
'  Pruefplatz_Auswertung  -  Grafische Auswertung und Analyse der Prueflings-
'  daten eines Pruefplatzes (CSV-Export, z. B. "Pruefplatz_2_1.csv")
'
'  Einstieg:
'     Start_Auswertung      CSV auswaehlen, importieren und komplett auswerten
'     Neu_Berechnen         Auswertung mit geaenderten Einstellungen wiederholen
'                           (ohne erneuten Import)
'     Startseite_Anlegen    legt das Blatt "Start" mit Schaltflaechen fuer die
'                           Makros an (nur noetig, wenn das Modul in eine leere
'                           Mappe importiert wurde)
'
'  Erzeugte Blaetter:
'     Uebersicht            Kennzahlen, Auffaelligkeiten, Fehler-Pareto
'     Einstellungen         Grenzwerte, Einheiten, Faktoren (frei aenderbar)
'     Daten                 aufbereitete Messwerte + eigene IO/NIO-Bewertung
'     Statistik             Kennwerte je Messkanal, Klassen, Korrelationen
'     Diagramme             alle Diagramme
'     NIO_Liste             alle auffaelligen Pruefungen mit Fehlergrund
'     Wochen                Wochenauswertung (Anzahl, NIO-Quote, Mediane)
'     Datenqualitaet        Importprotokoll (Formatabweichungen usw.)
'     Rohdaten / Hilfsdaten interne Tabellen (Hilfsdaten ist ausgeblendet)
'
'  Erwartetes CSV-Format (Semikolon, Dezimalkomma, UTF-8):
'     LfdN;5-7;5-8;6-7;6-8;6-12;5-11;5-6 R;5-6 C;Gesamtergebnis;Pruefzeit
'     Pruefzeit = Unix-Zeitstempel in Millisekunden
'
'  Voraussetzung: Excel 2010 oder neuer unter Windows.
'===============================================================================
Option Explicit

' ---- Spalten im CSV / Rohdaten -------------------------------------------------
Private Const N_KAN As Long = 8          ' Anzahl Messkanaele (Spalten 2..9)
Private Const COL_LFDN As Long = 1
Private Const COL_ERG As Long = 10
Private Const COL_ZEIT As Long = 11

' ---- Kanaltypen -----------------------------------------------------------------
Private Const TYP_ISO As String = "ISO"     ' Isolationswiderstand (Untergrenze)
Private Const TYP_KONT As String = "KONT"   ' Durchgang / Kontakt (Obergrenze)
Private Const TYP_INFO As String = "INFO"   ' nur Information

' ---- Blattnamen -----------------------------------------------------------------
Private Const WS_UEB As String = "Uebersicht"
Private Const WS_SET As String = "Einstellungen"
Private Const WS_DAT As String = "Daten"
Private Const WS_STA As String = "Statistik"
Private Const WS_DIA As String = "Diagramme"
Private Const WS_NIO As String = "NIO_Liste"
Private Const WS_WOC As String = "Wochen"
Private Const WS_DQ As String = "Datenqualitaet"
Private Const WS_RAW As String = "Rohdaten"
Private Const WS_HLP As String = "Hilfsdaten"
Private Const WS_START As String = "Start"

' ---- Farben (Kanaele 1..8) ------------------------------------------------------
Private KanFarbe(1 To N_KAN) As Long

' ---- Einstellungen (werden aus Blatt "Einstellungen" gelesen) -------------------
Private kName(1 To N_KAN) As String
Private kKurz(1 To N_KAN) As String
Private kTyp(1 To N_KAN) As String
Private kFaktor(1 To N_KAN) As Double
Private kEinheit(1 To N_KAN) As String
Private kUG(1 To N_KAN) As Variant
Private kOG(1 To N_KAN) As Variant
Private gOverrange As Double            ' Rohwert ab dem "Messbereichsende" gilt
Private gTzOffset As Double             ' Stunden zu UTC
Private gFenster As Long                ' Fenster gleitender Median
Private gHistVon As Double, gHistBis As Double, gHistSchritt As Double
Private gOverrangeIstIO As Boolean      ' Overrange bei ISO = IO
Private kFormat(1 To N_KAN) As String   ' Zahlenformat je Kanal (aus Groessenordnung der Werte)

' ---- Rohdaten im Speicher -------------------------------------------------------
Private nZ As Long
Private rLfdN() As Variant
Private rWert() As Variant              ' (1..nZ, 1..N_KAN) Rohwert oder Empty
Private rErg() As Variant
Private rZeit() As Variant              ' Excel-Datum oder Empty

' ---- Importprotokoll ------------------------------------------------------------
Private dqLog As Collection

'===============================================================================
'  OEFFENTLICHE MAKROS
'===============================================================================
Public Sub Start_Auswertung()
    Dim datei As Variant
    datei = Application.GetOpenFilename("CSV-Dateien (*.csv),*.csv,Alle Dateien (*.*),*.*", , _
                                        "Pruefplatz-CSV auswaehlen")
    If VarType(datei) = vbBoolean Then Exit Sub

    On Error GoTo Fehler
    Performance True
    Application.StatusBar = "Importiere " & datei & " ..."

    InitFarben
    EinstellungenAnlegen False
    If Not CsvImportieren(CStr(datei)) Then GoTo Ende
    EinheitenAutomatisch

    AuswertungAusfuehren CStr(datei)
    Worksheets(WS_UEB).Activate
    Worksheets(WS_UEB).Range("A1").Select
    Application.StatusBar = False
    Performance False
    MsgBox "Auswertung abgeschlossen." & vbCrLf & nZ & " Pruefungen ausgewertet.", vbInformation
    Exit Sub
Fehler:
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "Pruefplatz-Auswertung"
Ende:
    Application.StatusBar = False
    Performance False
End Sub

Public Sub Neu_Berechnen()
    If Not BlattExistiert(WS_RAW) Then
        MsgBox "Es wurden noch keine Daten importiert. Bitte zuerst 'Start_Auswertung' ausfuehren.", vbExclamation
        Exit Sub
    End If
    On Error GoTo Fehler
    Performance True
    InitFarben
    EinstellungenAnlegen False
    RohdatenLaden
    AuswertungAusfuehren CStr(Worksheets(WS_RAW).Range("N1").Value)
    Worksheets(WS_UEB).Activate
    Application.StatusBar = False
    Performance False
    MsgBox "Auswertung neu berechnet.", vbInformation
    Exit Sub
Fehler:
    Application.StatusBar = False
    Performance False
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "Pruefplatz-Auswertung"
End Sub

' Setzt das Blatt "Einstellungen" auf die Standardwerte zurueck
' Legt das Blatt "Start" mit Schaltflaechen fuer alle Makros an
Public Sub Startseite_Anlegen()
    Dim ws As Worksheet
    If BlattExistiert(WS_START) Then
        Set ws = ThisWorkbook.Worksheets(WS_START)
        Dim i As Long
        For i = ws.Buttons.Count To 1 Step -1
            ws.Buttons(i).Delete
        Next i
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = WS_START
    End If
    ws.Activate
    OhneGitternetz ws

    ws.Range("B2").Value = "Pruefplatz-Auswertung"
    ws.Range("B2").Font.Size = 20: ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(31, 78, 121)
    ws.Range("B3").Value = "Grafische Auswertung und Analyse der Pruefplatz-CSV-Daten"
    ws.Range("B3").Font.Color = RGB(118, 118, 118)

    StartKnopf ws, ws.Range("B5"), "1. CSV importieren und auswerten", "Start_Auswertung", _
               "CSV-Datei auswaehlen; alle Blaetter und Diagramme werden neu erzeugt."
    StartKnopf ws, ws.Range("B9"), "2. Neu berechnen", "Neu_Berechnen", _
               "Nach Aenderungen im Blatt 'Einstellungen' (Grenzwerte, Einheiten) - ohne erneuten Import."
    StartKnopf ws, ws.Range("B13"), "Einstellungen zuruecksetzen", "Einstellungen_Zuruecksetzen", _
               "Grenzwerte und Einheiten auf die Standardwerte zuruecksetzen."
    ws.Range("B17").Value = "Hinweis: Die Standard-Grenzwerte sind Annahmen und muessen an die Pruefspezifikation " & _
                            "angepasst werden (Blatt 'Einstellungen')."
    ws.Range("B17").Font.Italic = True
    ws.Range("B17").Font.Color = RGB(192, 0, 0)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:D").ColumnWidth = 14
    ws.Range("A1").Select
End Sub

Private Sub StartKnopf(ws As Worksheet, ziel As Range, ByVal text As String, ByVal makro As String, ByVal info As String)
    Dim b As Object
    Set b = ws.Buttons.Add(ziel.Left, ziel.Top, 260, 40)
    b.OnAction = makro
    b.Caption = text
    b.Font.Size = 12
    b.Font.Bold = True
    ws.Cells(ziel.Row + 1, ziel.Column + 5).Value = info
    ws.Cells(ziel.Row + 1, ziel.Column + 5).Font.Color = RGB(89, 89, 89)
End Sub

Public Sub Einstellungen_Zuruecksetzen()
    If MsgBox("Alle Einstellungen auf Standard zuruecksetzen?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    EinstellungenAnlegen True
    Worksheets(WS_SET).Activate
End Sub

'===============================================================================
'  ABLAUF
'===============================================================================
Private Sub AuswertungAusfuehren(ByVal quelle As String)
    EinstellungenLesen
    FormateBestimmen

    Dim bew() As Variant      ' (1..nZ): "IO" / "NIO"
    Dim grund() As String     ' Fehlergruende je Zeile
    Bewerten bew, grund

    Application.StatusBar = "Schreibe Datenblatt ..."
    DatenblattSchreiben bew, grund
    Application.StatusBar = "Berechne Statistik ..."
    StatistikSchreiben bew
    Application.StatusBar = "Wochenauswertung ..."
    WochenSchreiben bew
    Application.StatusBar = "NIO-Liste ..."
    NioListeSchreiben bew, grund
    Application.StatusBar = "Erzeuge Diagramme ..."
    HilfsdatenSchreiben
    Dim pareto As Range
    Set pareto = ParetoTabelle(grund)
    DiagrammeErzeugen pareto
    Application.StatusBar = "Uebersicht ..."
    UebersichtSchreiben quelle, bew, grund
    DatenqualitaetSchreiben

    BlattReihenfolge
End Sub

'===============================================================================
'  EINSTELLUNGEN
'===============================================================================
Private Sub EinstellungenAnlegen(ByVal zuruecksetzen As Boolean)
    Dim ws As Worksheet
    If BlattExistiert(WS_SET) And Not zuruecksetzen Then Exit Sub
    Set ws = BlattHolen(WS_SET, True)

    With ws
        .Range("A1").Value = "Einstellungen der Pruefplatz-Auswertung"
        .Range("A1").Font.Size = 14: .Range("A1").Font.Bold = True
        .Range("A2").Value = "Gelb hinterlegte Zellen koennen geaendert werden. Danach Makro 'Neu_Berechnen' ausfuehren. " & _
                             "Grenzwerte gelten in der Anzeigeeinheit. Leere Grenze = keine Pruefung."
        .Range("A2").Font.Italic = True

        .Range("A4:I4").Value = Array("Nr", "Kanal (CSV)", "Kurzname", "Typ", "Faktor Rohwert -> Anzeige", _
                                      "Anzeigeeinheit", "Untergrenze (UG)", "Obergrenze (OG)", "Bemerkung")
        Dim d As Variant
        ' Nr, Name, Kurz, Typ, Faktor, Einheit, UG, OG, Bemerkung
        d = Array( _
          Array(1, "5-7 (SL-GDS)", "5-7 SL-GDS", TYP_ISO, 0.000001, "MOhm", 5000, "", "Isolation Signalleitung/Schirm - Gehaeuse; Rohwert in Ohm"), _
          Array(2, "5-8 (SL-GDS)", "5-8 SL-GDS", TYP_ISO, 0.000001, "MOhm", 5000, "", "Isolation Signalleitung/Schirm - Gehaeuse; Rohwert in Ohm"), _
          Array(3, "6-7 (Masse-GDS)", "6-7 Masse-GDS", TYP_ISO, 0.000001, "MOhm", 5000, "", "Isolation Masse - Gehaeuse; Rohwert in Ohm"), _
          Array(4, "6-8 (Masse-GDS)", "6-8 Masse-GDS", TYP_ISO, 0.000001, "MOhm", 5000, "", "Isolation Masse - Gehaeuse; Rohwert in Ohm"), _
          Array(5, "6-12 (Masse)", "6-12 Masse", TYP_KONT, 1, "Ohm", "", 10, "Durchgang Masse; erwartet ~0 Ohm"), _
          Array(6, "5-11 (Signal)", "5-11 Signal", TYP_KONT, 1, "Ohm", "", 10, "Durchgang Signal; erwartet ~0 Ohm"), _
          Array(7, "5-6 Widerstand", "5-6 R", TYP_ISO, 0.000001, "MOhm", 100, "", "Widerstand 5-6; Rohwert in Ohm"), _
          Array(8, "5-6 Kapazitaet", "5-6 C", TYP_INFO, 1, "Rohwert (nF lt. CSV)", "", "", "Einheit im CSV unklar - Faktor ggf. anpassen"))
        Dim i As Long, j As Long
        For i = 0 To N_KAN - 1
            For j = 0 To 8
                .Cells(5 + i, 1 + j).Value = d(i)(j)
            Next j
        Next i

        .Range("A15").Value = "Allgemein"
        .Range("A15").Font.Bold = True
        .Range("A16:C16").Value = Array("Overrange-Schwelle (Rohwert)", 9900000000000#, "Rohwerte >= Schwelle gelten als 'Messbereichsende' (> Messbereich)")
        .Range("A17:C17").Value = Array("Overrange bei ISO-Kanaelen = IO", "JA", "JA: sehr hoher Isolationswert ist gut. NEIN: Overrange wird als NIO gewertet")
        .Range("A18:C18").Value = Array("Zeitzone (Stunden zu UTC)", 2, "Pruefzeit ist Unix-Zeit (UTC). 2 = MESZ, 1 = MEZ")
        .Range("A19:C19").Value = Array("Fenster gleitender Median (Pruefungen)", 50, "Glaettung fuer Trenddiagramm")
        .Range("A20:C20").Value = Array("Histogramm von (log10 Ohm)", 6, "Klassen fuer Isolations-Histogramm (Rohwert, logarithmisch)")
        .Range("A21:C21").Value = Array("Histogramm bis (log10 Ohm)", 13, "")
        .Range("A22:C22").Value = Array("Histogramm Schrittweite (Dekaden)", 0.25, "")

        .Range("A24").Value = "Hinweis: Die Grenzwerte (5000 MOhm Isolation, 10 Ohm Durchgang, 100 MOhm fuer 5-6) sind Annahmen " & _
                              "und muessen an die Pruefspezifikation angepasst werden."
        .Range("A24").Font.Color = RGB(192, 0, 0)

        .Range("E5:H12").Interior.Color = RGB(255, 242, 204)
        .Range("C5:D12").Interior.Color = RGB(255, 242, 204)
        .Range("B16:B22").Interior.Color = RGB(255, 242, 204)
        KopfFormat .Range("A4:I4")
        .Range("A4:I12").Borders.LineStyle = xlContinuous
        .Range("E5:E12").NumberFormat = "0.######"
        .Range("B16").NumberFormat = "0.00E+00"
        .Columns("A:I").AutoFit
        .Columns("A").ColumnWidth = 38
        .Columns("I").ColumnWidth = 60
    End With

    ' Gueltigkeitsliste fuer Typ
    With ws.Range("D5:D12").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="ISO,KONT,INFO"
    End With
    With ws.Range("B17").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="JA,NEIN"
    End With
End Sub

Private Sub EinstellungenLesen()
    Dim ws As Worksheet, i As Long
    Set ws = Worksheets(WS_SET)
    For i = 1 To N_KAN
        kName(i) = CStr(ws.Cells(4 + i, 2).Value)
        kKurz(i) = CStr(ws.Cells(4 + i, 3).Value)
        If kKurz(i) = "" Then kKurz(i) = kName(i)
        kTyp(i) = UCase$(Trim$(CStr(ws.Cells(4 + i, 4).Value)))
        If kTyp(i) <> TYP_ISO And kTyp(i) <> TYP_KONT Then kTyp(i) = TYP_INFO
        If IsNumeric(ws.Cells(4 + i, 5).Value) And Not IsEmpty(ws.Cells(4 + i, 5).Value) Then
            kFaktor(i) = CDbl(ws.Cells(4 + i, 5).Value)
        Else
            kFaktor(i) = 1
        End If
        If kFaktor(i) = 0 Then kFaktor(i) = 1
        kEinheit(i) = CStr(ws.Cells(4 + i, 6).Value)
        kUG(i) = GrenzeLesen(ws.Cells(4 + i, 7))
        kOG(i) = GrenzeLesen(ws.Cells(4 + i, 8))
    Next i
    gOverrange = ZahlOder(ws.Range("B16").Value, 9900000000000#)
    ' Overrange-Schwelle in Anzeigeeinheit (wird von der bedingten Formatierung verwendet)
    ws.Range("J4").Value = "Overrange (Anzeige)"
    For i = 1 To N_KAN
        ws.Cells(4 + i, 10).Value = gOverrange * kFaktor(i)
    Next i
    ws.Range("J5:J12").NumberFormat = "0.00E+00"
    ws.Range("J4:J12").Font.Color = RGB(128, 128, 128)
    gOverrangeIstIO = (UCase$(Trim$(CStr(ws.Range("B17").Value))) <> "NEIN")
    gTzOffset = ZahlOder(ws.Range("B18").Value, 0)
    gFenster = CLng(ZahlOder(ws.Range("B19").Value, 50))
    If gFenster < 3 Then gFenster = 3
    gHistVon = ZahlOder(ws.Range("B20").Value, 6)
    gHistBis = ZahlOder(ws.Range("B21").Value, 13)
    gHistSchritt = ZahlOder(ws.Range("B22").Value, 0.25)
    If gHistSchritt <= 0 Or gHistBis <= gHistVon Then gHistVon = 6: gHistBis = 13: gHistSchritt = 0.25
End Sub

Private Function GrenzeLesen(ByVal c As Range) As Variant
    If IsEmpty(c.Value) Then
        GrenzeLesen = Empty
    ElseIf IsNumeric(c.Value) And CStr(c.Value) <> "" Then
        GrenzeLesen = CDbl(c.Value)
    Else
        GrenzeLesen = Empty
    End If
End Function

Private Function ZahlOder(ByVal v As Variant, ByVal ersatz As Double) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then ZahlOder = CDbl(v) Else ZahlOder = ersatz
End Function

'===============================================================================
'  CSV-IMPORT
'===============================================================================
Private Function CsvImportieren(ByVal datei As String) As Boolean
    Dim txt As String, zeilen() As String, i As Long, f() As String
    Set dqLog = New Collection

    txt = DateiLesenUtf8(datei)
    If Len(txt) = 0 Then
        MsgBox "Die Datei ist leer oder konnte nicht gelesen werden.", vbExclamation
        Exit Function
    End If
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    If Left$(txt, 1) = ChrW(&HFEFF) Then txt = Mid$(txt, 2)
    zeilen = Split(txt, vbLf)

    ' Kopfzeile
    Dim kopf() As String
    kopf = Split(zeilen(0), ";")
    If UBound(kopf) < COL_ZEIT - 1 Then
        MsgBox "Unerwartetes Format: Die Kopfzeile hat nur " & (UBound(kopf) + 1) & _
               " Spalten (erwartet " & COL_ZEIT & ").", vbExclamation
        Exit Function
    End If

    ' Datenzeilen zaehlen
    Dim n As Long
    For i = 1 To UBound(zeilen)
        If Len(Trim$(zeilen(i))) > 0 Then n = n + 1
    Next i
    If n = 0 Then
        MsgBox "Keine Datenzeilen gefunden.", vbExclamation
        Exit Function
    End If

    nZ = n
    ReDim rLfdN(1 To nZ): ReDim rWert(1 To nZ, 1 To N_KAN)
    ReDim rErg(1 To nZ): ReDim rZeit(1 To nZ)

    Dim z As Long, k As Long, ok As Boolean, v As Double, roh As String
    z = 0
    For i = 1 To UBound(zeilen)
        If Len(Trim$(zeilen(i))) > 0 Then
            z = z + 1
            f = Split(zeilen(i), ";")
            If UBound(f) < COL_ZEIT - 1 Then
                DqAdd i + 1, "-", zeilen(i), "Zeile hat zu wenige Spalten (" & (UBound(f) + 1) & ")"
                ReDim Preserve f(0 To COL_ZEIT - 1)
            End If

            ' LfdN
            v = ZahlParsen(f(COL_LFDN - 1), ok, roh)
            If ok Then rLfdN(z) = v Else rLfdN(z) = z: DqAdd i + 1, "LfdN", f(COL_LFDN - 1), "LfdN ungueltig - durch Zeilennummer ersetzt"

            ' Messkanaele
            For k = 1 To N_KAN
                v = ZahlParsen(f(k), ok, roh)
                If ok Then
                    rWert(z, k) = v
                    If roh <> "" Then DqAdd i + 1, kopf(k), f(k), roh
                Else
                    rWert(z, k) = Empty
                    DqAdd i + 1, kopf(k), f(k), "Wert nicht lesbar - ignoriert"
                End If
            Next k

            ' Gesamtergebnis
            Select Case LCase$(Trim$(f(COL_ERG - 1)))
                Case "true", "wahr", "1", "io", "ok", "pass": rErg(z) = True
                Case "false", "falsch", "0", "nio", "nok", "fail": rErg(z) = False
                Case Else: rErg(z) = Empty: DqAdd i + 1, "Gesamtergebnis", f(COL_ERG - 1), "unbekannter Wert"
            End Select

            ' Pruefzeit (Unix ms -> Excel-Datum)
            v = ZahlParsen(f(COL_ZEIT - 1), ok, roh)
            If ok And v > 0 Then
                rZeit(z) = UnixMsZuDatum(v)   ' Zeitzone wird bei Verwendung addiert
            Else
                rZeit(z) = Empty
                DqAdd i + 1, "Pruefzeit", f(COL_ZEIT - 1), "Zeit nicht lesbar"
            End If
        End If
    Next i

    RohdatenSpeichern datei, kopf
    CsvImportieren = True
End Function

' Liest eine Textdatei als UTF-8 (fuer Sonderzeichen wie Omega); Fallback ANSI
Private Function DateiLesenUtf8(ByVal datei As String) As String
    Dim stm As Object
    On Error GoTo Ansi
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2              ' Text
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile datei
    DateiLesenUtf8 = stm.ReadText(-1)
    stm.Close
    Exit Function
Ansi:
    On Error GoTo 0
    Dim ff As Integer, z As String, s As String
    ff = FreeFile
    Open datei For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, z
        s = s & z & vbLf
    Loop
    Close #ff
    DateiLesenUtf8 = s
End Function

' Robuste Zahlenerkennung:
'   "2,64E+06"        -> Dezimalkomma
'   "1.404.108.875"   -> Tausenderpunkte
'   "5645613.5"       -> Dezimalpunkt (wird protokolliert)
' hinweis <> "" wenn das Format vom Standard (Dezimalkomma) abweicht
Private Function ZahlParsen(ByVal s As String, ByRef ok As Boolean, ByRef hinweis As String) As Double
    Dim t As String, nPunkte As Long, i As Long, ch As String
    ok = False: hinweis = ""
    t = Trim$(Replace(s, ChrW(160), ""))
    t = Replace(t, " ", "")
    If t = "" Then Exit Function

    nPunkte = Len(t) - Len(Replace(t, ".", ""))
    If InStr(t, ",") > 0 Then
        If nPunkte > 0 Then hinweis = "Tausenderpunkte entfernt"
        t = Replace(t, ".", "")
        t = Replace(t, ",", ".")
    ElseIf nPunkte > 1 Then
        t = Replace(t, ".", "")
        hinweis = "Tausenderpunkte entfernt"
    ElseIf nPunkte = 1 Then
        hinweis = "Dezimalpunkt statt Dezimalkomma"
    End If

    ' nur gueltige Zeichen zulassen
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If InStr("0123456789.eE+-", ch) = 0 Then Exit Function
    Next i
    If InStr("0123456789.-+", Left$(t, 1)) = 0 Then Exit Function

    Dim erg As Double
    On Error GoTo Ungueltig
    erg = Val(t)                   ' Val ist gebietsschemaunabhaengig (Punkt)
    ok = True
    If hinweis <> "" Then hinweis = hinweis & " -> " & CStr(erg)
    ZahlParsen = erg
    Exit Function
Ungueltig:
    ok = False
End Function

Private Function UnixMsZuDatum(ByVal ms As Double) As Double
    UnixMsZuDatum = CDbl(DateSerial(1970, 1, 1)) + ms / 86400000#
End Function

Private Sub DqAdd(ByVal zeile As Long, ByVal spalte As String, ByVal orig As String, ByVal info As String)
    If dqLog Is Nothing Then Set dqLog = New Collection
    dqLog.Add Array(zeile, spalte, orig, info)
End Sub

' Rohdaten in Blatt sichern, damit "Neu_Berechnen" ohne CSV funktioniert
Private Sub RohdatenSpeichern(ByVal datei As String, kopf() As String)
    Dim ws As Worksheet, a() As Variant, i As Long, k As Long
    Set ws = BlattHolen(WS_RAW, True)
    ReDim a(0 To nZ, 1 To COL_ZEIT)
    For k = 1 To COL_ZEIT
        a(0, k) = Trim$(kopf(k - 1))
    Next k
    a(0, COL_ZEIT) = "Pruefzeit (UTC)"
    For i = 1 To nZ
        a(i, 1) = rLfdN(i)
        For k = 1 To N_KAN
            a(i, 1 + k) = rWert(i, k)
        Next k
        a(i, COL_ERG) = rErg(i)
        a(i, COL_ZEIT) = rZeit(i)
    Next i
    ws.Range("A1").Resize(nZ + 1, COL_ZEIT).Value = a
    ws.Range("K2").Resize(nZ, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Range("B2").Resize(nZ, N_KAN).NumberFormat = "0.000E+00"
    KopfFormat ws.Range("A1").Resize(1, COL_ZEIT)
    ws.Range("M1").Value = "Quelldatei:"
    ws.Range("N1").Value = datei
    ws.Range("M2").Value = "Importiert:"
    ws.Range("N2").Value = Now
    ws.Range("N2").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Range("M3").Value = "Datensaetze:"
    ws.Range("N3").Value = nZ
    ws.Columns("A:N").AutoFit

    ' Importprotokoll sichern
    Dim wq As Worksheet
    Set wq = BlattHolen(WS_DQ, True)
    wq.Range("A3:D3").Value = Array("CSV-Zeile", "Spalte", "Originalwert", "Hinweis")
    wq.Range("F1").Value = dqLog.Count
    wq.Range("F1").Font.Color = RGB(255, 255, 255)
    If dqLog.Count > 0 Then
        Dim b() As Variant
        ReDim b(1 To dqLog.Count, 1 To 4)
        Dim e As Variant
        i = 0
        For Each e In dqLog
            i = i + 1
            b(i, 1) = e(0)
            b(i, 2) = CStr(e(1))
            b(i, 3) = "'" & CStr(e(2))     ' als Text, damit Excel nicht umwandelt
            b(i, 4) = CStr(e(3))
        Next e
        wq.Range("A4").Resize(dqLog.Count, 4).Value = b
    End If
End Sub

Private Sub RohdatenLaden()
    Dim ws As Worksheet, a As Variant, i As Long, k As Long
    Set ws = Worksheets(WS_RAW)
    nZ = CLng(Val(CStr(ws.Range("N3").Value)))
    If nZ < 1 Then Err.Raise vbObjectError + 1, , "Rohdaten sind leer."
    a = ws.Range("A2").Resize(nZ, COL_ZEIT).Value
    ReDim rLfdN(1 To nZ): ReDim rWert(1 To nZ, 1 To N_KAN)
    ReDim rErg(1 To nZ): ReDim rZeit(1 To nZ)
    For i = 1 To nZ
        rLfdN(i) = a(i, 1)
        For k = 1 To N_KAN
            If IsNumeric(a(i, 1 + k)) And Not IsEmpty(a(i, 1 + k)) Then rWert(i, k) = CDbl(a(i, 1 + k)) Else rWert(i, k) = Empty
        Next k
        rErg(i) = a(i, COL_ERG)
        If IsNumeric(a(i, COL_ZEIT)) And Not IsEmpty(a(i, COL_ZEIT)) Then rZeit(i) = CDbl(a(i, COL_ZEIT)) Else rZeit(i) = Empty
    Next i
    Set dqLog = Nothing
End Sub

'===============================================================================
'  BEWERTUNG
'===============================================================================
Private Sub Bewerten(ByRef bew() As Variant, ByRef grund() As String)
    Dim i As Long, k As Long, w As Double, g As String
    ReDim bew(1 To nZ): ReDim grund(1 To nZ)
    For i = 1 To nZ
        g = ""
        For k = 1 To N_KAN
            If IsEmpty(rWert(i, k)) Then
                If kTyp(k) <> TYP_INFO Then g = g & "; " & kKurz(k) & " fehlt"
            Else
                w = rWert(i, k) * kFaktor(k)
                If rWert(i, k) >= gOverrange Then
                    Select Case kTyp(k)
                        Case TYP_ISO
                            If Not gOverrangeIstIO Then g = g & "; " & kKurz(k) & " Overrange"
                        Case TYP_KONT
                            g = g & "; " & kKurz(k) & " offen (Overrange)"
                    End Select
                Else
                    If Not IsEmpty(kUG(k)) Then
                        If w < kUG(k) Then g = g & "; " & kKurz(k) & " < UG"
                    End If
                    If Not IsEmpty(kOG(k)) Then
                        If w > kOG(k) Then g = g & "; " & kKurz(k) & " > OG"
                    End If
                End If
            End If
        Next k
        If g = "" Then
            bew(i) = "IO"
        Else
            bew(i) = "NIO"
            grund(i) = Mid$(g, 3)
        End If
    Next i
End Sub

'===============================================================================
'  DATENBLATT
'===============================================================================
Private Sub DatenblattSchreiben(bew() As Variant, grund() As String)
    Dim ws As Worksheet, a() As Variant, i As Long, k As Long, nS As Long
    Set ws = BlattHolen(WS_DAT, True)
    nS = 3 + N_KAN + 4   ' LfdN, Datum, KW, Kanaele, CSV-Erg, Bewertung, Grund, Overrange-Anz
    ReDim a(0 To nZ, 1 To nS)

    a(0, 1) = "LfdN": a(0, 2) = "Pruefzeit": a(0, 3) = "KW"
    For k = 1 To N_KAN
        a(0, 3 + k) = kKurz(k) & " [" & kEinheit(k) & "]"
    Next k
    a(0, 4 + N_KAN) = "Ergebnis CSV"
    a(0, 5 + N_KAN) = "Bewertung"
    a(0, 6 + N_KAN) = "Fehlergrund"
    a(0, 7 + N_KAN) = "Anz. Overrange"

    Dim nOv As Long
    For i = 1 To nZ
        a(i, 1) = rLfdN(i)
        If Not IsEmpty(rZeit(i)) Then
            a(i, 2) = rZeit(i) + gTzOffset / 24#
            a(i, 3) = IsoKW(CDate(a(i, 2)))
        End If
        nOv = 0
        For k = 1 To N_KAN
            If Not IsEmpty(rWert(i, k)) Then
                a(i, 3 + k) = rWert(i, k) * kFaktor(k)
                If rWert(i, k) >= gOverrange Then nOv = nOv + 1
            End If
        Next k
        If IsEmpty(rErg(i)) Then a(i, 4 + N_KAN) = "" Else a(i, 4 + N_KAN) = IIf(rErg(i), "true", "false")
        a(i, 5 + N_KAN) = bew(i)
        a(i, 6 + N_KAN) = grund(i)
        a(i, 7 + N_KAN) = nOv
    Next i

    ws.Range("A1").Resize(nZ + 1, nS).Value = a
    ws.Range("B2").Resize(nZ, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    For k = 1 To N_KAN
        ws.Cells(2, 3 + k).Resize(nZ, 1).NumberFormat = kFormat(k)
    Next k

    ' Als Tabelle formatieren (Filter, Streifen)
    TabelleAnlegen ws.Range("A1").Resize(nZ + 1, nS), "tblDaten"

    DatenFormatieren ws

    ws.Columns("A:" & SpBuchst(nS)).AutoFit
    ws.Columns(SpBuchst(6 + N_KAN)).ColumnWidth = 45
    Fixieren ws, "B2"
End Sub

'===============================================================================
'  STATISTIK
'===============================================================================
Private Sub StatistikSchreiben(bew() As Variant)
    Dim ws As Worksheet, k As Long, r As Long, i As Long
    Set ws = BlattHolen(WS_STA, True)
    ws.Range("A1").Value = "Statistik je Messkanal"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Kennwerte in Anzeigeeinheit. 'gueltig' = ohne Overrange (Messbereichsende). " & _
                           "Bei Isolationswiderstaenden sind geometrisches Mittel und log-Statistik aussagekraeftiger als das arithmetische Mittel."
    ws.Range("A2").Font.Italic = True

    Dim kz As Variant
    kz = Array("Kanal", "Typ", "Einheit", "UG", "OG", "Anzahl", "fehlend", "Overrange", "gueltig", _
               "Minimum", "P5", "Q1 (P25)", "Median", "Q3 (P75)", "P95", "Maximum", _
               "Mittelwert", "Std.-Abw.", "Geom. Mittel", "Std.-Abw. log10", _
               "< UG", "> OG", "Verletzungen %", "Ppk (log, UG)", "Werte = 0")
    For r = 0 To UBound(kz)
        ws.Cells(4 + r, 1).Value = kz(r)
    Next r
    KopfFormat ws.Range("A4").Resize(UBound(kz) + 1, 1)

    Dim v() As Double, lg() As Double, nV As Long, nL As Long
    Dim nOv As Long, nMiss As Long, nUG As Long, nOG As Long, n0 As Long, w As Double
    For k = 1 To N_KAN
        ReDim v(1 To nZ): ReDim lg(1 To nZ)
        nV = 0: nL = 0: nOv = 0: nMiss = 0: nUG = 0: nOG = 0: n0 = 0
        For i = 1 To nZ
            If IsEmpty(rWert(i, k)) Then
                nMiss = nMiss + 1
            ElseIf rWert(i, k) >= gOverrange Then
                nOv = nOv + 1
            Else
                w = rWert(i, k) * kFaktor(k)
                nV = nV + 1: v(nV) = w
                If rWert(i, k) > 0 Then nL = nL + 1: lg(nL) = Log10(rWert(i, k)) Else n0 = n0 + 1
                If Not IsEmpty(kUG(k)) Then If w < kUG(k) Then nUG = nUG + 1
                If Not IsEmpty(kOG(k)) Then If w > kOG(k) Then nOG = nOG + 1
            End If
        Next i
        If kTyp(k) = TYP_KONT Then nOG = nOG + nOv

        Dim c As Long: c = 1 + k
        ws.Cells(4, c).Value = kKurz(k)
        ws.Cells(5, c).Value = kTyp(k)
        ws.Cells(6, c).Value = kEinheit(k)
        ws.Cells(7, c).Value = kUG(k)
        ws.Cells(8, c).Value = kOG(k)
        ws.Cells(9, c).Value = nZ
        ws.Cells(10, c).Value = nMiss
        ws.Cells(11, c).Value = nOv
        ws.Cells(12, c).Value = nV
        If nV > 0 Then
            ReDim Preserve v(1 To nV)
            QuickSort v, 1, nV
            ws.Cells(13, c).Value = v(1)
            ws.Cells(14, c).Value = Perzentil(v, 0.05)
            ws.Cells(15, c).Value = Perzentil(v, 0.25)
            ws.Cells(16, c).Value = Perzentil(v, 0.5)
            ws.Cells(17, c).Value = Perzentil(v, 0.75)
            ws.Cells(18, c).Value = Perzentil(v, 0.95)
            ws.Cells(19, c).Value = v(nV)
            ws.Cells(20, c).Value = Mittel(v, nV)
            ws.Cells(21, c).Value = StdAbw(v, nV)
        End If
        If nL > 1 And kTyp(k) <> TYP_KONT Then
            ReDim Preserve lg(1 To nL)
            Dim mL As Double, sL As Double
            mL = Mittel(lg, nL): sL = StdAbw(lg, nL)
            ws.Cells(22, c).Value = (10 ^ mL) * kFaktor(k)
            ws.Cells(23, c).Value = sL
            If Not IsEmpty(kUG(k)) And kTyp(k) = TYP_ISO And sL > 0 And kUG(k) > 0 Then
                ws.Cells(27, c).Value = (mL - Log10(kUG(k) / kFaktor(k))) / (3 * sL)
            End If
        End If
        ws.Cells(24, c).Value = nUG
        ws.Cells(25, c).Value = nOG
        ws.Cells(26, c).Value = (nUG + nOG) / nZ
        ws.Cells(28, c).Value = n0
    Next k
    KopfFormat ws.Range("B4").Resize(1, N_KAN)
    For k = 1 To N_KAN
        ws.Cells(13, 1 + k).Resize(10, 1).NumberFormat = kFormat(k)
        ws.Cells(7, 1 + k).Resize(2, 1).NumberFormat = kFormat(k)
    Next k
    ws.Range("B23").Resize(1, N_KAN).NumberFormat = "0.000"
    ws.Range("B26").Resize(1, N_KAN).NumberFormat = "0.0%"
    ws.Range("B27").Resize(1, N_KAN).NumberFormat = "0.00"
    ws.Range("A4").Resize(25, N_KAN + 1).Borders.LineStyle = xlContinuous
    ws.Range("A4").Resize(25, N_KAN + 1).Borders.Color = RGB(191, 191, 191)
    PpkAmpel ws.Range("B27").Resize(1, N_KAN)
    ws.Range("A30").Value = "Ppk (log, UG) = (Mittel(log10 R) - log10 UG) / (3 * s(log10 R)), nur fuer ISO-Kanaele mit UG. " & _
                            ">= 1,33 faehig; 1,0 - 1,33 bedingt; < 1,0 nicht faehig."
    ws.Range("A30").Font.Italic = True

    ' --- Bewertung CSV vs. eigene Bewertung (Kreuztabelle) ---------------------
    Dim r0 As Long: r0 = 33
    ws.Cells(r0, 1).Value = "Vergleich Gesamtergebnis CSV <-> eigene Bewertung"
    ws.Cells(r0, 1).Font.Bold = True: ws.Cells(r0, 1).Font.Size = 12
    ws.Cells(r0 + 1, 1).Resize(1, 4).Value = Array("CSV \ Bewertung", "IO", "NIO", "Summe")
    ws.Cells(r0 + 2, 1).Value = "true"
    ws.Cells(r0 + 3, 1).Value = "false"
    ws.Cells(r0 + 4, 1).Value = "leer/unbekannt"
    Dim m(1 To 3, 1 To 2) As Long, ri As Long, ci As Long
    For i = 1 To nZ
        If IsEmpty(rErg(i)) Then
            ri = 3
        ElseIf rErg(i) = True Then
            ri = 1
        Else
            ri = 2
        End If
        If bew(i) = "IO" Then ci = 1 Else ci = 2
        m(ri, ci) = m(ri, ci) + 1
    Next i
    For ri = 1 To 3
        ws.Cells(r0 + 1 + ri, 2).Value = m(ri, 1)
        ws.Cells(r0 + 1 + ri, 3).Value = m(ri, 2)
        ws.Cells(r0 + 1 + ri, 4).Value = m(ri, 1) + m(ri, 2)
    Next ri
    ws.Cells(r0 + 5, 1).Value = "Summe"
    ws.Cells(r0 + 5, 2).Value = m(1, 1) + m(2, 1) + m(3, 1)
    ws.Cells(r0 + 5, 3).Value = m(1, 2) + m(2, 2) + m(3, 2)
    ws.Cells(r0 + 5, 4).Value = nZ
    ws.Cells(r0 + 5, 1).Resize(1, 4).Font.Bold = True
    KopfFormat ws.Cells(r0 + 1, 1).Resize(1, 4)
    ws.Cells(r0 + 1, 1).Resize(5, 4).Borders.LineStyle = xlContinuous

    ' --- Korrelationsmatrix (log10) ---------------------------------------------
    r0 = 41
    ws.Cells(r0, 1).Value = "Korrelationsmatrix (Pearson auf log10 der Rohwerte, ohne Overrange und Nullwerte)"
    ws.Cells(r0, 1).Font.Bold = True: ws.Cells(r0, 1).Font.Size = 12
    Dim a As Long, b As Long
    For a = 1 To N_KAN
        ws.Cells(r0 + 1, 1 + a).Value = kKurz(a)
        ws.Cells(r0 + 1 + a, 1).Value = kKurz(a)
        For b = 1 To N_KAN
            ws.Cells(r0 + 1 + a, 1 + b).Value = KorrLog(a, b)
        Next b
    Next a
    KopfFormat ws.Cells(r0 + 1, 2).Resize(1, N_KAN)
    KopfFormat ws.Cells(r0 + 2, 1).Resize(N_KAN, 1)
    With ws.Cells(r0 + 2, 2).Resize(N_KAN, N_KAN)
        .NumberFormat = "0.00"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    Farbskala ws.Cells(r0 + 2, 2).Resize(N_KAN, N_KAN)

    ' --- Klassierung Isolation (log) ---------------------------------------------
    r0 = 53
    ws.Cells(r0, 1).Value = "Haeufigkeitsverteilung (Klassen in log10 des Rohwerts; Overrange in letzter Klasse)"
    ws.Cells(r0, 1).Font.Bold = True: ws.Cells(r0, 1).Font.Size = 12
    Dim nKl As Long, kl As Long, untere As Double
    nKl = CLng((gHistBis - gHistVon) / gHistSchritt)
    ws.Cells(r0 + 1, 1).Value = "Klasse ab (log10)"
    ws.Cells(r0 + 1, 2).Value = "Klasse ab (Rohwert)"
    For k = 1 To N_KAN
        ws.Cells(r0 + 1, 2 + k).Value = kKurz(k)
    Next k
    Dim hist() As Long
    ReDim hist(0 To nKl + 1, 1 To N_KAN)
    For k = 1 To N_KAN
        For i = 1 To nZ
            If Not IsEmpty(rWert(i, k)) Then
                If rWert(i, k) <= 0 Then
                    kl = 0
                Else
                    kl = Int((Log10(rWert(i, k)) - gHistVon) / gHistSchritt) + 1
                    If kl < 1 Then kl = 0
                    If kl > nKl Then kl = nKl
                End If
                hist(kl, k) = hist(kl, k) + 1
            End If
        Next i
    Next k
    ws.Cells(r0 + 2, 1).Value = "< " & gHistVon & " bzw. 0"
    ws.Cells(r0 + 2, 2).Value = "< " & Format$(10 ^ gHistVon, "0E+00")
    For k = 1 To N_KAN
        ws.Cells(r0 + 2, 2 + k).Value = hist(0, k)
    Next k
    For kl = 1 To nKl
        untere = gHistVon + (kl - 1) * gHistSchritt
        ws.Cells(r0 + 2 + kl, 1).Value = untere
        ws.Cells(r0 + 2 + kl, 2).Value = 10 ^ untere
        For k = 1 To N_KAN
            ws.Cells(r0 + 2 + kl, 2 + k).Value = hist(kl, k)
        Next k
    Next kl
    ws.Cells(r0 + 3, 2).Resize(nKl, 1).NumberFormat = "0.0E+00"
    KopfFormat ws.Cells(r0 + 1, 1).Resize(1, N_KAN + 2)
    ws.Cells(r0 + 1, 1).Resize(nKl + 2, N_KAN + 2).Borders.LineStyle = xlContinuous
    ws.Cells(r0 + 1, 1).Resize(nKl + 2, N_KAN + 2).Borders.Color = RGB(191, 191, 191)
    ws.Range("Z1").Value = r0 + 2     ' Startzeile Histogramm (fuer Diagramme)
    ws.Range("Z2").Value = nKl + 1
    KlassenEinzelkanal ws, N_KAN, r0 + nKl + 5
    ws.Range("Z1:Z4").Font.Color = RGB(255, 255, 255)

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B:" & SpBuchst(N_KAN + 2)).ColumnWidth = 16
End Sub

Private Sub PpkAmpel(rg As Range)
    On Error Resume Next
    rg.FormatConditions.Delete
    rg.FormatConditions.Add(xlCellValue, xlLess, "=1").Interior.Color = RGB(255, 199, 206)
    rg.FormatConditions.Add(xlCellValue, xlBetween, "=1", "=133/100").Interior.Color = RGB(255, 235, 156)
    rg.FormatConditions.Add(xlCellValue, xlGreater, "=133/100").Interior.Color = RGB(198, 239, 206)
End Sub

' Farbskala blau (-1) - weiss (0) - rot (+1) fuer die Korrelationsmatrix
Private Sub Farbskala(rg As Range)
    Dim cs As ColorScale
    On Error Resume Next
    Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber: cs.ColorScaleCriteria(1).Value = -1
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(90, 138, 198)
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber: cs.ColorScaleCriteria(2).Value = 0
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
    cs.ColorScaleCriteria(3).Type = xlConditionValueNumber: cs.ColorScaleCriteria(3).Value = 1
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
End Sub

Private Sub Datenbalken(rg As Range)
    On Error Resume Next
    rg.FormatConditions.AddDatabar
End Sub

' Haeufigkeitsklassen eines Kanals mit eigenem, aus den Daten bestimmtem Bereich
' (logarithmische Klassen in der Anzeigeeinheit). Startzeile/Anzahl -> Z3/Z4.
Private Sub KlassenEinzelkanal(ws As Worksheet, ByVal k As Long, ByVal r0 As Long)
    Dim lg() As Double, n As Long, i As Long
    ReDim lg(1 To nZ)
    For i = 1 To nZ
        If Not IsEmpty(rWert(i, k)) Then
            If rWert(i, k) > 0 And rWert(i, k) < gOverrange Then
                n = n + 1: lg(n) = Log10(rWert(i, k) * kFaktor(k))
            End If
        End If
    Next i
    ws.Range("Z3").Value = 0: ws.Range("Z4").Value = 0
    If n < 2 Then Exit Sub
    ReDim Preserve lg(1 To n)
    QuickSort lg, 1, n
    ' Bereich P0,5 .. P99,5; bei weniger als einer Dekade lineare Klassen, sonst logarithmische
    Dim lo As Double, hi As Double, schritt As Double, nKl As Long, kl As Long, linear As Boolean, x As Double
    lo = Perzentil(lg, 0.005): hi = Perzentil(lg, 0.995)
    linear = (hi - lo <= 1)
    If linear Then
        lo = 10 ^ lo: hi = 10 ^ hi
        nKl = 25
        schritt = (hi - lo) / nKl
        If schritt <= 0 Then schritt = Abs(lo) * 0.01 + 1E-300
    Else
        If hi - lo <= 3 Then schritt = 0.1 Else schritt = 0.2
        lo = Int(lo / schritt) * schritt
        nKl = Int((hi - lo) / schritt) + 1
        If nKl > 80 Then nKl = 80
    End If
    Dim h() As Long
    ReDim h(0 To nKl + 1)
    For i = 1 To n
        If linear Then x = 10 ^ lg(i) Else x = lg(i)
        kl = Int((x - lo) / schritt) + 1
        If x - lo < 0 Then kl = 0
        If kl > nKl Then
            If linear And x <= hi * (1 + 1E-12) Then kl = nKl Else kl = nKl + 1
        End If
        h(kl) = h(kl) + 1
    Next i
    ws.Cells(r0, 1).Value = "Haeufigkeitsverteilung " & kKurz(k) & " [" & kEinheit(k) & "] (" & _
                            IIf(linear, "lineare", "logarithmische") & " Klassen, ohne Werte <= 0)"
    ws.Cells(r0, 1).Font.Bold = True: ws.Cells(r0, 1).Font.Size = 12
    ws.Cells(r0 + 1, 1).Resize(1, 2).Value = Array("Klasse ab [" & kEinheit(k) & "]", "Anzahl")
    KopfFormat ws.Cells(r0 + 1, 1).Resize(1, 2)
    Dim a() As Variant
    ReDim a(1 To nKl + 2, 1 To 2)
    a(1, 1) = "kleiner": a(1, 2) = h(0)
    For kl = 1 To nKl
        If linear Then a(kl + 1, 1) = lo + (kl - 1) * schritt Else a(kl + 1, 1) = 10 ^ (lo + (kl - 1) * schritt)
        a(kl + 1, 2) = h(kl)
    Next kl
    a(nKl + 2, 1) = "groesser": a(nKl + 2, 2) = h(nKl + 1)
    ws.Cells(r0 + 2, 1).Resize(nKl + 2, 2).Value = a
    ws.Cells(r0 + 3, 1).Resize(nKl, 1).NumberFormat = kFormat(k)
    ws.Cells(r0 + 3, 1).Resize(nKl, 1).HorizontalAlignment = xlLeft
    ws.Cells(r0 + 1, 1).Resize(nKl + 3, 2).Borders.LineStyle = xlContinuous
    ws.Cells(r0 + 1, 1).Resize(nKl + 3, 2).Borders.Color = RGB(191, 191, 191)
    ws.Range("Z3").Value = r0 + 2
    ws.Range("Z4").Value = nKl + 2
End Sub

Private Function KorrLog(ByVal a As Long, ByVal b As Long) As Variant
    Dim i As Long, n As Long, x As Double, y As Double
    Dim sx As Double, sy As Double, sxx As Double, syy As Double, sxy As Double
    For i = 1 To nZ
        If Not IsEmpty(rWert(i, a)) And Not IsEmpty(rWert(i, b)) Then
            If rWert(i, a) > 0 And rWert(i, b) > 0 And rWert(i, a) < gOverrange And rWert(i, b) < gOverrange Then
                x = Log10(rWert(i, a)): y = Log10(rWert(i, b))
                n = n + 1
                sx = sx + x: sy = sy + y
                sxx = sxx + x * x: syy = syy + y * y: sxy = sxy + x * y
            End If
        End If
    Next i
    Dim nen As Double
    If n < 20 Then KorrLog = CVErr(xlErrNA): Exit Function   ' zu wenige Wertepaare
    nen = Sqr((n * sxx - sx * sx) * (n * syy - sy * sy))
    If nen <= 0 Then KorrLog = CVErr(xlErrNA) Else KorrLog = (n * sxy - sx * sy) / nen
End Function

'===============================================================================
'  WOCHENAUSWERTUNG
'===============================================================================
Private Sub WochenSchreiben(bew() As Variant)
    Dim ws As Worksheet, i As Long, j As Long, d As Double, n As Long
    Dim mo() As Double, dMin As Double, dMax As Double, hat As Boolean
    Set ws = BlattHolen(WS_WOC, True)

    ' Montag der Woche (Excel-Datum) je Pruefung
    ReDim mo(1 To nZ)
    For i = 1 To nZ
        If IsEmpty(rZeit(i)) Then
            mo(i) = -1
        Else
            d = Int(rZeit(i) + gTzOffset / 24#)
            mo(i) = d - (Weekday(d, vbMonday) - 1)
            If Not hat Then dMin = mo(i): dMax = mo(i): hat = True
            If mo(i) < dMin Then dMin = mo(i)
            If mo(i) > dMax Then dMax = mo(i)
        End If
    Next i
    If Not hat Then
        ws.Range("A1").Value = "Keine gueltigen Pruefzeiten vorhanden."
        ws.Range("Z1").Value = 0
        Exit Sub
    End If

    ' Wochen-Klassen (auch Wochen ohne Pruefung, damit die Zeitachse stimmt)
    n = CLng((dMax - dMin) / 7) + 1
    Dim keys() As Double, wo As Collection, neu As Collection
    Set wo = New Collection                         ' je Woche eine Collection mit Zeilennummern
    ReDim keys(1 To n)
    For j = 1 To n
        keys(j) = dMin + 7 * (j - 1)
        Set neu = New Collection
        wo.Add neu
    Next j
    For i = 1 To nZ
        If mo(i) >= 0 Then wo(CLng((mo(i) - dMin) / 7) + 1).Add i
    Next i

    ws.Range("A1").Value = "Wochenauswertung"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Auswertung je Kalenderwoche (die Pruefzeiten im CSV sind teilweise gerundet, " & _
                           "eine Auswertung je Tag waere ungenau). Wochen ohne Pruefung sind mit 0 enthalten."
    ws.Range("A2").Font.Italic = True

    Dim c As Long, k As Long
    ws.Range("A4:F4").Value = Array("KW", "Woche ab", "Anzahl", "IO", "NIO", "NIO-Quote")
    c = 7
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO Then
            ws.Cells(4, c).Value = "Median " & kKurz(k) & " [" & kEinheit(k) & "]"
            c = c + 1
        End If
    Next k
    Dim nSp As Long: nSp = c - 1

    Dim a() As Variant, col As Collection, nio As Long, it As Variant, v() As Double, nv As Long
    ReDim a(1 To n, 1 To nSp)
    For j = 1 To n
        Set col = wo(j)
        nio = 0
        For Each it In col
            If bew(it) = "NIO" Then nio = nio + 1
        Next it
        a(j, 1) = "KW " & Format$(IsoKW(CDate(keys(j))), "00") & "/" & Year(CDate(keys(j)) + 3)
        a(j, 2) = CDate(keys(j))
        a(j, 3) = col.Count
        a(j, 4) = col.Count - nio
        a(j, 5) = nio
        If col.Count > 0 Then a(j, 6) = nio / col.Count Else a(j, 6) = CVErr(xlErrNA)
        c = 7
        For k = 1 To N_KAN
            If kTyp(k) = TYP_ISO Then
                ReDim v(1 To col.Count + 1): nv = 0
                For Each it In col
                    If Not IsEmpty(rWert(it, k)) Then nv = nv + 1: v(nv) = rWert(it, k) * kFaktor(k)
                Next it
                If nv > 0 Then
                    ReDim Preserve v(1 To nv)
                    QuickSort v, 1, nv
                    a(j, c) = Perzentil(v, 0.5)
                Else
                    a(j, c) = CVErr(xlErrNA)
                End If
                c = c + 1
            End If
        Next k
    Next j
    ws.Range("A5").Resize(n, nSp).Value = a
    ws.Range("B5").Resize(n, 1).NumberFormat = "dd.mm.yyyy"
    ws.Range("F5").Resize(n, 1).NumberFormat = "0.0%"
    c = 7
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO Then
            ws.Cells(5, c).Resize(n, 1).NumberFormat = kFormat(k)
            c = c + 1
        End If
    Next k
    KopfFormat ws.Range("A4").Resize(1, nSp)
    ws.Range("A4").Resize(n + 1, nSp).Borders.LineStyle = xlContinuous
    ws.Range("A4").Resize(n + 1, nSp).Borders.Color = RGB(191, 191, 191)
    Datenbalken ws.Range("F5").Resize(n, 1)
    ws.Columns("A:" & SpBuchst(nSp)).AutoFit
    ws.Range("Z1").Value = n
    ws.Range("Z1").Font.Color = RGB(255, 255, 255)
End Sub

'===============================================================================
'  NIO-LISTE
'===============================================================================
Private Sub NioListeSchreiben(bew() As Variant, grund() As String)
    Dim ws As Worksheet, i As Long, n As Long, k As Long, a() As Variant, j As Long
    Set ws = BlattHolen(WS_NIO, True)
    For i = 1 To nZ
        If bew(i) = "NIO" Then n = n + 1
    Next i
    ws.Range("A1").Value = "Auffaellige Pruefungen (eigene Bewertung NIO): " & n & " von " & nZ
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A3").Resize(1, 3 + N_KAN).Value = KopfNio()
    KopfFormat ws.Range("A3").Resize(1, 3 + N_KAN)
    If n = 0 Then Exit Sub
    ReDim a(1 To n, 1 To 3 + N_KAN)
    For i = 1 To nZ
        If bew(i) = "NIO" Then
            j = j + 1
            a(j, 1) = rLfdN(i)
            If Not IsEmpty(rZeit(i)) Then a(j, 2) = rZeit(i) + gTzOffset / 24#
            a(j, 3) = grund(i)
            For k = 1 To N_KAN
                If Not IsEmpty(rWert(i, k)) Then a(j, 3 + k) = rWert(i, k) * kFaktor(k)
            Next k
        End If
    Next i
    ws.Range("A4").Resize(n, 3 + N_KAN).Value = a
    ws.Range("B4").Resize(n, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    For k = 1 To N_KAN
        ws.Cells(4, 3 + k).Resize(n, 1).NumberFormat = kFormat(k)
    Next k
    ws.Columns("A:" & SpBuchst(3 + N_KAN)).AutoFit
    ws.Columns("C").ColumnWidth = 55
    ws.Range("A3").Resize(n + 1, 3 + N_KAN).AutoFilter
End Sub

' Bedingte Formatierung des Datenblatts: Grenzwertverletzungen rot, Overrange grau.
' Grenzen werden ueber Zellbezuege auf "Einstellungen" angegeben (unabhaengig vom Gebietsschema).
' Reine Optik - Fehler brechen die Auswertung nicht ab.
Private Sub DatenFormatieren(ws As Worksheet)
    Dim rg As Range, fc As FormatCondition, zs As Long, k As Long
    On Error Resume Next
    For k = 1 To N_KAN
        Set rg = ws.Cells(2, 3 + k).Resize(nZ, 1)
        zs = 4 + k
        Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, _
                                         Formula1:="=" & WS_SET & "!$J$" & zs)
        fc.Interior.Color = RGB(217, 217, 217): fc.StopIfTrue = True
        If Not IsEmpty(kUG(k)) Then
            Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, _
                                             Formula1:="=" & WS_SET & "!$G$" & zs)
            fc.Interior.Color = RGB(255, 199, 206): fc.Font.Color = RGB(156, 0, 6)
        End If
        If Not IsEmpty(kOG(k)) Then
            Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, _
                                             Formula1:="=" & WS_SET & "!$H$" & zs)
            fc.Interior.Color = RGB(255, 199, 206): fc.Font.Color = RGB(156, 0, 6)
        End If
    Next k
    Set rg = ws.Cells(2, 5 + N_KAN).Resize(nZ, 1)
    Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NIO""")
    fc.Interior.Color = RGB(255, 199, 206): fc.Font.Color = RGB(156, 0, 6): fc.Font.Bold = True
    Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""IO""")
    fc.Interior.Color = RGB(198, 239, 206): fc.Font.Color = RGB(0, 97, 0)
End Sub

Private Sub TabelleAnlegen(rg As Range, ByVal nm As String)
    Dim lo As ListObject
    On Error GoTo NurKopf
    Set lo = rg.Worksheet.ListObjects.Add(xlSrcRange, rg, , xlYes)
    lo.Name = nm
    lo.TableStyle = "TableStyleLight9"
    Exit Sub
NurKopf:
    KopfFormat rg.Rows(1)
End Sub

Private Function KopfNio() As Variant
    Dim a() As Variant, k As Long
    ReDim a(1 To 3 + N_KAN)
    a(1) = "LfdN": a(2) = "Pruefzeit": a(3) = "Fehlergrund"
    For k = 1 To N_KAN
        a(3 + k) = kKurz(k) & " [" & kEinheit(k) & "]"
    Next k
    KopfNio = a
End Function

'===============================================================================
'  HILFSDATEN FUER DIAGRAMME
'===============================================================================
' Spalten:  A LfdN | B..I  Werte log-tauglich (#NV bei <=0/leer) |
'           J..Q gleitender Median | R UG-Linie Iso | S..Z Kapazitaets-Dekaden
Private Sub HilfsdatenSchreiben()
    Dim ws As Worksheet, a() As Variant, i As Long, k As Long, j As Long
    Dim lo As Long, hi As Long, h As Long, buf() As Double, nb As Long
    Set ws = BlattHolen(WS_HLP, True)
    ReDim a(0 To nZ, 1 To 1 + 2 * N_KAN + 1)
    a(0, 1) = "LfdN"
    For k = 1 To N_KAN
        a(0, 1 + k) = kKurz(k)
        a(0, 1 + N_KAN + k) = "Median" & gFenster & " " & kKurz(k)
    Next k
    a(0, 2 + 2 * N_KAN) = "UG Isolation"
    h = gFenster \ 2
    For i = 1 To nZ
        a(i, 1) = rLfdN(i)
        For k = 1 To N_KAN
            If IsEmpty(rWert(i, k)) Then
                a(i, 1 + k) = CVErr(xlErrNA)
            ElseIf rWert(i, k) <= 0 Then
                a(i, 1 + k) = CVErr(xlErrNA)
            Else
                a(i, 1 + k) = rWert(i, k) * kFaktor(k)
            End If
            ' gleitender Median (zentriert)
            lo = i - h: hi = i + h
            If lo < 1 Then lo = 1
            If hi > nZ Then hi = nZ
            ReDim buf(1 To hi - lo + 1): nb = 0
            For j = lo To hi
                If Not IsEmpty(rWert(j, k)) Then
                    If rWert(j, k) > 0 Then nb = nb + 1: buf(nb) = rWert(j, k) * kFaktor(k)
                End If
            Next j
            If nb > 0 Then
                ReDim Preserve buf(1 To nb)
                QuickSort buf, 1, nb
                a(i, 1 + N_KAN + k) = Perzentil(buf, 0.5)
            Else
                a(i, 1 + N_KAN + k) = CVErr(xlErrNA)
            End If
        Next k
        a(i, 2 + 2 * N_KAN) = UgIso()
    Next i
    ws.Range("A1").Resize(nZ + 1, 2 + 2 * N_KAN).Value = a

    ws.Visible = xlSheetHidden
End Sub

Private Function UgIso() As Variant
    Dim k As Long
    UgIso = CVErr(xlErrNA)
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO And Not IsEmpty(kUG(k)) Then
            If k <= 4 Then UgIso = kUG(k): Exit Function
        End If
    Next k
End Function

'===============================================================================
'  DIAGRAMME
'===============================================================================
Private Sub DiagrammeErzeugen(pr As Range)
    Dim ws As Worksheet, wh As Worksheet, wsS As Worksheet, wW As Worksheet
    Dim ch As Chart, s As Series, k As Long
    Const W As Double = 620, H As Double = 340, GAP As Double = 15
    Set ws = BlattHolen(WS_DIA, True)
    Set wh = Worksheets(WS_HLP)
    Set wsS = Worksheets(WS_STA)
    Set wW = Worksheets(WS_WOC)
    OhneGitternetz ws
    ws.Range("A1").Value = "Diagramme"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Logarithmische Achsen bei Widerstaenden. Punkte am oberen Rand = Messbereichsende (Overrange)."
    ws.Range("A2").Font.Italic = True

    Dim xR As Range
    Set xR = wh.Range("A2").Resize(nZ, 1)
    Dim pos As Long: pos = 0

    ' 1) Verlauf Isolationswiderstaende ------------------------------------------
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlXYScatter, "Isolationswiderstaende - Verlauf ueber LfdN")
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO And k <= 4 Then
            Set s = ch.SeriesCollection.NewSeries
            s.Name = kKurz(k)
            s.XValues = xR
            s.Values = wh.Cells(2, 1 + k).Resize(nZ, 1)
            PunktSerie s, KanFarbe(k), 3
        End If
    Next k
    UgLinie ch, xR, wh.Cells(2, 2 + 2 * N_KAN).Resize(nZ, 1)
    AchsenLog ch, "LfdN", "Widerstand [" & kEinheit(1) & "]"

    ' 2) Gleitender Median -------------------------------------------------------
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlXYScatterLinesNoMarkers, _
                           "Trend: gleitender Median (" & gFenster & " Pruefungen)")
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO Then
            Set s = ch.SeriesCollection.NewSeries
            s.Name = kKurz(k)
            s.XValues = xR
            s.Values = wh.Cells(2, 1 + N_KAN + k).Resize(nZ, 1)
            s.Format.Line.ForeColor.RGB = KanFarbe(k)
            s.Format.Line.Weight = 2
        End If
    Next k
    UgLinie ch, xR, wh.Cells(2, 2 + 2 * N_KAN).Resize(nZ, 1)
    AchsenLog ch, "LfdN", "Median [" & kEinheit(1) & "]"

    ' 3) Histogramm Isolation (log-Klassen) ---------------------------------------
    Dim hr As Long, hn As Long
    hr = wsS.Range("Z1").Value: hn = wsS.Range("Z2").Value
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlColumnClustered, "Verteilung Isolationswiderstaende (Klassen log10 Ohm)")
    For k = 1 To N_KAN
        If kTyp(k) = TYP_ISO And k <= 4 Then
            Set s = ch.SeriesCollection.NewSeries
            s.Name = kKurz(k)
            s.XValues = wsS.Cells(hr, 1).Resize(hn, 1)
            s.Values = wsS.Cells(hr, 2 + k).Resize(hn, 1)
            s.Format.Fill.ForeColor.RGB = KanFarbe(k)
        End If
    Next k
    ch.ChartGroups(1).GapWidth = 30
    ch.ChartGroups(1).Overlap = 0
    AchsenTitel ch, "Klasse ab log10(R / Ohm)", "Anzahl Pruefungen"
    ch.Axes(xlCategory).TickLabels.NumberFormat = "0.00"

    ' 4) Pareto Fehlergruende -----------------------------------------------------
    If Not pr Is Nothing Then
        Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlColumnClustered, "Pareto der Fehlergruende (NIO)")
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "Anzahl"
        s.XValues = pr.Columns(1)
        s.Values = pr.Columns(2)
        s.Format.Fill.ForeColor.RGB = RGB(192, 80, 77)
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "kumuliert %"
        s.XValues = pr.Columns(1)
        s.Values = pr.Columns(3)
        s.ChartType = xlLineMarkers
        s.AxisGroup = xlSecondary
        s.Format.Line.ForeColor.RGB = RGB(64, 64, 64)
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 5
        ch.Axes(xlValue, xlSecondary).MinimumScale = 0
        ch.Axes(xlValue, xlSecondary).MaximumScale = 1
        ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
        ch.ChartGroups(1).GapWidth = 40
        AchsenTitel ch, "", "Anzahl"
    End If

    ' 5) IO/NIO Anteil ------------------------------------------------------------
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlDoughnut, "Anteil IO / NIO (eigene Bewertung)")
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Bewertung"
    s.XValues = Worksheets(WS_STA).Range("B34:C34")
    s.Values = Worksheets(WS_STA).Range("B38:C38")
    Dim pt As Point
    Set pt = s.Points(1): pt.Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
    Set pt = s.Points(2): pt.Format.Fill.ForeColor.RGB = RGB(192, 80, 77)
    s.HasDataLabels = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = True
    s.DataLabels.ShowCategoryName = True
    s.DataLabels.Separator = vbLf
    ch.HasLegend = False

    ' 6) Wochentrend --------------------------------------------------------------
    Dim nW As Long
    nW = wW.Range("Z1").Value
    If nW > 0 Then
        Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlColumnStacked, "Pruefungen und NIO-Quote je Kalenderwoche")
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "IO": s.XValues = wW.Range("A5").Resize(nW, 1): s.Values = wW.Range("D5").Resize(nW, 1)
        s.Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "NIO": s.XValues = wW.Range("A5").Resize(nW, 1): s.Values = wW.Range("E5").Resize(nW, 1)
        s.Format.Fill.ForeColor.RGB = RGB(192, 80, 77)
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "NIO-Quote": s.XValues = wW.Range("A5").Resize(nW, 1): s.Values = wW.Range("F5").Resize(nW, 1)
        s.ChartType = xlLineMarkers
        s.AxisGroup = xlSecondary
        s.Format.Line.ForeColor.RGB = RGB(64, 64, 64)
        ch.Axes(xlValue, xlSecondary).MinimumScale = 0
        ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
        ch.ChartGroups(1).GapWidth = 40
        AchsenTitel ch, "Kalenderwoche", "Anzahl Pruefungen"

        ' 7) Median je Woche ------------------------------------------------------
        Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlLineMarkers, "Median Isolationswiderstand je Kalenderwoche")
        Dim c As Long: c = 7
        For k = 1 To N_KAN
            If kTyp(k) = TYP_ISO Then
                Set s = ch.SeriesCollection.NewSeries
                s.Name = kKurz(k)
                s.XValues = wW.Range("A5").Resize(nW, 1)
                s.Values = wW.Cells(5, c).Resize(nW, 1)
                s.Format.Line.ForeColor.RGB = KanFarbe(k)
                s.MarkerBackgroundColor = KanFarbe(k)
                s.MarkerForegroundColor = KanFarbe(k)
                c = c + 1
            End If
        Next k
        ch.Axes(xlValue).ScaleType = xlScaleLogarithmic
        AchsenTitel ch, "Kalenderwoche", "Median [" & kEinheit(1) & "]"
    End If

    ' 8) Einzelkanaele 5-6 R und 5-6 C ------------------------------------------
    For k = 7 To N_KAN
        Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlXYScatter, kKurz(k) & " - Verlauf ueber LfdN")
        Set s = ch.SeriesCollection.NewSeries
        s.Name = kKurz(k)
        s.XValues = xR
        s.Values = wh.Cells(2, 1 + k).Resize(nZ, 1)
        PunktSerie s, KanFarbe(k), 3
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "gleitender Median"
        s.XValues = xR
        s.Values = wh.Cells(2, 1 + N_KAN + k).Resize(nZ, 1)
        s.ChartType = xlXYScatterLinesNoMarkers
        s.Format.Line.ForeColor.RGB = RGB(64, 64, 64)
        s.Format.Line.Weight = 2
        If Not IsEmpty(kUG(k)) Then GrenzLinie ch, xR, kUG(k), "UG"
        If Not IsEmpty(kOG(k)) Then GrenzLinie ch, xR, kOG(k), "OG"
        AchsenLog ch, "LfdN", kKurz(k) & " [" & kEinheit(k) & "]"
    Next k

    ' 9) Kapazitaet - Dekadenverteilung ------------------------------------------
    Dim kr As Long, kn As Long
    kr = Val(CStr(wsS.Range("Z3").Value)): kn = Val(CStr(wsS.Range("Z4").Value))
    If kn > 0 Then
        Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlColumnClustered, kKurz(N_KAN) & " - Verteilung")
        Set s = ch.SeriesCollection.NewSeries
        s.Name = kKurz(N_KAN)
        s.XValues = wsS.Cells(kr, 1).Resize(kn, 1)
        s.Values = wsS.Cells(kr, 2).Resize(kn, 1)
        s.Format.Fill.ForeColor.RGB = KanFarbe(N_KAN)
        ch.ChartGroups(1).GapWidth = 20
        ch.HasLegend = False
        AchsenTitel ch, "Klasse ab [" & kEinheit(N_KAN) & "]", "Anzahl"
    End If

    ' 10) Korrelation 5-7 vs 5-8 und 6-7 vs 6-8 ----------------------------------
    KorrDiagramm ws, pos, W, H, GAP, wh, 1, 2
    KorrDiagramm ws, pos, W, H, GAP, wh, 3, 4

    ' 11) Durchgangspruefung -----------------------------------------------------
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlXYScatter, "Durchgangspruefung 6-12 / 5-11 (Werte > 0)")
    For k = 1 To N_KAN
        If kTyp(k) = TYP_KONT Then
            Set s = ch.SeriesCollection.NewSeries
            s.Name = kKurz(k)
            s.XValues = xR
            s.Values = wh.Cells(2, 1 + k).Resize(nZ, 1)
            PunktSerie s, KanFarbe(k), 5
        End If
    Next k
    AchsenLog ch, "LfdN", "Widerstand [Ohm]  (0 Ohm wird nicht dargestellt)"
End Sub

Private Function NeuesDiagramm(ws As Worksheet, ByRef pos As Long, ByVal W As Double, ByVal H As Double, _
                               ByVal GAP As Double, ByVal typ As XlChartType, ByVal titel As String) As Chart
    Dim L As Double, T As Double, co As ChartObject
    L = 10 + (pos Mod 2) * (W + GAP)
    T = 45 + (pos \ 2) * (H + GAP)
    Set co = ws.ChartObjects.Add(L, T, W, H)
    co.Name = "Diagramm" & (pos + 1)
    co.Chart.ChartType = typ
    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = titel
    co.Chart.ChartTitle.Font.Size = 12
    co.Chart.ChartTitle.Font.Bold = True
    co.Chart.HasLegend = True
    co.Chart.Legend.Position = xlLegendPositionBottom
    pos = pos + 1
    Set NeuesDiagramm = co.Chart
End Function

Private Sub PunktSerie(s As Series, ByVal farbe As Long, ByVal groesse As Long)
    s.ChartType = xlXYScatter
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = groesse
    s.MarkerBackgroundColor = farbe
    s.MarkerForegroundColor = farbe
    s.Format.Line.Visible = msoFalse
End Sub

Private Sub UgLinie(ch As Chart, xR As Range, yR As Range)
    Dim s As Series
    If IsError(UgIso()) Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "UG " & UgIso() & " " & kEinheit(1)
    s.XValues = xR
    s.Values = yR
    s.ChartType = xlXYScatterLinesNoMarkers
    s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    s.Format.Line.DashStyle = msoLineDash
    s.Format.Line.Weight = 1.75
End Sub

Private Sub GrenzLinie(ch As Chart, xR As Range, ByVal wert As Double, ByVal nm As String)
    Dim s As Series, xa(1 To 2) As Double, ya(1 To 2) As Double
    xa(1) = xR.Cells(1, 1).Value: xa(2) = xR.Cells(xR.Rows.Count, 1).Value
    ya(1) = wert: ya(2) = wert
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm & " " & wert
    s.XValues = xa
    s.Values = ya
    s.ChartType = xlXYScatterLinesNoMarkers
    s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    s.Format.Line.DashStyle = msoLineDash
    s.Format.Line.Weight = 1.75
End Sub

Private Sub KorrDiagramm(ws As Worksheet, ByRef pos As Long, ByVal W As Double, ByVal H As Double, _
                         ByVal GAP As Double, wh As Worksheet, ByVal a As Long, ByVal b As Long)
    Dim ch As Chart, s As Series
    Dim r As Variant, rT As String
    r = KorrLog(a, b)
    If Not IsError(r) Then rT = "  (r log = " & Format$(r, "0.00") & ")"
    Set ch = NeuesDiagramm(ws, pos, W, H, GAP, xlXYScatter, "Korrelation " & kKurz(a) & " / " & kKurz(b) & rT)

    ' nur Wertepaare, bei denen beide Werte > 0 sind, kompakt in "Hilfsdaten" ablegen
    ' (#NV in den X-Werten eines Punktdiagramms wuerde Excel die X-Achse verfaelschen)
    Dim c0 As Long, i As Long, n As Long, xy() As Variant
    c0 = 8 + 2 * N_KAN + 2 * pos          ' freie Spalten rechts der Pareto-Tabelle
    ReDim xy(1 To nZ, 1 To 2)
    For i = 1 To nZ
        If Not IsEmpty(rWert(i, a)) And Not IsEmpty(rWert(i, b)) Then
            If rWert(i, a) > 0 And rWert(i, b) > 0 Then
                n = n + 1
                xy(n, 1) = rWert(i, a) * kFaktor(a)
                xy(n, 2) = rWert(i, b) * kFaktor(b)
            End If
        End If
    Next i
    wh.Cells(1, c0).Value = kKurz(a)
    wh.Cells(1, c0 + 1).Value = kKurz(b)
    If n = 0 Then Exit Sub
    wh.Cells(2, c0).Resize(nZ, 2).Value = xy

    Set s = ch.SeriesCollection.NewSeries
    s.Name = kKurz(a) & " vs. " & kKurz(b)
    s.XValues = wh.Cells(2, c0).Resize(n, 1)
    s.Values = wh.Cells(2, c0 + 1).Resize(n, 1)
    PunktSerie s, KanFarbe(a), 3
    ch.HasLegend = False
    ch.Axes(xlCategory).ScaleType = xlScaleLogarithmic
    ch.Axes(xlValue).ScaleType = xlScaleLogarithmic
    ch.Axes(xlCategory).TickLabels.NumberFormat = kFormat(a)
    ch.Axes(xlValue).TickLabels.NumberFormat = kFormat(b)
    AchsenTitel ch, kKurz(a) & " [" & kEinheit(a) & "]", kKurz(b) & " [" & kEinheit(b) & "]"
End Sub

Private Sub AchsenLog(ch As Chart, ByVal xT As String, ByVal yT As String)
    With ch.Axes(xlValue)
        .ScaleType = xlScaleLogarithmic
        .TickLabels.NumberFormat = "General"
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    End With
    With ch.Axes(xlCategory)
        .MinimumScale = 0
        .HasMajorGridlines = False
    End With
    AchsenTitel ch, xT, yT
End Sub

Private Sub AchsenTitel(ch As Chart, ByVal xT As String, ByVal yT As String)
    If xT <> "" Then
        ch.Axes(xlCategory).HasTitle = True
        ch.Axes(xlCategory).AxisTitle.Text = xT
    End If
    If yT <> "" Then
        ch.Axes(xlValue).HasTitle = True
        ch.Axes(xlValue).AxisTitle.Text = yT
    End If
End Sub

' Erzeugt Pareto-Tabelle der Einzel-Fehlergruende auf dem Blatt "Hilfsdaten"
Private Function ParetoTabelle(grund() As String) As Range
    Dim wh As Worksheet, i As Long, teile() As String, t As Variant
    Dim idx As New Collection             ' Fehlergrund -> Position in k()/v()
    Dim n As Long, k() As String, v() As Long, j As Long, tmpS As String, tmpL As Long, sum As Double
    Dim g As String
    Set wh = Worksheets(WS_HLP)
    ReDim k(1 To 100): ReDim v(1 To 100)
    For i = 1 To nZ
        g = grund(i)
        If g <> "" Then
            teile = Split(g, "; ")
            For Each t In teile
                j = SchluesselPos(idx, CStr(t))
                If j = 0 Then
                    n = n + 1
                    If n > UBound(k) Then
                        ReDim Preserve k(1 To 2 * n)
                        ReDim Preserve v(1 To 2 * n)
                    End If
                    k(n) = CStr(t)
                    idx.Add n, CStr(t)
                    j = n
                End If
                v(j) = v(j) + 1
                sum = sum + 1
            Next t
        End If
    Next i
    If n = 0 Then Exit Function
    ' absteigend sortieren (n klein -> Bubblesort)
    For i = 1 To n - 1
        For j = i + 1 To n
            If v(j) > v(i) Then
                tmpL = v(i): v(i) = v(j): v(j) = tmpL
                tmpS = k(i): k(i) = k(j): k(j) = tmpS
            End If
        Next j
    Next i
    Dim c0 As Long: c0 = 4 + 2 * N_KAN    ' Spalte nach Hilfsdaten
    wh.Cells(1, c0).Resize(1, 3).Value = Array("Fehlergrund", "Anzahl", "kumuliert")
    Dim cum As Double
    For i = 1 To n
        cum = cum + v(i)
        wh.Cells(1 + i, c0).Value = k(i)
        wh.Cells(1 + i, c0 + 1).Value = v(i)
        wh.Cells(1 + i, c0 + 2).Value = cum / sum
    Next i
    Set ParetoTabelle = wh.Cells(2, c0).Resize(n, 3)
End Function

'===============================================================================
'  UEBERSICHT
'===============================================================================
Private Sub UebersichtSchreiben(ByVal quelle As String, bew() As Variant, grund() As String)
    Dim ws As Worksheet, i As Long, k As Long, nIO As Long, nNIO As Long
    Dim tMin As Double, tMax As Double, hatZeit As Boolean, nFalse As Long, nTrue As Long
    Set ws = BlattHolen(WS_UEB, True)
    OhneGitternetz ws

    For i = 1 To nZ
        If bew(i) = "IO" Then nIO = nIO + 1 Else nNIO = nNIO + 1
        If Not IsEmpty(rErg(i)) Then
            If rErg(i) = True Then
                nTrue = nTrue + 1
            Else
                nFalse = nFalse + 1
            End If
        End If
        If Not IsEmpty(rZeit(i)) Then
            If Not hatZeit Then tMin = rZeit(i): tMax = rZeit(i): hatZeit = True
            If rZeit(i) < tMin Then tMin = rZeit(i)
            If rZeit(i) > tMax Then tMax = rZeit(i)
        End If
    Next i

    With ws
        .Range("A1").Value = "Pruefplatz-Auswertung"
        .Range("A1").Font.Size = 20: .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(31, 78, 121)
        .Range("A2").Value = "Quelle: " & quelle & "   |   erstellt: " & Format$(Now, "dd.mm.yyyy hh:nn")
        .Range("A2").Font.Color = RGB(118, 118, 118)

        ' KPI-Kacheln
        KPI ws, "A4", "Pruefungen", Format$(nZ, "#,##0"), RGB(31, 78, 121)
        KPI ws, "D4", "IO (eigene Bewertung)", Format$(nIO, "#,##0"), RGB(84, 130, 53)
        KPI ws, "G4", "NIO (eigene Bewertung)", Format$(nNIO, "#,##0"), RGB(192, 0, 0)
        KPI ws, "J4", "Ausbeute (First Pass Yield)", Format$(nIO / nZ, "0.0%"), RGB(84, 130, 53)
        If hatZeit Then
            KPI ws, "M4", "Zeitraum", Format$(tMin + gTzOffset / 24#, "dd.mm.yy") & " - " & _
                Format$(tMax + gTzOffset / 24#, "dd.mm.yy"), RGB(89, 89, 89)
        End If

        ' Auffaelligkeiten
        Dim r As Long: r = 9
        .Cells(r, 1).Value = "Auffaelligkeiten und Hinweise (automatisch ermittelt)"
        .Cells(r, 1).Font.Bold = True: .Cells(r, 1).Font.Size = 13
        r = r + 1
        Dim txt As Collection: Set txt = New Collection

        If nFalse = nZ Then
            txt.Add "Das Feld 'Gesamtergebnis' ist im CSV bei ALLEN " & nZ & " Pruefungen 'false' - auch bei Pruefungen, " & _
                    "deren Messwerte unauffaellig sind. Das Feld scheint vom Pruefplatz nicht korrekt gesetzt zu werden. " & _
                    "Die Auswertung verwendet deshalb eine eigene Bewertung anhand der Grenzwerte (Blatt Einstellungen)."
        ElseIf nTrue + nFalse > 0 Then
            txt.Add "Gesamtergebnis lt. CSV: " & nTrue & " x true, " & nFalse & " x false. Vergleich mit eigener Bewertung: Blatt Statistik."
        End If

        ' Overrange je Kanal
        Dim s As String, nOv As Long
        s = ""
        For k = 1 To N_KAN
            nOv = 0
            For i = 1 To nZ
                If Not IsEmpty(rWert(i, k)) Then If rWert(i, k) >= gOverrange Then nOv = nOv + 1
            Next i
            If nOv > 0 Then s = s & ", " & kKurz(k) & ": " & nOv & " (" & Format$(nOv / nZ, "0.0%") & ")"
        Next k
        If s <> "" Then txt.Add "Messbereichsende (Rohwert >= " & Format$(gOverrange, "0.0E+00") & ") erreicht bei " & Mid$(s, 3) & _
                                ". Bei Isolationskanaelen bedeutet das 'sehr hochohmig' (" & IIf(gOverrangeIstIO, "als IO gewertet", "als NIO gewertet") & ")."

        ' Pruefungen mit (fast) allen Kanaelen im Overrange = vermutlich kein Pruefling
        Dim nLeer As Long, lst As String, cnt As Long
        For i = 1 To nZ
            cnt = 0
            For k = 1 To N_KAN
                If Not IsEmpty(rWert(i, k)) Then If rWert(i, k) >= gOverrange Then cnt = cnt + 1
            Next k
            If cnt >= N_KAN - 1 Then nLeer = nLeer + 1: lst = lst & ", " & rLfdN(i)
        Next i
        If nLeer > 0 Then txt.Add nLeer & " Pruefung(en) mit (nahezu) allen Kanaelen im Overrange - vermutlich ohne kontaktierten Pruefling: LfdN " & Mid$(lst, 3)

        ' Durchgangspruefung
        For k = 1 To N_KAN
            If kTyp(k) = TYP_KONT And Not IsEmpty(kOG(k)) Then
                cnt = 0
                For i = 1 To nZ
                    If Not IsEmpty(rWert(i, k)) Then If rWert(i, k) * kFaktor(k) > kOG(k) Then cnt = cnt + 1
                Next i
                If cnt > 0 Then txt.Add "Durchgang " & kKurz(k) & ": " & cnt & " Pruefung(en) hochohmig/offen (> " & kOG(k) & " " & kEinheit(k) & _
                                        ") - Hinweis auf Kontaktierungs- oder Leitungsfehler."
            End If
        Next k

        ' Kapazitaet: Dekaden
        ' Zweigipfligkeit: groesste Luecke in den sortierten log-Werten
        Dim lgC() As Double, nC As Long, n0 As Long, gPos As Long, jj As Long
        ReDim lgC(1 To nZ)
        For i = 1 To nZ
            If Not IsEmpty(rWert(i, N_KAN)) Then
                If rWert(i, N_KAN) > 0 And rWert(i, N_KAN) < gOverrange Then
                    nC = nC + 1: lgC(nC) = Log10(rWert(i, N_KAN) * kFaktor(N_KAN))
                Else
                    If rWert(i, N_KAN) <= 0 Then n0 = n0 + 1
                End If
            End If
        Next i
        If nC > 20 Then
            ' Klassen zu 0,1 Dekaden; gesucht ist ein Tal (3 Klassen mit <= 1 % der Werte),
            ' links und rechts davon jeweils mindestens 10 % der Werte
            ReDim Preserve lgC(1 To nC)
            QuickSort lgC, 1, nC
            Dim b0 As Double, nb As Long, hb() As Long, bb As Long, links As Long, mitte As Long, tal As Long
            b0 = Int(lgC(1) * 10) / 10
            nb = Int((lgC(nC) - b0) * 10) + 1
            If nb > 400 Then nb = 400
            ReDim hb(1 To nb)
            For jj = 1 To nC
                bb = Int((lgC(jj) - b0) * 10) + 1
                If bb > nb Then bb = nb
                hb(bb) = hb(bb) + 1
            Next jj
            For bb = 2 To nb - 1
                mitte = hb(bb - 1) + hb(bb) + hb(bb + 1)
                If mitte <= 0.01 * nC And links >= 0.1 * nC And nC - links - mitte >= 0.1 * nC Then
                    If tal = 0 Then tal = bb: gPos = links + hb(bb - 1)
                End If
                links = links + hb(bb - 1)
            Next bb
            If tal > 0 Then
                txt.Add kKurz(N_KAN) & " ist zweigipflig: " & gPos & " Werte um " & _
                        Format$(10 ^ lgC((1 + gPos) \ 2), kFormat(N_KAN)) & " " & kEinheit(N_KAN) & " und " & (nC - gPos) & _
                        " Werte um " & Format$(10 ^ lgC((gPos + 1 + nC) \ 2), kFormat(N_KAN)) & " " & kEinheit(N_KAN) & _
                        ". Moegliche Ursachen: Messbereichsumschaltung / Einheitenfehler im Pruefprogramm " & _
                        "oder zwei Produktvarianten - bitte pruefen."
            End If
        End If
        If n0 > 0 Then txt.Add kKurz(N_KAN) & ": " & n0 & " Messung(en) mit Wert 0 (keine Messung?)."

        ' Datenqualitaet
        Dim nDq As Long
        If BlattExistiert(WS_DQ) Then nDq = CLng(Val(CStr(Worksheets(WS_DQ).Range("F1").Value)))
        If nDq > 0 Then
            txt.Add nDq & " Formatabweichung(en) im CSV (z. B. Dezimalpunkt statt -komma, Tausenderpunkte) - korrigiert eingelesen, siehe Blatt Datenqualitaet."
        End If
        ' Zeitaufloesung pruefen (viele identische Zeitstempel = gerundete Werte)
        Dim zt() As Double, nzt As Long, nVersch As Long
        ReDim zt(1 To nZ)
        For i = 1 To nZ
            If Not IsEmpty(rZeit(i)) Then nzt = nzt + 1: zt(nzt) = rZeit(i)
        Next i
        If nzt > 0 Then
            QuickSort zt, 1, nzt
            nVersch = 1
            For i = 2 To nzt
                If zt(i) <> zt(i - 1) Then nVersch = nVersch + 1
            Next i
        End If
        If hatZeit And nVersch < nZ / 2 Then
            txt.Add "Nur " & nVersch & " unterschiedliche Pruefzeiten bei " & nZ & " Pruefungen: Die Pruefzeit ist im CSV gerundet " & _
                    "(vermutlich nur 6 signifikante Stellen, Aufloesung ca. 2,8 h - Datei wurde wohl zwischenzeitlich in Excel gespeichert). " & _
                    "Zeitbezogene Auswertungen erfolgen deshalb je Kalenderwoche."
        End If
        ' Plausibilitaet der Einheit: Spaltenkopf "MOhm", Werte aber in Ohm?
        Dim tmp() As Double, nt As Long
        ReDim tmp(1 To nZ)
        For i = 1 To nZ
            If Not IsEmpty(rWert(i, 1)) Then If rWert(i, 1) > 0 Then nt = nt + 1: tmp(nt) = rWert(i, 1)
        Next i
        If nt > 0 Then
            ReDim Preserve tmp(1 To nt)
            QuickSort tmp, 1, nt
            If Perzentil(tmp, 0.5) > 100000000# Then
                txt.Add "Die Spaltenueberschriften nennen 'MOhm', die Rohwerte liegen jedoch offensichtlich in Ohm vor (Median " & kKurz(1) & _
                        " = " & Format$(Perzentil(tmp, 0.5), "0.0E+00") & "). Umrechnung ueber den Faktor im Blatt Einstellungen."
            End If
        End If

        ' Kanal mit niedrigstem Ppk
        Dim wsS As Worksheet, minP As Double, minK As Long
        Set wsS = Worksheets(WS_STA)
        minP = 1E+99
        For k = 1 To N_KAN
            If IsNumeric(wsS.Cells(27, 1 + k).Value) And Not IsEmpty(wsS.Cells(27, 1 + k).Value) Then
                If wsS.Cells(27, 1 + k).Value < minP Then minP = wsS.Cells(27, 1 + k).Value: minK = k
            End If
        Next k
        If minK > 0 Then txt.Add "Kritischster Isolationskanal: " & kKurz(minK) & " mit Ppk(log) = " & Format$(minP, "0.00") & _
                                 IIf(minP < 1, " (nicht prozessfaehig gegen UG " & kUG(minK) & " " & kEinheit(minK) & ")", "") & "."

        Dim t As Variant
        For Each t In txt
            .Cells(r, 1).Value = ChrW(8226) & " " & t
            .Range(.Cells(r, 1), .Cells(r, 15)).Merge
            .Cells(r, 1).WrapText = True
            .Rows(r).RowHeight = 15 * (1 + Len(t) \ 150)
            .Cells(r, 1).VerticalAlignment = xlTop
            r = r + 1
        Next t

        ' Top-Fehlergruende
        r = r + 1
        .Cells(r, 1).Value = "Haeufigste Fehlergruende"
        .Cells(r, 1).Font.Bold = True: .Cells(r, 1).Font.Size = 13
        r = r + 1
        Dim wh As Worksheet, c0 As Long, j As Long
        Set wh = Worksheets(WS_HLP)
        c0 = 4 + 2 * N_KAN
        .Cells(r, 1).Resize(1, 3).Value = Array("Fehlergrund", "Anzahl", "Anteil an allen Pruefungen")
        KopfFormat .Cells(r, 1).Resize(1, 3)
        j = 2
        Do While wh.Cells(j, c0).Value <> "" And j <= 11
            .Cells(r + j - 1, 1).Value = wh.Cells(j, c0).Value
            .Cells(r + j - 1, 2).Value = wh.Cells(j, c0 + 1).Value
            .Cells(r + j - 1, 3).Value = wh.Cells(j, c0 + 1).Value / nZ
            .Cells(r + j - 1, 3).NumberFormat = "0.0%"
            j = j + 1
        Loop
        If j > 2 Then .Cells(r, 1).Resize(j - 1, 3).Borders.LineStyle = xlContinuous

        ' Kennwerte Kurzform
        r = r + j + 1
        .Cells(r, 1).Value = "Kennwerte (Median / P5 / geometr. Mittel)"
        .Cells(r, 1).Font.Bold = True: .Cells(r, 1).Font.Size = 13
        r = r + 1
        .Cells(r, 1).Resize(1, 7).Value = Array("Kanal", "Einheit", "Median", "P5", "Geom. Mittel", "Verletzungen", "Ppk (log)")
        KopfFormat .Cells(r, 1).Resize(1, 7)
        For k = 1 To N_KAN
            .Cells(r + k, 1).Value = kKurz(k)
            .Cells(r + k, 2).Value = kEinheit(k)
            .Cells(r + k, 3).Value = wsS.Cells(16, 1 + k).Value
            .Cells(r + k, 4).Value = wsS.Cells(14, 1 + k).Value
            .Cells(r + k, 5).Value = wsS.Cells(22, 1 + k).Value
            .Cells(r + k, 6).Value = wsS.Cells(24, 1 + k).Value + wsS.Cells(25, 1 + k).Value
            .Cells(r + k, 7).Value = wsS.Cells(27, 1 + k).Value
        Next k
        For k = 1 To N_KAN
            .Cells(r + k, 3).Resize(1, 3).NumberFormat = kFormat(k)
        Next k
        .Cells(r + 1, 7).Resize(N_KAN, 1).NumberFormat = "0.00"
        .Cells(r, 1).Resize(N_KAN + 1, 7).Borders.LineStyle = xlContinuous

        .Columns("A").ColumnWidth = 34
        .Columns("B:O").ColumnWidth = 13
        .Cells(r + N_KAN + 2, 1).Value = "Details: Blaetter Diagramme, Statistik, Wochen, NIO_Liste. Grenzwerte: Blatt Einstellungen -> danach Makro Neu_Berechnen."
        .Cells(r + N_KAN + 2, 1).Font.Italic = True
    End With
End Sub

Private Sub KPI(ws As Worksheet, ByVal adr As String, ByVal titel As String, ByVal wert As String, ByVal farbe As Long)
    Dim rg As Range
    Set rg = ws.Range(adr).Resize(1, 3)
    rg.Merge
    rg.Value = titel
    rg.Font.Color = RGB(255, 255, 255)
    rg.Font.Size = 10
    rg.Interior.Color = farbe
    rg.HorizontalAlignment = xlCenter
    Set rg = rg.Offset(1, 0).Resize(2, 3)
    rg.Merge
    rg.Value = "'" & wert
    rg.Font.Size = 20
    rg.Font.Bold = True
    rg.Font.Color = farbe
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.Interior.Color = RGB(242, 242, 242)
    ws.Range(adr).Resize(3, 3).BorderAround xlContinuous, xlThin, , farbe
End Sub

'===============================================================================
'  DATENQUALITAET
'===============================================================================
Private Sub DatenqualitaetSchreiben()
    Dim ws As Worksheet, n As Long
    Set ws = BlattHolen(WS_DQ, False)
    n = CLng(Val(CStr(ws.Range("F1").Value)))
    ws.Range("A1").Value = "Importprotokoll / Datenqualitaet: " & n & " Hinweis(e)"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A3:D3").Value = Array("CSV-Zeile", "Spalte", "Originalwert", "Hinweis")
    KopfFormat ws.Range("A3:D3")
    ws.Columns("A:D").AutoFit
    If ws.Columns("D").ColumnWidth > 80 Then ws.Columns("D").ColumnWidth = 80
End Sub

'===============================================================================
'  HILFSFUNKTIONEN
'===============================================================================
' Position eines Schluessels in einer Collection (0 = nicht vorhanden)
Private Function SchluesselPos(col As Collection, ByVal key As String) As Long
    On Error Resume Next
    SchluesselPos = 0
    SchluesselPos = col(key)
End Function

' Zahlenformat je Kanal aus der Groessenordnung der Werte (Median der Betraege > 0)
Private Sub FormateBestimmen()
    Dim k As Long, i As Long, v() As Double, n As Long, m As Double
    For k = 1 To N_KAN
        ReDim v(1 To nZ): n = 0
        For i = 1 To nZ
            If Not IsEmpty(rWert(i, k)) Then
                If rWert(i, k) <> 0 And Abs(rWert(i, k)) < gOverrange Then n = n + 1: v(n) = Abs(rWert(i, k) * kFaktor(k))
            End If
        Next i
        If n = 0 Then
            m = 1
        Else
            ReDim Preserve v(1 To n)
            QuickSort v, 1, n
            m = Perzentil(v, 0.5)
        End If
        kFormat(k) = FormatFuer(m)
    Next k
End Sub

Private Function FormatFuer(ByVal m As Double) As String
    If m >= 1000 Then
        FormatFuer = "#,##0"
    ElseIf m >= 10 Then
        FormatFuer = "#,##0.0"
    ElseIf m >= 0.1 Then
        FormatFuer = "#,##0.000"
    Else
        FormatFuer = "0.000E+00"
    End If
End Function

' Nach dem Import: Einheit der Kapazitaet automatisch erkennen, solange im Blatt
' "Einstellungen" noch die Standardwerte (Faktor 1, Einheit "Rohwert ...") stehen.
' Werte < 1E-3 sind offensichtlich in Farad -> Anzeige in nF.
Private Sub EinheitenAutomatisch()
    Dim ws As Worksheet, k As Long, i As Long, v() As Double, n As Long, m As Double
    Set ws = Worksheets(WS_SET)
    For k = 1 To N_KAN
        If InStr(1, CStr(ws.Cells(4 + k, 2).Value), "Kapazit", vbTextCompare) > 0 _
           And Val(CStr(ws.Cells(4 + k, 5).Value)) = 1 _
           And Left$(CStr(ws.Cells(4 + k, 6).Value), 7) = "Rohwert" Then
            ReDim v(1 To nZ): n = 0
            For i = 1 To nZ
                If Not IsEmpty(rWert(i, k)) Then
                    If rWert(i, k) > 0 And rWert(i, k) < gOverrangeStandard() Then n = n + 1: v(n) = rWert(i, k)
                End If
            Next i
            If n > 0 Then
                ReDim Preserve v(1 To n)
                QuickSort v, 1, n
                m = Perzentil(v, 0.5)
                If m < 0.001 Then
                    ws.Cells(4 + k, 5).Value = 1000000000#
                    ws.Cells(4 + k, 6).Value = "nF"
                    ws.Cells(4 + k, 9).Value = "Rohwert in Farad (automatisch erkannt) -> Anzeige in nF"
                End If
            End If
        End If
    Next k
End Sub

Private Function gOverrangeStandard() As Double
    gOverrangeStandard = ZahlOder(Worksheets(WS_SET).Range("B16").Value, 9900000000000#)
End Function

Private Sub InitFarben()
    KanFarbe(1) = RGB(31, 119, 180)
    KanFarbe(2) = RGB(255, 127, 14)
    KanFarbe(3) = RGB(44, 160, 44)
    KanFarbe(4) = RGB(148, 103, 189)
    KanFarbe(5) = RGB(214, 39, 40)
    KanFarbe(6) = RGB(140, 86, 75)
    KanFarbe(7) = RGB(23, 190, 207)
    KanFarbe(8) = RGB(188, 189, 34)
End Sub

Private Sub Performance(ByVal an As Boolean)
    Application.ScreenUpdating = Not an
    Application.EnableEvents = Not an
    Application.DisplayAlerts = Not an
    If an Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

Private Function BlattExistiert(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    BlattExistiert = Not ws Is Nothing
End Function

' Holt ein Blatt (legt es an) und leert es optional komplett
Private Function BlattHolen(ByVal nm As String, ByVal leeren As Boolean) As Worksheet
    Dim ws As Worksheet, co As ChartObject
    If BlattExistiert(nm) Then
        Set ws = ThisWorkbook.Worksheets(nm)
        If leeren Then
            ws.Visible = xlSheetVisible
            Dim i As Long
            For i = ws.ChartObjects.Count To 1 Step -1
                ws.ChartObjects(i).Delete
            Next i
            BlattAufraeumen ws
            ws.Cells.Clear
        End If
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Activate
    Set BlattHolen = ws
End Function

' Entfernt Tabellen, Filter, bedingte Formate und Gueltigkeiten (Fehler werden ignoriert,
' z. B. wenn ein Blatt nichts davon enthaelt)
Private Sub BlattAufraeumen(ws As Worksheet)
    Dim n As Long, i As Long
    On Error Resume Next
    n = ws.ListObjects.Count          ' bleibt 0, falls nicht verfuegbar
    For i = n To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.FormatConditions.Delete
    ws.Cells.Validation.Delete
    ws.Rows.RowHeight = ws.StandardHeight
End Sub

Private Sub BlattReihenfolge()
    Dim nm As Variant, vorher As Worksheet, i As Long
    Dim ordnung As Variant
    ordnung = Array(WS_START, WS_UEB, WS_DIA, WS_STA, WS_WOC, WS_NIO, WS_DAT, WS_SET, WS_DQ, WS_RAW, WS_HLP)
    If BlattExistiert(WS_HLP) Then Worksheets(WS_HLP).Visible = xlSheetVisible
    For i = UBound(ordnung) To 0 Step -1
        If BlattExistiert(CStr(ordnung(i))) Then
            Worksheets(CStr(ordnung(i))).Move Before:=ThisWorkbook.Worksheets(1)
        End If
    Next i
    Worksheets(WS_HLP).Visible = xlSheetHidden
    On Error Resume Next
    Worksheets(WS_UEB).Tab.Color = RGB(31, 78, 121)
    Worksheets(WS_DIA).Tab.Color = RGB(31, 119, 180)
    Worksheets(WS_SET).Tab.Color = RGB(255, 192, 0)
    Worksheets(WS_NIO).Tab.Color = RGB(192, 0, 0)
End Sub

Private Sub OhneGitternetz(ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub Fixieren(ws As Worksheet, ByVal adr As String)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(adr).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub KopfFormat(rg As Range)
    rg.Font.Bold = True
    rg.Font.Color = RGB(255, 255, 255)
    rg.Interior.Color = RGB(31, 78, 121)
    rg.WrapText = True
    rg.VerticalAlignment = xlCenter
End Sub

Private Function SpBuchst(ByVal n As Long) As String
    SpBuchst = Split(Cells(1, n).Address(True, False), "$")(0)
End Function

Private Function Log10(ByVal x As Double) As Double
    Log10 = Log(x) / Log(10#)
End Function

Private Function Mittel(v() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double
    For i = 1 To n: s = s + v(i): Next i
    Mittel = s / n
End Function

Private Function StdAbw(v() As Double, ByVal n As Long) As Double
    Dim i As Long, m As Double, s As Double
    If n < 2 Then Exit Function
    m = Mittel(v, n)
    For i = 1 To n: s = s + (v(i) - m) ^ 2: Next i
    StdAbw = Sqr(s / (n - 1))
End Function

' Perzentil (wie Excel QUANTIL.INKL) auf sortiertem Feld 1..n
Private Function Perzentil(v() As Double, ByVal p As Double) As Double
    Dim n As Long, h As Double, lo As Long
    n = UBound(v) - LBound(v) + 1
    If n = 1 Then Perzentil = v(LBound(v)): Exit Function
    h = (n - 1) * p + 1
    lo = Int(h)
    If lo >= n Then Perzentil = v(n): Exit Function
    Perzentil = v(lo) + (h - lo) * (v(lo + 1) - v(lo))
End Function

Private Sub QuickSort(v() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi
    p = v((lo + hi) \ 2)
    Do While i <= j
        Do While v(i) < p: i = i + 1: Loop
        Do While v(j) > p: j = j - 1: Loop
        If i <= j Then
            t = v(i): v(i) = v(j): v(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort v, lo, j
    If i < hi Then QuickSort v, i, hi
End Sub

' ISO-Kalenderwoche
Private Function IsoKW(ByVal d As Date) As Long
    Dim dRef As Date
    dRef = DateSerial(Year(d - Weekday(d - 1) + 4), 1, 3)
    IsoKW = Int((d - dRef + Weekday(dRef) + 5) / 7)
End Function
