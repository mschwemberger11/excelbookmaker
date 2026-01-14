Attribute VB_Name = "Berichte"
  '*******************************************************************************
  '* MAKRO Berichte *                                  * Version 19-3  8.2.2019 *
  'Aufruf: Mit Strg+b von der ARPROT I-Spalte aus über Sub Erfassen.             *
  '*******************************************************************************
  'Verwendet aus Modul KontenplanPflege die Info-Routinen:
     'Sub KontenplanStruktur Globale Strukturdaten des Kontenplans und
     'Sub KtoKennDat(KtoNr)  Zustandsdaten von Konten
  
  'Erstellungs- und Druck-Möglichkeit über Dialogfenster von bis zu 3 Arten
  'von Berichten durch die Routinen:
  
  '(1)Sub KontenStandTabelleErstellen()
  '   Kontentabelle (Monatsendstände der im Kontenplan aufgeführten Bestands-,
  '   Ausgabe- und Einnahmekonten vom Jahresbeginn bis zum Erstellungszeitpunkt),
  '   Erhält den Namen "KonTab-[Monat der letzten ArProt-Buchung]"
  
  '(2)Sub SaLdenListeErstellen
  '   SaLdenliste: Anfangs-und Endzustände der Konten und deren Änderung
  '   in einer vorzugebenden Periode,
  '   Erhält den Namen "SaLdLi[Anfangsmonat]-[Endmonat]"

  '(3)Sub PeriodenBerichtErstellen
  '   Bericht: Ausgaben, Einnahmen und Ergebnis der Konten oder nach Maßgabe
  '   des Kontenplans (in Spalten G und H zusammengefaßte Kontengruppen) in einer
  '   vorzugebenden Periode,
  '   Erhält den Namen "Bericht[Anfangsmonat]-[Endmonat]"
  
  'Die dazu benötigten Vorlagen KoSTaVorl, SaLiVorl und BeriVorl werden - nach den
  'Auftragsdialog-Abläufen - vor dem Füllen mit Daten auf jeden Fall bereitgestellt.
  'Sie werden aus generischen Vorlagen zurechtgestutzt, mit der Projekterkennungs-
  'farbe und dem Buchjahr versehen und solange bei späteren Aufrufen von
  'BerichteErstellen verwendet. wie die ExAcc-Version und die Kontenplan-Version
  'nicht geändert wurde.
  'Die generischen Vorlagen sind im Falle KoSTaVorl der Erfolgskonten-Teil des
  'Kontenplans (ohne etwa vorhandene Personenkonten), in den Fällen SaLiVorl und
  'BeriVorl in der ExAcc-Mappe befindliche Blätter SaLiVorlVorl bzw. BeriVorlVorl.
  'Positionsanker PosAnkK, PsoAnkS, PosAnkB sorgen für geordnete Reihenfolge der
  'Berichtsblätter des Buchungsprojektes.
  'Die Seitenformate einschließlich Fußzeilen und Kopfzeilen mit den im Konten-
  'plan angegebenen Texten richtet die Druckroutine AktBlattDrucken ein. Diese kann
  'auch unabhängig von BerichteErstellen mit Strg+d für alle Blätter in einem
  'ExAcc-Format verwenset werden.
    
  Option Explicit
  Public ExAccVersionNeu As Boolean, AlteExAccVersion As String
  Public Const MaxZeilenSaLdLi = 200 ', MaxSpaltenSaLdLi = 26 '(SVZE,SaLdLiStruktur)
  Public AbbruchBerichteErstellen As Boolean, AbbruchSaLdenListeErstellen As Boolean
  Public AbbruchPeriodenDialog As Boolean, AbbruchSaLiVorlBereitstellen As Boolean
  Public AbbruchPeriodenBerichtErstellen As Boolean
  Public AbbruchBeriVorlErzeugen As Boolean
  Public Erlaeuterung As String, JahrAnfBestand As Double
  Public LinkerH As String, RechterH As String
  Public PSaLdLiFertigText As String, SaLiVorlFertigText As String
  Public PBerichtFertigText As String, BeriVorlFertigText As String
  Public KoSTaVorlVorhanden As Boolean, _
         SaLiVorlVorhanden As Boolean, BeriVorlVorhanden As Boolean, _
         FABeVorlVorhanden As Boolean
  Public PosAnkK As String, PosAnkS As String, PosAnkB As String, PosAnkF As String, _
         PAKVorhanden As Boolean, PASVorhanden As Boolean, PABVorhanden As Boolean, _
         PAFVorhanden As Boolean
  Public LetzteBlattNummer As Integer
  Public BlaNam As String, VorBlaNam As String
  
'SaLdenlisten-Spalten-Struktur: (R = Zeile, C = Spalte)
  Public Const BereichNrC = 1, SaLiKtoC = 2, SaLiÜberC = 3, SaLiBezC = 3
  Public Const SaLBlaNamC = 4, SaLiKArtC = 5, SaLiBereiZeilC = 6
  Public Const SaLiAnfSaLdenC = 9, SaLiÄndSaLdenC = 12, SaLiEndSaLdenC = 15
  Public Const SaLiSumC = 16, SaLiAnfSumC = 17, SaLiSamlBezC = 18
'Berichts-Aus-/Eingabezeilen
  Public BBZZAE As Integer, ZZAE2Groesser As Integer
  
  Dim StartZeile As Long, StartSpalte As Long, StartBlatt As String
  Dim KoPlaZeile As Long, KtoNr As Long
  
  Dim Farbe As String, Gewünscht
  Dim PAnfMonat As Integer, PEndMonat As Integer
  Dim PADatZahl As Long, PEDatZahl As Long
  Dim PADatum As String, PEDatum As String ', PVorDatum As String
  Dim TransaktDatum As String, DefaultMonat As Integer
  
  Public KonTabName As String, PSaLdLiName As String, PBerichtName As String
'  Dim MappenName As String
  Dim AktBlattName As String ', AktZeile As Integer, AktSpalte As Integer
  Dim AnfSaLdo, Endsaldo, VorJahrAnfStand
  Dim AnfStandSoll, EndStandSoll, AnfStandHaben, EndStandHaben
  Dim GesamtJahresanfStand, GesamtPeriodenAnfStand
  
'SaLdLiZeilenstruktur, Ausgegeben von Sub SaLdLiStruktur
  Dim AuEDatumZeileS As Long, ADatumSpalteS As Long, EDatumSpalteS As Long
  Dim UntenZeileS As Long ', RechtsSpalteS As Long
  Dim ZeileErgebnisSaLdLi As Long, PrüfergebnisZeile As Long
'BerichtZeilenstruktur, Ausgegeben von Sub BerichtStruktur
  Dim AuEDatumZeileB As Long, ADatumSpalteB As Long, EDatumSpalteB As Long
  Dim UntenZeileB As Long, RechtsSpalteB As Long
  Dim BerErgebnisZ As Long, BerJahranfSaLdoZ As Long, SaLdLiVersAlt As Long
  Dim A, B, W, BlattNummer As Long
  Const TiT = "Berichte erstellen"

  Const TitSL = "Perioden-SaLdenliste erstellen"
  Const TitPB = "Periodenberichte erstellen"
'Zeilenpositionen der Blöcke im SaLdLiblatt von Sub SaLdBlockPos
  Public SaBloPosBestand As Integer, SaBloPosAusgaben As Integer, _
         SaBloPosEinnahmen As Integer, SaBloPosAusgaben2 As Integer, _
         SaBloPosEinnahmen2 As Integer, SaBloPosFonds As Integer, _
         SaBloPosVermögen As Integer, SaBloPosErgebnis As Integer, _
         SaBloPosErgebnis2 As Integer, SaBloPosKontrolle As Integer, _
         SaBloUntenZeile As Integer
'Zeilenpositionen der Blöcke im Berichtsblatt von Sub BeriBlockPos
  Public BeBloPosBestand As Integer, BeBloPosAusgaben As Integer, _
         BeBloPosAusgaben2 As Integer, BeBloPosFonds As Integer, _
         BeBloPosVermögen As Integer, BeBloPosErgebnis As Integer, _
         BeBloPosUnterschrift As Integer, BeriUntenZeile As Integer
  Public PerAnfSaLdo, PerEndSaLdo, PerDifSaLdo, _
         PerAnfSoll, PerEndSoll, PerDifSoll, _
         PerAnfHaben, PerEndHaben, PerDifHaben
  Public BerichtArt As String, AktPeriodenBlatt As String, _
         NeueKonTab As String, NeueSaLdLi As String, NeuerBericht As String
'------------------- Globale Zeilenstruktur aus Sub Kontenplanstruktur -----------
'Public KPKZBestand As Integer, KPKZAusgaben As Integer, KPKZEinnahmen As Integer, _
'       KPKZAusgaben2 As Integer, KPKZEinnahmen2 As Integer, KPKZFonds As Integer, _
'       KPKZVermögen As Integer, KPKZMitglieder As Integer, KPKZSpender As Integer, _
'       KPKZAE As Integer, KPKZEnde As Integer  '=KPKZ für nächsten (nicht
'       vorhandenen) Bereich
'Public KPVersion As Integer


Sub BerichteErstellen()
'======================= I. Benutzerkommunikation ========================
'1 BH --------------- Meldungssystem initialisieren ---------------------
  AktVorgang = "Berichte erstellen"
  MELDUNG = ""
  ABBRUCH = False
'2 BH ------------ ExAccVersions- u.Mappen-Namen, Ausgangszustand ------------
  With ActiveWindow
    Application.CutCopyMode = False
    ExAccVersion = ThisWorkbook.Name
    MappenName = ActiveWorkbook.Name
    BuchJahr = Sheets("Kontenplan").Cells(1, 5)
    Farbe = Sheets("Kontenplan").Cells(1, 7)
    AlteExAccVersion = Sheets("Kontenplan").Cells(1, 12)
    StartBlatt = ActiveSheet.Name
    StartZeile = ActiveCell.Row
    StartSpalte = ActiveCell.Column
    Cells(StartZeile - 4, StartSpalte).Activate 'V-Kasten
    StartZeile = ActiveCell.Row                 'oberhalb
'3 BH---------- V-Kasten für Vollzugsmeldung im ArProt vorbereiten ----------
    With Worksheets("ArProt")
      .Activate
      If Cells(1, 3) = 4 Then   'ArProt nur mit Dummy-Zeile
        Call AktBlattFärben(Sheets("Kontenplan").Cells(1, 7)) 'Farbkennzeichen
      End If
      Sheets("ArProt").Activate
'--------------------- V-Kasten-Parameter ----------------------------
      Range("I" & StartZeile & ":I" & StartZeile + 3 & "").Select
      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      Selection.Borders(xlInsideVertical).LineStyle = xlNone
      Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'4 BH ------------------- Datum oben im V-Kasten -----------------------
      Range("I" & StartZeile & ":I" & StartZeile + 3 & "").Select
      Cells(StartZeile, StartSpalte).Activate
      With ActiveCell
        .NumberFormat = "d/m/yy"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .ShrinkToFit = True
        .MergeCells = False
        .Value = Date
      End With
      Cells(StartZeile, StartSpalte).Activate
    End With 'Worksheets("ArProt")
'===================== II. Vorlagenbereitstellung =======================
'5 BH ---------------------- Vorlagen-Situation -------------------------
PositionsankerPrüfen:
    KoSTaVorlVorhanden = False
    SaLiVorlVorhanden = False
    BeriVorlVorhanden = False
'    FABeVorlVorhanden = False
    PAKVorhanden = False     'Positionsanker für Konten
    PASVorhanden = False     'Positionsanker für SaLi
    PABVorhanden = False     'Positionsanker für Beri
    PAFVorhanden = False     'Positionsanker für Finanzamtbescheinigungen
    LetzteBlattNummer = 0
'   Dim BlaNamS As String, BlaNamB As String
    With ActiveWorkbook
      For Each W In Worksheets
        BlaNam = W.Name
        LetzteBlattNummer = LetzteBlattNummer + 1  'für "ans Ende stellen
        '---------------------- Vorlagen vorhanden? ----------------
        If W.Name = "KoSTaVorl" Then
          KoSTaVorlVorhanden = True
        End If
        If W.Name = "SaLiVorl" Then
          SaLiVorlVorhanden = True
        End If
        If W.Name = "BeriVorl" Then
          BeriVorlVorhanden = True
        End If
'        If W.Name = "FABeVorl" Then
'          FABeVorlVorhanden = True
'        End If
'6 BH -------------------- Positionsanker vorhanden? --------------------
        If W.Name = "PosAnkK" Then
          PAKVorhanden = True
        End If
        If W.Name = "PosAnkS" Then
          PASVorhanden = True
        End If
        If W.Name = "PosAnkB" Then
          PABVorhanden = True
        End If
'        If W.Name = "PosAnkF" Then
'          PAFVorhanden = True
'        End If
      Next W
'7 BH ---------------- Vollständig vorhandene PosAnks belassen ---------
             'Es wird vorausgesetzt, dass die Anker gut plaziert sind.
      If (PAKVorhanden = True And PASVorhanden = True And _
          PABVorhanden = True) Then  'And PAFVorhanden = True) Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Positionsanker K, S, B vorhanden"
        GoTo VorlagenPrüfen  '10 BH -----------
      End If
'8 BH -------- Nur teilweise vorhandene PosAnk löschen -------------------
'              und SaLi-Vorlagen und Beri-Vorlagen gleich mit
      For Each W In Worksheets
        BlaNam = W.Name
        If BlaNam = "PosAnkK" Or BlaNam = "PosAnkS" Or BlaNam = "PosAnkB" Or _
        Left(BlaNam, 8) = "SaliVorl" Or Left(BlaNam, 8) = "SaLiVorl" Or _
        Left(BlaNam, 8) = "BeriVorl" Then
          Application.DisplayAlerts = False
          Sheets(BlaNam).Delete
          Application.DisplayAlerts = True
          LetzteBlattNummer = LetzteBlattNummer - 1
        End If
      Next W
      PAKVorhanden = False
      PASVorhanden = False
      PABVorhanden = False
      SaLiVorlVorhanden = False
      BeriVorlVorhanden = False
'9 BH ---------------- Drei PosAnk neu ans Ende setzen ----------------------
      Windows(ExAccVersion).Activate
      '---------- PositionsAnker für KontenstandsTabellenvorlage
      Sheets("PosAnkK").Select
      Sheets("PosAnkK").Copy _
             after:=Workbooks(MappenName).Sheets(LetzteBlattNummer)
      LetzteBlattNummer = LetzteBlattNummer + 1
      '---------- PositionsAnker für Saldenlistenvorlage ----------
      Windows(ExAccVersion).Activate
      Sheets("PosAnkS").Select
      Sheets("PosAnkS").Copy _
             after:=Workbooks(MappenName).Sheets(LetzteBlattNummer)
      LetzteBlattNummer = LetzteBlattNummer + 1
      '---------- PositionsAnker für Berichtvorlage --------------
      Windows(ExAccVersion).Activate
      Sheets("PosAnkB").Select
      Sheets("PosAnkB").Copy _
             after:=Workbooks(MappenName).Sheets(LetzteBlattNummer)
      LetzteBlattNummer = LetzteBlattNummer + 1
      '----------- PositionsAnker für Finanzambestätigung ---------
        'wird im Modul FABestätigungen weiter verwendet, hier nicht
'      Windows(ExAccVersion).Activate
'      Sheets("PosAnkF").Select
'      Sheets("PosAnkF").Copy _
'             after:=Workbooks(MappenName).Sheets(LetzteBlattNummer)
'      LetzteBlattNummer = LetzteBlattNummer + 1
      '--------------- Rückkehr zur Projektmappe -----------------
      Windows(MappenName).Activate
      Sheets("ArProt").Activate
      Cells(StartZeile, StartSpalte).Activate  'Eingabemodus verlassen
      MELDUNG = MELDUNG & Chr(10) & _
      "Positionsanker aus " & ExAccVersion & " gesetzt"

VorlagenPrüfen:
'10 BH ----------- Versionsvergleich ggf. löschen alter Vorlagen ------------
      If ExAccVersion <> AlteExAccVersion Then '= Kontenplan Cells(1, 12)
        ExAccVersionNeu = True
        Windows(MappenName).Activate
        Sheets("Kontenplan").Cells(1, 12) = ExAccVersion  'auf aktuellen Stand
      Else
        ExAccVersionNeu = False
      End If
'11 BH --------- KontenstandtabellenVorlage auf gültigen Stand -------------
StandKoSTaVorl:
      Windows(MappenName).Activate
      If KoSTaVorlVorhanden = True Then
        If ExAccVersionNeu = False And _
           Sheets("KoSTaVorl").Cells(1, 1) = Sheets("Kontenplan").Cells(1, 1) Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Vorhandene KoSTaVorl gültig"
          GoTo StandSaLiVorl
        End If
        Sheets("KoSTaVorl").Activate
        Application.DisplayAlerts = False
        Sheets("KoSTaVorl").Delete
        Application.DisplayAlerts = True
      End If
      Call KoSTaVorlBereitstellen
      MELDUNG = MELDUNG & Chr(10) & _
      "KoSTaVorl neu bereitgestellt"
'12 BH ---------------- SaLiVorl auf gültigen Stand ---------------
StandSaLiVorl:
      If SaLiVorlVorhanden = True Then
        If ExAccVersionNeu = False And _
           Sheets("SaLiVorl").Cells(1, 1) = Sheets("Kontenplan").Cells(1, 1) Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Vorhandene SaLiVorl gültig"
          GoTo StandBeriVorl
        End If
        Application.DisplayAlerts = False
        Sheets("SaLiVorl").Delete
        Application.DisplayAlerts = True
      End If
      Call SaLiVorlBereitstellen
      MELDUNG = MELDUNG & Chr(10) & _
      "SaliVorl neu bereitgestellt"
'13 BH ---------------- BeriVorl auf gültigen Stand ---------------
StandBeriVorl:
      If BeriVorlVorhanden = True Then
        If ExAccVersionNeu = False And _
           Sheets("BeriVorl").Cells(1, 1) = Sheets("Kontenplan").Cells(1, 1) Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Vorhandene BeriVorl gültig"
          GoTo VorlagenSindAktuell
        End If
        Application.DisplayAlerts = False
        Sheets("BeriVorl").Delete
        Application.DisplayAlerts = True
      End If
      Call BeriVorlBereitstellen
      MELDUNG = MELDUNG & Chr(10) & _
      "BeriVorl neu bereitgestellt"
'======================== III. Berichte erstellen ==========================
'14 BH --------- Drei Vorlagen bereit vor den Positionsankern --------------
VorlagenSindAktuell:
    NeueKonTab = ""
    NeueSaLdLi = ""
    NeuerBericht = ""
'15 BH---------------- Auftrag Kontenstandstabelle erstellen -----------------
    A = MsgBox(prompt:= _
            "Aktuelle Kontenstandstabelle erstellen?", _
            Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      BerichtArt = "KontenStandTabelle"
      Sheets("ArProt").Activate
      Call KontenStandTabelleErstellen
      Worksheets("ArProt").Activate
    End If
    If A = vbNo Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Etwa vorhandene ''KontenTabelle'' wird belassen."
    End If 'MeldeStufe >= 2
'16 BH--------------- Auftrag Perioden-SaLdenliste erstellen ------------------
    A = MsgBox(prompt:= _
            "Perioden-SaLdenliste erstellen?", Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      BerichtArt = "SaLdLi"
      Sheets("ArProt").Activate
      Call SaLdenListeErstellen
     End If
    If A = vbNo Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Keine Neue ''Saldenliste'' erstellt."
    End If
'17 BH------------------ Auftrag Perioden-Bericht erstellen --------------------
     A = MsgBox(prompt:= _
            "Perioden-Bericht erstellen?", _
            Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      BerichtArt = "Bericht"
      Sheets("ArProt").Activate
      Call PeriodenBerichtErstellen
    End If
    If A = vbNo Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kein neuer Periodenbericht erstellt."
    End If
'18 BH -------------------- Fertigmeldung ----------------------
  A = MsgBox(prompt:= _
            "Berichte sind erstellt." & Chr(10) & _
            "Folgende neu erstellte Blätter können gedruckt werden:" & Chr(10) _
            & "''" & NeueKonTab & "'', ''" & NeueSaLdLi & "'', ''" & _
            NeuerBericht & "''", _
            Buttons:=vbOKOnly, Title:=TiT)
  End With 'ActiveWindow
  If NeueKonTab <> "" Then
    Sheets(NeueKonTab).Activate
    A = MsgBox(prompt:= _
            "Neue Kontab drucken?" & Chr(10) & _
            "(Kann auch später mit Strg+d veranlasst werden)", _
            Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      Sheets(NeueKonTab).Activate
      Call AktBlattDrucken
    End If
  End If
  If NeueSaLdLi <> "" Then
    Sheets(NeueSaLdLi).Activate
    A = MsgBox(prompt:= _
            "Neue SaLdLi drucken?" & Chr(10) & _
            "(Kann auch später mit Strg+d veranlasst werden)", _
            Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      Sheets(NeueSaLdLi).Activate
      Call AktBlattDrucken
    End If
  End If
  If NeuerBericht <> "" Then
    Sheets(NeuerBericht).Activate
    A = MsgBox(prompt:= _
          "Neuen Bericht drucken?" & Chr(10) & _
            "(Kann auch später mit Strg+d veranlasst werden)", _
          Buttons:=vbYesNo, Title:=TiT)
    If A = vbYes Then
      Sheets(NeuerBericht).Activate
      Call AktBlattDrucken
    End If
  End If
End With
End Sub 'BerichteErstellen  (Hauptprogramm)
                                                                

'============================= KontenstandTabelle =============================
Sub KontenStandTabelleErstellen()
  'Prüft ob eine aktuelle KoSTaVorl vorhanden ist. Wenn nicht dann erstellt es
  'sie mit Sub KoSTaVorlErzeugen.
  
 'Als Vorlage zum Füllen der Kontenstandtabelle mit den Kontentänden dient bei
 'jeder Erstellung der Kontenplan

 Dim KoSTaVorhanden As Boolean, KoSTaVorlVorhanden As Boolean
 Dim LetztBuchung As String
 Const TitKS = "Kontenstandstabelle erstellen"
'1 KE------------------ KoSTaVorlage vorhanden? -------------------
    KoSTaVorlVorhanden = False
    For Each W In Worksheets
      If W.Name = "KoSTaVorl" Then
        KoSTaVorlVorhanden = True  'Diese Vorlage ist schon dem Kontenplan
        Exit For                  'angepasst, aber möglicherweise nicht der
      End If                      'aktuellen Version
    Next W
'2 KE ------------------ Wenn KoSTaVorlage veraltet: Löschen ----------------------
    If KoSTaVorlVorhanden = True Then
      With Sheets("KoSTaVorl")
        .Activate                 ' Versionsvergleich: KoSTaVorl aktuell?
        If Sheets("KoSTaVorl").Cells(1, 1) <> Sheets("Kontenplan").Cells(1, 1) Then
          Application.DisplayAlerts = False
          Sheets("KoSTaVorl").Delete
          Application.DisplayAlerts = True
          KoSTaVorlVorhanden = False
          LetzteBlattNummer = LetzteBlattNummer - 1  'für "ans Ende stellen"
        End If
      End With
    End If
'3 KE ---------------------- Aktuelle KoSTa-Vorlage erzeugen ----------------------
    If KoSTaVorlVorhanden = False Then
      Call KoSTaVorlBereitstellen '----> KoSTaVorl einheitliche Form für das Jahr,
    End If                       '      solange Kontenplan nicht geändert wird                    '
'4 KE -------------------  KonTab-Blatt erzeugen --------------------------
    Call KonTabBlattErzeugen
'5 KE ---------------------- KonTab-Blatt füllen ---------------------------------
    Call KonTabBlattFüllen
'6 KE ------------ Fertigmeldung mit umbenannter KonTab ----------------
    Worksheets("ArProt").Activate
    LetztBuchung = Cells(1, 2)
    Cells(StartZeile, StartSpalte).Activate
    ActiveCell.Offset(1, 0).Activate
    ActiveCell = "KonTab-" & MonatTZ(Right(LetztBuchung, 3)) & ""
               'Nur für Vollzugsmeldungskasten
         'Function MonatTZ(MText As String) As Integer
         'Nummer des Monats, der durch "MText" (3 Buchstaben,
         'z.B. "Jan") gegeben ist
    NeueKonTab = "KonTab"   'Für Druckangebot
    Sheets("Kontab").Activate
    A = MsgBox(prompt:="Kontenstand-Tabelle ''KonTab'' erstellt.", _
                 Buttons:=vbOKOnly, Title:="Berichte erstellen")
    GoTo EndeKTE
EndeKTE:
End Sub 'KontenStandTabelleErstellen()

Sub KoSTaVorlBereitstellen() '----------------------------------- 2.3.2017 ----------
  'Erzeugt es eine neue, mit dem aktuellen Kontenplan übereinstimmende
  'KoSTaVorl aus dem Erfolgskontenteil des Kontenplans und stellt sie vor
  'dem Positionsanker PosAnkK bereit.
  'Setzt voraus, dass keine KoSTaVorlage vorhanden ist.
  
  Dim KPZeile As Integer
  Dim KtoTabSpalte As Integer, A, W, K As Integer, DrukBreich As Range
  Dim BereichAnfZeile As Integer, BA As Integer, BE As Integer
  
'1 KTV -------------------- KoSTaVorl-Blatt erzeugen -------------------------
    Call KontenplanStruktur
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "KoSTaVorlErzeugen abgebrochen"
      GoTo EndeKoSTaVorlErzeugen
    End If
    Windows(MappenName).Activate
    Sheets("Kontenplan").Select
    Sheets("Kontenplan").Copy Before:=Sheets("PosAnkK")
    Sheets("Kontenplan (2)").Select
    Sheets("Kontenplan (2)").Name = "KoSTaVorl"
    Range("D1").Select
      ActiveCell.FormulaR1C1 = "KONTENSTANDTABELLE"
'2 KTV ------------------- KoSTaVorl-Blatt zuschneiden ------------------
  With Sheets("KoSTaVorl")
    Columns("F:T").Select
    Selection.Delete shift:=xlToLeft
    'eleganter: hier mit Range löschen statt mit Zeilen
    For K = 1 To KPKZEnde - KPLetzteErfolgsKtoZ
      Rows(KPLetzteErfolgsKtoZ + 1).Select 'aus KontenplanStruktur
      Selection.Delete shift:=xlUp
    Next K
    Cells(1, 3) = KPLetzteErfolgsKtoZ + 1 'EndZeile in Kopfzeile vermerken
'3 KTV ------------------- Spalte E bereinigen ----------------------
'    ActiveWindow.SmallScroll Down:=-33
    Range("E3:E" & KPLetzteErfolgsKtoZ & "").Select
    Selection.ClearContents
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
'4 KTV ------- Rahmen und Druckbereich für die Tabelle (wird expandiert)--------
    Cells(1, 3) = KPLetzteErfolgsKtoZ + 1
    Set DrukBreich = Range("A1:F" & KPLetzteErfolgsKtoZ + 1 & "")
    DrukBreich.Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A$1:$F$" & KPLetzteErfolgsKtoZ + 1 & ""
  End With 'Sheets("KoSTaVorl")
'5 KTV --------------- Monatsspalten und Summenspalte anfügen --------------------
    With Sheets("KoSTaVorl")
      .Activate
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "gleicher KtoArt"
      Call KonTabZelFormat
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Dez"  'eigentlich LetztBuchung
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "30.Nov"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Okt"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "30.Sep"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Aug"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Jul"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "30.Jun"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Mai"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "30.Apr"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Mrz"
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
      Call KonTabZelFormat
      Range("F2").Select
        If SchaltTag = 1 Then
          ActiveCell.FormulaR1C1 = "29.Feb"
        End If
        If SchaltTag = 0 Then
          ActiveCell.FormulaR1C1 = "28.Feb"
        End If
      Columns("F:F").Select
        Selection.Insert shift:=xlToRight
 '----------- Spaltenbreite justieren
       Columns("D:D").ColumnWidth = 32.44
       Columns("E:Q").Select
       Selection.ColumnWidth = 9
        
      Call KonTabZelFormat  '---------------------
      Range("F2").Select
        ActiveCell.FormulaR1C1 = "31.Jan"
      Columns("F:F").Select
      Call KonTabZelFormat
      Range("E2").Select
        ActiveCell.FormulaR1C1 = "1.Januar"
      Range("E4:E" & Cells(1, 3).Value & "").Select
      Call KonTabZelFormat
     Range("E1").Select
       Selection.NumberFormat = "#,##0_ ;[Red]-#,##0 "
       With Selection
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = True
         .ReadingOrder = xlContext
         .MergeCells = False
       End With
     Range("E1").Select
       With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
      End With
      Selection.NumberFormat = "0"
    End With 'Sheets("KontenTabelle")
'6KTV -------------- Bereichkopfrahmen für Kontenstandstabelle -----------------------
    Range("F" & KPKZBestand & ":R" & KPKZBestand & "").Select
    Call BKRahmenFormat   'Bereichkopf-Rahmenformat
    Range("F" & KPKZAusgaben & ":R" & KPKZAusgaben & "").Select
    Call BKRahmenFormat
    Range("F" & KPKZEinnahmen & ":R" & KPKZEinnahmen & "").Select
    Call BKRahmenFormat
    If KPKZAusgaben2 <> 0 Then
      Range("F" & KPKZAusgaben2 & ":R" & KPKZAusgaben2 & "").Select
      Call BKRahmenFormat
    End If
    If KPKZEinnahmen2 <> 0 Then
      Range("F" & KPKZEinnahmen2 & ":R" & KPKZEinnahmen2 & "").Select
      Call BKRahmenFormat
    End If
    If KPKZFonds <> 0 Then
      Range("F" & KPKZFonds & ":R" & KPKZFonds & "").Select
      Call BKRahmenFormat
    End If
    If KPKZVermögen <> 0 Then
      Range("F" & KPKZVermögen & ":R" & KPKZVermögen & "").Select
      Call BKRahmenFormat
    End If
'7 KTV --------- Erläuterung in Zeile 1 der Kontenstandstabelle ---------------------
  Cells(1, 10).Activate
  With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 1
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  With Selection.Font
    .Name = "Arial"
    .Size = 12
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ColorIndex = xlAutomatic
  End With
  Cells(1, 10) = "Leere Felder: Keine Veränderung im Monat; 0 oder Vormonatsstand gilt."
  Cells(1, 18).Activate
  With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  With Selection.Font
    .Name = "Arial"
    .Size = 12
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ColorIndex = xlAutomatic
  End With
  Cells(1, 18) = "Summen"
  Cells(1, 1).Activate
'8 KTV ------------ Summenformeln für Werte gleicher Kontoart ---------------------
  Call KontenplanStruktur
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "KoSTaVorlErzeugen abgebrochen"
    GoTo EndeKoSTaVorlErzeugen
  End If
  With Sheets("KoSTaVorl")
    .Activate
    Cells(KPKZBestand, 18).Activate
    ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZBestand & "]C[-1]:RC[-1]))"
    Cells(KPKZAusgaben, 18).Activate
    ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZAusgaben & "]C[-1]:RC[-1]))"
    Cells(KPKZEinnahmen, 18).Activate
    ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZEinnahmen & "]C[-1]:RC[-1]))"
    If KPKZAusgaben2 <> 0 Then
      Cells(KPKZAusgaben2, 18).Activate
      ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZAusgaben2 & "]C[-1]:RC[-1]))"
    End If
    If KPKZEinnahmen2 <> 0 Then
      Cells(KPKZEinnahmen2, 18).Activate
      ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZEinnahmen2 & "]C[-1]:RC[-1]))"
    End If
'    If KPKZFonds <> 0 Then          'da Fonds ein Teil des Bestands ist, dürfen die
'      Cells(KPKZFonds, 18).Activate 'Beträge nicht in dieser Spalte summiert werden
'      ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZFonds & "]C[-1]:RC[-1]))"
'    End If
    If KPKZVermögen <> 0 Then
      Cells(KPKZVermögen, 18).Activate
      ActiveCell.FormulaR1C1 = "=SUM((R[" & SLZZVermögen & "]C[-1]:RC[-1]))"
    End If
    Sheets("KoSTaVorl").Activate
    Call AktBlattFärben(Farbe)      'im Modul Jahreswechsel
  End With 'Sheets("KoSTaVorl")
EndeKoSTaVorlErzeugen:
End Sub 'KoSTaVorlErzeugen

Sub BKRahmenFormat()   'Hilfsroutine für KoSTaVorlageErzeugen                                        '9.4.2016
'Bereichkopf-Rahmenlinien in KoSTaVorl. (Hilfsroutine in '6 KTV----)
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
End Sub 'BKRahmenFormat

Sub KonTabBlattErzeugen()
  Dim WName As String
'1 KTE ---------------- Etwa vorhandene alte KonTab löschen --------------------
    For Each W In Worksheets
      WName = W.Name
      If W.Name = "KonTab" Then
        Application.DisplayAlerts = False
        Sheets("KonTab").Delete
        Application.DisplayAlerts = True
        MELDUNG = MELDUNG & Chr(10) & _
        "Bereits vorhandenes Blatt ''" & "KonTab" & "'' wurde gelöscht."
        Exit For
      End If
    Next W
'2 KTE---------- aktuelle KoSTaVorlage kopieren, "KonTab" nennen --------------
With Worksheets("KoSTaVorl")
      .Activate
      ActiveSheet.Copy Before:=ActiveWorkbook.Sheets("KoSTaVorl")
      ActiveSheet.Name = "KonTab"
    End With 'Worksheets("SaLiVorl")'
End Sub 'KonTabBlattErzeugen()

Sub KonTabBlattFüllen()
Dim LetztTA As Long, ArProtZeile As Long, LetztBuchung As String, LetztArProtZeile As Long
Dim KPZeile As Long, AktKtoBlattName As String, ErfolgsKontoEndZeile As Long
Dim KontoArt As Integer, KontoZeile As Long, MonatsEndStand As Double
Dim MonatsDatum As String, KontoLetztStand As Boolean, KtoTabSpalte As Integer
'Dim BereichAnfZeile As Long, BA As Long, BE As Long

'1 KT ----------------Buchungsstand ermitteln -------------------------
  With Sheets("ArProt")
'    .Activate
    LetztBuchung = Cells(1, 2)      'Datum letzter Buchung
  End With
'2 KT -------------------- Buchungsstand in Erste Zeile ---------------------
  With Sheets("KonTab")
      .Activate
    Cells(1, 6).Activate
    With Selection
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .HorizontalAlignment = xlLeft
      .AddIndent = False
      .IndentLevel = 1
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
    End With
    LetztBuchung = Sheets("ArProt").Cells(1, 2)
    Cells(1, 6) = "Buchungen bis " & LetztBuchung & " erfasst"

'4 KT --------------- Kontenplan-Schleife -------------------------
         '(enthält Monatsstand-Schleife im Kontenblatt)
    Call KontenplanStruktur
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "KonTabBlattFüllen abgebrochen"
      GoTo EndeKonTabBlattFüllen
    End If
    '------------
    For KPZeile = KPRErsteZeile To KPLetzteErfolgsKtoZ
      With Sheets("Kontenplan")
        If Sheets("Kontenplan").Cells(KPZeile, 2) = "" Then GoTo KontenPlanSchleifeNext
        Call KtoKennDat(Sheets("Kontenplan").Cells(KPZeile, 2))
        If ABBRUCH = True Then
          MELDUNG = MELDUNG & Chr(10) & _
          "KonTabBlattFüllen abgebrochen"
          GoTo EndeKonTabBlattFüllen
        End If
      End With
      Sheets("KonTab").Activate
      If AKtoBlatt = "" Or AKtoEinricht = "" Then
         GoTo KontenPlanSchleifeNext 'überspringen
      End If
'5 KT -------------- Monatsstand-Schleife im Kontenblatt ----------------
      With Sheets(AKtoBlatt)
        .Activate    '<-- später entfernen zur Beschleunigung
        ErfolgsKontoEndZeile = Cells(1, 1).Value + 2
        KontoArt = Cells(1, 3).Value
        KontoLetztStand = False
        '-------------
        For KontoZeile = 3 To Cells(1, 1).Value + 2
          Sheets(AKtoBlatt).Activate
          Cells(KontoZeile, 6).Activate
          If Cells(KontoZeile, 6) = "" Then GoTo KontoBlattschleifeNext
          If IsNumeric(Cells(KontoZeile, 6)) = True And _
             Cells(KontoZeile, 5).Value = "Kontostand" Then
            MonatsEndStand = Cells(KontoZeile, 6)  '.Value
            If AKtoArt = EingabKto Then
              MonatsEndStand = -MonatsEndStand  'Zeichenumkehr bei Zugängen
            End If
            If KontoZeile = 4 Then
'              MonatsDatum = "1.Januar"
              KtoTabSpalte = 5
            Else
'             MonatsDatum = Cells(KontoZeile, 3).Value
             KtoTabSpalte = Cells(KontoZeile - 1, 3) + 5
            End If
            If Cells(KontoZeile - 2, 2) = "***" Then
              KontoLetztStand = True
            End If
            Sheets("KonTab").Cells(KPZeile, KtoTabSpalte) = MonatsEndStand
            If KontoLetztStand = True Then
              Sheets("KonTab").Cells(KPZeile, 17).Value = MonatsEndStand
              Exit For
            End If
          End If 'IsNumeric=True und Kontenstand
KontoBlattschleifeNext:
        Next KontoZeile
      End With 'Sheets(AKtoBlatt)
KontenPlanSchleifeNext:
    Next KPZeile
    Sheets("KonTab").Activate
    Sheets("KonTab").Range("E6").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[" & SLZZBestand & "]C)"
  End With 'Sheets("KonTab")
EndeKonTabBlattFüllen:
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Kontenstandtabellenblatt füllen nichtgelungen"
  End If
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Kontenstandtabellenblatt gefüllt"
  End If
End Sub 'KonTabBlattFüllen


Sub KonTabZelFormat()  'Aufgerufen von KonTabErzeugen         'April 2016
'  Columns("G:G").Select
  Selection.ColumnWidth = 10.9
'  Range("G2:G" & KPLetzteErfolgsKtoZ & "").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
'  Range("G4:G" & KPLetzteErfolgsKtoZ & "").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    End With
End Sub 'KonTabZelFormat()  Aufgerufen von KonTabErzeugen
'=================== Ende Kontentabelle-Aktivitäten ========================


'======================== SaLdenliste Aktivitäten ==========================
Sub SaLdenListeErstellen()
  'Erstellung einer PeriodenSaLdLi in den Schritten:
  'setzt voraus, dass die generische (für alle Kontenplanstrukturen geeignete)
  '  Vorlage "SaLiVorlvorl" vorhanden ist, die (nur bezüglich Erkennungsfarbe
  '  und Buchungsjahr spezialisiert ist und) in der Regel während des gesamten
  '  Buchungsprojekts unverändert erhalten bleibt.
  'Leitet aus "SaLiVorlvorl" die Vorlage "SaLiVorl" (mit Erkennungsfarbe und
  '  Buchungsjahr) ab, die mit der aktuellen Kontenplanstruktur übereinstimmt,
  '  falls diese nicht schon vorhanden und noch aktuell ist (Sub SaLiVorlBereitstellen)
  'ermittelt in einem Benutzer-Dialog den gewünschten Anfangs- und End-Monat der
  '  SaLdLi-Periode und leitet von der "SaLiVorl" die periodenspezifische
  '  Vorlage ab durch Erzeugen eines Blattes mit Namen "SaLdLi" & Periode
  '  (Beisp. "SaLdLi1-12"), das als Vorlage für das Füllen mit den Daten der
  '  Periode gilt (Sub PeriodenDialog).
  'füllt das periodenspezifische Vorlagenblatt mit den Daten für die SaLdLi
  '  (Sub PeriodenSaLdLiFüllen()
  'gibt den SaLdLispezifischen Beitrag zur Vollzugsmeldung an den Ort der
  '  Veranlassung im Arbeitsprotokoll
'1 SE------------------ SaLiVorlage vorhanden? -------------------
    SaLiVorlVorhanden = False
    For Each W In Worksheets
      If W.Name = "SaLiVorl" Then
        SaLiVorlVorhanden = True  'Diese Vorlage ist schon dem Kontenplan
        Exit For                     'angepasst, aber möglicherweise nicht der
      End If                         'aktuellen Version
    Next W
'2 SE ------------------ Wenn SaLiVorlage veraltet: Löschen ----------------------
    If SaLiVorlVorhanden = True Then
      With Sheets("SaLiVorl")
        .Activate                 ' Versionsvergleich: SaLiVorl aktuell?
        If Sheets("SaLiVorl").Cells(1, 1) <> Sheets("Kontenplan").Cells(1, 1) Then
          Application.DisplayAlerts = False
          Sheets("SaLiVorl").Delete
          Application.DisplayAlerts = True
          SaLiVorlVorhanden = False
        End If
      End With
    End If
'3 SE ---------------------- Aktuelle SaLdLi-Vorlage erzeugen ----------------------
    If SaLiVorlVorhanden = False Then
      Call SaLiVorlBereitstellen '------> generische SaLiVorlage ohne Periode,
    End If                           'Blatt "SaLiVorl"
'4 SE -------------------- Periode festlegen --------------------------------------
    If SaLiVorlVorhanden = True Then
      Call PeriodenDialog '---------------> Namen des zu erzeugenden SaLdLiblatts
      If AbbruchPeriodenDialog = True Then 'mit Perioden-Angabe im Namen (für
        GoTo NichtGelungen       '          Berichtart "SaLdLi" oder "Bericht"
      End If
    End If
'5 SE -------------------  PeriodenSaLdLi-Blatt erzeugen --------------------------
    Call PerSaLdLiBlattErzeugen
'6 SE ---------------------- SaLdLi-Blatt füllen ---------------------------------
    Call PeriodenSaLdLiFüllen
'7 SE ---------------------- Fertigmeldung SaLdLi ----------------------------
    Sheets("ArProt").Cells(StartZeile, StartSpalte).Activate
    ActiveCell.Offset(2, 0).Activate
    ActiveCell = PSaLdLiName
    NeueSaLdLi = PSaLdLiName
    Sheets(PSaLdLiName).Activate
    Cells(SaBloPosKontrolle, 2).Activate   'Kontrollrechnung sichtbar
    A = MsgBox(prompt:="SaLdenliste ''" & PSaLdLiName & "'' erstellt.", _
                 Buttons:=vbOKOnly, Title:="Berichte erstellen")
    Sheets("ArProt").Activate
    GoTo EndePSE
NichtGelungen:
    AbbruchSaLdenListeErstellen = True
    PSaLdLiFertigText = _
    "SaLdLi-Erstellung ''" & PSaLdLiName & "'' nicht gelungen." & Chr(10) & _
    Erlaeuterung & ""
    Erlaeuterung = "wegen nicht gelungenem Periodendialog"
EndePSE:
End Sub 'PeriodenSaLdLiErstellen()

Sub SaLiVorlBereitstellen()             'eingefügt von Bericht und abgeändert
'Wird vom Hauptprogramm BerichteErstellen in der Vorlagenbereitstellphase
'aufgerufen, wenn kein Blatt "SaLiVorl" vorhanden ist oder dieses wegen
'Nichtübereinstimmung mit der Kontenplan-version gelöscht wurde.
'Kopiert das Blatt "SaLiVorlVorl",bei Vorhandensein vom Aus/Eingaben2-Bereich
'"Sali2VorlVorl", aus der Mappe ExAcc in die Anwendungsmappe,
'färbt es ein versieht es mit dem Buchjahr, stutzt es um die nicht gebrauchten
'Blöcke, erweitert sie um die SaLdLizeilen gemäß als globale Variable abgelegter
'Information vom Programm "KontoPlanStruktur".
'Ist das geschehen, erhält das Blatt den Namen "SaLiVorl" und die
'Kontenplan-Version, und kann so lange zur Erzeugung von PeriodenSaLdLi-
'Vorlagen verwendet werden, wie die Kontenplan-Version bleibt.
'Eine etwa vorhandene veraltete SaLiVorl wird gelöscht.

  Dim I As Integer, K As Long, ÜZeile As Long, SZeile As Long
  Dim EEiZ As Integer, LEiZ As Integer, EiZ As Long, ZEiZ As Integer
  Dim ZZAusGroesser As Boolean, ZZZwischen As Integer, ZZRest As Integer
  Dim BereichNr As Integer, BereichÜberschrift As String
  Dim A As VbMsgBoxStyle, W, WS, SaLdLiVers As Long, SaLiName As String
  Const TiT = "SaLiVorlBereitstellen"

'1 SVE -------- SaLiVorlVorl von ExAcc in die Anwendermappe kopieren ----------
'  Wenn ein zweiter Ein/Ausgabebereich im Projekt vorgegeben ist, wird aus ExAcc
'  statt der Urvorlage SaLiVorlVorl die Urvorlage Sali2VorlVorl zugrunde gelegt,
'  die sich nur in den Prüfrechnungen voneinander unterscheiden.
  Call KontenplanStruktur
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Kontenplanstruktur defekt. " & AktVorgang & " abgebrochen"
    GoTo EndeSaLiVorlBereitstellen
  End If
'**************************
'--------- SaliVorl aus ExAcc in die AnwenderMappe vor PosAnkS übertragen ------
'                      PosAnkS vorhanden vorausgesetzt
  If KPKZBereich2Vorhanden = False Then
'---------- PosAnkS als Zwischen-Hilfs-Blatt, wird PosAnkS (2) --------------
    Windows(MappenName).Activate
    Sheets("PosAnkS").Select
    Sheets("PosAnkS").Copy Before:=Sheets("PosAnkS")
    Windows(ExAccVersion).Activate
    Sheets("SaLiVorlVorl").Select
    Range("A1:N52").Select
    Selection.Copy
    Windows(MappenName).Activate
    Sheets("PosAnkS (2)").Select
    Range("A1:N52").Select
    ActiveSheet.Paste
    Windows(MappenName).Activate
    Sheets("PosAnkS (2)").Activate
    Sheets("PosAnkS (2)").Name = "SaliVorl"
  End If
  
'    Sheets("PosAnkS").Copy Before:=Sheets("PosAnkS")
'    Sheets("PosAnkS").Select
'    Sheets("PosAnkS").Copy Before:=Sheets("PosAnkS")
'    Sheets("PosAnkS (2)").Select
'    Windows("ExAcc2023-2.xlsm").Activate
'    ActiveWindow.SmallScroll Down:=-24
'    Range("A1:N52").Select
'    Selection.Copy
'    Windows("KiBu2019-6.xls").Activate
'    Sheets("PosAnkS (2)").Select
'    Range("A1:N52").Select
'    ActiveSheet.Paste
'---------------------------------------------------------
'    Sheets("SaLiVorlVorl").Copy Before:=Workbooks(MappenName).Sheets("PosAnkS")
'    Windows(MappenName).Activate
'    Sheets("SaLiVorlVorl").Activate 'hat durch Copy nicht zum workbook MappenName gewechselt!
'    Sheets("SaLiVorlVorl").Name = "SaLiVorl"
'  End If
'-------------------------------------------------
'  If KPKZBereich2Vorhanden = True Then
'    Windows(ExAccVersion).Activate
'    Sheets("SaLi2VorlVorl").Select  'unterscheidet sich von SaLiVorlVorl nur in der Prüfrechnung
'    Sheets("SaLi2VorlVorl").Copy Before:=Workbooks(MappenName).Sheets("PosAnkS"
'    Windows(MappenName).Activate
'    Sheets("SaLi2VorlVorl").Activate 'hat durch Copy nicht zum workbook MappenName gewechselt!
'    Sheets("SaLi2VorlVorl").Name = "SaLiVorl"
'    Sheets("SaLiVorl").Activate
'  End If
  
 If KPKZBereich2Vorhanden = True Then
'---------- PosAnkS als Zwischen-Hilfs-Blatt, wird PosAnkS (2) --------------
    Windows(MappenName).Activate
    Sheets("PosAnkS").Select
    Sheets("PosAnkS").Copy Before:=Sheets("PosAnkS")
    Windows(ExAccVersion).Activate
    Sheets("SaLi2VorlVorl").Select
    Range("A1:N52").Select
    Selection.Copy
    Windows(MappenName).Activate
    Sheets("PosAnkS (2)").Select
    Range("A1:N52").Select
    ActiveSheet.Paste
    Windows(MappenName).Activate
    Sheets("PosAnkS (2)").Activate
    Sheets("PosAnkS (2)").Name = "SaliVorl"
  End If
    
  Call AktBlattFärben(Farbe)
  Windows(MappenName).Activate
  Sheets("SaLiVorl").Cells(2, 11) = BuchJahr
  Sheets("SaLiVorl").Activate 'hat durch Copy nicht zum workbook MappenName gewechselt!
 '2 SVE ----------------- Nicht gebrauchte Blöcke entfernen ----------------------------
    Call KontenplanStruktur    'gibt die Kontenplankopfzeilen KPKZ^^^^-Werte
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenplanstruktur defekt. " & AktVorgang & " abgebrochen"
      GoTo EndeSaLiVorlBereitstellen
    End If
    Call SaLiBlockPos("SaLiVorl") 'Pos in der VorVorlage
    With Sheets("SaLiVorl")
      .Activate
      'KPKZBestand, KPKZAusgaben und KPKZEinnahmen sind nicht optional, _
      'werden daher nicht abgefragt
      If KPKZAusgaben2 = 0 Then
        Call SaLiBlockPos("SaLiVorl")
        Range("A" & SaBloPosAusgaben2 & ":M" & SaBloPosAusgaben2 + 4 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If KPKZEinnahmen2 = 0 Then
        Call SaLiBlockPos("SaLiVorl")
        Range("A" & SaBloPosEinnahmen2 & ":M" & SaBloPosEinnahmen2 + 4 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If KPKZAusgaben2 = 0 And KPKZEinnahmen2 = 0 Then
        Call SaLiBlockPos("SaLiVorl")
        Range("A" & SaBloPosErgebnis2 & ":M" & SaBloPosErgebnis2 + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If KPKZFonds = 0 Then
        Call SaLiBlockPos("SaLiVorl")
        Range("A" & SaBloPosFonds - 1 & ":M" & SaBloPosFonds + 3 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If KPKZVermögen = 0 Then
        Call SaLiBlockPos("SaLiVorl")
        Range("A" & SaBloPosVermögen & ":M" & SaBloPosVermögen + 4 & "").Select
        Selection.Delete shift:=xlUp
      End If
    End With 'Sheets("SaLiVorl")
'3 SVE -*-*-*-*-*-*- Verbliebene Blöcke dimensionieren: Bestandsblock -*-*-*-*-*-*-
    Call SaLiBlockPos("SaLiVorl") 'Pos der ggf. gestutzten Vorlage
    '------------------ Abbruch bei fehlendem Bestandsbereich --------------------
    If SaBloPosBestand = 0 Then
      Sheets("ArProt").Activate
      Call MsgBox(prompt:= _
           "Im Kontenplan fehlt der Bereich ''Bestand'' (Kontoart 1)" & _
           AktVorgang & " abgebrochen", _
           Buttons:=vbOKOnly, Title:="SaLiVorlBereitstellen")
      AbbruchSaLiVorlBereitstellen = True
      Exit Sub
    End If 'SaBloPosBestand = 0
    '------------------- Bestandsblock: Zeile entfernen ----------------------
    Sheets("SaLiVorl").Activate
    If SLZZBestand = 1 Then
      Range("A" & SaBloPosBestand + 2 & ":M" & SaBloPosBestand + 2 & "").Select
      Selection.Delete shift:=xlUp
    End If
    '------------ Bestandsblock: Zeilen hinzufügen ------------------------
    If SLZZBestand > 2 Then
      Range("A" & SaBloPosBestand + 2 & ":M" & _
      SaBloPosBestand + SLZZBestand - 1 & "").Select
        Selection.Insert shift:=xlDown
    End If 'SLZZBestand > 2
'4 SVE -*-*-*-*-*-*-*-*- Block Ausgaben dimensionieren -*-*-*-*-*-*-*-*-*-
    Call SaLiBlockPos("SaLiVorl") 'Pos der gestutzten u. erweiterten Vorlage
    '------------ Abbruch bei fehlendem Ausgabeblock ------------------------
    If SaBloPosAusgaben = 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
        "Im Kontenplan fehlt der Bereich ''Ausgaben'' (Kontoart 2)" & _
        AktVorgang & " abgebrochen"
      ABBRUCH = True
      GoTo EndeSaLiVorlBereitstellen
    End If 'SaBloPosBestand = 0
    '--------------- Block Ausgaben: Zeile entfernen -------------------------
    If SaBloPosAusgaben <> 0 Then
      If SLZZAusgaben = 1 Then
        Range("A" & SaBloPosAusgaben + 2 & ":M" & SaBloPosAusgaben + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
    '--------------- Block Ausgaben Zeilen hinzufügen -----------------------
      If SLZZAusgaben > 2 Then
        Range("A" & SaBloPosAusgaben + 2 & ":M" & _
        SaBloPosAusgaben + SLZZAusgaben - 1 & "").Select
          Selection.Insert shift:=xlDown
'-------------------- Linien im Block vervollständigen --------------------
        Range("A" & SaBloPosAusgaben + 2 & ":M" & _
            SaBloPosAusgaben + SLZZAusgaben - 1 & "").Select
          Selection.Borders(xlDiagonalDown).LineStyle = xlNone
          Selection.Borders(xlDiagonalUp).LineStyle = xlNone
          With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
          End With
          With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
          End With
          With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
          End With
          With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
          End With
          With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
          End With
        End If
    End If 'SaBloPosAusgaben <> 0
 '5 SVE -*-*-*-*-*-*-*-*- Block Einnahmen dimensionieren -*-*-*-*-*-*-*-*--
    Call SaLiBlockPos("SaLiVorl") 'Pos der gestutzten u. erweiterten Vorlage
    '------------ Abbruch bei fehlendem Einnahmenblock ------------------------
    If SaBloPosEinnahmen = 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
        "Im Kontenplan fehlt der Bereich ''Einnahmen'' (Kontoart 3)" & _
        AktVorgang & " abgebrochen"
      ABBRUCH = True
      GoTo EndeSaLiVorlBereitstellen
    End If 'SaBloPosBestand = 0
    '--------------- Block Einnahmen Zeile entfernen -------------------------
    If SaBloPosEinnahmen <> 0 Then
      If SLZZEinnahmen = 1 Then
        Range("A" & SaBloPosEinnahmen + 2 & ":M" & SaBloPosEinnahmen + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
    '--------------- Block Einnahmen Zeilen hinzufügen -------------------------
      If SLZZEinnahmen > 2 Then
        Range("A" & SaBloPosEinnahmen + 2 & ":M" & _
        SaBloPosEinnahmen + SLZZEinnahmen - 1 & "").Select
          Selection.Insert shift:=xlDown
      End If
      Range("A" & SaBloPosEinnahmen + 2 & ":M" & _
      SaBloPosEinnahmen + SLZZEinnahmen - 1 & "").Select
      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
    End If 'SaBloPosEinnahmen <> 0
'6 SVE -*-*-*-*-*-Verbliebenen Ausgaben2-Block dimensionieren: -*-*-*-*-*-*-
    Call SaLiBlockPos("SaLiVorl") 'Pos der gestutzten u. erweiterten Vorlage
    If SaBloPosAusgaben2 <> 0 Then
      If SLZZAusgaben2 = 1 Then
        Range("A" & SaBloPosAusgaben2 + 2 & ":M" & SaBloPosAusgaben2 + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If SLZZAusgaben2 > 2 Then
        Range("A" & SaBloPosAusgaben2 + 2 & ":M" & _
        SaBloPosAusgaben2 + SLZZAusgaben2 - 1 & "").Select
          Selection.Insert shift:=xlDown
      End If
    End If 'SaBloPosAusgaben2 <> 0
'7 SVE -*-*-*-*-*-Verbliebenen Einnahmen2-Block dimensionieren: -*-*-*-*-*-*-
    Call SaLiBlockPos("SaLiVorl")
    If SaBloPosEinnahmen2 <> 0 Then
      If SLZZEinnahmen2 = 1 Then
        Range("A" & SaBloPosEinnahmen2 + 2 & ":M" & SaBloPosEinnahmen2 + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If SLZZEinnahmen2 > 2 Then
        Range("A" & SaBloPosEinnahmen2 + 2 & ":M" & _
        SaBloPosEinnahmen2 + SLZZEinnahmen2 - 1 & "").Select
          Selection.Insert shift:=xlDown
      End If
    End If 'SaBloPosEinnahmen2 <> 0
'8 SVE -*-*-*-*-*-*- Etwa verblebenen Fondsblock dimensionieren -*-*-*-*-*-*-*-*----
      
    Call SaLiBlockPos("SaLiVorl") 'Pos der ggf. gestutzten Vorlage
    If SaBloPosFonds <> 0 Then
      If SLZZFonds = 1 Then
        Range("A" & SaBloPosFonds + 2 & ":M" & SaBloPosFonds + 2 & "").Select
        Selection.Delete shift:=xlUp
      End If
      '------------ Fondsblock: Zeilen hinzufügen ------------------------
      If SLZZFonds > 2 Then
        Range("A" & SaBloPosFonds + 2 & ":M" & _
        SaBloPosFonds + SLZZFonds - 1 & "").Select
          Selection.Insert shift:=xlDown
      End If 'SLZZFonds > 2
    End If 'SaBloPosFonds <> 0
'9 SVE ------------- Verbliebene Blöcke dimensionieren: Vermögensblock ------------
    Call SaLiBlockPos("SaLiVorl") 'Pos der ggf. um Blöcke gestutzten Vorlage
    If SaBloPosVermögen <> 0 Then
      If SLZZVermögen = 1 Then
        Range("A" & SaBloPosVermögen + 2 & ":M" & SaBloPosVermögen + 2 & "").Select
        Selection.Delete shift:=xlUp  'Summenzeile auch gelöscht, um Bezugsfehler
       End If                         'zu vermeiden: hier nicht, aber bei BeriVorlV
      If SLZZVermögen > 2 Then
        Range("A" & SaBloPosVermögen + 2 & ":M" & _
        SaBloPosBestand + SLZZVermögen - 2 & "").Select
        Selection.Insert shift:=xlDown
      End If
    End If 'SaBloPosVermögen <> 0
        
'10 SVE ------------------ Druckbereich, Seitenumbruch --------------------
    Call SaLiBlockPos("SaLiVorl")
    Range("$A$1:$M$" & SaBloUntenZeile & "").Activate
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A$1:$M$" & SaBloUntenZeile - 1 & ""
    ActiveSheet.Cells(3, 1) = SaBloUntenZeile   'für Drucken
     If SaBloPosAusgaben + SLZZAusgaben > 35 Then
'      Set ActiveSheet.HPageBreaks(1).Location = _
         Range("A" & SaBloPosEinnahmen & "")
    End If
'11 SVE ------------------ Versionseintrag, Fertignamen ---------------------------
    Sheets("SaLiVorl").Cells(1, 1) = KPVersion 'von KontenplanStruktur erzeugt
    Sheets("SaLiVorl").Name = "SaLiVorl"
    SaLiVorlVorhanden = True
'12 SVE ------------------ Vollzugsmeldung (bedingt) ---------------------------
EndeSaLiVorlBereitstellen:
    Sheets("SaLiVorl").Activate
    If ABBRUCH = False Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Blatt ''SaLiVorl'' nach Kontenplanversion " & KPVersion & " erstellt "
    End If
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Erstellung von Blatt ''SaLiVorl'' nach Kontenplanversion " & KPVersion & _
      " nicht gelungen"
    End If
End Sub 'SaLiVorlBereitstellen

Sub SaLiBlockPos(Blat)
'Beschreibt Vorhandensein und Position (Zeile) der Blöcke von Konten gleicher Art
'in SaLdLi-Formularen mittels der globalen Variablen SaBloPos^^^^^^.
  Dim Z As Integer, BlattName As String
'-------------- Alte Pos löschen -------------------------
  BlattName = Blat
  With Sheets(BlattName)
    .Activate
    SaBloPosBestand = 0
    SaBloPosAusgaben = 0
    SaBloPosEinnahmen = 0
    SaBloPosAusgaben2 = 0
    SaBloPosEinnahmen2 = 0
    SaBloPosFonds = 0
    SaBloPosVermögen = 0
    SaBloPosErgebnis = 0
    SaBloPosErgebnis2 = 0
    SaBloPosKontrolle = 0
'--------------------- Berichtsformular-Untenzeile -------------------
    For Z = 1 To 200
      If Cells(Z, 1) = "***" Then
        SaBloUntenZeile = Z
        Exit For
      End If
    Next Z
'------------------- Pos in die globalen Variablen -----------------
    For Z = 6 To SaBloUntenZeile
      If Cells(Z, 1) = 1 Then
        SaBloPosBestand = Z
      End If
      If Cells(Z, 1) = 2 Then
        SaBloPosAusgaben = Z
      End If
      If Cells(Z, 1) = 3 Then
        SaBloPosEinnahmen = Z
      End If
      If Cells(Z, 1) = 4 Then
        SaBloPosAusgaben2 = Z
      End If
      If Cells(Z, 1) = 5 Then
        SaBloPosEinnahmen2 = Z
      End If
      If Cells(Z, 1) = 6 Then
        SaBloPosFonds = Z
      End If
      If Cells(Z, 1) = 7 Then
        SaBloPosVermögen = Z
      End If
      If Cells(Z, 1) = 12 Then
        SaBloPosErgebnis = Z
      End If
      If Cells(Z, 1) = 13 Then
        SaBloPosErgebnis2 = Z
      End If
      If Cells(Z, 1) = 14 Then
        SaBloPosKontrolle = Z
      End If
    Next Z
  End With '
End Sub 'SaLiBlockPos

Sub PerSaLdLiBlattErzeugen() 'AnfMonat As Integer, EndMonat As Integer)
'Erzeugt aus den vom Hauptprogramm vorgegebenen Namen "PSaLdLiName" bzw. "PBerichtName",
'in denen die Periode erscheint, ein Blatt für den Eintrag der Werte der angegebenen
'Periode durch Kopieren der Vorlagen, Plazieren an geeignete Stelle der Blattleiste der
'Mappe und Eintragen der Periodenüberschrift.
'Gegebenenfalls wird die globale Variable 'Abbruch' auf TRUE gesetzt.
'Sub Setzt voraus, dass Formularvorlagen mit den Namen "SaLiVorl" und "BeriVorl"
'existieren, die in den ersten drei Spalten der ersten unbenutzten Zeile die Spalte und
'die Zeile des Anfangs- und Enddatums enthalten.
  Dim PSaLdLiBlattText As String, PerBerVorh As Boolean
  Const TiT = "PeriodBlattErzeugen"

'1 PBE ------------------------ Perioden-SaLdLi schon vorhanden? -----------------------
      PerBerVorh = False
      For Each W In Worksheets
        If W.Name = PSaLdLiName Then
          PerBerVorh = True
          Exit For
        End If
      Next W
'2 PBE --------- etwa vorhandenen PSaLdLi löschen.  ---------------------------------
      If PerBerVorh = True Then
        Application.DisplayAlerts = False
        Sheets(PSaLdLiName).Delete
        Application.DisplayAlerts = True
        If MeldeStufe >= 2 Then
          A = MsgBox(prompt:= _
            "Bereits vorhandenes Blatt ''" & PSaLdLiName & "'' wurde gelöscht.", _
            Buttons:=vbOKOnly, Title:=TiT)
        End If
      End If 'PerBerVorh = True
'3 PBE----------------- SaLdLi-Vorlage kopieren, "SaLdLin-n" nennen -------------------
PeriodenSaLdLiBlatt:
    With Worksheets("SaLiVorl")
      .Activate
      With ActiveWorkbook
        Sheets("SaLiVorl").Activate
        ActiveSheet.Copy Before:=ActiveWorkbook.Sheets("SaLiVorl")
        ActiveSheet.Name = PSaLdLiName   'ADatumZelle, EDatumZelle, RechtsUntenZelle
      End With 'ActiveWorkbook
    End With 'Worksheets("SaLiVorl")
'4 PBE ----------- Anfangs- und End-Datum in SaLdLi-Kopf eintragen -------------------
    With Sheets(PSaLdLiName)
      .Activate
      Cells(2, 10).Value = LangDatum(PADatum)
      Cells(2, 12).Value = LangDatum(PEDatum)
      Cells(2, 13).Value = BuchJahr
    End With 'Sheets(PSaLdLiName)
    MELDUNG = MELDUNG & Chr(10) & _
    "Das Blatt ''" & PSaLdLiName & "'' kann gefüllt werden."
 End Sub 'PerSaLdLiBlattErzeugen

Sub PeriodenSaLdLiFüllen()
  Dim KPZeile As Integer, KPBereichA As Integer, KPBereichE As Integer
  Dim BereichAnf As Integer, BereichEnd As Integer, BeschrText As String
  
  If ABBRUCH = True Then GoTo EndePeriodenSaLdLiFüllen
'1 PSF ----------------- Übertragungsdaten für alle Bereiche -------------------
  AktPeriodenBlatt = "SaLdLi" & PAnfMonat & "-" & PEndMonat & ""
BereichsAnfuEnde:
  Call KontenplanStruktur
  If ABBRUCH = True Then GoTo EndePeriodenSaLdLiFüllen
  Call SaLiBlockPos(AktPeriodenBlatt)
  Sheets("Kontenplan").Activate
'2 PSF ----------------- Bereich "Bestand" übertragen ------------------------
  If SaBloPosBestand <> 0 Then
    '-------------- Position im leeren SaLdLiBlatt
    BereichAnf = SaBloPosBestand + 1           'Zeile des 1.Kontos
    BereichEnd = SaBloPosBestand + SLZZBestand 'Zeile des letzten Kontos
    '-------------- Position im Kontenplan
    KPBereichA = KPKZBestand + 1               'Zeile des 1.Kontos
    KPBereichE = KPKZBestand + SLZZBestand     'Zeile des letzten Kontos
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If
'3 PSF ----------------- Bereich "Ausgaben" übertragen ------------------------
  If SaBloPosAusgaben <> 0 Then
    BereichAnf = SaBloPosAusgaben + 1
    BereichEnd = SaBloPosAusgaben + SLZZAusgaben
    KPBereichA = KPKZAusgaben + 1
    KPBereichE = KPKZAusgaben + SLZZAusgaben
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosAusgaben <> 0
'4 PSF ----------------- Bereich "Einnahmen" übertragen ------------------------
  If SaBloPosEinnahmen <> 0 Then
    BereichAnf = SaBloPosEinnahmen + 1
    BereichEnd = SaBloPosEinnahmen + SLZZEinnahmen
    KPBereichA = KPKZEinnahmen + 1
    KPBereichE = KPKZEinnahmen + SLZZEinnahmen
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosEinnahmen <> 0
  '5 PSF ----------------- Bereich "Ausgaben2" übertragen ----------------
  If SaBloPosAusgaben2 <> 0 Then
    BereichAnf = SaBloPosAusgaben2 + 1
    BereichEnd = SaBloPosAusgaben2 + SLZZAusgaben2
    KPBereichA = KPKZAusgaben2 + 1
    KPBereichE = KPKZAusgaben2 + SLZZAusgaben2
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosAusgaben2 <> 0
'6 PSF ----------------- Bereich "Einnahmen2" übertragen -----------------
  If SaBloPosEinnahmen2 <> 0 Then
    BereichAnf = SaBloPosEinnahmen2 + 1
    BereichEnd = SaBloPosEinnahmen2 + SLZZEinnahmen2
    KPBereichA = KPKZEinnahmen2 + 1
    KPBereichE = KPKZEinnahmen2 + SLZZEinnahmen2
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosEinnahmen2 <> 0
'7 PSF --------------- Bereich "Fonds" übertragen -------------------
  If SaBloPosFonds <> 0 Then
    BereichAnf = SaBloPosFonds + 1
    BereichEnd = SaBloPosFonds + SLZZFonds
    KPBereichA = KPKZFonds + 1
    KPBereichE = KPKZFonds + SLZZFonds
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosFonds <> 0
 '8 PSF -------------- Bereich "Vermögen" übertragen ------------------
  If SaBloPosVermögen <> 0 Then
    BereichAnf = SaBloPosVermögen + 1
    BereichEnd = SaBloPosVermögen + SLZZVermögen
    KPBereichA = KPKZVermögen + 1
    KPBereichE = KPKZVermögen + SLZZVermögen
    Call SaLdLiBereichFüllen(BereichAnf, BereichEnd, KPBereichA, KPBereichE)
  End If 'SaBloPosVermögen <> 0
'9 PSF -----------Jahresanfangstand in den Ergebnisblock ---------------
    Call BestandSummeJahrAnf
    Call SaLiBlockPos(PSaLdLiName)
    Sheets(PSaLdLiName).Activate
    Cells(SaBloPosErgebnis + 1, 7).Activate
    ActiveCell = JahrAnfBestand 'von Sub BestandSummeJahrAnf
'10 PSF --------------------- Kontrollrechnung --------------------------

'10 PSF ------------------ Druckbereich festlegen ------------------
    Call SaLiBlockPos(PSaLdLiName)
    Range("$A$1:$O$" & SaBloUntenZeile & "").Activate
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A$1:$O$" & SaBloUntenZeile - 1 & ""
'10 PSF ----------------- Vollzugsvermerk in ArProt --------------------------
  With Sheets("ArProt")
    .Activate
    Cells(StartZeile + 2, StartSpalte) = PSaLdLiName
  End With
EndePeriodenSaLdLiFüllen:
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "PeriodenSaLdLi füllen nicht gelungen"
  End If
  If ABBRUCH = False Then
    MELDUNG = MELDUNG & Chr(10) & _
    "PeriodenSaLdLi gefüllt"
  End If
End Sub 'PeriodenSaLdLiFüllen()

Sub SaLdLiBereichFüllen(BreichAnf As Integer, BreichEnd As Integer, _
                        KPBreichA As Integer, KPBreichE As Integer)
  Dim KPZeile As Integer, SaFoZ As Integer, AKtoZiBerst As Integer
    AKtoBeriZeileAlt = 0
    AKtoZiBerst = 1   'Zeile im Bereich
    For KPZeile = KPBreichA To KPBreichE
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus > KtoGanzLeer Then  'also Blatt vorhanden
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          SaFoZ = BreichAnf + AKtoZiB - 1 'AKtoZiB = Zeile im KP-Bereich
          Cells(SaFoZ, 2) = AKtoNr
          Cells(SaFoZ, 3) = AKtoBeschr
          Cells(SaFoZ, 4) = AKtoBlatt
          Cells(SaFoZ, 5) = PerAnfSoll
          Cells(SaFoZ, 6) = PerAnfHaben
          Cells(SaFoZ, 7) = PerAnfSaLdo
          Cells(SaFoZ, 10) = PerDifSaLdo
          Cells(SaFoZ, 11) = PerEndSoll
          Cells(SaFoZ, 12) = PerEndHaben
          Cells(SaFoZ, 13) = PerEndSaLdo
'3 PSF ----------------- Berichtszeilenverteilung ----------------------
          If AKtoBeriZeile = AKtoBeriZeileAlt Then
            Cells(BreichAnf - 1 + AKtoZiBerst, 14) = _
            Cells(BreichAnf - 1 + AKtoZiBerst, 14) + PerEndSaLdo
          End If
          '--------- Berichtszeilentrennung -----------
          If AKtoBeriZeile <> AKtoBeriZeileAlt Then
            AKtoZiBerst = AKtoZiB
            Cells(SaFoZ, 14) = PerEndSaLdo
            Range("B" & SaFoZ & ":N" & SaFoZ & "").Select
            With Selection.Borders(xlEdgeTop)
              .LineStyle = xlContinuous
              .ColorIndex = xlAutomatic
              .TintAndShade = 0
              .Weight = xlThin
            End With
            Range("B" & SaFoZ + 1 & "").Select
          End If
          AKtoBeriZeileAlt = AKtoBeriZeile
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
 End Sub 'SaLdLiBereichFüllen
 '========================= Ende SaLdenlistenaktivitäten =============================

'================================= Periodenbericht ===================================
Sub PeriodenBerichtErstellen()           'Original!   SaLdLi entsprechend. Kopiert
  'Erstellung eines Periodenberichtes in den Schritten:
  'setzt voraus, dass die generische (für alle Kontenplanstrukturen geeignete)
  '  BeriVorlage "BeriVorlVo" vorhanden ist, die nur bezüglich Erkennungsfarbe
  '  und Buchungsjahr spezialisiert ist und während des gesamten Buchungsprojekts
  '  unverändert erhalten bleibt
  'leitet aus "BeriVorlVo" die Vorlage "BeriVorl" ab, die mit der aktuellen
  '  Kontenplanstruktur übereinstimmt, falls diese nicht schon vorhanden und noch
  '  aktuell ist (Sub BeriVorlErzeugen)
  'ermittelt in einem Benutzer-Dialog den gewünschten Anfangs- und End-Monat der
  '  Berichts-Periode und leitet von der "BeriVorl" die periodenspezifische
  '  Vorlage ab durch Erzeugen eines Blattes mit Namen "Bericht" & Periode
  '  (Beisp. "Bericht1-12"), das als Vorlage für das Füllen mit den Daten der
  '  Periode gilt (Sub PeriodenDialog)
  'füllt das periodenspezifische Vorlagenblatt mit den Daten für den Bericht
  '  Sub PeriodenBerichtFüllen()
  'gibt den Berichtspezifischen Beitrag zur Vollzugsmeldung an den Ort der
  '  Veranlassung im Arbeitsprotokoll
'1 BE------------------ BeriVorlage vorhanden? -------------------
    BeriVorlVorhanden = False
    For Each W In Worksheets
      If W.Name = "BeriVorl" Then
        BeriVorlVorhanden = True  'Diese Vorlage ist schon dem Kontenplan
        Exit For                     'angepasst, aber möglicherweise nicht der
      End If                         'aktuellen Version
    Next W
'2 BE ------------------ Wenn BeriVorlage veraltet: Löschen ----------------------
    If BeriVorlVorhanden = True Then
      With Sheets("BeriVorl")
        .Activate                 ' Versionsvergleich: BeriVorl aktuell?
        If Sheets("BeriVorl").Cells(1, 1) <> Sheets("Kontenplan").Cells(1, 1) Then
          Application.DisplayAlerts = False
          Sheets("BeriVorl").Delete
          Application.DisplayAlerts = True
          BeriVorlVorhanden = False
        End If
      End With
    End If
'3 BE ------------------ Aktuelle BeriVorlage erzeugen -------------------------
    If BeriVorlVorhanden = False Then
      Call BeriVorlBereitstellen '------> generische BeriVorlage ohne Periode,
    End If                     '        Blatt "BeriVorl"
'4 BE ------------------------- Periode Festlegen ---------------------------------
    If BeriVorlVorhanden = True Then
      Call PeriodenDialog '-----------> Namen des zu erzeugenden Berichtsblatts
      If AbbruchPeriodenDialog = True Then '          mit Perioden-Angabe im Namen
        GoTo NichtGelungen   '          (für Berichtart "SaLdLi" oder "Bericht"
      End If
    End If
'5 BE ----------------------- Berichtblatt erzeugen ------------------------------
    Call PerBerichtBlattErzeugen
'6 BE ----------------------- Berichtblatt füllen -------------------------------
    Call PeriodenBerichtFüllen
'7 SE ---------------------- Fertigmeldung Bericht ----------------------------
    Worksheets("ArProt").Activate
    Cells(StartZeile, StartSpalte).Activate
    ActiveCell.Offset(3, 0).Activate
    ActiveCell = PBerichtName
    NeuerBericht = PBerichtName
    Sheets(PBerichtName).Activate
    A = MsgBox(prompt:="Periodenbericht ''" & PBerichtName & "'' erstellt.", _
                 Buttons:=vbOKOnly, Title:="Berichte erstellen")
    Sheets("ArProt").Activate
    GoTo EndePBE
NichtGelungen:
    AbbruchPeriodenBerichtErstellen = True
    PSaLdLiFertigText = _
    "SaLdLi-Erstellung ''" & PSaLdLiName & "'' nicht gelungen." & Chr(10) & _
    Erlaeuterung & ""
    Erlaeuterung = "wegen nicht gelungenem Periodendialog"
EndePBE:
End Sub 'PeriodenBerichtErstellen()
                                                              'Februar 2019
Sub BeriVorlBereitstellen()         'Original!
'Wird vom Hauptprogramm BerichtErstellen in der Vorlagenbereitstellphase
'aufgerufen, wenn kein Blatt "BeriVorl" vorhanden ist oder dieses wegen
'Nichtübereinstimmung mit der Kontenplanversion gelöscht wurde.
'Kopiert das Blatt "BeriVorlVorl" aus der Mappe ExAcc in die Anwendungsmappe
'vor den Positionsanker "PosAnkB", färbt es ein, versieht es mit Buchjahr
'und stutzt es gemäß Kontenplaninformation zurecht, versieht es mit der
'Kontenplanversionsnummr (in Zelle A1) und benennt es in "BeriVorl" um.

  Dim I As Integer, K As Long, ÜZeile As Long, SZeile As Long
  Dim EEiZ As Integer, LEiZ As Integer, EiZ As Long, ZEiZ As Integer
  Dim ZZAusGroesser As Boolean, ZZZwischen As Integer, ZZRest As Integer
  Dim BereichNr As Integer, BereichÜberschrift As String, AltBerichtName As String
  Dim A As VbMsgBoxStyle, W, WS, BerichtVers As Long
  Dim AltBeriVorlAufbewahrt As Boolean
  Dim UeberCriftLinks As String, UeberCriftRechts As String
  Const TiT = "BeriVorlErzeugen"
  
  With ActiveWindow
 '1 BVE -------- BeriVorlVorl aus ExAcc in Anwendermappe kopieren -------------
    Windows(ExAccVersion).Activate
    Sheets("BeriVorlVorl").Select
 'Beri***********************
 '--------- BeriVorl aus ExAcc in die AnwenderMappe vor PosAnkB übertragen ------
'                      PosAnkB vorhanden vorausgesetzt
'  If KPKZBereich2Vorhanden = False Then
'---------- PosAnkS als Zwischen-Hilfs-Blatt, wird PosAnkS (2) --------------
    Windows(MappenName).Activate
    Sheets("PosAnkB").Select
    Sheets("PosAnkB").Copy Before:=Sheets("PosAnkB")
    Windows(ExAccVersion).Activate
    Sheets("BeriVorlVorl").Select
    Range("A1:K42").Select
    Selection.Copy
    Windows(MappenName).Activate
    Sheets("PosAnkB (2)").Select
    Range("A1:K42").Select
    ActiveSheet.Paste
    Windows(MappenName).Activate
    Sheets("PosAnkB (2)").Activate
    Sheets("PosAnkB (2)").Name = "BeriVorl"
' End If
'--------------------------------
''   ActiveWorkbook.Save
''    ActiveWorkbook.Save
'    Application.Left = 37
'    Application.Top = 64
'    Windows(MappenName).Activate
'    '-----------------------
'    Workbooks(MappenName).Sheets("BeriVorlVorl").Select
'    Sheets("BeriVorlVorl").Name = "BeriVorl"
'------------------------ BeriVorl an aktuellen Fall anpassen ------------------------
    With Windows(MappenName)
      Sheets("BeriVorl").Activate 'hat durch Copy nicht zum workbook MappenName gewechselt!
      Call AktBlattFärben(Farbe)      'Färben
'------------------------------Blatt-Überschriften ---------------------------------
      Sheets("BeriVorl").Cells(3, 8).Value = BuchJahr        'Buchjahr
      Sheets("Kontenplan").Activate
      UeberCriftLinks = Sheets("Kontenplan").Cells(1, 9)    'Überschrift links
      UeberCriftRechts = Sheets("Kontenplan").Cells(1, 11) 'Überschrift rechts
      Sheets("BeriVorl").Activate
      Cells(1, 2).Value = UeberCriftLinks
      Cells(1, 8).Value = UeberCriftRechts
    End With
'2 SVE ----------------- Nicht gebrauchte Blöcke entfernen ----------------------------
23    Call KontenplanStruktur    'gibt die Kontenplankopfzeilen KPKZ^^^^-Werte
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenplanstruktur defekt. " & AktVorgang & " abgebrochen"
      GoTo EndeBeriVorlBereitstellen
    End If
    Call BeriBlockPos("BeriVorl") 'Pos in der VorVorlage ohne nicht aktuelle Kästen
    With Sheets("BeriVorl")
      .Activate
      If KPKZVermögen = 0 Then   'von Kontenplanstruktur
        Call BeriBlockPos("BeriVorl")
        Sheets("BeriVorl").Activate
        Range("A" & BeBloPosVermögen & ":K" & BeBloPosVermögen + 4 & "").Select
        Selection.Delete shift:=xlUp
      End If
      If KPKZFonds = 0 Then
        Call BeriBlockPos("BeriVorl")
        Range("A" & BeBloPosFonds - 1 & ":K" & BeBloPosFonds + 3 & "").Select
        Selection.Delete shift:=xlUp
      End If
      'KPKZAusgaben und KPKZBestand sind nicht optional, werden daher nicht gefragt
      If KPKZAusgaben2 = 0 Then
        Call BeriBlockPos("BeriVorl")
        Range("A" & BeBloPosAusgaben2 & ":K" & BeBloPosAusgaben2 + 4 & "").Select
        Selection.Delete shift:=xlUp
      End If
    End With 'Sheets("BeriVorl")
'3 BVE ------------- Verbliebene Blöcke dimensionieren: Vermögensblock ---------------
    Call BeriBlockPos("BeriVorl") 'Pos der ggf. um Blöcke gestutzten Vorlage
    If KPKZVermögen <> 0 Then
      If BBZZVermögen = 1 Then
        '------------------ eine Zeile Löschen -------------------------------
        Range("A" & BeBloPosVermögen + 2 & ":K" & BeBloPosVermögen + 3 & "").Select
        Selection.Delete shift:=xlUp  'Summenzeile auch gelöscht, durch Linie ersetzt
        '--------------- Blockumramdung ergänzen ---------------------
        Range("B" & BeBloPosVermögen + 1 & ":J" & BeBloPosVermögen + 1 & "").Select
        With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlMedium
        End With
      End If
      If BBZZVermögen > 2 Then
        '------------------ Zeilen einfügen ---------------------------
        Range("A" & BeBloPosVermögen + 2 & ":K" & _
        BeBloPosVermögen + BBZZVermögen - 2 & "").Select
        Selection.Insert shift:=xlDown
      End If
    End If 'BeBloPosVermögen <> 0
'4 BVE -*-*-*-*-*-*- Verbliebene Blöcke dimensionieren: Bestandsblock -*-*-*-*-*-*-
    Call BeriBlockPos("BeriVorl") 'Pos der ggf. gestutzten Vorlage
    If BeBloPosBestand = 0 Then
      Call MsgBox(prompt:= _
           "Im Kontenplan fehlt der Bereich ''Bestand'' (Kontoart 1)" & _
           AktVorgang & "abgebrochen", _
           Buttons:=vbOKOnly, Title:="BeriVorlErzeugen")
      AbbruchBeriVorlErzeugen = True
      Exit Sub
    End If 'BeBloPosBestand = 0
    If BBZZBestand = 1 Then
      Range("A" & BeBloPosBestand + 2 & ":K" & BeBloPosBestand + 3 & "").Select
      Selection.Delete shift:=xlUp
    End If '
'4.1 BVE ------------ Bestandsblock: Zeilen hinzufügen ------------------------
    If BBZZBestand > 2 Then
      Range("A" & BeBloPosBestand + 2 & ":J" & _
      BeBloPosBestand + BBZZBestand - 1 & "").Select
        Selection.Insert shift:=xlDown
    End If 'BBZZBestand > 2
'4.2 BVE ------------- Bestandsblock: waagerechte Linien -------------------------
    Call BeriBlockPos("BeriVorl")
    Range("B" & BeBloPosBestand & _
          ":J" & BeBloPosBestand + BBZZBestand + 1 & "").Select
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
 '5 BVE -*-*-*-*-*-*- Etwa verbliebenen Fondsblock dimensionieren -*-*-*-*-*-*-*-*-------
    Call BeriBlockPos("BeriVorl") 'Pos der ggf. gestutzten Vorlage
    If BeBloPosFonds <> 0 Then
      If BBZZFonds = 1 Then
        Range("A" & BeBloPosFonds + 2 & ":K" & BeBloPosFonds + 3 & "").Select
        Selection.Delete shift:=xlUp
        Range("B" & BeBloPosFonds + 1 & ":J" & BeBloPosFonds + 1 & "").Select
        With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlMedium
        End With
         With Selection.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlMedium
        End With
      End If
'5.1 BVE ------------ Fondsblock: Zeilen hinzufügen ------------------------
      If BBZZFonds > 2 Then
        Range("A" & BeBloPosFonds + 2 & ":J" & _
        BeBloPosFonds + BBZZFonds - 1 & "").Select
          Selection.Insert shift:=xlDown
'5.2 BVE ------------- Fondsblock: Waagerechte Linien -------------------------
        Sheets(PBerichtName).Activate
        Call BeriBlockPos(PBerichtName)
        Range("B" & BeBloPosFonds & ":J" & BeBloPosFonds + BBZZFonds + 1 & "").Select
        With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
        End With
      End If 'BBZZFonds > 2
    End If 'BeBloPosFonds <> 0
'6 BVE -*-*-*-*-*- Blöcke Ausgaben und Einnahmen dimensionieren -*-*-*-*-*-*-*-*--
    Call BeriBlockPos("BeriVorl") 'Pos der gestutzten u. erweiterten Vorlage
    If BeBloPosAusgaben = 0 Then
      Call MsgBox(prompt:= _
           "Im Kontenplan fehlt der Bereich ''Ausgaben'' (Kontoart 2)" & _
           AktVorgang & " abgebrochen", _
           Buttons:=vbOKOnly, Title:="BeriVorlErzeugen")
      AbbruchBeriVorlErzeugen = True
      Exit Sub
    End If 'BeBloPosBestand = 0
    If BeBloPosAusgaben <> 0 Then
      If BBZZAusgaben = 1 Then
        Range("A" & BeBloPosAusgaben + 2 & _
              ":F" & BeBloPosAusgaben + 3 & "").Select
        Selection.Delete shift:=xlUp
        Range("A" & BeBloPosAusgaben + 3 & _
              ":F" & BeBloPosAusgaben + 4 & "").Select
        Selection.Insert shift:=xlDown
      End If
'6.1 BVE ------ Nebeneinanderliegende Blöcke; Welcher hat mehr Zeilen? -------------
      If BBZZAusgaben > 2 Or BBZZEinnahmen > 2 Then
        If BBZZAusgaben >= BBZZEinnahmen Then
         BBZZAE = BBZZAusgaben
        Else
          BBZZAE = BBZZEinnahmen
        End If
      End If
'6.2 BVE -------- Zeilen gemäß größerem Block hinzufügen ----------
        Range("A" & BeBloPosAusgaben + 2 & ":J" & _
        BeBloPosAusgaben + BBZZAE - 1 & "").Select
          Selection.Insert shift:=xlDown
'''        EEiZ = BeBloPosAusgaben + 2        'erste eingefügte Zeile
'6.3 BVE --------- Waagerechte Linien im Ausgabeblock-Block -----------
        Range("A" & BeBloPosAusgaben + 1 & ":E" & _
          BeBloPosAusgaben + BBZZAE & "").Select
        With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
        End With
'
'6.7 BVE --------- Waagerechte Linien im Einnahmen-Block -----------------
        Call BeriBlockPos("BeriVorl")
        Range("G" & BeBloPosAusgaben & ":J" & _
              BeBloPosAusgaben + BBZZAE + 1 & "").Select
          Selection.Borders(xlDiagonalDown).LineStyle = xlNone
          Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
        End With

      End If 'BeBloPosAusgaben <> 0
'7 BVE -*-*-*-*-*-Verbliebenen Ausgaben2-u.Einnahmen2-Block dimensionieren: -*-*-*-
    Call BeriBlockPos("BeriVorl") 'Pos der gestutzten bzw. erweitertn Vorlage
 '      Range("A" & BeBloPosAusgaben2 + 2 & ":F" & BeBloPosAusgaben2 + 2 & "").Select
      If BBZZAusgaben2 = 1 And BBZZEinnahmen2 = 1 Then
        Range("A" & BeBloPosAusgaben2 + 2 & ":K").Select
        Selection.Delete shift:=xlUp
        GoTo DruckbereichKorrigieren
      End If
      If BBZZAusgaben2 > 2 Or BBZZEinnahmen2 > 2 Then
        If BBZZAusgaben2 >= BBZZEinnahmen2 Then
          Range("A" & BeBloPosAusgaben2 + 2 & ":K" & _
          BeBloPosAusgaben2 + BBZZAusgaben2 - 1 & "").Select
        Else
          Range("A" & BeBloPosAusgaben2 + 2 & ":K" & _
          BeBloPosAusgaben2 + BBZZEinnahmen2 - 1 & "").Select
       End If
       Selection.Insert shift:=xlDown
     End If
        '8 BVE ------------------ Druckbereich korrigieren ---------------------------
DruckbereichKorrigieren:
    Call BeriBlockPos("BeriVorl")
    Sheets("BeriVorl").Cells(4, 1) = BeriUntenZeile
    Range("$A$1:$K$" & BeriUntenZeile & "").Activate
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$" & BeriUntenZeile - 1 & ""
    ActiveSheet.Cells(4, 1) = BeriUntenZeile
'9 BVE ------------------ Versionseintrag, Fertignamen ---------------------------
    Sheets("BeriVorl").Cells(1, 1) = KPVersion 'von KontenplanStruktur erzeugt
    Sheets("BeriVorl").Name = "BeriVorl"
    BeriVorlVorhanden = True
'10 BVE ------------------ Vollzugsmeldung (bedingt) ---------------------------
EndeBeriVorlBereitstellen:
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Erstellen der Vorlage ''BeriVorl'' nach Kontenplanversion " _
      & KPVersion & " nicht gelungen."
    End If
    If ABBRUCH = False Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Vorlage ''BeriVorl'' nach Kontenplanversion " & KPVersion & " erstellt."
    End If
  End With   '?
  End Sub 'BeriVorlErzeugen
'============================== Ende BeriVorl erzeugen =======================

                                                                'April 2016
Sub BeriBlockPos(Blat)
'Beschreibt Vorhandensein und Position (Zeile) der Blöcke von Konten gleicher
'Art in Berichtsformularen mittels der globalen Variablen BeBloPos^^^^^^.
  Dim Z As Integer, BlattName As String
'-------------- Alte Pos löschen -------------------------
  BlattName = Blat
  With Sheets(BlattName)
    .Activate
    BeBloPosBestand = 0
    BeBloPosAusgaben = 0     'gilt auch für Einnahmen (daneben liegend)
    BeBloPosAusgaben2 = 0
    BeBloPosFonds = 0
    BeBloPosVermögen = 0
    BeBloPosErgebnis = 0
    BeBloPosUnterschrift = 0
    BeriUntenZeile = 0
'--------------------- Berichtsformular-Untenzeile -------------------
    For Z = 1 To 200
      If Cells(Z, 1) = "***" Then
        BeriUntenZeile = Z
        Exit For
      End If
    Next Z
'------------------- Pos in die globalen Variablen -----------------
    For Z = 1 To BeriUntenZeile
      If Cells(Z, 1) = 7 Then
        BeBloPosVermögen = Z
      End If
      If Cells(Z, 1) = 1 Then
        BeBloPosBestand = Z
      End If
      If Cells(Z, 1) = 6 Then
        BeBloPosFonds = Z
      End If
      If Cells(Z, 1) = 2 Then
        BeBloPosAusgaben = Z
      End If
      If Cells(Z, 1) = 4 Then
        BeBloPosAusgaben2 = Z
      End If
      If Cells(Z, 1) = 12 Then
        BeBloPosErgebnis = Z
      End If
      If Cells(Z, 1) = 13 Then
        BeBloPosUnterschrift = Z
      End If
    Next Z
  End With '
End Sub 'BeriBlockPos
                                                                'Februar 2017
Sub PeriodenDialog()
  'Liefert den ersten und letzten Tag einer als Monatszahlen eingegebenen Periode als
  'Zahlen (Tagnr.im Jahr) PADatZahl und PEDatZahl und als Datumtext PADatum, und PEDatum
  '(im Format TT:MMM) und erzeugt den periodenspezifischen Namen
  '"PSaLdLiName" (BerichtArt = "SaLdLi") bzw. PBerichtName (BerichtArt = "Bericht")
  'BerichtArt = "SaLdLi")
  'bzw. "PBerichtName"  (BerichtArt = "Bericht")
  Const TiT = "Periodendialog"  'für / "Bericht"
  TransaktJahr = Sheets("Kontenplan").Cells(1, 5)
  AbbruchPeriodenDialog = False
'1 PD --------------- Default-Perioden-Monat aus ArProt-Startzelle ---------------
  Sheets("ArProt").Activate
  TransaktDatum = Cells(StartZeile, APCDatum).Value
  DefaultMonat = MonatZZ(DatumTZ(TransaktDatum))
'2 PD ----------- Bericht-Periode von SaLdLi-Periode Übernehmen ---------------
  If BerichtArt = "SaLdLi" Then
    PEDatZahl = 0
  End If
  If BerichtArt = "Bericht" And PEDatZahl <> 0 Then
    GoTo BerichtPeriodeBestätigen
  Else
    GoTo PAM
  End If
'2PD --------------- Anfangs-Monat der Periode aus Input-Box ---------------
PAM: PAnfMonat = _
    Application.InputBox(prompt:="Bitte Anfangsmonat der Periode eingeben:" & _
                         Chr(10) & Chr(10) & "(Eine Zahl 1 bis 12)", _
                         Title:=BerichtArt & "VorlKopie" & ": Perioden-Anfang", _
                         Default:=DefaultMonat, Type:=1, Left:=300, Top:=200)
    If PAnfMonat = 0 Or PAnfMonat = False Then
      AbbruchPeriodenDialog = True
      GoTo BEENDEN
    End If
    If IsNumeric(PAnfMonat) = False Or PAnfMonat < 1 Or PAnfMonat > 12 Then
      A = MsgBox("Nur Zahl von 1 bis 12 zulässig. Wiederholen?", vbRetryCancel, _
                 Title:=BerichtArt & "VorlKopie" & ": Eingabe Anfangsmonat")
      If A = vbRetry Then GoTo PAM
      If A = vbCancel Then
        AbbruchPeriodenDialog = True
        GoTo BEENDEN
      End If
    End If 'IsNumeric(PAnfMonat) = False
'3 PD --------------- Ende-Monat der Periode aus Input-Box ---------------
PEM: PEndMonat = _
    Application.InputBox(prompt:="Bitte Endmonat der Periode eingeben:" & _
                         Chr(10) & Chr(10) & "(Eine Zahl 1 bis 12)", _
                         Title:=BerichtArt & "VorlKopie" & ": Perioden-Ende", _
                         Default:=DefaultMonat, Type:=1, Left:=300, Top:=200)
    If PEndMonat = 0 Or PEndMonat = False Then
      AbbruchPeriodenDialog = True
      GoTo BEENDEN
    End If
    If IsNumeric(PEndMonat) = False Or PEndMonat < 1 Or PEndMonat > 12 Then
      A = MsgBox("nur Zahl von 1 bis 12 zulässig. Wiederholen?", vbRetryCancel, _
               BerichtArt & "VorlKopie" & ": Eingabe Endmonat")
      If A = vbRetry Then GoTo PEM
      If A = vbCancel Then
        AbbruchPeriodenDialog = True
        GoTo BEENDEN
      End If
    End If
    If PEndMonat < PAnfMonat Then
      A = MsgBox("Endmonat darf nicht vor Afangsmonat liegen. Wiederholen?", 36, _
               BerichtArt & "VorlKopie" & ": Eingabe Endmonat")
      If A = vbYes Then GoTo PEM
      If A = vbNo Then
        AbbruchPeriodenDialog = True
        GoTo BEENDEN
      End If
    End If
'4 PD ----------- Anfangs- und End-Datum der Periode ermitteln, bestätigen -------------
    PADatZahl = MonatsErster(PAnfMonat)    'Integer, Tag-Nr. des Jahres
    PEDatZahl = MonatsLetzter(PEndMonat)   'Integer
    PADatum = DatumZT(PADatZahl)           'String im Format TT.MMM.JJJJ
    PEDatum = DatumZT(PEDatZahl)           'String
'    PVorDatum = DatumZT(PADatZahl - 1)     'String Monatsletzter Vormonat
    If BerichtArt = "SaLdLi" Then
      Gewünscht = _
        MsgBox("Zeitraum von " & PADatum & " bis " & PEDatum, vbYesNo, _
          "Summen- und SaLdenliste" & ": Periode bestätigen")
      If Gewünscht = vbYes Then
          PSaLdLiName = "SaLdLi" & PAnfMonat & "-" & PEndMonat
          GoTo Ende
      End If
      If Gewünscht = vbNo Then
        GoTo PeriodeneingabeWiederholen
      End If
    End If 'BerichtArt = "SaLdLi"
BerichtPeriodeBestätigen:
    If BerichtArt = "Bericht" Then
      Gewünscht = _
        MsgBox("Zeitraum von " & PADatum & " bis " & PEDatum, vbYesNo, _
          "Bericht" & ": Periode bestätigen")
      If Gewünscht = vbYes Then '
        PBerichtName = "Bericht" & PAnfMonat & "-" & PEndMonat
        GoTo Ende
      End If
      If Gewünscht = vbNo Then
        GoTo PeriodeneingabeWiederholen
      End If
    End If 'BerichtArt = "Bericht"
PeriodeneingabeWiederholen:
      A = MsgBox(prompt:="Periodeneingabe wiederholen?", _
           Buttons:=vbYesNo, Title:=TiT)
      If A = vbYes Then
        GoTo PAM
      Else
        AbbruchPeriodenDialog = True
        GoTo BEENDEN
      End If
BEENDEN:
    If AbbruchPeriodenDialog = True Then
      With Worksheets("ArProt")
        .Activate
        Cells(StartZeile, 9).Activate
      End With
      MsgBox ("Eingabe der Periode " & PAnfMonat & "-" & PEndMonat & Chr(10) & _
              "schlug fehl.  Gegebenenfalls neu starten.")
    End If 'Abbruch = True
Ende:
 End Sub 'Periodendialog


Sub PerBerichtBlattErzeugen() 'AnfMonat As Integer, EndMonat As Integer)
'Erzeugt aus den vom Hauptprogramm vorgegebenen Namen "PSaLdLiName" bzw. "PBerichtName",
'in denen die Periode erscheint, ein Blatt für den Eintrag der Werte der angegebenen
'Periode durch Kopieren der Vorlagen, Plazieren an geeignete Stelle der Blattleiste der
'Mappe und Eintragen der Periodenüberschrift.
'Gegebenenfalls wird die globale Variable 'Abbruch' auf TRUE gesetzt.
'Sub Setzt voraus, dass Formularvorlagen mit den Namen "SaLiVorl" und "BeriVorl"
'existieren, die in den ersten drei Spalten der ersten unbenutzten Zeile die Spalte und
'die Zeile des Anfangs- und Enddatums enthalten.
  Dim FertigText As String, PerBerVorh As Boolean
  Const TiT = "PeriodBlattErzeugen"

'1 PBE ------------------------ Perioden-Bericht schon vorhanden? -----------------------
      PerBerVorh = False
      For Each W In Worksheets
        If W.Name = PBerichtName Then
          PerBerVorh = True
          Exit For
        End If
      Next W
'2 PBE --------- etwa vorhandenen PBericht löschen.  ---------------------------------
      If PerBerVorh = True Then
        Application.DisplayAlerts = False
        Sheets(PBerichtName).Delete
        Application.DisplayAlerts = False
        FertigText = _
          "Bereits vorhandenes Blatt ''" & PBerichtName & "'' wurde gelöscht."
      End If 'PerBerVorh = True
'3 PBE----------------- Bericht-Vorlage kopieren, "Berichtn-n" nennen -------------------
PeriodenBerichtBlatt:
    With Worksheets("BeriVorl")
      .Activate
      With ActiveWorkbook
        Sheets("BeriVorl").Activate
        ActiveSheet.Copy Before:=ActiveWorkbook.Sheets("BeriVorl")
        ActiveSheet.Name = PBerichtName   'ADatumZelle, EDatumZelle, RechtsUntenZelle
      End With 'ActiveWorkbook
    End With 'Worksheets("BeriVorl")
'4 PBE ----------- Anfangs- und End-Datum in Bericht-Kopf eintragen -------------------
    With Sheets(PBerichtName)
      .Activate
      Cells(3, 3).Value = LangDatum(PADatum)
      Cells(3, 7).Value = LangDatum(PEDatum)
      Cells(3, 8).Value = BuchJahr
    End With 'Sheets(PBerichtName)
    If MeldeStufe >= 2 Then
      A = MsgBox(prompt:=FertigText, Buttons:=vbOKOnly, Title:=TiT)
    End If
End Sub 'PerBerichtBlattErzeugen =======================================

Sub PeriodenBerichtFüllen() '============================================
  Dim KPZeile As Integer, AktPeriodenBlatt As String, BeFoZ As Integer
  Dim Wert As Double, Wert2 As Double, Proz As Double, Proz2 As Double
  Dim Bas As Double, Bas2 As Double
  Dim BeschrText As String
  Dim SummenZeile As Integer, SummenZeile2 As Integer
  Dim BBZZAE As Integer
  
  If ABBRUCH = True Then GoTo EndePeriodenBerichtFüllen

    
'1 PBF ----------------- Übertragungsdaten für alle Bereiche -------------------
  AktPeriodenBlatt = "Bericht" & PAnfMonat & "-" & PEndMonat & ""
BereichsAnfuEnde:
  Call KontenplanStruktur
  If ABBRUCH = True Then GoTo EndePeriodenBerichtFüllen
  Call BeriBlockPos(AktPeriodenBlatt)   'Periodenblatt)
  With Sheets("Kontenplan")
    .Activate
'2 PBF ----------------- Bereich "Vermögen" übertragen --------------------
BereichVermoegeen:
  If KPKZVermögen = 0 Then GoTo BereichBestand
    For KPZeile = KPKZVermögen + 1 To KPKZVermögen + SLZZVermögen
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus < 2 Then
        Call KontoBlattEinrichten(Cells(KPZeile, 2))
      End If
      If AKtoStatus >= 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosVermögen + AKtoBeriZeile
          If Cells(BeFoZ, 7) <> "" Then
            Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & ", "
          End If
          Cells(BeFoZ, 2) = AKtoBeriText   'Bei KtoArt Fonds keine
          Cells(BeFoZ, 3) = PerAnfSaLdo    'Summierung in den einzelnen
          Cells(BeFoZ, 8) = PerEndSaLdo    'Berichtszeilen vorgesehen
          'Bei längerem Text Zeile erhöhen --------------------
          BeschrText = Cells(BeFoZ, 2)
          If Len(BeschrText) > 27 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
'2.1 PBF ---------- Prozentwerte Bereich Vermögen eintragen ------------------
    If BBZZVermögen > 1 Then
      Bas = Cells(BeBloPosVermögen + BBZZVermögen + 1, 3).Value
      Bas2 = Cells(BeBloPosVermögen + BBZZVermögen + 1, 8).Value
      For BeFoZ = BeBloPosVermögen + 1 To BeBloPosVermögen + BBZZVermögen
        Wert = Cells(BeFoZ, 3).Value
        Wert2 = Cells(BeFoZ, 8).Value
        Proz = Wert * 100 / Bas
        Proz2 = Wert2 * 100 / Bas2
        Cells(BeFoZ, 4) = Proz
        Cells(BeFoZ, 5) = "%"
        Cells(BeFoZ, 9) = Proz2
        Cells(BeFoZ, 10) = "%"
      Next BeFoZ
    End If 'BBZZVermögen > 1
'Ende BereichVermögen
'3 PBF ----------------- Bereich "Bestand" übertragen ------------------
BereichBestand:
  If KPKZBestand = 0 Then GoTo BereichFonds
    For KPZeile = KPKZBestand + 1 To KPKZBestand + SLZZBestand
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoNr = 0 Then GoTo NächsteKPZ
      If AKtoArt <> BestandKto Then GoTo ProzentWerteBestand
      If AKtoStatus >= 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosBestand + AKtoBeriZeile
          If Cells(BeFoZ, 2) <> "" Then
            Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & ", "
          End If
          Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & AKtoBeriText
          Cells(BeFoZ, 3) = Cells(BeFoZ, 3) + PerAnfSaLdo
          Cells(BeFoZ, 7) = Cells(BeFoZ, 7) + PerDifSaLdo
          Cells(BeFoZ, 8) = Cells(BeFoZ, 8) + PerEndSaLdo
          'Bei längerem Text Zellenhöhe vergrößern --------------------
          BeschrText = Cells(BeFoZ, 2)
          If Len(BeschrText) > 25 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
          If Len(BeschrText) > 35 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 46.8
          End If
        End With
      End If 'AKtoStatus > 2
NächsteKPZ:
    Next KPZeile
'3.1 PBF -------------- Prozentwerte Bestand eintragen ---------------
ProzentWerteBestand:
    Sheets(AktPeriodenBlatt).Activate
    Bas = Cells(BeBloPosBestand + BBZZBestand + 1, 3).Value
    Bas2 = Cells(BeBloPosBestand + BBZZBestand + 1, 8).Value
    For BeFoZ = BeBloPosBestand + 1 To BeBloPosBestand + BBZZBestand
      Wert = Cells(BeFoZ, 3).Value
      Wert2 = Cells(BeFoZ, 8).Value
      Proz = Wert * 100 / Bas
      Proz2 = Wert2 * 100 / Bas2
      Cells(BeFoZ, 4) = Proz
      Cells(BeFoZ, 5) = "%"
      Cells(BeFoZ, 9) = Proz2
      Cells(BeFoZ, 10) = "%"
    Next BeFoZ
'Ende BereichBestand  KPKZBestand <> 0
'4 PBF ----------------- Bereich "Fonds" übertragen ---------------------
BereichFonds:
  If KPKZFonds = 0 Then GoTo BereichAusgaben
    For KPZeile = KPKZFonds + 1 To KPKZFonds + SLZZFonds
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus > 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosFonds + AKtoBeriZeile
          If Cells(BeFoZ, 7) <> "" Then
            Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & ", "
          End If
          Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & AKtoBeriText
          Cells(BeFoZ, 3) = Cells(BeFoZ, 3) + PerAnfSaLdo
          Cells(BeFoZ, 8) = Cells(BeFoZ, 8) + PerEndSaLdo
          'Bei längerem Text Zeile erhöhen --------------------
          BeschrText = Cells(BeFoZ, 2)
          If Len(BeschrText) > 20 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
'4.1 PBF ----- Prozentwerte Fonds PBeginn bezogen auf Bestand -------
    Bas = Cells(BeBloPosBestand + BBZZBestand + 1, 3)
    If Bas = 0 Then GoTo ProzEndFond
    For BeFoZ = BeBloPosFonds + 1 To BeBloPosFonds + BBZZFonds
      Wert = Cells(BeFoZ, 3).Value
      If Wert = 0 Then GoTo FondBeginNextBeFoZ
      Proz = Wert * 100 / Bas
      Cells(BeFoZ, 4) = Proz
      Cells(BeFoZ, 5) = "%"
FondBeginNextBeFoZ:
    Next BeFoZ
'4.2 PBF ----- Prozentwerte Fonds PEnde bezogen auf Bestand ------
ProzEndFond:
    Bas2 = Cells(BeBloPosBestand + BBZZBestand + 1, 8).Value
    If Bas2 = 0 Then GoTo BereichAusgaben
    For BeFoZ = BeBloPosFonds + 1 To BeBloPosFonds + BBZZFonds
      Wert2 = Cells(BeFoZ, 8).Value
      If Wert2 = 0 Then GoTo FondEndNextBeFoZ
      Proz2 = Wert2 * 100 / Bas2
      Cells(BeFoZ, 9) = Proz2
      Cells(BeFoZ, 10) = "%"
FondEndNextBeFoZ:
    Next BeFoZ
'  Ende Bereich Fonds KPKZFonds <> 0
'5 PBF ----------------- Bereich "Ausgaben" übertragen ------------------------
BereichAusgaben:
  If KPKZAusgaben = 0 Then GoTo BereichEinnahmen  'KPKZ ist Kontenplan-Bereichskopfzeile
    For KPZeile = KPKZAusgaben + 1 To KPKZAusgaben + SLZZAusgaben
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoBeriZeile <> "" And AKtoBeriZeile = 0 Then GoTo NextKPZeile
      If AKtoStatus > 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosAusgaben + AKtoBeriZeile
          If Cells(BeFoZ, 2) <> "" Then
            Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & ", "
          End If
          Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & AKtoBeriText
         Cells(BeFoZ, 3) = Cells(BeFoZ, 3) + PerEndSaLdo - PerAnfSaLdo
          'Bei längerem Text Zeile erhöhen --------------------
          BeschrText = Cells(BeFoZ, 2)
          If Len(BeschrText) > 27 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
          If Len(BeschrText) > 55 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 46.8
          End If
        End With
      End If 'AKtoStatus > 2
NextKPZeile:
    Next KPZeile
'5.1 PBF --------- Prozentwerte Bereich Ausgaben eintragen-------
    If BBZZAusgaben >= BBZZEinnahmen Then
      BBZZAE = BBZZAusgaben
    Else
      BBZZAE = BBZZEinnahmen
    End If
    Bas = Cells(BeBloPosAusgaben + BBZZAE + 1, 3)
    If Bas = 0 Then GoTo BereichEinnahmen 'Verzicht auf Prozentangaben
    For BeFoZ = BeBloPosAusgaben + 1 To BeBloPosAusgaben + BBZZAE
      If Cells(BeFoZ, 3) = 0 Then GoTo NBeFoZ2 'Verzicht diese Zeile
      Wert = Cells(BeFoZ, 3)
      Proz = (Wert / Bas) * 100
      Cells(BeFoZ, 4) = Proz
      Cells(BeFoZ, 5) = "%"
NBeFoZ2:
    Next BeFoZ
'  Ende BereichAusgaben  'KPKZAusgaben <> 0
'6 PBF ----------------- Bereich "Einnahmen" übertragen ------------------------
BereichEinnahmen:
  If KPKZEinnahmen = 0 Then GoTo BereichAusgaben2
  With Sheets("Kontenplan")
    .Activate
    For KPZeile = KPKZEinnahmen + 1 To KPKZEinnahmen + SLZZEinnahmen
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus >= 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosAusgaben + AKtoBeriZeile
          If Cells(BeFoZ, 7) <> "" Then
            Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & ", "
          End If
          Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & AKtoBeriText
          Cells(BeFoZ, 8) = Cells(BeFoZ, 8) + PerEndSaLdo - PerAnfSaLdo
          'Bei längerem Text Zellenhöhe vergrößern --------------------
          BeschrText = Cells(BeFoZ, 7)
          If Len(BeschrText) > 27 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
'6.1 PBF --------- Prozentwerte Bereich Einnahmen eintragen -----------
    With Sheets(AktPeriodenBlatt)
      .Activate
      If BBZZAusgaben >= BBZZEinnahmen Then
        BBZZAE = BBZZAusgaben
      Else
        BBZZAE = BBZZEinnahmen
      End If
      Bas = Cells(BeBloPosAusgaben + BBZZAE + 1, 8)
      If Bas = 0 Then GoTo BereichAusgaben2 'Verzicht auf Prozente
      For BeFoZ = BeBloPosAusgaben + 1 To BeBloPosAusgaben + BBZZAE
        If Cells(BeFoZ, 8) = 0 Then GoTo NBeFoZ3 'Diese Zeile ohne Proz
        Wert = Cells(BeFoZ, 8).Value '*** Typen unverträglich
        Proz = (Wert / Bas) * 100 '*** Überlauf
        Cells(BeFoZ, 9) = Proz
        Cells(BeFoZ, 10) = "%"
NBeFoZ3:
      Next BeFoZ
    End With
'  Ende KPKZf 'KPKZEinnahmen <> 0
'7 PBF ----------------- Bereich "Ausgaben2" übertragen ---------------
BereichAusgaben2:
If KPKZAusgaben2 = 0 Then GoTo BereichEinnahmen2
    For KPZeile = KPKZAusgaben2 + 1 To KPKZAusgaben2 + SLZZAusgaben2
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus > 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosAusgaben2 + AKtoBeriZeile
          If Cells(BeFoZ, 2) <> "" Then
            Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & ", "
          End If
          Cells(BeFoZ, 2) = Cells(BeFoZ, 2) & AKtoBeriText
          Cells(BeFoZ, 3) = Cells(BeFoZ, 3) + PerEndSaLdo - PerAnfSaLdo
          'Bei längerem Text Zeile erhöhen --------------------
          BeschrText = Cells(BeFoZ, 2)
          If Len(BeschrText) > 20 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
'7.1 PBF ---------- Prozentwerte Bereich Ausgaben2 eintragen ------------
    Sheets(AktPeriodenBlatt).Activate
    If BBZZAusgaben2 >= BBZZEinnahmen2 Then
      BBZZAE = BBZZAusgaben2
    Else
      BBZZAE = BBZZEinnahmen2
    End If
    Bas = Cells(BeBloPosAusgaben2 + BBZZAE + 1, 3)
    For BeFoZ = BeBloPosAusgaben2 + 1 To BeBloPosAusgaben2 + BBZZAE
      If Cells(BeFoZ, 3) = 0 Then GoTo NeBefoz
      Wert = Cells(BeFoZ, 3).Value
      If Bas <> 0 Then
        Proz = (Wert / Bas) * 100
      End If
      Cells(BeFoZ, 4) = Proz
      Cells(BeFoZ, 5) = "%"
NeBefoz:
    Next BeFoZ
'  Ende BereichAusgaben2   KPKZAusgaben2 <> 0
'8 PBF ----------------- Bereich "Einnahmen2" übertragen -----------------
BereichEinnahmen2:
  If KPKZEinnahmen2 = 0 Then GoTo BereichErgebnis
      For KPZeile = KPKZEinnahmen2 + 1 To KPKZEinnahmen2 + SLZZEinnahmen2
      Sheets("Kontenplan").Activate
      Cells(KPZeile, 2).Activate
      Call KtoKennDat(Cells(KPZeile, 2))
      If AKtoStatus > 2 Then
        Call PeriodenSaLden(AKtoNr)  'PerAnfSaLdo,PerEndSaLdo
        With Sheets(AktPeriodenBlatt)
          .Activate
          BeFoZ = BeBloPosAusgaben2 + AKtoBeriZeile
          If Cells(BeFoZ, 7) <> "" Then
            Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & ", "
          End If
          Cells(BeFoZ, 7) = Cells(BeFoZ, 7) & AKtoBeriText
          Cells(BeFoZ, 8) = Cells(BeFoZ, 8) + PerEndSaLdo - PerAnfSaLdo
          'Bei längerem Text Zellenhöhe vergrößern --------------------
          BeschrText = Cells(BeFoZ, 7)
          If Len(BeschrText) > 27 Then
            Range("B" & BeFoZ & "").Select
            Rows("" & BeFoZ & ":" & BeFoZ & "").RowHeight = 31.8
          End If
        End With
      End If 'AKtoStatus > 2
    Next KPZeile
'8.1 PBF --------- Prozentwerte Bereich Einnahmen2 eintragen ---------------
    If BBZZAusgaben2 >= BBZZEinnahmen2 Then
      BBZZAE = BBZZAusgaben2
    Else
      BBZZAE = BBZZEinnahmen2
    End If
    Sheets(AktPeriodenBlatt).Activate
    Bas = Cells(BeBloPosAusgaben2 + BBZZAE + 1, 8)  '.Value
    For BeFoZ = BeBloPosAusgaben2 + 1 To BeBloPosAusgaben2 + BBZZAE
      If Cells(BeFoZ, 8) = 0 Then GoTo NBeFoZ
      Wert = Cells(BeFoZ, 8).Value '*** Typen unverträglich
      Proz = (Wert / Bas) * 100 '*** Überlauf
      Cells(BeFoZ, 9) = Proz
      Cells(BeFoZ, 10) = "%"
NBeFoZ:
    Next BeFoZ
'  Ende BereichEinnahmen2  'KPKZEinnahmen2 <> 0
  End With 'With Sheets(AktVorlage)
'9 PBF --------------- Jahresanfangstand in den Ergebnisblock -----------
BereichErgebnis:
    Call BestandSummeJahrAnf
    Call BeriBlockPos(PBerichtName)
  With Sheets(PBerichtName)
    Sheets(PBerichtName).Cells(BeBloPosErgebnis + 3, 3) = JahrAnfBestand
'10 PBF ------------------ Änderung in % in den Ergebnisblock ------------
  With Sheets(PBerichtName)
  Bas = Cells(BeBloPosErgebnis + 2, 3)
  Wert = Cells(BeBloPosErgebnis + 2, 7)
  If Bas <> 0 Then
    Proz = (Wert / Bas) * 100
    Cells(BeBloPosErgebnis + 2, 9) = Proz
    Cells(BeBloPosErgebnis + 2, 10) = "%"
  End If
  Bas2 = Cells(BeBloPosErgebnis + 3, 3)
  Wert2 = Cells(BeBloPosErgebnis + 3, 7)
  If Bas2 <> 0 Then
    Proz = (Wert2 / Bas) * 100
    Cells(BeBloPosErgebnis + 3, 9) = Proz
    Cells(BeBloPosErgebnis + 3, 9) = "&"
  End If
End With 'Sheets(PBerichtName)
End With 'Kontenplan
EndePeriodenBerichtFüllen:
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Periodenbericht füllen nicht gelungen"
  End If
  If ABBRUCH = False Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Periodenbericht gefüllt "
  End If
End With  '?
End Sub 'PeriodenBerichtFüllen() ========================================
  
Sub BestandSummeJahrAnf()
  Dim Z As Integer
  
  If ABBRUCH = True Then GoTo EndeBestandSummeJahrAnf
  Call KontenplanStruktur
  If ABBRUCH = True Then GoTo EndeBestandSummeJahrAnf
  With Sheets("Kontenplan").Activate
    JahrAnfBestand = 0
    For Z = KPKZBestand + 1 To KPKZBestand + SLZZBestand
      Cells(Z, 2).Activate
      Call KtoKennDat(ActiveCell)
      If AKtoStatus >= 3 Then
        JahrAnfBestand = JahrAnfBestand + Sheets(AKtoBlatt).Cells(4, 6)
      End If
    Next Z
  End With
EndeBestandSummeJahrAnf:
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "BestandSumme zum JahresAnfang nicht ermittelt"
  End If
  If ABBRUCH = False Then
    MELDUNG = MELDUNG & Chr(10) & _
    "BestandSumme zum JahresAnfang ermittelt"
  End If
End Sub

Sub PeriodenSaLden(KtoNummer) ' PAnf, PEnd)
'Aufgerufen von den Füll-Programmen SaldLiBereichFüllen u. PeriodenBerichtFüllen
  'in: 'PADatZahl = MonatsErster(PAnfMonat)    'Integer
       'PEDatZahl = MonatsLetzter(PEndMonat)   'Integer
       'PADatum = DatumZT(PADatZahl)           'String
       'PEDatum = DatumZT(PEDatZahl)           'String
  'out:'PerAnfSaLdo, PerAnfSoll, PerAnfHaben
       'PerEndSaLdo, PerEndSoll, PerEndHaben
       'PerDifSaLdo, PerDifSoll, PerDifHaben
  Dim KoZ As Integer, AnfBlat As String, AnfZeil As Integer, _
      AnfSpal As Integer, LezMonatMEintr As Integer
  Dim PAnfSZeile As Integer, PEndSZeile As Integer, KZPAnf As Integer, _
      KZPEnd As Integer
'1 PS ------------------ Anfangssituation -----------------------------
  AnfBlat = ActiveSheet.Name
  AnfZeil = ActiveCell.Row
  AnfSpal = ActiveCell.Column
'2 PS ---------------- Löschen vorhandener SaLden ----------------------
  PerAnfSaLdo = 0
  PerAnfSoll = 0
  PerAnfHaben = 0
  PerEndSaLdo = 0
  PerEndSoll = 0
  PerEndHaben = 0
'2 PS -------------------- Daten des Kontos ----------------------------
  Call KtoKennDat(KtoNummer)
'3 PS ----------------- Letzter Monat mit Eintrag ----------------------
  With Sheets(AKtoBlatt)
    .Activate
    LezMonatMEintr = Cells(Cells(1, 1), KoCDatum).Offset(1, 1)
'4 PS ----------------- PAnfSaLdoZeile ermitteln -----------------
PAnfSchleife:
    If PAnfMonat = 1 Then
      PAnfSZeile = 4
      GoTo PEndSchleife
    End If
    For KoZ = 6 To Cells(1, 1) + 2
      Cells(KoZ, 5).Activate
      If Cells(KoZ - 1, 5) = "Umsatz Periode" And Cells(KoZ, 5) = "Kontostand" Then
        If Cells(KoZ - 1, 3) < PAnfMonat - 1 Then   'für den Fall keiner weiteren
          PAnfSZeile = KoZ                          'Monate mit Einträgen
        End If                                      'Wird ggf. überschrieben
        If Cells(KoZ - 1, 3) <= PAnfMonat - 1 And _
           Cells(KoZ - 2, 2) = "***" Then
          PAnfSZeile = KoZ
          Exit For
        End If
        If Cells(KoZ - 1, 3) = PAnfMonat - 1 Then
          PAnfSZeile = KoZ
          Exit For
        End If
      End If
    Next KoZ
'5 PS ------------- PAnfSaLdoZeile und PEndSaLdoZeile suchen -------------
PEndSchleife:
    For KoZ = 6 To Cells(1, 1) + 2
      Cells(KoZ, 5).Activate
      If Cells(KoZ - 1, 5) = "Umsatz Periode" And Cells(KoZ, 5) = "Kontostand" Then
        If Cells(KoZ - 1, 3) < PEndMonat Then       'für den Fall keiner weiteren
          PEndSZeile = KoZ                          'Monate mit Einträgen
        End If                                      'Wird ggf. überschrieben
        If Cells(KoZ - 1, 3) <= PEndMonat And _
           Cells(KoZ - 2, 2) = "***" Then
          PEndSZeile = KoZ
          Exit For
        End If
        If Cells(KoZ - 1, 3) = PEndMonat Then
          PEndSZeile = KoZ
          Exit For
        End If
      End If
    Next KoZ
PeriodenAnfEndSaLdo:
    PerAnfSaLdo = Cells(PAnfSZeile, 6).Value
    PerAnfSoll = Cells(PAnfSZeile, 7).Value
    PerAnfHaben = Cells(PAnfSZeile, 8).Value
    PerEndSaLdo = Cells(PEndSZeile, 6).Value
    PerEndSoll = Cells(PEndSZeile, 7).Value
    PerEndHaben = Cells(PEndSZeile, 8).Value

'8 PS ----------- Differenz, Vorzeichenumkehr bei Eingabekonten ----------------
PeriodenDifferenz:
    PerDifSaLdo = PerEndSaLdo - PerAnfSaLdo
    If AKtoArt = EingabKto Or AKtoArt = Eingab2Kto Then   'bei Eingabekonten
      PerAnfSaLdo = -PerAnfSaLdo                          'Vorzeichen umkehren
      PerEndSaLdo = -PerEndSaLdo
      PerDifSaLdo = -PerDifSaLdo
    End If
  End With 'Sheets(AKtoBlatt)
'9 PS -------------------- Anfangssituation ----------------------------------
  Sheets(AnfBlat).Activate
  Cells(AnfZeil, AnfSpal).Activate
End Sub 'PeriodenSaLden

'========================================================================
'----------------- Aktive Seite einrichten und drucken ------------------
Sub AktBlattDrucken()  ' AktBlattDrucken Makro   Tastenkombination: Strg+d
Attribute AktBlattDrucken.VB_ProcData.VB_Invoke_Func = "d\n14"
  'Druckt das aktive Blatt sensitiv nach den in ExAcc-Projekten vorkommenden
  'Blattformaten Kontenplan, ArProt, Konto, Kontenstandtabelle,SaldLi, Bericht.
  
  Dim BlattName As String, EndZeile As Integer 'für Druckbereich
  Dim KopfZeileLinks As String, KopfZeileRechts As String
  Dim FussZeileLinks As String, FussZeileMitte As String, FussZeileRechts As String
  
With ActiveWindow
  BlattName = ActiveSheet.Name
  ExAccVersion = ThisWorkbook.Name
  KopfZeileLinks = " " & Sheets("Kontenplan").Cells(1, 9) & "  "   'Blank vorweg sonst
  KopfZeileRechts = " " & Sheets("Kontenplan").Cells(1, 11) & "  " 'wird eine Zahl dem
  FussZeileLinks = "&8&HErstellt mit " & ExAccVersion               'Schriftgrad zugeschlagen
  FussZeileMitte = "&8&HSeite &P von &N "
  FussZeileRechts = "&8&D  &U  Uhr"
  
'------------------ Alte Kopf- und Fußzeilen löschen -----------------------------
    Application.PrintCommunication = True
    With ActiveSheet.PageSetup
        .LeftHeader = "                                          "
        .CenterHeader = "                                        "
        .RightHeader = "                                         "
        .LeftFooter = "                                          "
        .CenterFooter = "                                        "
        .RightFooter = "                                         "
    End With
'    Application.PrintCommunication = True
  
'--------------------------- Blatt-spezifisches Einrichten --------------------
  If BlattName = "Kontenplan" Or Cells(1, 4) = "Kontenplan" Then
    GoTo KontenplanDrucken
  End If
 If BlattName = "ArProt" Or Cells(1, 6) = "ARBEITSPROTOKOLL" Then
    GoTo ArbeitsprotokollDrucken
  End If
  If Cells(1, 2) = "Kontoart" Then
    GoTo KontoBlattDrucken   'auch für KntoVorl
  End If
  If BlattName = "KonTab" Or Cells(1, 4) = "Kontenstandtabelle" Then
    GoTo KontenStandTabelleDrucken
  End If
  If Left(BlattName, 6) = "SaLdLi" Or Cells(2, 2) = "Summen- und SaLdenliste" Then
    GoTo SaLdLiDrucken   'auch für SaLiVorl und SaLiVorlVorl
  End If
  If Left(BlattName, 7) = "Bericht" Or Cells(2, 2) = "Bericht des Schatzmeisters" Then
    GoTo BerichtDrucken
  End If
  A = MsgBox("Seiteneingerichtetes Drucken für " & BlattName & " nicht implementiert." _
      & Chr(10) & "Von ExCel angebotenes Drucken verwenden!", _
      vbOKOnly, "Drucken mit ExAcc")
  GoTo EndeDrucken
  
KontenplanDrucken:   'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(1, 3)
    Sheets(BlattName).Select
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$3"
        .PrintTitleColumns = ""
    End With
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$" & EndZeile
'    Application.PrintCommunication = False
        .LeftHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
        .CenterHeader = ""
        .RightHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "erstellt mit " & ExAccVersion & ""
        .CenterFooter = "Seite &P von &N"
        .RightFooter = "&8&D &U Uhr"
        .LeftFooter = FussZeileLinks    '"&8&HErstellt mit " & ExAccVersion
        .CenterFooter = FussZeileMitte  '"&8&HSeite &P von &N "
        .RightFooter = FussZeileRechts  '"&8&D  &U  Uhr"
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
        '-------------------- Druckoptionen Kontenplan ----------
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintSheetEnd
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
    End With
'    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken
    
ArbeitsprotokollDrucken:  'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(1, 3) + 4
    Range("C" & EndZeile & "").Select
    Sheets(BlattName).Select
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$2"
        .PrintTitleColumns = ""
    End With
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$K$" & EndZeile
'    Application.PrintCommunication = False
    KopfZeileLinks = " " & Sheets("Kontenplan").Cells(1, 9)   'Blank vorweg sonst
    KopfZeileRechts = " " & Sheets("Kontenplan").Cells(1, 11) 'wird eine Zahl dem
    With ActiveSheet.PageSetup                          'Schriftgrad 14 angehängt
        .LeftHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
        .CenterHeader = ""
        .RightHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "Erstellt mit " & ExAccVersion & ""
        .CenterFooter = "Seite &P von &N"
        .RightFooter = "&8&D &U Uhr"
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
        '--------------- Druckoptionen Arbeitsprotokoll -------------
        .PrintHeadings = False   'Was ist hiermit gemeint?
        .PrintGridlines = True
        .PrintComments = xlPrintSheetEnd
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 10
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
    End With
 '   Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken
    
KontoBlattDrucken:  'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(1, 1) + 4
    Sheets(BlattName).Select
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$2"
        .PrintTitleColumns = ""
    End With
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$I$" & EndZeile
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
        .CenterHeader = ""
        .RightHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "&""Arial""&8" & FussZeileLinks & ""
'        .CenterFooter = "&8" & "Seite &P von &N "
        .CenterFooter = "&8&HSeite &P von &N "
        .RightFooter = "&8&D &U Uhr"
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
 '       .PrintHeadings = False   '--------- Kontoblatt -------
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 10
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
    End With
'    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken

KontenStandTabelleDrucken:  'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(1, 3)
'    Range("C" & EndZeile & "").Select
    Sheets(BlattName).Select
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$R$" & EndZeile
'    Application.PrintCommunication = False
'------------ Header und Footer Kontenstandtabelle ----------------------
    With ActiveSheet.PageSetup
        .LeftHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
        .CenterHeader = ""
        .RightHeader = _
        "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "&""Arial""&8" & FussZeileLinks & ""
        .CenterFooter = ""
        .RightFooter = "&8&D &U Uhr"
'-------------------- Margins Kontenstandtabelle --------------------
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)  '(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.118110236220472) '(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
'----------------- Print options Kontenstandtabelle ------------------
 '       .PrintHeadings = True '<--- versuchsweise auf true abgeändert
        .PrintGridlines = True
        .PrintComments = xlPrintSheetEnd
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
    End With
 '   Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken
    
SaLdLiDrucken:  'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(3, 1)
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$4"
        .PrintTitleColumns = ""
    End With
    
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$1:$N$" & EndZeile & ""
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
         .LeftHeader = _
         "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
         .CenterHeader = ""
         .RightHeader = _
         "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "erstellt mit " & ExAccVersion & ""
        .CenterFooter = ""
        .RightFooter = "&8&D &U Uhr"
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.275590551181102)
        .HeaderMargin = Application.InchesToPoints(0.393700787401575)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
        '---------------------- Druckoptionen SaLdLi ----------------
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
     End With
'    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken
    
BerichtDrucken:  'Application.PrintCommunication = True im Kopf- und Fuß-Lösch-Vorspann
    EndZeile = Cells(4, 1) - 1
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$4"
        .PrintTitleColumns = ""
    End With
'    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$1:$K$" & EndZeile & ""
'    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
   '     "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileLinks & ""
        .CenterHeader = ""
        .RightHeader = ""
 '      "&""Bookman Old Style,Fett Kursiv""&14" & KopfZeileRechts & ""
        .LeftFooter = "&8erstellt mit " & ExAccVersion & ""
        .CenterFooter = ""
        .RightFooter = "&8&D &U Uhr"
        .LeftMargin = Application.InchesToPoints(0.196850393700787) '(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787) '(0.78740157480315)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0.118110236220472) '(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.118110236220472)
        '--------------------- Druckoptionen Bericht -----------------
        .PrintHeadings = False  'welche Headings sind gemeint?
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False                '85 geht in der Breite auf ein Blatt
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False       'True war ursprünglich
        .AlignMarginsHeaderFooter = False
     End With
'    Application.PrintCommunication = True
 '   Range("D13:D14").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    GoTo EndeDrucken
''End With
EndeDrucken:
End Sub
    

