Attribute VB_Name = "ArProtSchreiben"

'******************************************************************************
'* Modul   ERFASSEN (ArProtSchreiben)   Shortcut: Strg+b                      *
'******************************************************************************
'Änderungen 6.1.2002:
'  Umstellung auf "Text-Datum", weil Excel ab Dezember das Datum im Format
'  "d.mmm" falsch interpretiert. (Modul DatumTextScan)
'Änderungen Januar 2004:
'  Buchungs-Identnummer eingeführt; ermöglicht, bei Korrekturbuchungen auch die
'  Ta-Nr. zu ändern.
'Änderungen März 2016:
'  Jahreswechsel automatisierter
'Änderungen Mai-Juli 2017
'  Umstellung auf Informations- und Format-Prüf-Routinen KontenplanStruktur,
'  KtoKennDat und ABZParam, Stornieren unbedingt vollständig
'Überarbeitung Januar 2019:
'  Schreibhilfen für Sollkonto- und Habenkonto-Spalten
'==============================================================================
'Bediener-Schnittstelle zum Bearbeiten  einer Arbeitsprotokoll-Zeile
'(Zeile schreiben, Buchen, Stornieren) oder zur Erstellung von Berichten
'(Kontenstandstabelle, Summen- und Saldenliste, Periodenbericht).
'Bearbeiten nach Maßgabe, in welcher Spalte sich die aktive Zelle befindet:
'Werte aus vorhergehender Arbeitsprotokollzeile übertragen, oder hochzählen
'oder aus einer Kontoplanzelle auffüllen oder eine von zwei Konten-verändernde
'Aktionen starten, nämlich Buchen von der Spalte "gebucht"(8) aus und Stornieren
'von der Spalte "TA"(1) aus.
'Diese beiden Aktionen werden mit allen notwendigen Parametern von der Routine
'"Sub ABZParam" versorgt, die in einem Block von Globalvariablen abgelegt sind
'und alle mit "ABZ" beginnen.
'Sub ABZParam verwendet ihrerseits die Routine "Sub KtoKennDat", die sich ihrer-
'seits auf Sub KontenplanStruktur stützt, beide aus dem Modul KontenplanPflege.
'Sie nehmen vielfältige Formatprüfungen an den Kontoblättern vor.
'==============================================================================
'Ein Buchungsauftrag erfordert mindestens 2, höchstens 4 Zugriffe auf diese
'Konten; ein Stornierungsauftrag ebenso. Es ist unerlässlich, das dies konsistent
'geschieht, d.h. alle erforderlichen Einträge müssen vollständig vollzogen
'werden, alle Löschungen ebenso.
'Die Verwendung der Inforoutinen "KontoplanStruktur", "KtoKennDat" und "ABZParameter"
'hat den Zweck, mögliche Abbruchursachen so früh wie möglich, möglichst noch vor
'Zugriff auf ein Konto, zu erkennen und so Inkonsistenzen zu vermeiden. Deshalb
'schreiben die Inforoutinen ihre Ergebnisse in den globalen ABZ-Block.
'ABZ = Aktuelle Buchungs-Zeile; alle Variablen des Blocks beginnen mit "ABZ" und
'enthalten die Werte des letzten ABZParam-Aufrufs; sie werden vom nächsten Aufruf
'überschrieben.
'                Begleitende Berichterstattung bei größeren Vorgängen
'Die shortcut-Routinen (die durch Tippen einer Tastenkombination aufgerufen werden,
'wie z.B. ERFASSEN durch strg+b von ArProt aus) setzen die globalen Variablen MELDUNG
'(String, wird zu Beginn gelöscht) und ABBRUCH (Boolan, wird zu Beginn false gesetzt)
'zurück. Die Vorgangs-Routinen (durch Verzweigung der shortcut-Routine aufgerufenen,
'wie z.B."Buchen", "Stornieren" oder "BerichteErstellen") schreiben zu Beginn den
'Vorgang in die globale Stringvariable VORGANG und anschließend an neuralgischen
'Stellen kumulierend (konkateniert) einen Text in MELDUNG. Wenn dieser Text eine
'Situation beschreibt, die ein Abbrechen erfordert, setzt die Routine zusätzlich
'ABBRUCH = True und springt dann zum Routinenende und von da in die aufrufende
'Routine. Jede aufrufende Routine prüft nach Rückkehr und jede 'aufgerufene Routine
'zu Beginn, ob ABBRUCH = True ist und springt in diesem Fall an 'ihr Ende, das
'ähnlich wie bei allen anderen Routinen gebaut ist. Dieses fortgesetzte Überspringen
'setzt sich fort bis zur Hauptroutine (der Shortcut-Aufrufebene).
'Nur diese gibt in ihrem Fertigmeldungsteil einen Text aus und zwar ergänzt um den
'Text in VORGANG und den Text "ausgeführt" bzw. "abgebrochen", je nachdem ABBRUCH
'falsch oder wahr ist. Ist ABBRUCH wahr, wird auch die kumulierte MELDUNG ausgegeben.
'Auf diese Weise führt ein Abbruch genau wie der fehlerfreie Auftragsablauf durch
'alle Ebenen hinauf zum Abschluss in der Hauptroutine.
'==============================================================================
'Benutzte Info-Routinen:
'Sub ABZParam(Aktive Zelle in ArProt)        (im Modul Erfassen enthalten)
'   Stellt alle Parameter einer ArProt-Zeile geprüft und, wenn nötig und
'   möglich, verbessert im Block "ABZ.." bereit
'Sub KontenplanStruktur()                    (im Modul KontenplanPflege)
'   Stellt die (hauptsächlich für Berichte notwendigen) Strukturdaten des
'   Kontenplans geprüft im Block "KP.." zusammen
'Sub KtoKennDat(KontoNr)                       (im Modul KontenplanPflege)
'   Stellt alle Daten eines Kontoblatts auf Struktur geprüft und, wenn möglich
'   und nötig auch korrigiert, im Block "AKto.." zusammen
'Ablaufroutinen:
' Sub Buchen
'   Sub AktKtoBuchen
'     Sub EintragsOrt, Sub Eintragen
' Sub EinträgeLöschen
' Hilfsroutinen

Option Explicit
'_____________________________________________________________________________
'ArProt-Spaltenstruktur  (APC = ArProt-Column):
Public Const APCTANr = 1, APCDatum = 2, APCBeleg = 3, APCSollkto = 4, _
             APCHabenkto = 5, APCText = 6, APCBetrag = 7, APCgebucht = 8, _
             APCBer = 9, APCBuID = 10, APCBetrofKto = 11
Public MeldeStufe As Integer, TiT As String
Public AktBlatt As String, AktKonto As Integer, _
       AktZeile As Integer, AktSpalte As Integer
       
'Für Vorgangsbegleitenden kumulierenden Text
Public ABBRUCH As Boolean, _
       MELDUNG As String, Hinweis As String
       
'Betroffene Konten: Anzeigetext Buchungsstand in ArProt
Public BeKoBlätter As String, BeKoBlatt As String, RestString As String, _
       AnzahlStorni As Integer
 
'Werte der aktuellen Buchungszeile, hauptsächlich von Sub ABZParam
'Buchungszeile-bezogen
Public ABZeilNr As Integer, ABZSpalte As Integer, _
       ABZTADatumText As String, ABZBeleg As String, _
       ABZTADatumTag As Integer, ABZTADatumMonat As Integer, _
       ABZTaNr As Integer, ABZVorTANr As Integer, _
       ABZText As String, ABZBetrag As Double, ABZeichen, ABZeichNächst, _
       ABZNormalBuchg As Boolean, ABZKorrekturBuchg As Boolean, ABZKopierBuchg As Boolean, _
       ABZBeri As String, ABZBuID As Integer, ABZBetrofKto As String, _
       ABZAnzErfordEintr As Integer, ABZAnzAusgefEintr As Integer
'Temporär eines der Konten innerhalb der Buchung
Public ABZAkKto As Integer, ABZAkBlat As String, ABZAkGegKto As Integer, _
       ABZAkSoHa As String, ABZAk3StZ As Integer, ABZAk3StZAend As Integer, _
       ABZAkEintragsZeile As Integer, ABZAkEintrMonat As String, _
       ABZAkAnzGelöcZeil As Integer
'Konten-bezogen (bis zu 4 Konten bei einer Buchung)
Public ABZSoKto As Integer, _
       ABZSoKArt As Integer, ABZSoBlat As String, ABZSoKto3StZ As Integer, _
       ABZSoEintragsZeile As Integer
Public ABZSoSamlKto As Integer, _
       ABZSoSamlKArt As Integer, ABZSoSamlBlat As String, ABZSoSaml3StZ As Integer, _
       ABZSoSamlEintragsZeile As Integer
Public ABZHaKto As Integer, _
       ABZHaKArt As Integer, ABZHaBlat As String, ABZHaKto3StZ As Integer, _
       ABZHaEintragsZeile As Integer
Public ABZHaSamlKto As Integer, _
       ABZHaSamlKArt As Integer, ABZHaSamlBlat As String, ABZHaSaml3StZ As Integer, _
       ABZHaSamlEintragsZeile As Integer
Public ABZStreunKto As Integer, _
       ABZStreunKArt As Integer, ABZStreunBlat As String, ABZStreun3StZ As Integer, _
       ABZStreunEintragsZeile As Integer
Public Meldung1 As String  'verwendet von EinträgeLöschen
'Aus alter Version Buchen
Public Startadresse As Range, ArProZeile As Long, BuID As Long, LetztBuId As Long
Dim Austext As String, DruckBereich As String

'=============================================================================
       Sub ERFASSEN()
  'Die von der aktiven Zelle einer ArProt-Zeile aus durch Tippen von 'Strg+b'
  'spaltenspezifisch zu bewirkenden Aktionen; insbesondere Buchen, Stornieren
  'und Berichteerstellen.
  Dim APSpaltenNr As Integer, NächsteZelle As Range, EndText As String
  Dim SollKtoKz As Integer, HabenKtoKz As Integer, KtoKz As Integer
  Dim KtoArtH As Integer, KtoArtS As Integer, Beschreibung As String
  Dim BlattStelle, AuszugsStelle, ASt, BSt, A, B, APTransaktNr As Integer
  Dim Beleg As Variant
  Dim Rechts2VonBeleg As Variant, TypBel As Variant, R2TypBel As Variant
  Dim BuchZeichen, BetroffeneKonten, VorZeileIstÜberschrift As Boolean
  Dim BeschreibH, BeschreibS, MSt As Integer
  
 With ActiveWindow
    Application.CutCopyMode = False
'    Application.MacroOptions Macro:="ERFASSEN", Description:="", HasMenu:=False, _
'       MenuText:="", HasShortcutKey:=True, ShortcutKey:="b", Category:=14, _
'       StatusBar:="", HelpContextID:="0", HelpFile:=""
'ERF1 ------------------------ Blatt ArProt erzwingen -------------------------
    TiT = ThisWorkbook.Name & "  Modul ERFASSEN  Aktion "
    If ActiveSheet.Name <> "ArProt" Then
      A = MsgBox(prompt:="kann nur vom Blatt ''ArProt'' aus verwendet werden." & _
                       Chr(10) & "Dorthin Wechseln?", _
                Buttons:=vbYesNo, Title:=TiT & "Tastenkombination ''Strg+b''")
      If A = vbYes Then                       'kein Aktivieren von ArProt, wenn
        Worksheets("ArProt").Activate        'mit Abbrechen quittiert wird
        Cells(Cells(1, 1).Value + 2, 2).Activate  'nur zum positionieren
      End If
      If A = vbNo Then
        Exit Sub
      End If
    End If
'ERF2 ------------------------ ArProt Anfangszustand -------------------------
    With Sheets("ArProt")
      AktBlatt = ActiveSheet.Name     'nur zum wiederherstellen
      AktZeile = ActiveCell.Row
      AktSpalte = ActiveCell.Column
      If AktSpalte > 9 Or AktZeile < 3 Then
        Call MsgBox("Strg+b an dieser Stelle wirkungslos.", vbOKOnly, _
                  TiT & "Erfassen")
        Exit Sub
      End If
    End With 'Sheets("ArProt")
'ERF3 -------------- Buchungsjahr, CalTag, MeldeStufr --------------------
    BuchJahr = Sheets("Kontenplan").Cells(1, 5)
    CalTag = Sheets("Kontenplan").Cells(1, 6)
    MeldeStufe = Sheets("ArProt").Cells(1, 7)
'ERF4 ---------- Erinnerung bei Jahreswechsel: Kontenplan in Ordnung? --------------
    Call KontenplanStruktur
    If ABBRUCH = True Then
      GoTo EndeErfassen
    End If
'ERF5 ------------------ Meldestufe feststellen oder ggf. festlegen ------------
'MeldeStufeBestätigen:
'    Dim BeschrMSt As String
'    MeldeStufe = Sheets("ArProt").Cells(1, 7)
'    If MeldeStufe = 1 Then
'      BeschrMSt = "''Keine Ablaufmeldungen''"
'    End If
'    If MeldeStufe > 1 Then
'      BeschrMSt = "''Fertigmeldung enthält Ablaufmeldungen''"
'    End If
'    If Sheets("ArProt").Cells(1, 7) > 0 Then
'      A = MsgBox(prompt:="Die Meldestufe ist " & BeschrMSt & Chr(10) & _
'                       "Ist das so gewünscht?", Buttons:=vbYesNo, _
'                       Title:="Tastenkombination ''Strg+b'' ")
'      If A = vbNo Then
'        If MeldeStufe = 1 Then
'          MeldeStufe = 2
'        Else
'          MeldeStufe = 1
'        End If
'      End If
'      B = MsgBox(prompt:="Diese Frage auch künftig stellen?", _
'          Buttons:=vbYesNo, Title:="Tastenkombination ''Strg+b'' ")
'      If B = vbNo Then
'        Sheets("ArProt").Cells(1, 8) = 0
'      End If 'B = vbno
'      If B = vbYes Then
'        Sheets("ArProt").Cells(1, 8) = 1
'      End If 'B = vbno
'    End If  'Meldestufe gewünscht oder 1
'ERF6 ------------------Erlaubten Eingabebereich erzwingen -------------------
    With Worksheets("ArProt")
      .Activate
      If AktSpalte > APCBer Or AktZeile < 3 Then ' _
         Or AktZeile > Cells(1, 3) + 15 Then    '+15 wegen Kopierbuchungen
        MsgBox ("Tastenkombination ''Strg+b'' hier wirkungslos")
        Exit Sub
      End If
    End With 'Worksheets("ArProt")
  
'ERF7 ==================== Spaltenabhängige Aktionen =========================
ASpalte: ' (TA-Nummer)--------- Buchungszeile stornieren -----------------------
    With Worksheets("ArProt")
      If AktSpalte = APCTANr Then
        Call Stornieren
        Application.CutCopyMode = False
        GoTo EndeErfassen
      Else
        GoTo BSpalte
      End If
    End With 'Worksheets("ArProt")
BSpalte: ' --------------------- TA-Datum hochzählen --------------------------
    With Worksheets("ArProt")
      If AktSpalte <> APCDatum Then GoTo CSpalte
        If ActiveCell.Value = "" Or ActiveCell.Value = " " Then
          If VorZeileIstÜberschrift = True Then
            ActiveCell.Offset(-2, 0).Copy Destination:=ActiveCell
          Else
            ActiveCell.Offset(-1, 0).Copy Destination:=ActiveCell
          End If
        Else 'Zelle nicht leer
          ActiveCell.Value = DatumZT(DatumTZ(ActiveCell.Value) + 1)
 '        If DatRoutFehler > 0 Then Exit Sub
        End If 'ActiveCell.Value = ""
        Exit Sub
      'End If 'aktspalte = APCDatum
    End With 'Worksheets("ArProt")
CSpalte: ' -------------------Belegnr. hochzählen ---------------
    With Worksheets("ArProt")
      If AktSpalte <> APCBeleg Then GoTo DSpalte
        Beleg = ActiveCell.Value
        If Beleg = "" Then
          If VorZeileIstÜberschrift = True Then
            ActiveCell.Offset(-2, 0).Copy Destination:=ActiveCell
          Else
            ActiveCell.Offset(-1, 0).Copy Destination:=ActiveCell
          End If
          Exit Sub
        End If 'Beleg = ""
        If Beleg <> "" Then
          If IsNumeric(Beleg) = True Then
            ActiveCell.Value = Beleg + 1
            Exit Sub
          End If
          If IsNumeric(Right(Beleg, 1)) = True Then 'And _
'           IsNumeric(Left(Beleg, Len(Beleg) - 1)) = False Then
            ActiveCell.Value = Left(Beleg, Len(Beleg) - 1) & CStr(Right(Beleg, 1) + 1)
            GoTo BelegZählEnde
          End If
        End If 'Beleg <> ""
        If TypeName(Beleg) = "String" Then
          BlattStelle = CStr(CInt(Right(Beleg, 1)) + 1)
          If CInt(BlattStelle) > 5 And Left(Rechts2VonBeleg, 1) = "." Then
            BlattStelle = 1
            ASt = Left(Beleg, Len(Beleg) - 2)
            If IsNumeric(Right(ASt, 2)) = True Then
              AuszugsStelle = Right(ASt, 2) + 1
              ActiveCell.Value = Left(Beleg, Len(Beleg) - 4) & AuszugsStelle & "." & _
              BlattStelle
            Else   'Auszugsnummer einstellig
              AuszugsStelle = Right(ASt, 1) + 1
              ActiveCell.Value = Left(Beleg, Len(Beleg) - 3) & AuszugsStelle & "." & BlattStelle
            End If
          Else
            ActiveCell.Value = Left(Beleg, Len(Beleg) - 1) & BlattStelle
          End If
        End If 'TypeName(Beleg) = "String"
BelegZählEnde:
      'End If 'aktspalte = APCBeleg
      Exit Sub
    End With 'Worksheets("ArProt")
DSpalte: ' ------ Ältere Zeile mit gleichem Sollkto hierher kopieren, -------
'                 wenn in Spalte D die KtoNr eines Ausgabektos steht
    Dim ZNr As Integer, StaRow As Integer
    With Worksheets("ArProt")
      If AktSpalte <> APCSollkto Then GoTo ESpalte
      If ActiveCell = "" Then Exit Sub
      Call KtoKennDat(ActiveCell)
      If ABBRUCH = True Then Exit Sub
      If AKtoArt = AusgabKto Or AKtoArt = Ausgab2Kto Then
        StaRow = ActiveCell.Row
        For ZNr = StaRow - 1 To 3 Step -1
          Cells(ZNr, APCSollkto).Activate
          If Cells(ZNr, APCSollkto) = Cells(StaRow, APCSollkto) Then
            Range("D" & ZNr & ":G" & ZNr).Select
            Selection.Copy
            Range("D" & StaRow & ":G" & StaRow).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Cells(StaRow, APCgebucht) = "***"
            Cells(StaRow, APCgebucht).Activate
            Exit For
          End If
        Next ZNr
      End If
    End With
    Exit Sub
ESpalte: ' ---------- Alphanum. HabnKto in numerisches wandeln --------------
'                   wie DSpalte, wenn in Espalte ein EingabeKto
    With Worksheets("ArProt")
      If AktSpalte <> APCHabenkto Then GoTo FSpalte
      If ActiveCell = "" Then Exit Sub
      If IsNumeric(ActiveCell.Value) = True Then GoTo Numerisch
      Call SucheKonto(ActiveCell.Value)   'im Modul M2KontPlan
      Exit Sub
Numerisch:
      Call KtoKennDat(ActiveCell)
      If ABBRUCH = True Then Exit Sub
      If AKtoArt = EingabKto Or _
         AKtoArt = Eingab2Kto Then
        StaRow = ActiveCell.Row
        For ZNr = StaRow - 1 To 3 Step -1
          Cells(ZNr, APCHabenkto).Activate
          If Cells(ZNr, APCHabenkto) = Cells(StaRow, APCHabenkto) Then
            Range("D" & ZNr & ":G" & ZNr).Select
            Selection.Copy
            Range("D" & StaRow & ":G" & StaRow).Select
            ActiveSheet.Paste
            Cells(StaRow, APCgebucht) = "***"
            Exit For
          End If
        Next ZNr
        Application.CutCopyMode = False
        Cells(StaRow, APCBetrofKto).Activate
        Cells(StaRow, APCHabenkto).Activate
      End If
    End With 'Worksheets("ArProt")
    Exit Sub
FSpalte: ' -----------Beschreibung: Verschiedene Textvorschläge --------------
    With Worksheets("ArProt")
      If AktSpalte <> APCText Then GoTo GSpalte
        Beschreibung = Cells(AktZeile, APCText).Value
        If ActiveCell.Value = "" Then
          Call KtoKennDat(Cells(AktZeile, APCSollkto))
          SollKtoKz = AKtoNr
          KtoArtS = AKtoArt
          BeschreibS = AKtoBeschr
          Call KtoKennDat(Cells(AktZeile, APCHabenkto))
          HabenKtoKz = AKtoNr
          KtoArtH = AKtoArt
          BeschreibH = AKtoBeschr
          Exit Sub
        End If
     '--------- HabenKto = Mitglieds- oder Spenderkonto --------------
          If KtoArtH = MitgliedKto Then
            Beschreibung = "Beitrag " & BeschreibH
            Exit Sub
          End If
          If KtoArtH = SpenderKto Then
            Beschreibung = "Spende " & BeschreibH
            Exit Sub
          End If
      '--------- nicht Mitglieds- oder Spenderkonto -----------
          If KtoArtH <> MitgliedKto And KtoArtH <> SpenderKto Then
            If KtoArtH = BestandKto Then
              Beschreibung = BeschreibH
              Exit Sub
            End If
            If KtoArtS = BestandKto Then
              Beschreibung = BeschreibS
              Exit Sub
            End If
            If KtoArtS <> BestandKto And KtoArtH <> BestandKto Then
              Beschreibung = AKtoBeschr
              Exit Sub
            End If
          End If 'KtoArtH <> MitgliedKto
          ActiveCell.Value = Beschreibung
       ' End If 'ActiveCell.Value = ""
      'End If 'aktspalte = APCText
    End With 'Worksheets("ArProt")
GSpalte: ' ---------------- Spalte Betrag: Nur Kopie von Vorzeile ----------------------
    With Worksheets("ArProt")
      If AktSpalte <> APCBetrag Then GoTo HSpalte
        If VorZeileIstÜberschrift = False Then
          ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
        Else
          ActiveCell.Value = ActiveCell.Offset(-2, 0).Value
        End If
        Exit Sub
      'End If
    End With 'Worksheets("ArProt")
HSpalte: '----- Spalte Buchungsdatum: Aufruf von Buchen ----------------
    If AktSpalte <> APCgebucht Then GoTo ISpalte
    Call Buchen
    GoTo EndeErfassen
ISpalte:  '--------- Spalte I (Berichte): Aufruf von Bericht ------------
    If AktSpalte = APCBer And AktZeile > 3 Then
      Call BerichteErstellen 'BerichteErstellen
      GoTo EndeErfassen
    End If
  End With 'ActiveWindow
EndeErfassen: '--------------- Ende-Meldung -----------------------------
  If ABBRUCH = True Then
    EndText = MELDUNG & Chr(10) & Chr(10) & AktVorgang & _
    " ArProt-Zeile " & ABZeilNr & " abgebrochen"
  End If
  If ABBRUCH = False And MeldeStufe > 1 Then
    EndText = MELDUNG & Chr(10) & Chr(10) & AktVorgang & _
    " ArProt-Zeile " & ABZeilNr & " ausgeführt" _
    & Chr(10) & Chr(10) & Hinweis
  End If
  If ABBRUCH = False And MeldeStufe <= 1 Then
    EndText = AktVorgang & " ausgeführt" & _
    Chr(10) & Chr(10) & Hinweis
  End If
  ABBRUCH = False
  Call MsgBox(EndText, vbOKOnly, TiT & "Buchen")
  AktVorgang = ""
End Sub 'Erfassen


Sub Buchen() '============ Setzt Aufruf von ABZParam voraus ===========================
' Buchen einer Arbeitsprotokollzeile, veranlaßt von Shortcut Strg+b von der Spalte H aus.
' Dass die weiteren Voraussetzungen für die Buchung erfüllt sind, stellt der vorausgehende
' Durchlauf des Sub ABZParam sicher.
' Vorgänge: Eintrag der vom Sub ABZParam in ABZ-Block geschriebenen Daten der
' Arbeitsprotokollzeile in die bis zu 4 Kontoblätter
' mittels Finden (mittels Function Kontoblatt) der den Kontokennzahlen in der Arbeits-
' zeile zugehörigen Kontoblätter und Eintrag (mittels Sub Eintrag und zwar dort
' mittels Sub EintragsOrt an der Transaktionsdatum-chronologischen Stelle.
' An einem Monatsübergang wird automatisch ein Summenblock eingefügt mit den Monatssummen
' und den Summen seit Jahresbeginn, und zwar der zum Monatsende des Transaktionsdatums
' gehörige und alle etwa fehlenden vorausgehenden, sodass spätere Buchungen älteren
' Transaktionsdatums richtig eingefügt werden.
' Sonderbehandlungen zu Konto-Einträgen:
' Ist im Kontenplan zu einem angesprochenen Konto in der Spalte "zugeord. SamlKto"
' (Spalte 6) ein Konto vermerkt, so wird parallel zu dem angesprochenen Konto die
' Buchung in dieses Sammelkonto eingetragen.
' Ist ein Konto von der Art 6 (Fondskonto) oder der Art 3 (Durchlaufposten), so wird
' parallel zu ihm ein Bestands-Konto (Kontoart 1) fortgeschrieben und zwar nach
' Maßgabe der durch eine InputBox verlangten Angabe der Realkontonummer.
' Als Durchführungsmeldung werden im Arbeitsprotokoll das aktuelle Datum in die
' Startzelle geschrieben. Außerdem werden in der Arbeitsprotokoll-Zeile die
' betroffenen Konten und eine Buchungs-Ident-Nummer geschrieben und, im Falle einer
' Normalbuchung (***), die Zelle A der nächsten Zeile aktiviert.

  Dim ArProtBlatt As String, DruckBereich As Range
  Dim VerwendetBlatt(1 To 6) As String
  Dim A, B, I As Integer
'  Dim NormalBuchung As Boolean, KorrekturBuchung As Boolean, _
'      ABZKopierBuchg As Boolean
  Dim ErgebnisInputBox As Variant, EingabeDatum As String
  Dim DezimalZeichen, BuchungsZahl As Integer
  
'1 Bu ---------------- Spezialisierung auf die ArProt-Zeile ---------------------
  AktVorgang = "Buchen"
  MELDUNG = ""       'Rücksetzen des kumulierenden Meldestrings
  ABBRUCH = False
'  BuchungsZahl = 1  'Zahl der hintereinander durchzuführenden (Kopier-)Buchungen
  AktZeile = ActiveCell.Row
  AktSpalte = ActiveCell.Column
WeitereZeileBuchen:
  With ActiveWindow
    Sheets("ArProt").Activate
    ArProtBlatt = ActiveSheet.Name  '"ArProt" garantiert das aufrufende ERFASSEN
'2 Bu ----------------- ArProt-Ende-Speicher fortschreiben ----------------------
  If AktZeile > Cells(1, 1) Then  'falls neue Zeile
    Cells(1, 1) = AktZeile
  End If
'3 Bu ------------------------- Druckbereich ArProt festlegen --------------------------
  With ActiveSheet
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$A1:$K" & Cells(1, 1) + 5 & ""
    Cells(1, 1).Select
    Cells(AktZeile, AktSpalte).Select
  End With
'4 Bu ------------------ alle Parameter der ArProt-Zeile-------------------------
  Call ABZParam        'alle Parameter der ArProt-Zeile im globalen ABZ-Block
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
       "ArProtzeile " & ABZeilNr & " vor erneutem Buchungsauftrag sanieren!"
    GoTo EndeBuchen
  End If
  With Sheets("ArProt")
    If ABZeichen = "***" Then ABZNormalBuchg = True
    If ABZeichen = "****" Then ABZKorrekturBuchg = True
    If ABZeichen = "*****" Then ABZKopierBuchg = True
    If IsDate(ABZeichen) = True Or ABZBuID <> 0 Or ABZBetrofKto <> "" Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Diese ArProt-Zeile ist schon gebucht." & Chr(10) & _
      "Vor neuerlichem Buchen erst Stornieren!"
      ABBRUCH = True
      GoTo BuchenEnd
    End If
'2 Bu ---------------- Spalte Betroffene Konten initialisieren -----------------
    Austext = ""
    For I = 1 To 4                   'AusText und alle VerwendetBlatt
      VerwendetBlatt(I) = ""         'vorsorglich leeren
    Next I                           '(wichtig im Falle des Sprungs hierher)
'3 Bu ------------- etwa Vorhandenen Durchstrich beseitigen -------------------
    Cells(ABZeilNr, APCTANr).Range("A1:G1").Select
    With Selection
      .Font.Strikethrough = False
    End With 'Selection
'3 Bu ---------------- Buchungs-Identnummer fortschreiben ---------------------
    LetztBuId = Cells(1, APCBuID).Value + 1
    ABZBuID = LetztBuId
    Cells(1, APCBuID).Value = LetztBuId
    Cells(ABZeilNr, APCBuID).Value = LetztBuId
  End With 'Sheets("ArProt")
'4 Bu ---------------- AktKtoBuchen mit Sollkto-Parameter ----------------------
  If ABZSoKto <> 0 Then
    ABZAkKto = ABZSoKto     'belegen der generischen Parameter mit den Werten
    ABZAkBlat = ABZSoBlat   'des Sollkontos zur Benutzung durch die Subs
    ABZAkGegKto = ABZHaKto  'AktKtoBuchen, Eintragsort, SldoBlockErzeugen,
    ABZAkSoHa = "Soll"        'Eintrag und EintragLöschen
    Call AktKtoBuchen(ABZSoKto, ABZSoBlat, ABZAkGegKto)
    If ABBRUCH = True Then
      GoTo EndeBuchen
    End If
  End If 'ABZSoKto <> 0
'5 Bu ---------------- AktKtoBuchen mit SollSammelkto-Parameter ---------------
SollSamlKtoBuchen:
  If ABZSoSamlKto <> 0 Then
    ABZAkKto = ABZSoSamlKto     'belegen der generischen Parameter
    ABZAkBlat = ABZSoSamlBlat   'mit den Werten des SollSammelkontos
    ABZAkGegKto = ABZHaKto      'zur Benutzung durch die Subs
    ABZAkSoHa = "Soll"            'AktKtoBuchen, Eintragsort, SldoBlockErzeugen
    ABZAk3StZ = ABZSoSaml3StZ     'Eintrag und EintragLöschen
    Call AktKtoBuchen(ABZSoSamlKto, ABZSoSamlBlat, ABZAkGegKto)
    If ABBRUCH = True Then
      GoTo EndeBuchen
    End If
  End If 'ABZSoSamlKto <> 0
'6 Bu ------------------ AktKtoBuchen mit Habenkonto-Parameter ------------------
  If ABZHaKto <> 0 Then
    ABZAkKto = ABZHaKto         'belegen der generischen Parameter
    ABZAkBlat = ABZHaBlat       'mit den Werten des Habenkontos
    ABZAkGegKto = ABZSoKto      'zur Benutzung durch die Subs
    ABZAkSoHa = "Habn"            'AktKtoBuchen, Eintragsort, SldoBlockErzeugen
    ABZAk3StZ = ABZHaKto3StZ    'Eintrag und EintragLöschen
    Call AktKtoBuchen(ABZHaKto, ABZHaBlat, ABZAkGegKto)
    If ABBRUCH = True Then
      GoTo EndeBuchen
    End If
  End If 'ABZHaKto <> 0
'7 Bu ---------- AktKtoBuchen mit HabenSamlkto-Parameter ----------------------
  If ABZHaSamlKto <> 0 Then
    ABZAkKto = ABZHaSamlKto     'belegen der generischen Parameter
    ABZAkBlat = ABZHaSamlBlat   'mit den Werten des Sollkontos
    ABZAkGegKto = ABZSoKto      'zur Benutzung durch die Subs
    ABZAkSoHa = "Habn"            'AktKtoBuchen, Eintragsort, SldoBlockErzeugen
    ABZAk3StZ = ABZHaKto3StZ    'Eintrag und EintragLöschen
    Call AktKtoBuchen(ABZHaSamlKto, ABZHaSamlBlat, ABZAkGegKto)
    If ABBRUCH = True Then
      GoTo EndeBuchen
    End If
  End If 'ABZHaSamlKto <> 0
  GoTo EndeBuchen
'8 Bu -------------- Endmaßnahmen bei abgebrochener Buchung ---------------------
EndeBuchen:
'8.1 Bu ----------------- Meldungstexte bei Abbruch -----------------
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & Chr(10) & _
      "Buchung abgebrochen. Beanstandeten Fehler korrigieren, " & _
      "dann zuerst Buchungszeile Stornieren und danach neu buchen."
'8.2 Bu ----------------- Hinterlassener ArProt-Zustand -------------------------
    With Worksheets("ArProt")
      .Activate
      Cells(ABZeilNr, APCgebucht).Activate
      If Cells(1, 1) < ActiveCell.Row Then
        Cells(1, 1) = ActiveCell.Row  'ArProt-Ende im ArProt-Kopf
      End If
    End With 'Worksheets("ArProt")
    GoTo BuchenEnd
  End If 'ABBRUCH = True
'9 Bu --------------- Endmaßnahmen bei erfolgreicher Buchung -------------------------
    With Worksheets("ArProt")
'9.1 Bu -------------- Datum als Vollzugsvermerk -> ArProt-Zeile ----------------
      .Activate
      Cells(ABZeilNr, ABZSpalte).Activate
      ActiveCell.Value = Date              'Systemdatum
'9.2 Bu ------------------- Austext abrunden -----------------------------------
      Austext = Left(Austext, Len(Austext) - 2)   'letztes "+" im Austext löschen
      Cells(ABZeilNr, APCBetrofKto) = Austext
'9.3 Bu ------------- Jüngstes Transaktionsdatum -->  ArProt A2 ----------------
      Dim LBuDa As String, APLBuDa As String   'Versuch, das ExCel-DatumFormat
      LBuDa = Cells(ABZeilNr, APCDatum)        'zu umgehen
      APLBuDa = Cells(1, 2)
      If DatumTZ(LBuDa) > DatumTZ(APLBuDa) Then
        Cells(1, 2) = Cells(ABZeilNr, APCDatum)   'Speicher jüngstes Transaktonsdatum
      End If
'9.4 Bu ------- Vorbereitung nächste Buchung bei Normalbuchung -------------------
      Dim Mspr As Integer
      Sheets("ArProt").Activate
'9.5 Bu ------------------- Nächste Buchungszeile ------------------------------
      If ABZeichNächst = "***" Then
        Mspr = 1
      End If
      If Cells(ABZeilNr + 1, APCgebucht) = "gebucht" Then   'Monatsüberschrift
        Mspr = 2
      Else
        Mspr = 1
      End If
      If ABZNormalBuchg = True Then
        If Cells(ABZeilNr + Mspr, APCgebucht) = "" And _
          Cells(ABZeilNr + Mspr, APCTANr) = 0 Then
          Range("A" & ABZeilNr + Mspr & ":K" & ABZeilNr + Mspr & "").Select
          Selection.Insert shift:=xlDown
          Cells(ABZeilNr + Mspr, APCTANr) = ABZTaNr + 1  'Nächste TA-Nr.
          Cells(ABZeilNr + Mspr, APCgebucht) = "***"     'Nächstes Buchkennzeichen
          Cells(ABZeilNr + Mspr, APCDatum).Select
'9.6 Bu -------- Fortschreiben der ArProt-Headerzellen A1,A3 ----------------------
          Sheets("ArProt").Cells(1, 1) = _
                        Cells(ActiveCell.Row, 1) 'Nächste TAN -> ArProt-Kopfzelle
          Sheets("ArProt").Cells(1, 3) = Cells(ABZeilNr + Mspr, 1).Row 'nächste Zeile
          GoTo BuchenEnd
        End If
      End If 'ABZNormalBuchg = True
'9.7 Bu ------------ Korrekturbuchung ändert nichts im ArProt-Header --------------
      If ABZKorrekturBuchg = True Then GoTo BuchenEnd
'9.8 Bu ------------ KopierBuchg: Nächste Zeile noch zu Buchen ? ---------------
      If ABZKopierBuchg = True Then
        If Cells(ABZeilNr + Mspr, APCgebucht) = "*****" Then
          Cells(ABZeilNr + Mspr, APCgebucht).Activate
          GoTo WeitereZeileBuchen
        End If
      End If
    End With 'Worksheets("ArProt")
  End With 'ActiveWindow
BuchenEnd:
End Sub 'Buchen
'=================================================================================
Sub Stornieren()
'Sichertellen, dass die Einträge aller von dieser ArProt-Zeile betroffenen Konten,
'einschließlich der etwa vorhandenen Sammelkonten, definiert durch die TA-Nr. , nicht
'mehr vorhanden sind und dann die ArProt-Zeile die korrekte Storno-Form hat.
'Ausführung auch bei anfänglich fehlender BuID und fehlenden betroffenen Konten in
'den Spalten J und K.
  Dim A, B, Titel As String, UndSoSaml As String, UndHa As String, UndHaSaml
 '---------------------- Anfangs- Datenzustand ---------------------------------
  AktVorgang = "Stornieren"
  Titel = TiT & "Stornieren"
  MELDUNG = ""  'Rücksetzen des kumulierenden Meldestrings
  ABBRUCH = False
  Call ABZParam 'Eingangsdaten der ArProt-Zeile mit der aktiven TA-Nr.-Zelle.
                'Liefert alle Daten dieser ArProt-Zeile u. der in ihr genannten
                'Konten in einen Block von Globalvariablen, alle mit ABZ beginnend,ab. Auch bei entdeckten Fehlern stornieren
  If ABBRUCH = True Then GoTo EndeStornieren
'---------------------- Eröffnungsdialog ---------------------------------'
  With Sheets("ArProt")
    .Activate
    If ABZSoSamlBlat <> "" Then            'für MsgBox-Aufzählungen
      UndSoSaml = "''  und  ''" & ABZSoSamlBlat
    End If
    If ABZHaBlat <> "" Then
      UndHa = "''  und  ''" & ABZHaBlat
    End If
    If ABZHaSamlBlat <> "" Then
      UndHaSaml = "''  und  ''" & ABZHaSamlBlat
    End If
    Cells(ABZeilNr, APCTANr).Range("A1:H1").Select
    A = MsgBox("Diese Buchung rückgängig machen? " & Chr(10) & _
    "(d.h. alle Einträge mit der Transaktionsnummer " & ABZTaNr & " aus den" & Chr(10) & _
    "Kontoblättern  ''" & ABZSoBlat & UndSoSaml & UndHa & UndHaSaml & "''," & Chr(10) & _
    "soweit vorhanden, entfernen?)", vbYesNo, Titel)
    If A = vbNo Then
      Cells(ABZeilNr, 9).Activate
      Cells(ABZeilNr, 1).Activate  'zum Löschen der Zeilenhervorhebung
      ABBRUCH = True
      GoTo EndeStornieren
    End If
'--------------------- Löschen verbliebener Einträge -----------------------
LöschAktion:
    Cells(ABZeilNr, 9).Activate   'wozu?
    Cells(ABZeilNr, 1).Activate  'zum Löschen der Zeilenhervorhebung
    Call EinträgeLöschen   '
'----------------------------- ArProt aufräumen --------------------------
    Cells(ABZeilNr, APCTANr).Range("A1:G1").Select
    With Selection
      .Font.Strikethrough = True
    End With 'Selection
    Cells(ABZeilNr, APCBuID).Activate   'löschen der Selection
    Cells(ABZeilNr, APCgebucht) = "****"
    Cells(ABZeilNr, APCgebucht).Activate
    Application.CutCopyMode = False
    GoTo EndeStornieren
EndeStornieren:
    If ABBRUCH = True Then
      Cells(ABZeilNr, APCTANr).Activate
      MELDUNG = MELDUNG & Chr(10) & "Stornieren wurde abgebrochen."
    End If
    If ABBRUCH = False Then
      If MeldeStufe >= 1 Then
        Call MsgBox(MELDUNG, vbOKOnly, TiT)  ' & Tit2)
      End If
    End If
  End With 'Sheets("ArProt")
End Sub 'BuchungZeileStornieren



'=========================Hilfsprogramme für Buchung ===========================
Sub ABZParam()
  'Aufgerufen von ERFASSEN im Falle einer aktiven Buchungsauftragsspalte
  '(Spalte A=APCTaNr: Rückgängig machen und Spalte H=APCgebucht: Buchen)
  'Speichert die Werte der aktiven ArProtzeile in den globalen ABZ-Block
  'einschliesslich abgeleiteter Werte, insbesondere Interna von Kontenblättern.
  'Sammelt die Meldungen über zu korrigierende oder korrigierte Buchzeilenwerte,
  'Bricht jedoch bei Entdeckung unbrauchbarer, weiter benötigter Werte ab.
  'Hinterläßt in ABZurteilGut eine Information, ob mit dem Buchen fortgefahren
  'werden kann, ohne die Buchungszeile oder die betroffenen Konten zu sanieren.
  
  Dim VorZeile As Integer, VorZeilOffset As Integer
  Dim TADatumText As String, DezimalZeichen As String
  Dim ParamAbbruch As Boolean

  With Sheets("ArProt")
    .Activate
    ParamAbbruch = False  'keine Beeinflussung der ABBRUCH-Situation
 '   MELDUNG = ""     'das darf nur das Hauptprogramm (die shortcut-Routine)
    AktZeile = ActiveCell.Row
    AktSpalte = ActiveCell.Column
    '---------------- ABZSpalte nur 1 oder 8 ------------------------
    ABZSpalte = AktSpalte
     If Not (ABZSpalte = 1 Or ABZSpalte = 8) Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Für Buchen muss die Zelle in der H-Spalte," & Chr(10) & _
      "für Stornieren oder TA-NrnOrdnen in der A-Spalte aktiviert sein."
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    '----------    ------ ABZeilNr ------------------------
    ABZeilNr = AktZeile
    '----------------ABZTANr-------------------------------
    ABZTaNr = Cells(ABZeilNr, APCTANr)
     '------------VorTAnr registrieren, (beurteilen im aufrufenden Programm------------
VorausgehendeTaNr:
    If ABZeilNr <= 3 Then
      ABZeilNr = 3
      ABZVorTANr = 0
    End If
    If ABZeilNr > 3 Then
      If Cells(ABZeilNr - 1, APCTANr) = "TA" Then
        VorZeilOffset = 2
      End If
      If IsNumeric(Cells(ABZeilNr - 1, APCTANr)) = True Then
        VorZeilOffset = 1
      End If
      VorZeile = ABZeilNr - VorZeilOffset
      ABZVorTANr = Cells(VorZeile, APCTANr)
   End If
  '--------------------- Ta-Nr-Sequenz normal? ----------------------
   If ABZTaNr <= ABZVorTANr Then
     MELDUNG = MELDUNG & Chr(10) & _
     "Unbrauchbare Ta-Nr. in ArProt-Zeile " & ABZeilNr & "." & Chr(10) & _
     "Vor weiterem Buchungsauftrag Zelle A" & ABZVorTANr & " aktivieren und" & _
     "''Strg+n'' tippen (Ta-Nummern ordnen)"
     ParamAbbruch = True
     GoTo EndeABZParam
   End If
   If Not (ABZTaNr = ABZVorTANr + 1) Then    'gelegentliche Meldung
     Hinweis = Hinweis & Chr(10) & _
     "Die Ta-Nr. in Zeile " & ABZeilNr & " sollte normalerweise " & ABZVorTANr + 1 & " sein. " & Chr(10) & _
     "Empfohlen: Mit Strg+n aus Zelle (" & ABZVorTANr & "," & APCTANr & ") korrigieren, falls" & Chr(10) & _
     "keine für Wiederverwendung vorgesehene Stornozeile vorhanden ist."
    End If
  '-------------- ABZTADatumText, ABZDatumTag und ABZDatumMonat -------------------
   Cells(ABZeilNr, APCDatum).Activate  'Versorgung von Sub DatumSpalten
   Call DatumSpalten               'ergibt Public DatumTag, DatumMonat
   ABZTADatumText = ActiveCell
   If ABBRUCH = True Then
     MELDUNG = MELDUNG & Chr(10) & _
     "Unbrauchbare Datumsangabe in ArProt-Zeile " & ABZeilNr & "."
     ParamAbbruch = True
     GoTo EndeABZParam
   End If
   ABZTADatumTag = DatumTag                '
   ABZTADatumMonat = DatumMonat
  '----------- ABZBeleg nicht als Datum interpretierbar---------------
BuZBeleg:
    ABZBeleg = Sheets("ArProt").Cells(ABZeilNr, APCBeleg)
'------------- ABZ-SoKto, -SoKart, -SoBlat, -SoKto3StZVorBu ------------
    
    ABZSoKto = Sheets("ArProt").Cells(ABZeilNr, APCSollkto)
    If ABZSoKto = 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Sollkonto fehlt"
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    If IsNumeric(ABZSoKto) = False Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Sollkonto unzulässig (nichtnumerisch)"
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    ABZAnzErfordEintr = 1
SoKKDAufruf:
    Call KtoKennDat(ABZSoKto)   'prüft auch Format des Sollkontos
      If ABBRUCH = True Then
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
    If AKtoStatus = KtoUnbekannt Then         'KtoUnbekannt = 0
      MELDUNG = MELDUNG & Chr(10) & _
      "Konto " & ABZSoKto & " unbekannt." & Chr(10) & _
      "Wenn kein Schreibfehler, dann mit ''Strg+k / EINFÜGEN'' im Kontenplan ergänzen."
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    If AKtoStatus = KtoBlattFehlt Then        'KtoBlattFehlt = 1
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenblatt  " & AKtoBlatt & " für Konto " & ABZSoKto & " wird eingerichtet."
'      Sheets("Kontenplan").Cells(AKtoKPZeil, 2).Activate 'kann nicht ausgeführt werden. Warum?
      Call KontoBlattEinrichten(ABZSoKto)   'mit Daten aus dem vorangegangenen KtoKennDat
      GoTo SoKKDAufruf
    End If
    If ABBRUCH = True Then   'Urteil von KtoKennDat
      Sheets(AKtoBlatt).Activate
      MELDUNG = MELDUNG & Chr(10) & MELDUNG & Chr(10) & _
       "Nicht behobener Strukturfehler in Konto " & ABZSoKto & "." & Chr(10) & _
       "Abbruchgrund. Vor Wiederholung Konto sanieren!"
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    ABZSoKto3StZ = AKto3SternZeile
    ABZSoKArt = AKtoArt
    ABZSoBlat = AKtoBlatt
'----------------ABZ-SoSamlKto, -SoSamlKart, -SoSamlBlat ----------
    ABZSoSamlKto = AKtoSamlKto
    If ABZSoSamlKto = 0 Then
      ABZSoSamlKArt = 0
      ABZSoSamlBlat = ""
    Else
      ABZAnzErfordEintr = ABZAnzErfordEintr + 1
SoSamlKKDAufruf:
      Call KtoKennDat(ABZSoSamlKto)
      If ABBRUCH = True Then
        MELDUNG = MELDUNG & Chr(10) & "Konto " & ABZSoSamlKto & ""
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
      If AKtoStatus = KtoUnbekannt Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Konto " & ABZSoSamlKto & " unbekannt." & Chr(10) & _
        "Wenn kein Schreibfehler, dann mit ''Strg+k / EINFÜGEN"" im Kontenplan ergänzen."
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
      If AKtoStatus = KtoBlattFehlt Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Kontenblatt  " & AKtoBlatt & " für Konto " & ABZSoSamlKto & " wird eingerichtet."
 '       Sheets("Kontenplan").Cells(AKtoKPZeil, 2).Activate
        Call KontoBlattEinrichten(ABZSoSamlKto)   'mit Daten aus dem vorangegangenen KtoKennDat
        GoTo SoSamlKKDAufruf
      End If
      If ABBRUCH = True Then   'Urteil von KtoKennDat
        MELDUNG = MELDUNG & Chr(10) & _
        "Nicht behobener Strukturfehler in Konto " & ABZSoSamlKto & "." & Chr(10) & _
        "Auftrag abgebrochen. Vor Wiederholung Konto sanieren!"
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
      ABZSoSaml3StZ = AKto3SternZeile
      ABZSoSamlKArt = AKtoArt
      ABZSoSamlBlat = AKtoBlatt
    End If
'---------------ABZ-HaKto, -HaKart, -HaBlat ----------------------
    ABZHaKto = Sheets("ArProt").Cells(ABZeilNr, APCHabenkto)
    If ABZHaKto = 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Habenkonto fehlt"
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
      
    If IsNumeric(ABZHaKto) = False Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Habenkonto unzulässig (nichtnumerisch)"
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    ABZAnzErfordEintr = ABZAnzErfordEintr + 1
HaKKDAufruf:
    Call KtoKennDat(ABZHaKto)
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Prüfen der ArProt-Buchungszeile " & ABZeilNr & ". Struktur des Ktos " & ABZHaKto & " fehlerhaft."
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    If AKtoArt = 6 Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Konto " & ABZHaKto & " ist ein nachrichtliches Konto und ist deshalb" & _
      "als Habenkonto unzulässig. Buchung wird nicht akzeptiert."
      If AktVorgang <> "STORNIEREN" Then  'Im falle Stornieren nicht abbrechen
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
    End If
    If AKtoStatus = KtoUnbekannt Then   'KtoUnbekannt = 0
      MELDUNG = MELDUNG & Chr(10) & _
      "Konto " & ABZHaKto & " unbekannt." & Chr(10) & _
      "Wenn kein Schreibfehler, dann mit ''Strg+k / EINFÜGEN'' im Kontenplan ergänzen."
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    If AKtoStatus = KtoBlattFehlt Then   'KtoBlattFehlt = 1
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenblatt  " & AKtoBlatt & " für Konto " & ABZHaKto & " wird eingerichtet."
      With Sheets("Kontenplan")
        Cells(AKtoKPZeil, 2).Activate
      End With
      Call KontoBlattEinrichten(ABZHaKto)  'mit Daten aus dem vorangegangenen KtoKennDat
      GoTo HaKKDAufruf
    End If
    If ABBRUCH = True Then   'Urteil von KtoKennDat
      MELDUNG = MELDUNG & Chr(10) & _
      "Nicht behobener Strukturfehler in Konto " & ABZHaKto & "."
      ParamAbbruch = True
      GoTo EndeABZParam
    End If
    ABZHaKto3StZ = AKto3SternZeile
    ABZHaKArt = AKtoArt
    ABZHaBlat = AKtoBlatt
'--------------ABZ-HaSamlKto, -HaSamlKart, -HaSamlBlat -----------
    ABZHaSamlKto = AKtoSamlKto
    If ABZHaSamlKto = 0 Then
      ABZHaSamlKArt = 0
      ABZHaSamlBlat = ""
    Else
      ABZAnzErfordEintr = ABZAnzErfordEintr + 1
HaSamlKKDAufruf:
      Call KtoKennDat(ABZHaSamlKto)
      If AKtoStatus = KtoUnbekannt Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Konto " & ABZHaSamlKto & " in ArProt-Zeile " & ABZeilNr & _
        " unbekannt." & Chr(10) & _
        "Wenn kein Schreibfehler, dann mit ''Strg+k / EINFÜGEN'' " & _
        "im Kontenplan ergänzen!"
        ABBRUCH = True
        ParamAbbruch = True
        GoTo EndeABZParam
      End If
      If AKtoStatus = KtoBlattFehlt Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Kontenblatt  " & AKtoBlatt & " für Konto " & ABZHaSamlKto & " wird eingerichtet."
 '       Sheets("Kontenplan").Cells(AKtoKPZeil, 2).Activate
        Call KontoBlattEinrichten(ABZHaSamlKto)  'mit Daten aus dem vorangegangenen KtoKennDat
        GoTo HaSamlKKDAufruf
      End If
      If ABBRUCH = True Then   'Urteil von KtoKennDat
        MELDUNG = MELDUNG & Chr(10) & _
        "Nicht behobener Strukturfehler in Konto " & ABZHaSamlKto & "."
        ParamAbbruch = True
        GoTo Text
      End If
      ABZHaSaml3StZ = AKto3SternZeile
      ABZHaSamlKArt = AKtoArt
      ABZHaSamlBlat = AKtoBlatt
    End If 'ABZHaSamlKto = 0
    '---------------- ABZText -----------------------------
Text:
    ABZText = Cells(ABZeilNr, APCText)
    '-------------- ABZBetrag mit Kommasicherung ----------------------
    Dim BuZBetrag, Vorkomma As String, Nachkomma As String, LängeBB As Integer
    ABZBetrag = Sheets("ArProt").Cells(ABZeilNr, APCBetrag)
    LängeBB = Len(ABZBetrag)
    If LängeBB < 4 Or IsNumeric(ABZBetrag) = True Then
      GoTo EndeKommaPrüfung  'ABZBetrag ohne Komma geschrieben
    End If
      DezimalZeichen = Left(Right(Cells(ABZeilNr, APCBetrag).Value, 3), 1)
      If DezimalZeichen <> "," Then
        BuZBetrag = ABZBetrag
        LängeBB = Len(BuZBetrag)
        Vorkomma = Left(BuZBetrag, LängeBB - 3)
        Nachkomma = Right(BuZBetrag, 2)
        BuZBetrag = Vorkomma & "," & Nachkomma
        ABZBetrag = BuZBetrag
        Cells(ABZeilNr, APCBetrag) = ABZBetrag
        MELDUNG = MELDUNG & Chr(10) & _
          "Betrag in Zelle (" & ABZeilNr & "," & APCBetrag & ") war mit " & _
          "Punkt statt Komma geschrieben." & Chr(10) & _
          "Wurde verbessert."
       End If
EndeKommaPrüfung:
    '---------------- ABZeichen -----------------------------
    ABZeichen = Sheets("ArProt").Cells(ABZeilNr, APCgebucht)
        If ABZeichen = "***" Then ABZNormalBuchg = True
        If ABZeichen = "****" Then ABZKorrekturBuchg = True
        If ABZeichen = "*****" Then ABZKopierBuchg = True
    '---------------- ABZeichen nächste Zeile -----------------------------
    ABZeichNächst = Sheets("ArProt").Cells(ABZeilNr + 1, APCgebucht)
      '----------------- ABZBeri -------------------------------
    ABZBeri = Sheets("ArProt").Cells(ABZeilNr, APCBer)
    '----------------- ABZBuID -------------------------------
    ABZBuID = Sheets("ArProt").Cells(ABZeilNr, APCBuID)
   '----------------- ABZBetrofKto Gesamttext ---------------------------
    ABZBetrofKto = Sheets("ArProt").Cells(ABZeilNr, APCBetrofKto)
   '-------------------- Vollständigkeit -----------------------
    If ABZSpalte = APCgebucht Then  'Bei Buchungs-Auftrag
      With Sheets("ArProt")
        If ABZTaNr = 0 Or IstEinDatum(Cells(ABZeilNr, APCDatum)) = False Or _
          ABZBeleg = "" Or ABZSoKto = 0 Or ABZHaKto = 0 Or ABZText = "" Or _
          ABZBetrag = 0 Or ABZeichen = "" Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Buchungszeile " & ABZeilNr & " unvollständig. Auftrag abgebrochen." & _
          Chr(10) & "Vor Wiederholung Buchungszeile ergänzen!"
          ParamAbbruch = True
          GoTo EndeABZParam
        End If
   '--------------------- Unvollständig gelöschte Buchung ------------------------
        If (ABZeichen = "***" Or ABZeichen = "****" Or ABZeichen = "*****") _
           And (ABZBuID > 0 Or ABZBetrofKto <> "") Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Eine vorausgehende Buchung dieser ArProt-Zeile wurde " & Chr(10) & _
          "unvollständig abgebrochen und unvollständig gelöscht." & Chr(10) & _
          "Vor neuerlicher Buchung Löschen beauftragen mit" & Chr(10) & _
          "aktivieren der Ta-Nr.-Zelle und Tippen von Strg+b ! "
          ParamAbbruch = True
          GoTo EndeABZParam
        End If
      End With 'Sheets("ArProt")
    End If 'ABZSpalte = APCgebucht
    GoTo EndeABZParam
EndeABZParam:
    If ParamAbbruch = True Then
      ABBRUCH = True
      MELDUNG = "Das Prüfen der Daten von ArProt-Zeile " & ABZeilNr & _
      "  hat folgenden Abbruchgrund ergeben" & Chr(10) & MELDUNG
    End If
    Sheets("ArProt").Activate
    Cells(ABZeilNr, ABZSpalte).Activate
  End With '?
End Sub 'ABZParam

'=========================================================================================
Sub AktKtoBuchen(Konto As Integer, Blatt As String, GegenKto As Integer)
'Findet (mittels Sub EintragsOrt) den gemäß ABZTaDatumMonat und ABZTaDatumTag
'chronologisch richtigen Eintragsort im Kontoblatt "Blatt" für das "Konto"

Dim ZeilenNr As Long, A
  Call KtoKennDat(Konto)
    If AKtoStatus < KtoGanzLeer Then
      Worksheets("Kontenplan").Cells(ABZeilNr, 2).Activate
      Call KontoBlattEinrichten(Konto)
      Blatt = AKtoBlatt
    End If
  With Worksheets(Blatt)
    .Activate
    Call EintragsOrt(ABZAkKto, ABZAkBlat) 'Liefert ABZAkEintragsZeile
    ZeilenNr = ABZAkEintragsZeile
    If ABBRUCH = True Then
      Worksheets("ArProt").Activate
      Cells(ABZeilNr, ABZSpalte).Activate
      MELDUNG = MELDUNG & Chr(10) & _
      "Kein Eintrag in das Konto " & Konto & " getätigt"
      ABBRUCH = True
      GoTo EndeAktKtoBuchen
    End If
    Call Eintrag(Blatt, GegenKto)
    If ABBRUCH = True Then GoTo EndeAktKtoBuchen
    Worksheets(Blatt).Activate
    MELDUNG = MELDUNG & Chr(10) & "Konto ''" & Blatt & "'' gebucht."
  End With 'Worksheets(Blatt)
EndeAktKtoBuchen:
End Sub 'AktKtoBuchen

'==================================================================================
Sub EintragsOrt(Konto As Integer, Blatt As String)
  'out: Verändert Public-Variable: ABZAkEintragsZeile, ABZAbbruch
  'Von einem Eintragsvorhaben mit dem in ABZParam in ABZTaDatumTag und
  'ABZTaDatumMonat gespaltenen Transaktionsdatum wird anhand der Monats-
  'Zahl der Bereich zwischen zwei Saldoblöcken im Kontoblatt gesucht,
  'in dem der Eintrag geschehen muss, und anschliessend in diesem Bereich
  'anhand der Tageszahl die chronologisch richtige Eintragszeile.
  'Wenn nötig, wird ein neuer Monatsbereich mit einem Saldoblock erzeugt.
  'ABZAkEintragsZeile ermittelt. Eintragsort liefert nur die Information.
  'Der Eintrag selbst wird im Sub Eintrag durch eine insert-shift-down-Aktion
  'vorgenommen.
  'Eintragsort setzt voraus, dass die Struktur des Kontenblatts von der
  'Info-Routine KtoKennDat geprüft ist.
  
  Dim AnfSheet As String, AnfRow As Integer, AnfColumn As Integer
  Dim Prüfzelle As Range, A, PrüfVar As Variant, PZv As Integer, N As Integer
  Dim Saldierart As String
  Dim DSZ As Integer, DreiSternZeile As Integer, DatumsZahl As Integer
  Dim FolgeMonatLetzterT As String, FolgeMonatLetzterZ As Integer
 ' Dim BremsDatum As Integer, StartZeilenInkrement As Integer
  Dim KoMonatAnfZ As Integer, KoMonatEndZ As Integer, MonZimSalBlock As Integer
'
  AnfSheet = ActiveSheet.Name
  AnfRow = ActiveCell.Row
  AnfColumn = ActiveCell.Column
'1 EO-------------- Ausgangswerte vor Bereichswerten (Januar) ----------------
  'ABBRUCH = False ' darf eine shortcut-Routine nicht
  With Sheets(Blatt)
    .Activate
'2 EO ----------------- Geprüfte 3Sternzeile des Kontos ----------------------
    Call KtoKennDat(Konto) '(Stellt seine Aufrufsituation wieder her)
    If AKtoArt = KtoUnbekannt Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Konto ''" & Konto & "'' unbekannt." & Chr(10) & _
      "Wenn kein Schreibfehler, mit Strg+k/EINFÜGEN im Kontenplan ergänzen."
      ABBRUCH = True
      GoTo EndeEintragsort
    End If
    ABZAk3StZ = AKto3SternZeile
    Cells(5, KoCBeleg).Activate   'Datumsspalte der ersten Buchungseintrags-Zeile
 '3 EO --------------------- Monatssuchschleife --------------------------
    KoMonatAnfZ = 5
MonatsSchleife:
    For DSZ = KoMonatAnfZ To ABZAk3StZ
      Cells(DSZ, KoCBeleg).Activate
      Set Prüfzelle = ActiveCell
      KoMonatEndZ = Prüfzelle.Offset(0, 0).Row
      If Prüfzelle.Offset(1, 2) = "Umsatz Periode" And _
         Prüfzelle.Offset(2, 2) = "Kontostand" Then
        KoMonatEndZ = Prüfzelle.Offset(0, 0).Row      'Zeile vor Saldoblock
        MonZimSalBlock = Prüfzelle.Offset(1, 0)         'durchsuchter Monat als Zahl
        If MonZimSalBlock = ABZTADatumMonat Then        'gesuchter Monat
          KoMonatEndZ = Prüfzelle.Offset(0, 0).Row
          GoTo TagesSchleife  'mit KoMonatAnfZ To KoMonatEndZ
        End If
        If MonZimSalBlock < ABZTADatumMonat Then   'nächsten Monat
          If Prüfzelle.Offset(0, -1) <> "***" Then  'Nachfolgemonat vorhanden
            KoMonatAnfZ = KoMonatEndZ + 3
  '         KoMonatEndZ = Prüfzelle.Offset(0, 0).Row
            GoTo MonatsSchleife      'nächsten Monat ab neuem KoMonatAnfZ untersuchen
          End If
 '---------------'Neuer Saldoblock Umrisse mit Leerzele vor und nach:
          If Prüfzelle.Offset(0, -1) = "***" Then    'neuen Saldoblock erzeugen
            Range("A" & DSZ & ":H" & DSZ + 3 & "").Select
            Selection.Copy
            Range("A" & DSZ + 4 & ":H" & DSZ + 7 & "").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Range("B" & DSZ + 3 & ":D" & DSZ + 8 & "").Select
            With Selection
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlBottom
            End With
            Range("F" & DSZ + 3 & ":H" & DSZ + 6 & "").Select
            Selection.NumberFormat = "#,##0.00"
  '------------------ Formeln für den neuen Saldoblock --------------------            'Enddatum des angehängten Monats (3sternzelle-bezogen)
            Cells(DSZ + 5, 3) = MonZimSalBlock + 1
            Cells(DSZ + 6, 3) = MonatsLetzText(MonZimSalBlock + 1)
            Range("H" & DSZ + 5 & "").Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
          'Haben-Stand Monatsende
            Range("H" & DSZ + 6 & "").Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)"
          'Soll-Umsätze im Monat
            Range("G" & DSZ + 5 & "").Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)"
          'Soll-Stand Monatsende
            Range("G" & DSZ + 6 & "").Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)"
          'Differenz Soll-Haben im Monat
            Range("F" & DSZ + 5 & "").Select
            ActiveCell.FormulaR1C1 = "=RC[+1]-RC[+2]"
          'Saldo
            Range("F" & DSZ + 6 & "").Select
            ActiveCell.FormulaR1C1 = "=R[-4]C+R[-1]C"
          'Dreisternzeile fortschreiben
            Cells(DSZ, KoCDatum) = ""
            Cells(1, 1) = DSZ + 4
            ABZAk3StZ = DSZ + 4    'ab hier neue DSZ !
            KoMonatAnfZ = ABZAk3StZ - 1
            GoTo MonatsSchleife   '---> Bedingung für weitersuchen hergestellt
          End If 'Prüfzelle.Offset(0, -1) = "***"    'ermöglicht mehrere neue Saldoblöcke
        End If 'KoMonatEndZ < ABZTADatumMonat
      End If 'Prüfzelle.Offset(1, 2) = "Umsatz Periode"
      KoMonatEndZ = Prüfzelle.Offset(0, 0).Row
    Next DSZ
TagesSchleife:
    If Cells(KoMonatAnfZ, KoCDatum) = "" Then
  '    KoMonatAnfZ = KoMonatAnfZ + 1
    End If
    For DSZ = KoMonatAnfZ To KoMonatEndZ
      ABZAkEintrMonat = Right(Cells(KoMonatEndZ + 2, 3), 3)
      Cells(DSZ, KoCDatum).Activate
      If Cells(DSZ, KoCDatum) = "***" Or _
         (Cells(DSZ, KoCDatum) = "" And _
         Cells(DSZ + 1, 5) = "Umsatz Periode" And _
         Cells(DSZ + 2, 5) = "Kontostand") Then
        ABZAkEintragsZeile = DSZ   'Eintragsort ist DreisternZeile oder letzte
        GoTo EndeEintragsort         'Zeile im Monatsbereich
      End If
      If ActiveCell = "" Then GoTo NDSZ
      Call DatumSpalten    'liefert Tag in DatumTag und Monat in DatumMonat
      If ABBRUCH = True Then GoTo EndeEintragsort
      If DatumTag > ABZTADatumTag Or DSZ = KoMonatEndZ Then
        ABZAkEintragsZeile = DSZ               '---> Eintragszeile gefunden
        Exit For
      End If
NDSZ:
    Next DSZ
  End With 'Worksheets(Blatt)
EndeEintragsort:
  If ABBRUCH = True Then
    If MeldeStufe >= 1 Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Das Ermitteln des Eintragsortes im Kontenblatt " & ABZAkBlat & "  wurde abgebrochen."
    End If
  End If
  If ABBRUCH = False Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Als Eintragsort in Blatt " & ABZAkBlat & " wurde Zeile " & ABZAkEintragsZeile & _
    " im Monatsbereich " & ABZAkEintrMonat & " ermittelt"
  End If
End Sub 'EintragsOrt

Sub Eintrag(ByVal Blat As String, GegenKonto As Integer)
  'Fügt in dem ausgewählten Kontoblatt "Blat" an der durch "KtoZeile" gegebenen
  'Stelle eine Leerzeile ein und kopiert die Werte der aktuellen ArProt-Zeile
  'dort hinein. Führt die Dreisternzeile (1,1) nach, vermerkt das Blatt in der
  'ArProt-Spalte "Betroffene Konten"
  Dim A, KontoNamen As String, DreiSternZeile As Integer, SoH As String
  Dim KtoZeile As Integer
With ActiveWindow
 ' ABBRUCH = False  'darf nur das Hauptprogramm
'1 Ei--------------------------- Kontoeintrag vorbereiten ------------------------
  With Worksheets(Blat)
    .Activate
      KtoZeile = ABZAkEintragsZeile   'von EintragsOrt ermittelt
      Cells(KtoZeile, 1).EntireRow.Insert shift:=xlDown
      Cells(1, 1).Value = Cells(1, 1).Value + 1 '***-Zeilenstand weiterzählen
      KontoNamen = Cells(1, 5).Value          'nur für Meldetext verwendet
      ABZAk3StZ = Cells(1, 1).Value
      ABZAkEintragsZeile = KtoZeile
      MELDUNG = MELDUNG & Chr(10) & _
      "Daten der ArProt-Zeile " & ABZeilNr & " wurden in Blatt " & ABZAkBlat & _
      ", Zeile " & ABZAkEintragsZeile & " im Monatsbereich " & ABZAkEintrMonat & " eingetragen"
    With Worksheets("ArProt")
      .Activate
      Austext = Austext & Blat & " + "
      Cells(ABZeilNr, APCBetrofKto) = Austext
    End With
'2 Ei------------------- Arbeitsprotokollzeile übertragen ------------------------
'   Kontoblatt-Spaltenstruktur (im Modul Kontenplanpflege definiert):
'   Public Const KoCTANr=1, KoCDatum=2, KoCBeleg=3, KoCBlockDatum=3, KoCGegKto=4, _
'             KoCBeschr=5, KoCSaldo=6, KoCSoll=7, KoCHaben=8, KoCBuID=9
    With Sheets(Blat)
      .Activate
      Cells(KtoZeile, KoCTANr) = ABZTaNr
      Cells(KtoZeile, KoCDatum) = ABZTADatumText
      Cells(KtoZeile, KoCBeleg) = ABZBeleg
      Cells(KtoZeile, KoCGegKto) = GegenKonto
      Cells(KtoZeile, KoCBeschr) = ABZText
      If ABZAkSoHa = "Soll" Then
        Cells(KtoZeile, KoCSoll) = ABZBetrag
      End If
      If ABZAkSoHa = "Habn" Then
        Cells(KtoZeile, KoCHaben) = ABZBetrag
      End If
      Cells(KtoZeile, KoCBuID) = ABZBuID
      Cells(KtoZeile, KoCTANr).Range("A1:H1").Select
      With Selection                              'Im Falle einer Korrekturbuchung
        .Font.Strikethrough = False               'ist die mitkopierte Durchstreichung,
        .Borders(xlEdgeTop).LineStyle = xlNone    'ausserdem etwa vorhandene
        .Borders(xlEdgeBottom).LineStyle = xlNone 'Rahmen im Kontoblatt zu beseitigen
      End With
    End With 'Sheets(Blat)
'4 Ei -------------------- Bedingte Eintragsmeldung -------------------------------
    With Sheets("ArProt")
      .Activate                     'praktische Positionierung
      Cells(ABZeilNr, 6).Activate   'auf dem Bildschirm
    End With
  End With 'Worksheets(Blat)
End With 'ActiveWindow
GoTo EndeEintrag
EndeEintrag:
End Sub 'Eintrag

'========================================================== im Modul ERFASSEN ===
Sub EinträgeLöschen()  'Daten von Sub ABZParam
  'Versucht jeden Eintrag unter der aktuellen Transaktionsnummer in den in
  'der Soll-, Haben- und Sammelkonto-Spalten der beteiligten Konten zu
  'löschen, selbst wenn er in einem Konto mehrmals vorhanden ist. Stützt
  'sich auf die Informationen des vorweg vom letzten ABZParam-Aufruf
  'gefüllten ABZ-Blocks und für die Blätter auf die jeweiligen KtoKennDat-
  'Aufrufe. Löscht in allen (bis zu 4) in der ABZeilNr angesprochenen
  'Kontoblättern mittels Sub AktKtoZeileLöschen alle noch zur
  'Transaktionsnummer gehörenden Zeilen, soweit sie vorhanden sind.
  'Zum Schluss wird noch geprüft, ob in der ArProt-
  'Spalte Betrofkto noch Konten vermerkt sind, die möglicherweise das Löschen
  'einer Zeile erfordern.
  'Nichtvorhandensein einer Zeile mit der TA-Nr. ist kein Abbruchgrund. So
  'kann ein konsistenter Zustand nach vorhergehenden unvollständigen
  'Buchungen oder Überbuchungen wiederhergestellt werden.
  'Die Daten der Buchung sind dann nur noch im ArProt bekannt.
  'Betroffene Konten und BuID werden in der Arprot-Zeile erst bei vollständigem
  'EinträgeLöschen-Durchlauf vollständig gelöscht.
  'Die Konsistenz der Kontoblätter wird von Sub AktKtoZeileLöschen sichergestellt.

  Dim A, I As Integer, Z As Integer
  Dim AbortKorrigieren As Boolean
  Dim Titel As String
  Dim UndSoSaml As String, UndHa As String, UndHaSaml As String
  
  Titel = TiT & "  Einträge Löschen "
  With Sheets("ArProt")
    .Activate
    If ABZeichen = "****" And ABZBuID = 0 And ABZBetrofKto = "" Then
      A = MsgBox( _
      "Die Buchung Ta-Nr. " & ABZTaNr & " ist schon storniert." & Chr(10) & _
      "Soll trotzdem noch einmal Löschen wiederholt werden?", vbYesNo, Titel)
      If A = vbNo Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Die Buchung Ta-Nr. " & ABZTaNr & " war schon storniert." & Chr(10) & _
        "Hier kein weiteres Löschen unternommen."
        GoTo ArProtAufräumen
      End If
      If A = vbYes Then GoTo KontenEinträgeLöschen
    End If
    If ABZeichen = "***" And ABZBuID = 0 And ABZBetrofKto = "" Then
      A = MsgBox( _
      "Die Buchung Ta-Nr. " & ABZTaNr & " ist noch nicht geschehen." & Chr(10) & _
      "Soll trotzdem das Fehlen von Einträgen unter der TA-Nr. " & ABZTaNr & _
      " sichergestellt werden?", vbYesNo, Titel)
      If A = vbNo Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Die Buchung Ta-Nr. " & ABZTaNr & " ist noch nicht geschehen." & Chr(10) & _
        "Keine Löschung durchgeführt."
        GoTo ArProtAufräumen
      End If
      If A = vbYes Then GoTo KontenEinträgeLöschen
    End If
  End With 'Sheets("ArProt")
  'Bei ABZBuID <> 0 kann Zeilensuche mittels BuID geschehen
  'Hier müssten die betroffenen Konten gesund sein (KtoKennDat-Prüfung)

KontenEinträgeLöschen:
'1 So -------------------- Sollkonto-Eintrag zurücksetzen -------------------------
  If ABZSoKto <> 0 Then        'Aktuelle Parameter für KontoZeileLöschen
    Call KtoKennDat(ABZSoKto)
    ABZAkKto = ABZSoKto      'vom SoKto
    ABZAkBlat = ABZSoBlat
    ABZAkGegKto = ABZHaKto
    ABZAkSoHa = "Soll"
    ABZAk3StZ = ABZSoKto3StZ
    Call AktKtoZeileLöschen    'Löschen der Buchungszeile
    If ABBRUCH = True Then GoTo EndeEinträgeLöschen
  End If 'ABZSoKto <> 0
'2 SoSaml ----------------- SoSamlKto-Eintrag zurücksetzen -------------------------
  If ABZSoSamlKto <> 0 Then  'Aktuelle Parameter für AktKtoZeileLöschen
    Call KtoKennDat(ABZSoSamlKto)
    ABZAkKto = ABZSoSamlKto      'vom SoSamlKto
    ABZAkBlat = ABZSoSamlBlat
    UndSoSaml = "'' und ''" & ABZSoSamlBlat  'für Meldetext
    ABZAkGegKto = ABZHaKto
    ABZAkSoHa = "Soll"
    ABZAk3StZ = ABZSoSaml3StZ
    Call AktKtoZeileLöschen          'Löschen einer Kontozeile Buchung
    If ABBRUCH = True Then GoTo EndeEinträgeLöschen
   End If
'3 Ha -------------------- Habenkonto-Eintrag zurücksetzen -------------------------
  If ABZHaKto <> 0 Then    'Aktuelle Parameter für AktKtoZeileLöschen
    Call KtoKennDat(ABZHaKto)
    ABZAkKto = ABZHaKto      'vom HaKto
    ABZAkBlat = ABZHaBlat
    Sheets(ABZAkBlat).Activate
    UndHa = "'' und ''" & ABZHaBlat    'für Meldetext
    ABZAkGegKto = ABZSoKto
    ABZAkSoHa = "Habn"
    ABZAk3StZ = ABZHaKto3StZ
    Call AktKtoZeileLöschen    'Löschen der Buchungszeile
    If ABBRUCH = True Then GoTo EndeEinträgeLöschen
  End If
'4 HaSaml -------------Haben-Sammelkto-Eintrag zurücksetzen -----------------------
  If ABZHaSamlKto <> 0 Then        'Aktuelle Parameter für AktKtoZeileLöschen
    Call KtoKennDat(ABZHaSamlKto)
    ABZAkKto = ABZHaSamlKto      'vom HaSamlKto
    ABZAkBlat = ABZHaSamlBlat
    UndHaSaml = "'' und ''" & ABZHaSamlBlat
    ABZAkGegKto = ABZHaKto
    ABZAkSoHa = "Habn"
    ABZAk3StZ = ABZHaSaml3StZ
    Call AktKtoZeileLöschen    'Löschen der Buchungszeile
    If ABBRUCH = True Then GoTo EndeEinträgeLöschen
  End If
'5 -------------------- Noch ein Eintrag zu löschen? --------------------
  If Austext <> "" Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Es konnten möglicherweise nicht alle zur TA " & ABZTaNr & " gehörenden Einträge" & _
    "gelöscht werden." & Chr(10) & _
    "Manuell prüfen, ob im Blatt/in denBlättern " & Austext & " noch Reste der TA " & _
    ABZTaNr & " vorhanden sind!"
    ABBRUCH = True
    GoTo EndeEinträgeLöschen
  End If
'KontoNrSuchen:
'    With Sheets("Kontenplan")
'      .Activate
'      For Z = 6 To Cells(1, 3)
'        If Cells(Z, KPCBlattname) = Austext Then
'          ABZStreunKto = Cells(Z, KPCKonto)
'          GoTo KontoNrGefunden
'        End If
'      Next Z
'      MELDUNG = MELDUNG & Chr(10) & _
'      "KontoNr zum Blatt " & ABZBetrofKto & " im Kontenplan nicht gefunden" & _
'      "Manuell prüfen, ob noch Reste der TA " & ABZTaNr & " in " & ABZBetrofKto & " vorhanden sind!"
'      ABBRUCH = True
'      GoTo EndeEinträgeLöschen
'    End With 'Sheets("Kontenplan")
'KontoNrGefunden:
'    Sheets("ArProt").Activate
'    Call KtoKennDat(ABZStreunKto)
'    ABZAkKto = ABZStreunKto      'vom HaSamlKto
'    ABZAkBlat = Austext
'    UndHaSaml = "'' und ''" & ABZStreunBlat
'    ABZAkGegKto = ABZSoKto
'    ABZAkSoHa = "Soll"
'    ABZAk3StZ = ABZStreun3StZ
'    Call AktKtoZeileLöschen    'Löschen der Buchungszeile
'    If ABBRUCH = True Then GoTo EndeEinträgeLöschen
'  End If 'ABZBetrofKto <> ""
ArProtAufräumen:
 '4EL -------------- Gelöschte Buchung in ArProt kennzeichnen ------------------
  With Sheets("ArProt")
    .Activate
    Cells(ABZeilNr, APCBuID) = ""     'BuID löschen
'    Austext = ""                    'für den Fall, dass schon gelöscht war
    Cells(ABZeilNr, APCBetrofKto) = Austext
    Cells(ABZeilNr, APCTANr).Range("A1:G1").Select
    With Selection.Font
      .Strikethrough = True
    End With
    Cells(ABZeilNr, APCgebucht).Value = "****"
    MELDUNG = Meldung1 & MELDUNG & Chr(10) & _
    "Einträge der ArProt-Zeile " & ABZeilNr & " in den Kontoblättern ''" & _
    ABZSoBlat & UndSoSaml & UndHa & UndHaSaml & "''" & _
    "nicht mehr vorhanden."
    Cells(ABZeilNr, APCgebucht).Activate 'Selektiertes Feld verkleinern
  End With 'Sheets("ArProt")
  GoTo EndeEinträgeLöschen
'5EL --------------- Änderungsvergleich der Kontenlängen ----------------------
EndeEinträgeLöschen:
End Sub 'EinträgeLöschen

Sub AktKtoZeileLöschen() 'von Sub EinträgeLöschen parametrisiert, aufgerufen
  'Löscht zu dem in ABZAkKto genannten Konto gehörigen Blatt ABZAkBlat
  'alle Zeilen mit der in ABZTaNr genannten TAN (normalerweise nur 1 Zeile),
  'dekrementiert dementsprechend den ***-Positionszeiger (Zelle A1 des Blatts)
  'und löscht den Blattnamen aus dem String in der Betroffene-Konten-Spalte der
  'ArProt-Zeile.
  
  Dim Z As Integer, BName As String, LängeBName As Integer, I As Integer
  Dim LängeAusText As Integer, PrüfText As String, VorText As String
  Dim NachText As String, BetrofKtoLen As Integer, AkBlaNamLen As Integer
  Dim BKSP As Integer 'BetrofKto-Stringscan-Position
  Dim AnzGelöcZeilen As Integer
  
  Call KtoKennDat(ABZAkKto)
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Struktur des Kontoblatts " & ABZAkBlat & " ungeeignet für Eintraglöschen."
    If Cells(1, 1) <> ABZAk3StZ Then
      MELDUNG = MELDUNG & Chr(10) & _
          "Kontenblatt " & ABZAkBlat & " hat in Zelle A1 eine" & _
          " falsche ***-Zeilen-Information. Zeile löschen nicht möglich."
    End If
    GoTo EndeAKLöschen
  End If 'ABBRUCH = True
  With Sheets(AKtoBlatt)
    .Activate
    If Cells(1, 1) = ABZAk3StZ Then
      AnzGelöcZeilen = 0
    End If
'---------- Suchen nach BuID und TANr: höhere Chance mit OR ----------------------
    For Z = 6 To ABZAk3StZ  'Konto-Schleife
      Cells(Z, KoCTANr).Activate    'für Verfolgung beim Test
      If Cells(Z, KoCTANr) = "" And Cells(Z, KoCDatum) = "" And _
               Z < ABZAk3StZ Then GoTo NächsteZeile
Vergleich:
      If Cells(Z, KoCDatum) = "***" Then
        If AnzGelöcZeilen = 0 Then
          MELDUNG = MELDUNG & Chr(10) & _
          "In Blatt ''" & AKtoBlatt & "'' keine zu löschende Zeile " & _
          "mit der Ta-Nr. " & ABZTaNr & " gefunden."
        Else
          MELDUNG = MELDUNG & Chr(10) & _
          "In Blatt ''" & AKtoBlatt & "'' wurden " & AnzGelöcZeilen & " Zeilen " & _
          "mit der Ta-Nr. " & ABZTaNr & " gelöscht."
        End If
        GoTo BetroffKtoLöschen
      End If
'------------------ Suche über TAN oder BuId (auch mehrmals)---------------------
      If Cells(Z, KoCBuID) = ABZBuID Or Cells(Z, KoCTANr) = ABZTaNr Then
        MELDUNG = MELDUNG & Chr(10) & _
        "In Kontoblatt ''" & ABZAkBlat & "'' wurde die Zeile " & Z & " gelöscht"
        Sheets(ABZAkBlat).Activate
        Cells(Z, 1).EntireRow.Select
        Selection.Delete shift:=xlUp        'Zeile Löschen
        Cells(1, 1) = Cells(1, 1) - 1       '3Sternzeile, vorbehaltlich Stutzen
        AnzGelöcZeilen = AnzGelöcZeilen + 1    'für Meldung
        ABZAk3StZ = ABZAk3StZ - 1
'        Cells(1, 9) = Cells(1, 9) + 1  'Storno-Zähler verzichtet
        If Cells(Z, KoCDatum) = "***" Then Exit For
        If Cells(Z, KoCTANr) = ABZTaNr Or _
           Cells(Z, KoCBuID) = ABZBuID Then GoTo Vergleich
      End If 'Cells(Z, KoCBuID) = ABZBuID
NächsteZeile:
    Next Z

    '----------- Blattname in ArProt-Betroffene-Konten ausschneiden --------------
BetroffKtoLöschen:
    With Sheets("ArProt")
      .Activate
      Austext = Cells(ABZeilNr, APCBetrofKto)
      If Austext = "" Then GoTo EndeBlattnameLöschen
      If Left(Austext, 2) = "+ " Then
        Austext = Right(Austext, Len(Austext) - 2)
      End If
      If Right(Austext, 1) = " " Then
        Austext = Left(Austext, Len(Austext) - 1)
      End If
      LängeAusText = Len(Austext)
      LängeBName = Len(ABZAkBlat)
      '----- Scan der Betroffene-Konten-Zelle in ArProt ---------------
      For I = 1 To LängeAusText + 1 - LängeBName
        PrüfText = Mid(Austext, I, LängeBName)
        If PrüfText = ABZAkBlat Then
          VorText = Left(Austext, I - 1)
          If LängeAusText - Len(VorText) - LängeBName < 3 Then
            NachText = ""
          Else
            NachText = Right(Austext, (LängeAusText - I - LängeBName))
          End If
          If Left(NachText, 2) = "+ " Then
            NachText = Right(NachText, Len(NachText) - 2)
          End If
          '---- Ausschneiden des Blattnamens -------------
          Austext = VorText & NachText
          Cells(ABZeilNr, APCBetrofKto) = Austext
          If Right(Austext, 3) = " + " Then
            Austext = Left(Austext, Len(Austext) - 3)
          End If
          If Left(Austext, 2) = "+ " Then
            Austext = Right(Austext, Len(Austext) - 2)
          End If
          Cells(ABZeilNr, APCBetrofKto) = Austext
          GoTo EndeAKLöschen
        End If 'PrüfText = ABZAkBlat
      Next I
EndeBlattnameLöschen:
      MELDUNG = MELDUNG & Chr(10) & _
      "In Spalte K von ArProt-Zeile " & Z & " wurde kein Blattname ''" & ABZAkBlat & _
      "'' zum löschen vorgefunden"
    End With
  End With 'Sheets(ABZAkBlat)
EndeAKLöschen:
End Sub 'AktKtoZeileLöschen

'Sub BetroffKtoBlatt(BeKoBlätter As String)   -> nach Modul NummernTAOrdnen verschoben
'Liest den ersten Kontoblattnamen aus dem Text "BeKoBlätter" und speichert ihn
'in die globale Vaiable "BekoBlatt", den verbleibenden Text in "RestString".
'BeKoBlätter enthält keinen, einen oder mehrere Namen, die durch den String " + "
'getrennt sind und endet mit einem Blank. Der Kontoblattname darf "+" nicht
'enthalten.  Vom Text RestString sind die ggf. führenden Zeichen " + " bereits
'weggeschnitten.
'Voraussetzung: Die Namen sind zusammenhängend (one Blank) und enthalten kein "+"
'Dim I As Integer, Länge As Integer, Zeichen As String
'  Länge = Len(BeKoBlätter)
'  If Länge = 0 Or Länge = 1 And BeKoBlätter = " " Then 'leerer oder 1-Blank-String
'    BeKoBlatt = ""
'    RestString = ""
'    Exit Sub
'  End If
'  RestString = BeKoBlätter
'  BeKoBlatt = ""
'  I = 0
'  Do
'    I = I + 1   'I ist Scanzeiger
'    If I > Länge Or Len(RestString) > 0 And Left(RestString, 1) = " " Then
'      BeKoBlatt = BeKoBlätter  'letzter Name
'      If Right(BeKoBlätter, 1) = " " Then
'        BeKoBlatt = Left(BeKoBlatt, Länge - 1)
'        RestString = ""
'        Exit Do
'      End If
'    End If
'    Zeichen = Mid(BeKoBlätter, I, 1)
'    If Zeichen = " " Then
'      BeKoBlatt = Left(BeKoBlätter, I - 1)
'      RestString = Right(BeKoBlätter, Länge - I)
'      If Left(RestString, 2) = "+ " Then
'        RestString = Right(BeKoBlätter, Länge - I - 2)
'        Exit Do
'      End If
'    End If
'  Loop
'End Sub 'BetroffKtoBlatt


'     zur Zeit nicht verwendet. Ändern auf: nicht bis zum Leeren stutzen!!!
Sub KontoBlattStutzen(KontoName As String)
'Die Dreisternzeile liegt unmittelbar hinter der chronologisch letzten
'Eintragungszeile. Danach folgt der zum Monat gehörige Saldoblock.
'Diese Ordnung kann beim Einträgelöschen eine Veränderung erfahren.
'Wenn alle Einträge des letzten Monats gelöscht sind, wird der
'Saldoblock überflüssig und die DreiSternZeile muß über den
'darüberliegenden Saldoblock gehoben werden.

  Dim DreiSternZeile As Integer, KontoNummer As Integer
  With Sheets(KontoName)
    .Activate
    KontoNummer = ABZAkKto
    Call KtoKennDat(ABZAkKto)
    Do
      DreiSternZeile = Cells(1, 1).Value
      If Cells(DreiSternZeile, KoCDatum).Offset(-2, 3) = "Kontostand" Then
        Range("B" & DreiSternZeile - 4 & ":B" & DreiSternZeile - 1 & "").Select
        Selection.Delete shift:=xlUp
        Cells(1, 1) = Cells(1, 1) - 4
        ABZAk3StZAend = ABZAk3StZAend - 4
        Range("C" & DreiSternZeile + 1 & ":H" & DreiSternZeile + 4 & "").Select
        Selection.Delete shift:=xlUp
      Else
        Exit Sub  'kein (weiteres) Stutzen
      End If
    Loop While DreiSternZeile > 6
  End With 'Sheets(KontoName)
End Sub 'KontoBlattStutzen

'=====================================================================================
'Function ArProtEnde()   'im Modul ArProtSchreiben (ERFASSEN)
'  Dim AnfBlat As String, Anfzeil As Integer, AnfSpalt As Integer
'  AnfBlat = ActiveSheet.Name
'  Anfzeil = ActiveCell.Row
'  AnfSpalt = aktspalte
'  With Sheets("ArProt")
'    .Activate
'    Cells(1, 3).Activate
'    ArProtEnde = ActiveCell.Value
'  End With
'  Sheets(AnfBlat).Activate
'  Cells(Anfzeil, AnfSpalt).Activate
'End Function 'ArProtEnde
'---------------------------------------------------------------------

Sub SucheKonto(SuchText As String)
'Aufruf von ERFASSEN aus mit ArProt als aktives Blatt
'Sucht im Kontenplan in der Spalte "Beschreibung", und wenn dort nicht gefunden,
'in der Spalte "Blattname" nach dem String "SuchText", jeweils von der ersten Zeile
'der Spalte an. Bietet in einer MessageBox das gefundene Konto an mit der Wahl,
'es zu akzeptieren, weiterzusuchen oder abzubrechen. Wird das Konto akzeptiert,
'Schreibt SucheKonto die Kontonummer in die aktive Zelle (des Blattes ArProt)und
'aktiviert die rechts daneben liegende Zelle


Dim Länge As Integer, AktBlatt As String, AktZell As Range, Zeile As Integer
Dim Kto As Integer, Becr As String, Blat As String, Art As Integer, A
Dim SuchAnfZeileB As Integer, SuchAnfZeileN As Integer
  With ActiveSheet          'Aufbewahren Aufrufsituation
    AktBlatt = ActiveSheet.Name
    Set AktZell = ActiveCell
  End With                  '---------------------------
  Länge = Len(SuchText)
  SuchAnfZeileB = 5   'für Suche in Spalte Beschreibung
  SuchAnfZeileN = 5   'für Suche in Spalte Blattname
Such:
  With Worksheets("Kontenplan")
    .Activate
    For Zeile = SuchAnfZeileB To Cells(1, 3) Step 1
      If Left(Cells(Zeile, KPCBeschr), Länge) = SuchText Then
        Exit For
      End If
    Next Zeile
    SuchAnfZeileB = Zeile + 1
    If Zeile >= Cells(1, 3).Value Then  'Weiterer Versuch mit Blattnamen-Spalte
      For Zeile = SuchAnfZeileN To Cells(1, 3) Step 1
        If Left(Cells(Zeile, KPCBlattname), Länge) = SuchText Then
          Exit For
        End If
      Next Zeile
      SuchAnfZeileN = Zeile + 1
    End If
    If Zeile < Cells(1, 3).Value Then
      Kto = Cells(Zeile, KPCKonto)
      Call KtoKennDat(Kto)
      Becr = AKtoBeschr
      Blat = AKtoBlatt
      Art = AKtoArt
    Else
      Kto = 0
      Becr = ""
      Blat = ""
      Art = 0
    End If
'    SuchAnfZeile = Zeile + 1
  End With
  With Worksheets(AktBlatt) 'Wiederherstellen Aufrufsituation
    .Activate
    AktZell.Activate
    If Kto <> 0 Then
      A = MsgBox("     " & Kto & "    " & Becr & "     " & Blat & Chr(10) & Chr(10) & _
                 "Schaltfläche ''Nein'', falls weitergesucht werden soll", 35, _
                 "Ist das Konto gemeint?")
      If A = vbCancel Then Exit Sub
      If A = vbYes Then
        ActiveCell.Value = Kto
 '       Call KtoKennDat(Kto)
        ActiveCell.Offset(0, 1).Activate
        If ActiveCell.Column = APCText And Art = 10 Then
          ActiveCell = "Beitrag " & AKtoBeschr & ""
        End If
        If ActiveCell.Column = APCText And Art = 11 Then
          ActiveCell = "Spende " & AKtoBeschr & ""
        End If
        ActiveCell.Offset(0, 1).Activate
        Exit Sub
      End If
      If A = vbNo Then GoTo Such
    End If
    If Kto = 0 Then
      A = MsgBox("Mit einer anderen Zeichenfolge versuchen" & Chr(10) & Chr(10), 0, _
                 "Kein Konto gefunden")
    End If
  End With                  '
End Sub  'SucheKonto
'=================================================================================
'============================================================================================
'============================================================================
