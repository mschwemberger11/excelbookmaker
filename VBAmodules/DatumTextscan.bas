Attribute VB_Name = "DatumTextscan"
     '************************************************************************
     'Modul DatumTextscan
     '1. Datums-Darstellung und Umwandlung vom Format d.mmm als Text in Integer
     '   (Zahl zwischen 1 und 365 des aktuellen Jahres) und umgekehrt. Die Zahl
     '   0 repräsentiert den 31.Dez des Vorjahres, der als Text im Format
     '   dd.mm.yyyy dargestellt wird. 11 Routinen.
     '2. Abtasten eines Textes bis zu einem vorgegebenen Endzeichen-String.
     '   1 Routine. Liefert die globalen Variablen TextStück (String),
     '   LängeTextstück (Long), Lesezeiger (Long), TextEndeErreicht (Boolean)
     '************************************************************************
    'Enthält die Routinen (hauptsächlich Datumsroutinen "DR-"):
        
  '2 'Function Schalttag
        'Liefert von dem im Kontenplan in Cells(1,5) vermerkten Jahr eine 1, wenn
        'es ein Schaltjahr ist, sonst eine 0, und liefert dies in der Kontenplan-
        'Kopfzeile F1 ab und speichert es in der globalen Variablen CalTag, hier
        'im Modul DatumTextscan definiert.
        'Funktioniert bis Jahr 2099.
        
  '3 'Sub DatumAlsDMMM()  Aufruf durch Strg+d
        'Datumsharmonisierung in einer Spalte durch Darstellung als Text
        'vom Format d.mmm  Sehr speziell auf eine Spalte bezogen
        
  '4 'Function DatumTZ(DText As String) As Integer
        'Datum-Zahl (Integer zwischen (0), 1 .. 365) aus Datum-Text "DText"
  '5 'Function DatumZT(DZahl As Integer) As String
        'Datum-Text (im Format d.mmm) aus Datum-Zahl "DZahl"
        
  '6 'Function MonatZZ(DZahl As Long) As Integer
        'Nummer des Monats, in dem das durch "DZahl" gegebene Datum liegt
  '7 'Function MonatTZ(MText As String) As Integer
        'Nummer des Monats, der durch "MText" (3 Buchstaben, z.B. "Jan") gegeben ist
  '8 'Function MonatZT(DZahl As Long) As String
        'Name des Monats (3Buchst), in dem das durch "DZahl" gegebene Datum liegt
  '9 'Function LangDatum(KurzDatum As String) As String
        'Verlängert im Datumstring, wie von MonatZT geliefert, den Monatsnamen
     
  '10'Function ErstTagIM(DZahl As Long) As Integer
        'Datum-Zahl vom ersten Tag des Monats, in dem der durch "DZahl" gegebene
        'Tag liegt
  '11'Function LetzTagIM(DZahl As Long) As Long
        'Datum-Zahl vom letzten Tag des Monats, in dem der durch "DZahl" gegebene
        'Tag liegt.
  '12'Function MonatsErster(MNumr As Integer) As Long
        'Datum-Zahl des ersten Tags vom Monat, der durch MNumr gegeben ist.
  '13'Function MonatsLetzter(MNumr As Integer) As Long
        'Datum-Zahl des letzten Tags vom Monat, der durch MNumr gegeben ist.
        
  '14'Function IstEinDatum(PrüfWert) As Boolean
        'Wahr, wenn die letzten 3 Zeichen von "Prüfwert" ein Monatsbezeichner
        '(Jan, Feb, Mrz, ... Dez) sind, sonst Falsch.
        
  '15'Function IstNumDatum(Prüfling As Variant) As Boolean               '22.6.2017
        'Gibt 'True' zurück, wenn der Prüfling als String das Format d.m. oder
        'dd.mm. hat und dd eine Zahl >=1 & <=31 und mm eine Zahl >=1 & <=12 ist.
      
  '16'Sub TextScan(Text As String, StartStelle As Long, EndString As String)
        'out: TextStück (String), LängeTextstück (Long), Lesezeiger (Long),
        '     TextEndeErreicht (Boolean)
        'Analyse des Strings "Text" von "StartStelle" bis Literal "EndString"
     '
     'Einige dieser Routinen verwenden die Function "Schalttag" aus dem Modul
     'M2Kontplan
   
    Option Explicit   'Modul 2
    Public CalTag As Integer
    Public TextStück As String, LängeTextstück As Long, Lesezeiger As Long
    Public DatumTag As Integer, DatumMonat As Integer
'    Public AktBlattStanDatm As String, _
'           AktZeilStanDatm As Integer, AktSpalStanDatm As Integer
    Dim EndString As String, LängeEndString As Long
    Dim TextEndeErreicht As Boolean
    Dim TagesZahl As Long, MonatsZahl As Integer

    
Function SchaltTag() As Integer
  'Errechnet von dem im Kontenplan in Cells(1,5) vermerkten Jahr eine 1, wenn
  'es ein Schaltjahr ist, sonst eine 0. Funktioniert bis Jahr 2099.
  Dim AktBlatt As String, AktZell As Range, Jahr, I As Long
  With ActiveWindow
'1 CT ----------------------- Aufbewahren Aufrufsituation --------------------------
    With ActiveSheet
      AktBlatt = ActiveSheet.Name
      Set AktZell = ActiveCell
    End With
'2 CT --------------- Schalttag (0 oder 1) ermitteln und --> Kontenplan F1 --------------
    With Sheets("Kontenplan")
      .Select
      Jahr = Cells(1, 5).Value
      For I = 4 To 24 Step 4   'funktioniert bis Jahr 2099
        If Jahr - I = 2004 Then
          SchaltTag = 1
          Exit For
        End If
        If Jahr - I < 2004 Then
          SchaltTag = 0
          Exit For
        End If
      Next I
'      Sheets("Kontenplan").Cells(1, 6) = SchaltTag
'      CalTag = SchaltTag
    End With 'Sheets("Kontenplan")
'3 CT --------------------- Wiederherstellen Aufrufsituation ------------------------
    With Worksheets(AktBlatt)
      .Activate
      AktZell.Activate
    End With
  End With 'ActiveWindow
End Function 'SchaltTag
   

  Function DatumTZ(DText As String) As Long
    'Liefert von einem Datumsausdruck in Textform die Integerzahl, die den Tag im
    'Transaktionsjahr repräsentiert, wobei der 1.Januar durch die Zahl 1 und der
    '31.Dezember durch die Zahl 365 wiedergegeben wird (in einem Schaltjahr 366).
    'Der 31.Dezember des Vorjahres hat den Zahlenwert 0. Andere Daten werden nicht
    'behandelt. Für Dtext sind nur die Formate d.mmm , d.mmm. und dd.mm.yyyy zulässig.
    'Eine nicht dargestellte oder nur 2stellig dargestellte Jahreszahl wird als
    'Transaktionsjahr (in der Regel das laufende Jahr) interpretiert. Sie wird der
    'Globalen Variablen TransaktJahr entnommen, die ihrerseits in einem Aufruf von
    'Kontenplanstruktur dem Kontenplan, Cells(1,5), entnommen wird.
    'Die Umwandlung einer Datumszahl in die Textform leistet Function DatumZT.
    Dim AktBlattName As String, zzeile As Integer, zspalte As Integer
    With ActiveWindow
      If DText = "" Then
        DatumTZ = 0
        Exit Function
      End If
      If Len(DText) > 6 Then   'ggf. Jahreszahl abschneiden
        DText = Left(DText, 6)
      End If
      Lesezeiger = 1
      '------------------- Tageszahl --------------------
      EndString = "."
      Call TextScan(DText, Lesezeiger, EndString)
      If LängeTextstück <= 3 Then   'hier locker wegen etwa führender Blanks
        TagesZahl = CInt(TextStück)
      Else
        With ActiveSheet
          zzeile = ActiveCell.Row
          zspalte = ActiveCell.Column
          MELDUNG = MELDUNG & Chr(10) & _
            "in Zelle (" & zzeile & "," & zspalte & ") fand Function DatumTZ" _
            & Chr(10) & "keine zulässige Tagesangabe"
          ABBRUCH = True
          Exit Function
        End With
      End If
      '-------------------Monatszahl ---------------------
      LängeEndString = Len(EndString)
      Call TextScan(DText, Lesezeiger, ".")
      '------------
      If LängeTextstück >= 2 And LängeTextstück <= 3 And _
         IsNumeric(Left(TextStück, LängeTextstück - 1)) = True Then
        MonatsZahl = CInt(TextStück)
      End If
      '------------
      If LängeTextstück = 3 And IsNumeric(TextStück) = False Then
        MonatsZahl = MonatTZ(TextStück)
        If MonatsZahl = 0 Then
          With ActiveSheet
            zzeile = ActiveCell.Row
            zspalte = ActiveCell.Column
            MELDUNG = MELDUNG & Chr(10) & _
              "in Zelle (" & zzeile & "," & zspalte & ") fand Function DatumTZ" _
              & Chr(10) & "keine zulässige Monatssangabe"
            ABBRUCH = True
            GoTo DatumTZEnde
          End With
        End If
      End If
      '-------------
      If TextEndeErreicht = True Then
        GoTo Abschluss
      End If
      '---------- ggf. Jahreszahl aus Globalvariable TransaktJahr ------------------
      Call TextScan(DText, Lesezeiger, ".") ' "." müsste ohne Belang sein
      If LängeTextstück = 4 Then
        If CInt(Right(TextStück, 2)) = CInt(Right(TransaktJahr, 2)) - 1 And MonatsZahl = 12 Then
          DatumTZ = 0
          Exit Function   '--------------->
        Else
          With ActiveSheet
            zzeile = ActiveCell.Row
            zspalte = ActiveCell.Column
 '          DatRoutMeldungen = DatRoutMeldungen & Chr(10) & _
 '              "in Zelle (" & zzeile & "," & zspalte & ") fand Function DatumTZ " & Chr(10) & _
 '              "keine zulässige Jahresangabe"
 '           DatRoutFehler = DatRoutFehler + 1
            Exit Function
          End With
        End If
      End If
      '-------------- Monat und Tag zusammensetzen --------------
DatumTZEnde:
'      If DatRoutFehler > 0 Then
'        Call MsgBox(DatRoutMeldungen, vbOKOnly, "Datumsroutinen")
'      Else
Abschluss:
        DatumTZ = MonatsErster(MonatsZahl) + TagesZahl - 1
'        If DatRoutFehler > 0 Then GoTo DatumTZEnde
'      End If
    End With 'ActiveWindow
  End Function 'DatumTZ
    
  Function DatumZT(DZahl As Long) As String
    'Stellt eine Datumszahl (Bereich 1 ... 365) in einen String vom Format d.mmm dar,
    'wobei die fehlende Jahresangabe als zum Transaktionsjahr gehörend zu verstehen ist.
    'Die Datumszahl folgt den Regeln, die die Function DatumTZ verwendet.
    'Setzt den gültigen Wert der globalen Variablen TransaktJahr - also einen voraus-
    'gegangenen Aufruf von Kontenplanstruktur - voraus.
    Dim MonaZ As Integer, MonaT As String, Ergebnis As String
    If DZahl = 0 Then
      DatumZT = "31.12." & Str(CInt(TransaktJahr) - 1) & ""
    End If
    If DZahl > 365 + CalTag Then
      MELDUNG = MELDUNG & Chr(10) & _
         "aus der Tageszahl " & DZahl & "konnte Function DatumZT " & Chr(10) & _
         "kein Datum im Format d.mmm ableiten"
      ABBRUCH = True
      Exit Function
    End If
    MonaZ = MonatZZ(DZahl)
    MonaT = MonatZT(DZahl)
    Ergebnis = Str(DZahl - MonatsErster(MonaZ) + 1) & "." & MonaT
    If Left(Ergebnis, 1) = " " Then Ergebnis = Right(Ergebnis, Len(Ergebnis) - 1)
    DatumZT = Ergebnis
  End Function 'DatumZT

Sub DatumSpalten() 'out: Public DatumTag, DatumMonat als Integer
    'Liefert von der als Text aufgefassten Datumsinformation in der aktiven Zelle
    'zwei Integerzahlen als getrennte Public-Variablen, "DatumTag" für den Tag im
    'Monat, und "DatumMonat" für den Monat und schreibt die Datuminformation als
    'String in der standardisierten Form d.mmm (Anfangsbuchstaben) in die aktive Zelle.
    'Für die Datumsinformation sind die Formate "d.mmm", "d.mmm.", "dd.mm.yyyy" (alpha)
    'und "t.m.", "tt.mm.", "tt.mm.jj", "tt.mm.jjjj" zulässig. Die Jahresangabe
    'hinter dem zweiten Punkt ist ohne Belang und wird ignoriert.
    'Das Jahr ist implizit das aktuelle Transaktionsjahr, hat aber im vorliegenden
    'Kontext keine Bedeutung.
    'Prüft, ob Tages- und Monatszahl im zulässigen Bereich sind und setzt ggf.
    'ABZAbbruch = True
    '
    'Benutzt Sub TextScan und Function MonatTZ
    '
  Dim SpAktBlatt As String, SpAktZeile As Integer, SpAktSpalte As Integer
  Dim Datum As String, LetzTag As Integer
'1------------------ Aufrufsituation ------------------------------
  SpAktBlatt = ActiveSheet.Name
  SpAktZeile = ActiveCell.Row
  SpAktSpalte = ActiveCell.Column
'2------------------------ Eingabeparameter prüfen ------------------------------
  Datum = ActiveCell.Value  'Cells(SpAktZeile, SpAktSpalte)
  If Datum = "" Then
    DatumTag = 0
    DatumMonat = 0
    MELDUNG = MELDUNG & Chr(10) & _
        "Der Datumtext ''" & Datum & "'' in Zelle B" & SpAktZeile & _
        " des Blatts " & SpAktBlatt & " ist unbrauchbar." & Chr(10) & _
        "Abbruchgrund."
        ABBRUCH = True
     GoTo EndeDatumSpalten
  End If
  Lesezeiger = 1
  '------------------- Tageszahl --------------------
  EndString = "."
  Call TextScan(Datum, Lesezeiger, EndString)
  If LängeTextstück <= 3 Then   'hier locker wegen etwa führender Blanks
    DatumTag = CInt(TextStück)
  Else
    MELDUNG = MELDUNG & Chr(10) & _
    "Der Datumtext ''" & Datum & "'' in Zelle B" & SpAktZeile & " des " _
    & "Blatts ''" & SpAktBlatt & "'' Hat keine brauchbare Tagesangabe" _
    & Chr(10) & "Abbruchgrund."
    ABBRUCH = True
    GoTo EndeDatumSpalten
  End If
  '-------------------Monatszahl ---------------------
  LängeEndString = Len(EndString)
  Call TextScan(Datum, Lesezeiger, EndString)
  '------------ Fall: Monat als Zahl angegeben ------------------------------
  If LängeTextstück >= 1 And LängeTextstück <= 3 And _
         IsNumeric(Left(TextStück, LängeTextstück)) = True Then
    DatumMonat = CInt(TextStück)
    If DatumMonat < 1 Or DatumMonat > 12 Then
      MELDUNG = MELDUNG & Chr(10) & _
      "in Zelle (" & SpAktZeile & "," & SpAktSpalte & ") keine zulässige Monatsangabe" _
      & "gefunden " & Chr(10) & "(Sub DatumSpalten)"
      ABBRUCH = True
      GoTo EndeDatumSpalten
    End If
  End If
  '------------ Fall Monat als 3stelliger Text angegeben ---------------------------
  If LängeTextstück = 3 And IsNumeric(TextStück) = False Then
    DatumMonat = MonatTZ(TextStück)                      'Public-Variable
  '------------- Tageszahl in gültigem Bereich? ----------------------------------
    If IsNumeric(DatumMonat) = True And DatumMonat >= 1 And DatumMonat <= 12 Then
      If (DatumMonat = 1 Or DatumMonat = 3 Or DatumMonat = 5 Or DatumMonat = 7 Or _
          DatumMonat = 8 Or DatumMonat = 10 Or DatumMonat = 12) Then
        LetzTag = 31
      End If
      If (DatumMonat = 4 Or DatumMonat = 6 Or DatumMonat = 9 Or DatumMonat = 11) Then
        LetzTag = 30
      End If
      If DatumMonat = 2 Then
        LetzTag = 28 + CalTag
      End If
    End If
    If DatumTag >= 1 And DatumTag <= LetzTag Then
      GoTo EndeDatumSpalten
    Else
      MELDUNG = MELDUNG & Chr(10) & _
      "in Zelle (" & SpAktZeile & "," & SpAktSpalte & ") keine zulässige Tagesangabe" _
      & "gefunden " & Chr(10) & "(Sub DatumSpalten)"
      ABBRUCH = True
      GoTo EndeDatumSpalten
    End If
  End If
  GoTo EndeDatumSpalten
  '------------------- Ausgangszustand wiederherstellen -------------------------
EndeDatumSpalten:
    Sheets(SpAktBlatt).Activate
    Cells(SpAktZeile, SpAktSpalte).Activate
    ActiveCell = CStr(DatumTag) & "." & MonatMzT(DatumMonat)
  End Sub 'DatumSpalten

  Function MonatZZ(DZahl As Long) As Integer
      MonatZZ = 0
      If DZahl >= 0 And DZahl <= 31 Then
        MonatZZ = 1
        GoTo ZZgef
      End If
      If DZahl > 31 And DZahl <= 59 + CalTag Then
        MonatZZ = 2
        GoTo ZZgef
      End If
      If DZahl > 59 + CalTag And DZahl <= 90 + CalTag Then
        MonatZZ = 3
        GoTo ZZgef
      End If
      If DZahl > 90 + CalTag And DZahl <= 120 + CalTag Then
        MonatZZ = 4
        GoTo ZZgef
      End If
      If DZahl > 120 + CalTag And DZahl <= 151 + CalTag Then
        MonatZZ = 5
        GoTo ZZgef
      End If
      If DZahl > 151 + CalTag And DZahl <= 181 + CalTag Then
        MonatZZ = 6
        GoTo ZZgef
      End If
      If DZahl > 181 + CalTag And DZahl <= 212 + CalTag Then
        MonatZZ = 7
        GoTo ZZgef
      End If
      If DZahl > 212 + CalTag And DZahl <= 243 + CalTag Then
        MonatZZ = 8
        GoTo ZZgef
      End If
      If DZahl > 243 + CalTag And DZahl <= 273 + CalTag Then
        MonatZZ = 9
        GoTo ZZgef
      End If
      If DZahl > 273 + CalTag And DZahl <= 304 + CalTag Then
        MonatZZ = 10
        GoTo ZZgef
      End If
      If DZahl > 304 + CalTag And DZahl <= 334 + CalTag Then
        MonatZZ = 11
        GoTo ZZgef
      End If
      If DZahl > 334 + CalTag And DZahl <= 365 + CalTag Then MonatZZ = 12
ZZgef:
'      If MonatZZ = 0 Then
'        DatRoutMeldungen = DatRoutMeldungen & Chr(10) & _
'               "zur Tageszahl " & DZahl & "fand Function MonatZZ keine Monatszahl"
'        DatRoutFehler = DatRoutFehler + 1
'      End If
    End Function 'MonatZZ
    
    Function MonatTZ(MText As String) As Integer
      MonatTZ = 0
      If MText = "Jan" Then MonatTZ = 1
      If MText = "Feb" Then MonatTZ = 2
      If MText = "Mrz" Then MonatTZ = 3
      If MText = "Apr" Then MonatTZ = 4
      If MText = "Mai" Then MonatTZ = 5
      If MText = "Jun" Then MonatTZ = 6
      If MText = "Jul" Then MonatTZ = 7
      If MText = "Aug" Then MonatTZ = 8
      If MText = "Sep" Then MonatTZ = 9
      If MText = "Okt" Then MonatTZ = 10
      If MText = "Nov" Then MonatTZ = 11
      If MText = "Dez" Then MonatTZ = 12
     End Function 'MonatTZ
     
    Function MonatMzT(MZahl)
      MonatMzT = ""
      If MZahl = 1 Then MonatMzT = "Jan"
      If MZahl = 2 Then MonatMzT = "Feb"
      If MZahl = 3 Then MonatMzT = "Mrz"
      If MZahl = 4 Then MonatMzT = "Apr"
      If MZahl = 5 Then MonatMzT = "Mai"
      If MZahl = 6 Then MonatMzT = "Jun"
      If MZahl = 7 Then MonatMzT = "Jul"
      If MZahl = 8 Then MonatMzT = "Aug"
      If MZahl = 9 Then MonatMzT = "Sep"
      If MZahl = 10 Then MonatMzT = "Okt"
      If MZahl = 11 Then MonatMzT = "Nov"
      If MZahl = 12 Then MonatMzT = "Dez"
    End Function
    
    Function MonatZT(DZahl As Long) As String
      MonatZT = ""
      If DZahl > 0 And DZahl <= 31 Then
        MonatZT = "Jan"
        GoTo ZTgef
      End If
      If DZahl > 31 And DZahl <= 59 + CalTag Then
        MonatZT = "Feb"
        GoTo ZTgef
      End If
      If DZahl > 59 + CalTag And DZahl <= 90 + CalTag Then
        MonatZT = "Mrz"
        GoTo ZTgef
      End If
      If DZahl > 90 + CalTag And DZahl <= 120 + CalTag Then
        MonatZT = "Apr"
        GoTo ZTgef
      End If
      If DZahl > 120 + CalTag And DZahl <= 151 + CalTag Then
        MonatZT = "Mai"
        GoTo ZTgef
      End If
      If DZahl > 151 + CalTag And DZahl <= 181 + CalTag Then
        MonatZT = "Jun"
        GoTo ZTgef
      End If
      If DZahl > 181 + CalTag And DZahl <= 212 + CalTag Then
        MonatZT = "Jul"
        GoTo ZTgef
      End If
      If DZahl > 212 + CalTag And DZahl <= 243 + CalTag Then
        MonatZT = "Aug"
        GoTo ZTgef
      End If
      If DZahl > 243 + CalTag And DZahl <= 273 + CalTag Then
        MonatZT = "Sep"
        GoTo ZTgef
      End If
      If DZahl > 273 + CalTag And DZahl <= 304 + CalTag Then
        MonatZT = "Okt"
        GoTo ZTgef
      End If
      If DZahl > 304 + CalTag And DZahl <= 334 + CalTag Then
        MonatZT = "Nov"
        GoTo ZTgef
      End If
      If DZahl > 334 + CalTag And DZahl <= 365 + CalTag Then MonatZT = "Dez"
ZTgef:
  '    If MonatZT = "" Then
  '      DatRoutMeldungen = DatRoutMeldungen & Chr(10) & _
  '             "zur Datumszahl " & DZahl & "fand Function MonatZT keinen Monat"
  '      DatRoutFehler = DatRoutFehler + 1
  '    End If
    End Function 'MonatZT
    
    Function LangDatum(KurzDatum As String) As String
    'Verlängert im Datumstring, wie von MonatZT geliefert, den Monatsnamen
    Dim Tag As String
       Tag = Left(KurzDatum, Len(KurzDatum) - 3)
      If Right(KurzDatum, 3) = "Jan" Then
        LangDatum = Tag & " Januar"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Feb" Then
        LangDatum = Tag & " Februar"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Mrz" Then
        LangDatum = Tag & " März"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Apr" Then
        LangDatum = Tag & " April"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Mai" Then
        LangDatum = Tag & " Mai"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Jun" Then
        LangDatum = Tag & " Juni"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Jul" Then
        LangDatum = Tag & " Juli"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Aug" Then
        LangDatum = Tag & " August"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Sep" Then
        LangDatum = Tag & " September"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Okt" Then
        LangDatum = Tag & " Oktober"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Nov" Then
        LangDatum = Tag & " November"
        Exit Function
      End If
      If Right(KurzDatum, 3) = "Dez" Then
        LangDatum = Tag & " Dezember"
        Exit Function
      End If
      MELDUNG = MELDUNG & Chr(10) & _
      "aus dem Kurzdatum " & KurzDatum & "konnte Function LangDatum kein" _
      & Chr(10) & "Langdatum ableiten"
      ABBRUCH = True
    End Function
    
    Function ErstTagIM(DZahl As Long) As Long
    'Liefert die Datum-Zahl vom ersten Tag des Monats, in dem DZahl liegt
      ErstTagIM = 0
      If DZahl > 0 And DZahl <= 31 Then ErstTagIM = 1
      If DZahl > 31 And DZahl <= 59 + CalTag Then ErstTagIM = 32
      If DZahl > 59 + CalTag And DZahl <= 90 + CalTag Then ErstTagIM = 60 + CalTag
      If DZahl > 90 + CalTag And DZahl <= 120 + CalTag Then ErstTagIM = 91 + CalTag
      If DZahl > 120 + CalTag And DZahl <= 151 + CalTag Then ErstTagIM = 121 + CalTag
      If DZahl > 151 + CalTag And DZahl <= 181 + CalTag Then ErstTagIM = 152 + CalTag
      If DZahl > 181 + CalTag And DZahl <= 212 + CalTag Then ErstTagIM = 182 + CalTag
      If DZahl > 212 + CalTag And DZahl <= 243 + CalTag Then ErstTagIM = 213 + CalTag
      If DZahl > 243 + CalTag And DZahl <= 273 + CalTag Then ErstTagIM = 244 + CalTag
      If DZahl > 273 + CalTag And DZahl <= 304 + CalTag Then ErstTagIM = 274 + CalTag
      If DZahl > 304 + CalTag And DZahl <= 334 + CalTag Then ErstTagIM = 305 + CalTag
      If DZahl > 334 + CalTag And DZahl <= 365 + CalTag Then ErstTagIM = 335 + CalTag
      If ErstTagIM = 0 Then
        MELDUNG = MELDUNG & Chr(10) & _
        "zur Datumszahl " & DZahl & " ErstTagImMonat nicht gefunden"
        ABBRUCH = True
      End If
    End Function 'ErstTagIM
      
    Function LetzTagIM(DZahl As Long) As Long
    'Liefert die Zahl des letzten Tags vom Monat, in dem auch DZahl liegt.
      LetzTagIM = 0
      If DZahl > 0 And DZahl <= 31 Then LetzTagIM = 31
      If DZahl > 31 And DZahl <= 59 + CalTag Then LetzTagIM = 59 + CalTag
      If DZahl > 59 + CalTag And DZahl <= 90 + CalTag Then LetzTagIM = 90 + CalTag
      If DZahl > 90 + CalTag And DZahl <= 120 + CalTag Then LetzTagIM = 120 + CalTag
      If DZahl > 120 + CalTag And DZahl <= 151 + CalTag Then LetzTagIM = 151 + CalTag
      If DZahl > 151 + CalTag And DZahl <= 181 + CalTag Then LetzTagIM = 181 + CalTag
      If DZahl > 181 + CalTag And DZahl <= 212 + CalTag Then LetzTagIM = 212 + CalTag
      If DZahl > 212 + CalTag And DZahl <= 243 + CalTag Then LetzTagIM = 243 + CalTag
      If DZahl > 243 + CalTag And DZahl <= 273 + CalTag Then LetzTagIM = 273 + CalTag
      If DZahl > 273 + CalTag And DZahl <= 304 + CalTag Then LetzTagIM = 304 + CalTag
      If DZahl > 304 + CalTag And DZahl <= 334 + CalTag Then LetzTagIM = 334 + CalTag
      If DZahl > 334 + CalTag And DZahl <= 365 + CalTag Then LetzTagIM = 365 + CalTag
      If LetzTagIM = 0 Then
        MELDUNG = MELDUNG & Chr(10) & _
        "zur Datumszahl " & DZahl & " LetzTagImMonat nicht gefunden"
        ABBRUCH = True
      End If
    End Function 'LetzTagIM
    
    Function MonatsLetzter(MNumr As Integer) As Long
    'Liefert die Zahl des letzten Tag vom Monat, dessen Nummer durch MNumr
    'gegeben ist.
      MonatsLetzter = 0
      If MNumr = 1 Then MonatsLetzter = 31
      If MNumr = 2 Then MonatsLetzter = 59 + CalTag
      If MNumr = 3 Then MonatsLetzter = 90 + CalTag
      If MNumr = 4 Then MonatsLetzter = 120 + CalTag
      If MNumr = 5 Then MonatsLetzter = 151 + CalTag
      If MNumr = 6 Then MonatsLetzter = 181 + CalTag
      If MNumr = 7 Then MonatsLetzter = 212 + CalTag
      If MNumr = 8 Then MonatsLetzter = 243 + CalTag
      If MNumr = 9 Then MonatsLetzter = 273 + CalTag
      If MNumr = 10 Then MonatsLetzter = 304 + CalTag
      If MNumr = 11 Then MonatsLetzter = 334 + CalTag
      If MNumr = 12 Then MonatsLetzter = 365 + CalTag
      If MonatsLetzter = 0 Then
        MELDUNG = MELDUNG & Chr(10) & _
        "zur Monatsnummer " & MNumr & "MonatsLetzer nicht gefunden"
        ABBRUCH = True
      End If
    End Function 'MonatsLetzter
     
    Function MonatsLetzText(MNumr As Integer) As String
      If MNumr = 1 Then MonatsLetzText = "31.Jan"
      If MNumr = 2 Then MonatsLetzText = CStr(28 + CalTag) & ".Feb"
      If MNumr = 3 Then MonatsLetzText = "31.Mrz"
      If MNumr = 4 Then MonatsLetzText = "30.Apr"
      If MNumr = 5 Then MonatsLetzText = "31.Mai"
      If MNumr = 6 Then MonatsLetzText = "30.Jun"
      If MNumr = 7 Then MonatsLetzText = "31.Jul"
      If MNumr = 8 Then MonatsLetzText = "31.Aug"
      If MNumr = 9 Then MonatsLetzText = "30.Sep"
      If MNumr = 10 Then MonatsLetzText = "31.Okt"
      If MNumr = 11 Then MonatsLetzText = "30.Nov"
      If MNumr = 12 Then MonatsLetzText = "31.Dez"
    End Function
    
    
    Function MonatsErster(MNumr As Integer) As Long
    'Liefert die Zahl des letzten Tag vom Monat dessen Nummer durch MNumr
    'gegeben ist. (CalTag = 1 oder 0, je ob Schaltjahr oder nicht)
      MonatsErster = 0
      If MNumr = 1 Then MonatsErster = 1
      If MNumr = 2 Then MonatsErster = 32
      If MNumr = 3 Then MonatsErster = 60 + CalTag
      If MNumr = 4 Then MonatsErster = 91 + CalTag
      If MNumr = 5 Then MonatsErster = 121 + CalTag
      If MNumr = 6 Then MonatsErster = 152 + CalTag
      If MNumr = 7 Then MonatsErster = 182 + CalTag
      If MNumr = 8 Then MonatsErster = 213 + CalTag
      If MNumr = 9 Then MonatsErster = 244 + CalTag
      If MNumr = 10 Then MonatsErster = 274 + CalTag
      If MNumr = 11 Then MonatsErster = 305 + CalTag
      If MNumr = 12 Then MonatsErster = 335 + CalTag
      If MonatsErster = 0 Then
        MELDUNG = MELDUNG & Chr(10) & _
        "zur Monatsnummer " & MNumr & "MonatsErster nicht gefunden"
        ABBRUCH = True
      End If
    End Function 'MonatsErster
    
    Function IstEinDatum(Prüfwert As Variant) As Boolean
      'Gibt 'True' zurück, wenn die letzten 3 Zeichen von Prüfwert einer der
      'Monatsnamen "Jan", "Feb", "Mrz", ... "Dez" ist.
      Dim P As String
      IstEinDatum = False
      If Prüfwert = "" Then Exit Function
      P = Right(Prüfwert, 3)
      If P = "Jan" Or P = "Feb" Or P = "Mrz" Then IstEinDatum = True
      If P = "Apr" Or P = "Mai" Or P = "Jun" Then IstEinDatum = True
      If P = "Jul" Or P = "Aug" Or P = "Sep" Then IstEinDatum = True
      If P = "Okt" Or P = "Nov" Or P = "Dez" Then IstEinDatum = True
    End Function
      
    Function IstNumDatum(Prüfling As String) As Boolean               '22.6.2017
      'Gibt 'True' zurück, wenn der Prüfling als String das Format d.m. oder
      'dd.mm. hat und dd eine Zahl >=1 & <=31 und mm eine Zahl >=1 & <=12 ist.
      Dim Q As String, LängeText As Integer, TagZahl As Integer, MonZahl As Integer
      Const Punkt = "."
      Lesezeiger = 1
      IstNumDatum = False
      If Prüfling = "" Then GoTo DRFehler
      LängeText = Len(Prüfling)
      If LängeText > 6 Then GoTo DRFehler
      Call TextScan(Prüfling, Lesezeiger, Punkt)
      If LängeTextstück > 3 Then GoTo DRFehler  'tritt auf, wenn Endzeichen <> Punkt
      If LängeTextstück <= 3 And LängeTextstück >= 2 Then  'Endzeichen zählt mit
        TagZahl = CInt(TextStück)
      End If
      Call TextScan(Prüfling, Lesezeiger, Punkt)
      If LängeTextstück > 2 Then GoTo DRFehler
      MonZahl = CInt(TextStück)
      If MonZahl > 12 Then Exit Function
      If (MonZahl = 1 Or MonZahl = 3 Or MonZahl = 5 Or MonZahl = 7 Or _
         MonZahl = 8 Or MonZahl = 10 Or MonZahl > 12) _
         And TagZahl > 31 Then GoTo DRFehler
      If MonZahl = 2 And TagZahl > 28 + CalTag Then GoTo DRFehler
      If (MonZahl = 4 Or MonZahl = 6 Or MonZahl = 9 Or MonZahl = 11) _
         And TagZahl > 30 Then GoTo DRFehler
      If TextEndeErreicht = False Then GoTo DRFehler
      IstNumDatum = True
      Exit Function
DRFehler:
      MELDUNG = MELDUNG & Chr(10) & _
      "Tages- und/oder Monatszahl in " & Prüfling & " nicht erkennbar"
      ABBRUCH = True
    End Function 'IstNumDatum
      
 
    
    
    Sub TextScan(Text, StartStelle As Long, EndString As String)
      'out: TextStück (String), LängeTextstück, Lesezeiger (Long),
      '     TextEndeErreicht (Boolean)
      
      'Läuft von der Stelle von "Text", die durch die übergeordnete Variable "LeseZeiger"
      'gegeben ist über die Zeichenfolge aus dem String "Text"  bis der vorgegebene Endstring
      '(einzelnes Zeichen oder String) oder das Ende von "Text" gefunden ist oder das Textende
      'erreicht ist und liefert die bis dorthin gefundenen Zeichen in der übergeordneten
      'Variablen "TextStück" und deren Anzahl in "LängeTextstück". Verstellt den Lesezeiger
      'auf die erste Stelle von "Text" hinter dem Endstring des gefundenen Textstücks bzw. auf
      'die Stelle hinter dem Gesamt-String, falls das gefundene Textstück die letzten Zeichen
      'im Gesamt-String "Text" darstellten. In diesem Falle wird die übergeordnete Boolesche
      'Variable "TextEndeErreicht" auf True gesetzt; sonst ist sie False.
    
    Dim I As Long, LängeText As Long, LängeEndString As Long, Start As Long
    
      Start = StartStelle
      TextEndeErreicht = False
      LängeText = Len(Text)
      LängeEndString = Len(EndString)
      TextStück = ""
      For I = Start To LängeText + 1 Step 1
        If (Mid(Text, I, LängeEndString) = EndString) Or (I >= LängeText + 1) Then
          TextStück = Mid(Text, Start, I - Start)
          Lesezeiger = I
          LängeTextstück = I - Start
          GoTo Aussprung
        End If
      Next I
Aussprung:
      If I < LängeText Then
        Lesezeiger = I + LängeEndString
      Else
        TextEndeErreicht = True
      End If
    End Sub 'TextScan
    
