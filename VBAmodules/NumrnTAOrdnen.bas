Attribute VB_Name = "NumrnTAOrdnen"

'*********************************************************************************
'* Makro NummernTAOrdnen  *                                                      *
'*********************************************************************************
' Tastenkombination: Strg+n
' Wird über Strg+n aufgerufen, wenn Blatt "ArProt" im aktiven Fenster ist und eine
' Zelle in der Spalte 1 aktiv ist.
'
' Benutzt die im Modul "APSchreiben" befindliche Routine
  'Sub BetroffKtoBlatt(BeKoBlätter As String)
  'Liest aus dem Parameter "BeKoBlätter" (ArProt-Spalte "Betroffene Konten") das erste
  'Konto heraus und speichert es als globale Stringvariable "ABZAkBlat". Die weiteren im
  'BeKoBlätter vorhandenen Konten werden als globale Stringvariable "RestString"
  
  'gespeichert.
Option Explicit
'


Sub TANrnOrdnen()
Attribute TANrnOrdnen.VB_ProcData.VB_Invoke_Func = "n\n14"
  Dim A, StartZeile As Long, AnfangsZelle As Range, Zeile As Long, Z As Long, KontoZ As Long
  Dim StartBlatt As String, TANummer, TANummerVor As Integer, AlteTANummer
  Dim BuchungVorhanden As Boolean, BuchungsKonten As String, BuID As Long, Gebucht
  Const TiT = "TA-Nummern ordnen"
With ActiveWindow
  Application.CutCopyMode = False
'TAO 1 ------------------------ Blatt richtig? ----------------------------
  StartBlatt = ActiveSheet.Name
  StartZeile = ActiveCell.Row
  If StartBlatt <> "ArProt" Then
    A = MsgBox("kann nur vom Blatt ''ArProt'' aus verwendet werden.", vbOKCancel, _
                "Tastenkombination ''Strg+n'' " & TiT)
    If A = vbOK Then                 'kein Aktivieren von ArProt, wenn
      Worksheets("ArProt").Activate  'mit Abbrechen quittiert wird
      Cells(ABZeilNr, 1).Activate
    End If
    
    Exit Sub
  End If
'TAO 2 ------------------------ Spalte richtig? ----------------------------
  With Sheets("ArProt")
    If ActiveCell.Column <> APCTANr Then
      A = MsgBox("kann nur von Spalte ''TA'' aus verwendet werden.", vbOKCancel, _
                 "Tastenkombination ''Strg+n'' " & TiT)
      If A = vbCancel Then Exit Sub      'kein Aktivieren einer Zelle in Spalte
      If A = vbOK Then                   'APCTANr, wenn mit Cancel quittiert wird
        Cells(StartZeile, APCTANr).Activate
      End If
    End If
'TAO 3 -------------------- Monatszeile überspringen ----------------------
    If ActiveCell.Value = "TA" And StartZeile > 3 Then
      StartZeile = StartZeile - 1
    End If
'TAO 4 ---------------------- Auftrag bestätigen ---------------------------
    Cells(StartZeile, APCTANr).Activate
    TANummer = ActiveCell.Value
    A = MsgBox("Sollen von der TA-Nr.  " & TANummer & "  in ArProt abwärts" & Chr(10) & _
               "alle TA-Nummern aufsteigend neu numeriert werden?", vbYesNo, _
                 "Tastenkombination ''Strg+n'' " & TiT)
    If A = vbNo Then Exit Sub
    If StartZeile = 3 And TANummer <> 1 Then
      A = MsgBox("Die TA-Nr. der ersten Buchungszeile wird auf 1 gesetzt." & Chr(10) & _
               "Einverstanden?", vbYesNo, _
                 "Tastenkombination ''Strg+n'' " & TiT)
      If A = vbNo Then
        StartZeile = StartZeile + 1
      End If
    End If
  End With  'Sheets("ArProt")
'TAO 5 ------------------ ArProt-Zeilen-Schleife -------------------------
  With Sheets("ArProt")
    .Activate
    For Z = StartZeile To Cells(1, 3).Value
      Sheets("ArProt").Activate
      Cells(Z, APCTANr).Activate
      
      AlteTANummer = ActiveCell.Value
      If Z = 3 Then        'erste ArProt-Zeile erhält mangels vor-
        TANummer = 1       'hergehender Nummern die TANummer 1
        GoTo TANrErmittelt
      End If
      If ActiveCell.Value = "TA" Then  'Monatsüberschriftzeile
        Z = Z + 1
        Cells(Z, APCTANr).Activate
        TANummerVor = Cells(Z - 2, APCTANr)
      Else
        If Cells(Z - 1, APCTANr).Value = "TA" Then
          TANummerVor = Cells(Z - 2, APCTANr)
        Else
          TANummerVor = Cells(Z - 1, APCTANr)
        End If
      End If
      If Cells(Z, APCBuID).Value = 0 And Cells(Z + 1, APCTANr) = "" Then
        TANummer = TANummerVor + 1
        ActiveCell.Value = TANummer  'ArProt-Ende in Kopfzellen:
        Cells(1, 3) = Z              'die nächste Buchungs-Zeile
        Cells(1, 1) = TANummer       'die vorbereitete nächste TANummer
        GoTo EndeOrdnen
      End If
      If ActiveCell = TANummerVor + 1 Or _
         TANummerVor + 1 = AlteTANummer Then  'Monatsüberschriftzeile überspringen
        GoTo NächsteZeile
      Else
        TANummer = TANummerVor + 1
        ActiveCell.Value = TANummer
      End If
'TAO 5 ----------------------------- Endekriterium -------------------------------
TANrErmittelt:
      MELDUNG = ""
      Cells(Z, APCTANr).Activate
'      Call ABZParam
      BuID = Cells(Z, APCBuID).Value
      Gebucht = Cells(Z, APCgebucht).Value
      If BuID = 0 And Gebucht = "" Then GoTo EndeOrdnen
      BuchungsKonten = Cells(Z, APCBetrofKto).Value
'TAO 6 ------------------------ Betroffene-Konten-Schleife -----------------------
'Sub BetroffKtoBlatt(BeKoBlätter As String)
  'liest aus dem Parameter "BeKoBlätter" (ArProt-Spalte "Betroffene Konten") das erste
  'Konto heraus und speichert es als globale Stringvariable "ABZAkBlat". Die weiteren im
  'BeKoBlätter vorhandenen Konten werden als globale Stringvariable "RestString"
  'gespeichert.
TAInKontoblattÄndern:
      If BuchungsKonten = "" Then GoTo NächsteZeile
      Call BetroffKtoBlatt(BuchungsKonten)  'aus Modul ArProtSchreiben
      ABZAkBlat = BeKoBlatt
      BuchungsKonten = RestString
      If ABZAkBlat <> "" Then
        With Sheets(ABZAkBlat)
          .Activate
'TAO 7 --------------------- BuID-Suche im betroffenen Konto ---------------------
          For KontoZ = 6 To Cells(1, 1).Value
            Cells(KontoZ, KoCBuID).Activate
            If Cells(KontoZ, KoCBuID) = BuID Then
              Cells(KontoZ, KoCTANr) = TANummer
              Exit For
            End If
            If KontoZ >= Cells(1, 1).Value Then
              A = MsgBox("In Kontoblatt ''" & ABZAkBlat & "'' fehlt BuID ''" & BuID & _
                           Chr(10) & "Trotzdem weiter?", vbOKCancel, TiT)
              If A = vbCancel Then Exit Sub
'              If A = vbOK Then GoTo TAInKontoblattÄndern
            End If
          Next KontoZ
        End With 'Sheets(ABZAkBlat)
      End If 'ABZAkBlat <> ""
      GoTo TAInKontoblattÄndern
'    End If 'BuID = 0 And Gebucht <> "" And
NächsteZeile:
    Next Z  'ArProt-Zeilenschleife
EndeOrdnen:
    If Z >= Cells(1, 3) Then
      A = MsgBox(prompt:="ArProt-Ende erreicht. TA-Nummern bis Nummer  " & _
          Chr(10) & TANummer & "  neu geordnet.", Buttons:=vbOKOnly, _
          Title:="TA-Nummern ")
    End If
  End With 'Sheets("ArProt")
End With 'ActiveWindow
End Sub 'TANrnOrdnen()

Sub BetroffKtoBlatt(BeKoBlätter As String)
'Liest den ersten Kontoblattnamen aus dem Text "BeKoBlätter" und speichert ihn
'in die globale Vaiable "BekoBlatt", den verbleibenden Text in "RestString".
'BeKoBlätter enthält keinen, einen oder mehrere Namen, die durch den String " + "
'getrennt sind und endet mit einem Blank. Der Kontoblattname darf "+" nicht
'enthalten.  Vom Text RestString sind die ggf. führenden Zeichen " + " bereits
'weggeschnitten.
'Voraussetzung: Die Namen sind zusammenhängend (one Blank) und enthalten kein "+"
Dim I As Integer, Länge As Integer, Zeichen As String
  Länge = Len(BeKoBlätter)
  If Länge = 0 Or Länge = 1 And BeKoBlätter = " " Then 'leerer oder 1-Blank-String
    BeKoBlatt = ""
    RestString = ""
    Exit Sub
  End If
  RestString = BeKoBlätter
  BeKoBlatt = ""
  I = 0
  Do
    I = I + 1   'I ist Scanzeiger
    If I > Länge Or Len(RestString) > 0 And Left(RestString, 1) = " " Then
      BeKoBlatt = BeKoBlätter  'letzter Name
      If Right(BeKoBlätter, 1) = " " Then
        BeKoBlatt = Left(BeKoBlatt, Länge - 1)
        RestString = ""
        Exit Do
      End If
    End If
    Zeichen = Mid(BeKoBlätter, I, 1)
    If Zeichen = " " Then
      BeKoBlatt = Left(BeKoBlätter, I - 1)
      RestString = Right(BeKoBlätter, Länge - I)
      If Left(RestString, 2) = "+ " Then
        RestString = Right(BeKoBlätter, Länge - I - 2)
        Exit Do
      End If
    End If
  Loop
End Sub 'BetroffKtoBlatt

