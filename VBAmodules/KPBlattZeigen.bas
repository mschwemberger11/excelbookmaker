Attribute VB_Name = "KPBlattZeigen"

'****************************************************************************
' Modul M3KoBlaZeigen                                                       *
'****************************************************************************
' Blendet das Kontoblatt ein, auf dessen Kontonummer in Kontenplan, ArProt,
' oder einem Kontoblatt gezeigt wird.
' Tastenkombination: Strg+z
' Benutzt die im Modul Kontenplanpflege definierten Routinen
' Sub KontenplanStruktur()
' Sub KtoKennDat(Kontonummer)
Sub KoBlaZeigen()
Attribute KoBlaZeigen.VB_ProcData.VB_Invoke_Func = "z\n14"
Dim AnfBlatt, AnfZeile, AnfSpalte, ZeilenNr, SpaltenNr, KtoNr, BlattName, A
Dim ZielBlatt As String
'Public KtoNr
With ActiveWindow
  With ActiveSheet          'Aufbewahren Aufrufsituation
    AnfBlatt = ActiveSheet.Name
    AnfZeile = ActiveCell.Row
    AnfSpalte = ActiveCell.Column
    KtoNr = Cells(AnfZeile, AnfSpalte)
  End With
  Call KontenplanStruktur
  If ABBRUCH = True Then Exit Sub
  If KPSAbbruch = True Then
    Call MsgBox(kpsmeldung & Chr(10) & "Auftrag ''Kontoblatt zeigen'' abgebrochen", 0, _
    TiT & " KoBlaZeigen")
    Exit Sub 'oberste Ebene
  End If
'2 ---------------------- aktive Zelle eine Kontonummer? ------------------
  If Not (AnfBlatt = "Kontenplan" Or AnfBlatt = "ArProt") Then
    GoTo IstEinKonto
  End If
  If AnfBlatt = "Kontenplan" And AnfZeile >= 5 And AnfZeile < KPKZEnde Then
    AnfSpalte = 2
    GoTo KontoBlattFinden
  End If '----------------------------
  If AnfBlatt = "ArProt" And AnfZeile >= 3 And _
     (AnfSpalte >= 4 And AnfSpalte <= 5) And _
     IsNumeric(Cells(AnfZeile, AnfSpalte)) = True Then
    GoTo KontoBlattFinden
  End If  '------------------------------
IstEinKonto:
  With Sheets(AnfBlatt)
    .Activate
    If IsNumeric(Cells(AnfZeile, AnfSpalte)) = True Then
      A = MsgBox(prompt:="Ist der Inhalt der aktiven Zelle" & Chr(10) & _
                "eine Kontonummer?", Buttons:=vbYesNo, _
                Title:="Tastenkombination ''Strg+z'' Kontoblatt zeigen")
      If A = vbYes Then
        GoTo KontoBlattFinden
      End If
    Else
      GoTo WirkungsLos
    End If
  End With
KontoBlattFinden:
  With Worksheets(AnfBlatt)
    .Activate
    Call KtoKennDat(Cells(AnfZeile, AnfSpalte).Value)
 '   If abbruch = True Then      'auch Fehlerhafte Blätter zeigen!
 '     Call MsgBox("Vorgang Kontoblatt zeigen abgebrochen", 0, "KtoKennDat in KoBlaZeigen")
 '     Exit Sub
 '   End If
    If AKtoEinricht <> "E" Then
      A = MsgBox(prompt:="Für diese Kontonummer, " & KtoNr & _
                ", ist kein Blatt eingerichtet" & Chr(10) & _
                "Soll es eingerichtet werden?", Buttons:=vbYesNo, _
                Title:="Kontoblatt zeigen")
      If A = vbYes Then
        Call KontoBlattEinrichten(Cells(AnfZeile, AnfSpalte)) 'Daten von KtoKennDat
        Call KtoKennDat(KtoNr)
      End If
      If A = vbNo Then
        GoTo Ausgang
      End If
    End If
  End With 'Worksheets(AnfBlatt)
  With Sheets(AKtoBlatt)
    .Activate
    ZielBlatt = InputBox _
        (prompt:="Wechseln zum angegebenen Blatt: Button OK drücken!", _
                 Title:="Kontoblatt zeigen", Default:=AnfBlatt) ', _
                 'Left:=450, Top:=300)
                 'Type:=2 = String, Type:=4 = Boolean,Type:=6 =
    If ZielBlatt <> "" Then
      Sheets(ZielBlatt).Activate
    End If
   Exit Sub
  End With
WirkungsLos:
  Call MsgBox(prompt:= _
       "Tastenkombination ''Strg+z'' hier wirkungslos" & Chr(10) & _
       "Zelle mit Kontonummer aktivieren", Buttons:=vbOKOnly, _
       Title:="ExAcc Kontoblatt zeigen")
Ausgang:
    With Sheets(AnfBlatt)
      .Activate
      Cells(AnfZeile, AnfSpalte).Activate
    End With
  End With 'ActiveWindow
End Sub 'KoPlaZeigen
'===============================================================================
        
