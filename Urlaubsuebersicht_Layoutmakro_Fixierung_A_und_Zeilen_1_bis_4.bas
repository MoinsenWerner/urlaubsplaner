Attribute VB_Name = "UrlaubsuebersichtLayout"
Option Explicit

Private Const AUSGABEBLATT_NAME As String = "Detaillierung"

' False: Das ursprüngliche Blatt wird nach erfolgreicher Umwandlung gelöscht.
' True:  Das ursprüngliche Blatt bleibt als sehr ausgeblendete Sicherung erhalten.
Private Const ORIGINALBLATT_ALS_BACKUP_BEHALTEN As Boolean = False

' False sorgt dafür, dass sämtliche Datumsspalten sichtbar bleiben.
' Es werden keine vergangenen Monate ausgeblendet.
Private Const VERGANGENE_MONATE_AUSBLENDEN As Boolean = False

' False sorgt dafür, dass auch die erste Jahresspalte sichtbar bleibt.
Private Const ERSTE_JAHRESSPALTE_AUSBLENDEN As Boolean = False

' Fixiert ausschließlich Spalte A sowie die Zeilen 1 bis 4.
Private Const KOPFZEILEN_FIXIEREN As Boolean = True

Private Const QUELL_DATUMSZEILE As Long = 1
Private Const QUELL_FERIENZEILE As Long = 2
Private Const QUELL_ERSTE_PERSONENZEILE As Long = 3

Private Const ZIEL_MONATSZEILE As Long = 2
Private Const ZIEL_KWZEILE As Long = 3
Private Const ZIEL_TAGESZEILE As Long = 4
Private Const ZIEL_ERSTE_PERSONENZEILE As Long = 5

Private Const TRENNZEILE_MARKER As String = "{TRENNZEILE}"

Public Sub UrlaubsuebersichtWieVorlageFormatieren()

    Dim wb As Workbook
    Dim quellBlatt As Worksheet
    Dim zielBlatt As Worksheet
    Dim altesBlattName As String
    Dim tempBlattName As String

    Dim alteBerechnung As XlCalculation
    Dim alterBildschirmstatus As Boolean
    Dim alteEreignisse As Boolean
    Dim alteWarnungen As Boolean

    Dim letzteQuellSpalte As Long
    Dim legendenZeile As Long
    Dim letztePersonenZeile As Long
    Dim letzteQuellZeile As Long

    Dim quellSpalten() As Long
    Dim zielSpalten() As Long
    Dim datumswerte() As Date
    Dim datumsAnzahl As Long

    Dim jahresSpalten() As Long
    Dim jahresWerte() As Long
    Dim jahresAnzahl As Long

    Dim namensZeilen As Object
    Dim bereitsHinzugefuegt As Object
    Dim ausgabeReihenfolge As Collection

    Dim gruppeOben As Variant
    Dim gruppeUnten As Variant

    Dim letzteZielSpalte As Long
    Dim letzteDatenZeile As Long
    Dim ersteLegendenZeile As Long
    Dim letzteLegendenZeile As Long

    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim zielSpalte As Long
    Dim vorherigesJahr As Long
    Dim personName As String
    Dim eintrag As Variant
    Dim datumWert As Date
    Dim fehlerNummer As Long
    Dim fehlerText As String
    Dim erstellungErfolgreich As Boolean

    alterBildschirmstatus = Application.ScreenUpdating
    alteEreignisse = Application.EnableEvents
    alteWarnungen = Application.DisplayAlerts
    alteBerechnung = Application.Calculation

    On Error GoTo Fehler

    If ActiveWorkbook Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Es ist keine Arbeitsmappe aktiv."
    End If

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Err.Raise vbObjectError + 1001, , "Das aktive Blatt ist kein Excel-Arbeitsblatt."
    End If

    Set wb = ActiveWorkbook
    Set quellBlatt = ActiveSheet
    altesBlattName = quellBlatt.Name

    If BlattExistiert(wb, AUSGABEBLATT_NAME) Then
        If StrComp(quellBlatt.Name, AUSGABEBLATT_NAME, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1002, , _
                "In der Arbeitsmappe existiert bereits ein Blatt namens """ & _
                AUSGABEBLATT_NAME & """. Bitte dieses Blatt zuerst umbenennen oder löschen."
        End If
    End If

    letzteQuellSpalte = quellBlatt.Cells(QUELL_DATUMSZEILE, quellBlatt.Columns.Count).End(xlToLeft).Column
    If letzteQuellSpalte < 2 Then
        Err.Raise vbObjectError + 1003, , "In Zeile 1 wurden keine Datumsspalten gefunden."
    End If

    legendenZeile = ZeileMitExaktemText(quellBlatt, 1, "Legende")
    If legendenZeile = 0 Then
        Err.Raise vbObjectError + 1004, , "In Spalte A wurde keine Zelle mit dem Text ""Legende"" gefunden."
    End If

    letztePersonenZeile = legendenZeile - 1
    letzteQuellZeile = quellBlatt.Cells(quellBlatt.Rows.Count, 1).End(xlUp).Row

    ReDim quellSpalten(1 To letzteQuellSpalte - 1)
    ReDim datumswerte(1 To letzteQuellSpalte - 1)

    For c = 2 To letzteQuellSpalte
        datumWert = DatumAusKopf(quellBlatt.Cells(QUELL_DATUMSZEILE, c).Value2)

        If datumWert <> 0 Then
            datumsAnzahl = datumsAnzahl + 1
            quellSpalten(datumsAnzahl) = c
            datumswerte(datumsAnzahl) = datumWert
        End If
    Next c

    If datumsAnzahl = 0 Then
        Err.Raise vbObjectError + 1005, , _
            "Die Datumsüberschriften konnten nicht ausgewertet werden. Erwartet wird zum Beispiel ""Do 01.01.2026""."
    End If

    ReDim Preserve quellSpalten(1 To datumsAnzahl)
    ReDim Preserve datumswerte(1 To datumsAnzahl)
    ReDim zielSpalten(1 To datumsAnzahl)
    ReDim jahresSpalten(1 To datumsAnzahl)
    ReDim jahresWerte(1 To datumsAnzahl)

    zielSpalte = 1
    vorherigesJahr = -1

    For i = 1 To datumsAnzahl
        If Year(datumswerte(i)) <> vorherigesJahr Then
            jahresAnzahl = jahresAnzahl + 1
            zielSpalte = zielSpalte + 1

            jahresSpalten(jahresAnzahl) = zielSpalte
            jahresWerte(jahresAnzahl) = Year(datumswerte(i))
            vorherigesJahr = Year(datumswerte(i))
        End If

        zielSpalte = zielSpalte + 1
        zielSpalten(i) = zielSpalte
    Next i

    ReDim Preserve jahresSpalten(1 To jahresAnzahl)
    ReDim Preserve jahresWerte(1 To jahresAnzahl)
    letzteZielSpalte = zielSpalte

    Set namensZeilen = CreateObject("Scripting.Dictionary")
    namensZeilen.CompareMode = vbTextCompare

    For r = QUELL_ERSTE_PERSONENZEILE To letztePersonenZeile
        personName = Trim$(CStr(quellBlatt.Cells(r, 1).Value2))

        If Len(personName) > 0 And StrComp(personName, "Admin Benutzer", vbTextCompare) <> 0 Then
            If namensZeilen.Exists(personName) Then
                Err.Raise vbObjectError + 1006, , _
                    "Der Name """ & personName & """ kommt mehrfach in Spalte A vor."
            End If

            namensZeilen.Add personName, r
        End If
    Next r

    gruppeOben = Array( _
        "Heiko Koch", _
        "Alexander Hälter", _
        "Bianca Stark", _
        "Jakob Würfel", _
        "Robert Bartsch", _
        "Simge Tepekesici", _
        "Nadine Bauer (TZ)", _
        "Stefanie Mahr (TZ)", _
        "Barbara Koch (TZ) HO" _
    )

    gruppeUnten = Array( _
        "Nils Gerber", _
        "Felix Kaiser", _
        "Laura Löffler" _
    )

    Set bereitsHinzugefuegt = CreateObject("Scripting.Dictionary")
    bereitsHinzugefuegt.CompareMode = vbTextCompare
    Set ausgabeReihenfolge = New Collection

    For Each eintrag In gruppeOben
        personName = CStr(eintrag)

        If namensZeilen.Exists(personName) Then
            ausgabeReihenfolge.Add personName
            bereitsHinzugefuegt.Add personName, True
        End If
    Next eintrag

    ' Neue oder unbekannte Namen bleiben erhalten und werden vor der Trennzeile
    ' in ihrer ursprünglichen Reihenfolge eingefügt.
    For r = QUELL_ERSTE_PERSONENZEILE To letztePersonenZeile
        personName = Trim$(CStr(quellBlatt.Cells(r, 1).Value2))

        If Len(personName) > 0 _
           And StrComp(personName, "Admin Benutzer", vbTextCompare) <> 0 _
           And Not bereitsHinzugefuegt.Exists(personName) _
           And Not WertIstInArray(personName, gruppeUnten) Then

            ausgabeReihenfolge.Add personName
            bereitsHinzugefuegt.Add personName, True
        End If
    Next r

    For Each eintrag In gruppeUnten
        If namensZeilen.Exists(CStr(eintrag)) Then
            ausgabeReihenfolge.Add TRENNZEILE_MARKER
            Exit For
        End If
    Next eintrag

    For Each eintrag In gruppeUnten
        personName = CStr(eintrag)

        If namensZeilen.Exists(personName) Then
            ausgabeReihenfolge.Add personName
            bereitsHinzugefuegt.Add personName, True
        End If
    Next eintrag

    If ausgabeReihenfolge.Count = 0 Then
        Err.Raise vbObjectError + 1007, , "Zwischen der Ferienzeile und der Legende wurden keine Personen gefunden."
    End If

    letzteDatenZeile = ZIEL_ERSTE_PERSONENZEILE + ausgabeReihenfolge.Count - 1
    ersteLegendenZeile = letzteDatenZeile + 2
    letzteLegendenZeile = ersteLegendenZeile + (12 - 1) * 2

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Urlaubsübersicht wird neu aufgebaut ..."

    quellBlatt.Calculate

    tempBlattName = EindeutigerBlattname(wb, "__Detaillierung_neu")
    Set zielBlatt = wb.Worksheets.Add(After:=quellBlatt)
    zielBlatt.Name = tempBlattName

    GrundlayoutAnlegen _
        quellBlatt, _
        zielBlatt, _
        quellSpalten, _
        zielSpalten, _
        datumswerte, _
        datumsAnzahl, _
        jahresSpalten, _
        jahresWerte, _
        jahresAnzahl, _
        letzteZielSpalte, _
        ausgabeReihenfolge, _
        namensZeilen, _
        letzteDatenZeile, _
        ersteLegendenZeile, _
        legendenZeile, _
        letzteQuellZeile

    zielBlatt.Activate

    ' Sicherheitshalber sämtliche Spalten einblenden.
    zielBlatt.Columns.Hidden = False

    If KOPFZEILEN_FIXIEREN Then
        With ActiveWindow
            .FreezePanes = False
            .SplitRow = 0
            .SplitColumn = 0

            ' Zeilen 1 bis 4 über die gesamte Tabellenbreite fixieren.
            .SplitRow = 4

            ' Spalte A über die gesamte Tabellenhöhe fixieren.
            .SplitColumn = 1

            .FreezePanes = True
        End With
    End If

    ActiveWindow.Zoom = 130
    ActiveWindow.DisplayGridlines = True

    If ORIGINALBLATT_ALS_BACKUP_BEHALTEN Then
        quellBlatt.Name = EindeutigerBlattname(wb, "__Original_Backup")
        quellBlatt.Visible = xlSheetVeryHidden
    Else
        quellBlatt.Delete
    End If

    zielBlatt.Name = AUSGABEBLATT_NAME
    erstellungErfolgreich = True

    ZielAnsichtAufAktuellenMonat _
        zielBlatt, _
        zielSpalten, _
        datumswerte, _
        datumsAnzahl

    zielBlatt.Cells(1, 1).Select

Aufraeumen:
    Application.StatusBar = False
    Application.Calculation = alteBerechnung
    Application.DisplayAlerts = alteWarnungen
    Application.EnableEvents = alteEreignisse
    Application.ScreenUpdating = alterBildschirmstatus

    If erstellungErfolgreich Then
        MsgBox _
            "Die Übersicht wurde in das Layout der Datei ""gewünschtes-ergebnis.xlsx"" umgewandelt." & vbCrLf & _
            "Die vorhandenen Einträge und sichtbaren Zellfarben wurden aus dem ursprünglichen Blatt übernommen." & vbCrLf & vbCrLf & _
            "Die Arbeitsmappe wurde nicht automatisch gespeichert.", _
            vbInformation, _
            "Umwandlung abgeschlossen"
    End If

    Exit Sub

Fehler:
    fehlerNummer = Err.Number
    fehlerText = Err.Description

    On Error Resume Next

    If Not zielBlatt Is Nothing Then
        If Not erstellungErfolgreich Then
            zielBlatt.Delete
        End If
    End If

    On Error GoTo 0

    Application.StatusBar = False
    Application.Calculation = alteBerechnung
    Application.DisplayAlerts = alteWarnungen
    Application.EnableEvents = alteEreignisse
    Application.ScreenUpdating = alterBildschirmstatus

    MsgBox _
        "Die Umwandlung wurde abgebrochen." & vbCrLf & vbCrLf & _
        "Fehler " & CStr(fehlerNummer) & ": " & fehlerText, _
        vbCritical, _
        "Fehler"

End Sub

Private Sub GrundlayoutAnlegen( _
    ByVal quellBlatt As Worksheet, _
    ByVal zielBlatt As Worksheet, _
    ByRef quellSpalten() As Long, _
    ByRef zielSpalten() As Long, _
    ByRef datumswerte() As Date, _
    ByVal datumsAnzahl As Long, _
    ByRef jahresSpalten() As Long, _
    ByRef jahresWerte() As Long, _
    ByVal jahresAnzahl As Long, _
    ByVal letzteZielSpalte As Long, _
    ByVal ausgabeReihenfolge As Collection, _
    ByVal namensZeilen As Object, _
    ByVal letzteDatenZeile As Long, _
    ByVal ersteLegendenZeile As Long, _
    ByVal quellLegendenZeile As Long, _
    ByVal letzteQuellZeile As Long)

    Dim gesamterBereich As Range
    Dim bereich As Range
    Dim quellZelle As Range
    Dim zielZelle As Range

    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim ausgabeZeile As Long
    Dim monatsStart As Long
    Dim istMonatsEnde As Boolean
    Dim personName As String
    Dim eintrag As Variant
    Dim stichtag As Date
    Dim jahresSpalte As Long
    Dim allesVorStichtag As Boolean
    Dim ersterSichtbarerDatumsspalte As Long

    Set gesamterBereich = zielBlatt.Range( _
        zielBlatt.Cells(1, 1), _
        zielBlatt.Cells(ersteLegendenZeile + 22, letzteZielSpalte) _
    )

    With gesamterBereich
        .Font.Name = "Arial"
        .Font.Size = 8
        .VerticalAlignment = xlCenter
    End With

    zielBlatt.Columns(1).ColumnWidth = 20.7109375

    For i = 1 To jahresAnzahl
        zielBlatt.Columns(jahresSpalten(i)).ColumnWidth = 7.42578125
    Next i

    For i = 1 To datumsAnzahl
        zielBlatt.Columns(zielSpalten(i)).ColumnWidth = 3.28515625
    Next i

    zielBlatt.Rows("1:" & CStr(ersteLegendenZeile + 22)).RowHeight = 15.75
    zielBlatt.Rows(ZIEL_KWZEILE).RowHeight = 15

    Set bereich = zielBlatt.Range( _
        zielBlatt.Cells(1, 1), _
        zielBlatt.Cells(1, letzteZielSpalte) _
    )
    RahmenSetzen bereich, xlEdgeBottom, xlMedium

    zielBlatt.Cells(ZIEL_MONATSZEILE, 1).Value = "Monat"
    zielBlatt.Cells(ZIEL_KWZEILE, 1).Value = "KW"
    zielBlatt.Cells(ZIEL_TAGESZEILE, 1).Value = "Ferienzeit"

    With zielBlatt.Range( _
        zielBlatt.Cells(ZIEL_MONATSZEILE, 1), _
        zielBlatt.Cells(ZIEL_TAGESZEILE, 1) _
    )
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    RahmenSetzen zielBlatt.Cells(ZIEL_MONATSZEILE, 1), xlEdgeLeft, xlMedium
    RahmenSetzen zielBlatt.Cells(ZIEL_KWZEILE, 1), xlEdgeLeft, xlMedium
    RahmenSetzen zielBlatt.Cells(ZIEL_TAGESZEILE, 1), xlEdgeLeft, xlMedium

    Set quellZelle = ErsteZelleMitTextInZeile(quellBlatt, QUELL_FERIENZEILE, 2, "Ferien")
    If Not quellZelle Is Nothing Then
        SichtbareFarbenKopieren quellZelle, zielBlatt.Cells(ZIEL_TAGESZEILE, 1)
    Else
        zielBlatt.Cells(ZIEL_TAGESZEILE, 1).Interior.Color = RGB(221, 235, 247)
    End If

    For i = 1 To jahresAnzahl
        jahresSpalte = jahresSpalten(i)

        Set bereich = zielBlatt.Range( _
            zielBlatt.Cells(ZIEL_MONATSZEILE, jahresSpalte), _
            zielBlatt.Cells(letzteDatenZeile, jahresSpalte) _
        )

        bereich.Merge
        bereich.Value = jahresWerte(i)

        With bereich
            .Interior.Color = RGB(64, 64, 64)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 24
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 90
        End With

        RahmenAussen bereich, xlMedium
    Next i

    monatsStart = 1

    For i = 1 To datumsAnzahl
        istMonatsEnde = False

        If i = datumsAnzahl Then
            istMonatsEnde = True
        ElseIf Month(datumswerte(i + 1)) <> Month(datumswerte(i)) _
            Or Year(datumswerte(i + 1)) <> Year(datumswerte(i)) Then
            istMonatsEnde = True
        End If

        If istMonatsEnde Then
            Set bereich = zielBlatt.Range( _
                zielBlatt.Cells(ZIEL_MONATSZEILE, zielSpalten(monatsStart)), _
                zielBlatt.Cells(ZIEL_MONATSZEILE, zielSpalten(i)) _
            )

            bereich.Merge
            bereich.Value = MonatsnameDeutsch(Month(datumswerte(i)))

            With bereich
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            RahmenAussen bereich, xlMedium
            monatsStart = i + 1
        End If
    Next i

    For i = 1 To datumsAnzahl
        Set zielZelle = zielBlatt.Cells(ZIEL_KWZEILE, zielSpalten(i))

        With zielZelle
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        If Weekday(datumswerte(i), vbMonday) = 1 Then
            zielZelle.Value = ISOJahreswoche(datumswerte(i))
            zielZelle.Interior.Color = RGB(191, 191, 191)
        End If

        Set zielZelle = zielBlatt.Cells(ZIEL_TAGESZEILE, zielSpalten(i))

        With zielZelle
            .Value = datumswerte(i)
            .NumberFormat = "dd"
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        SichtbareFarbenKopieren _
            quellBlatt.Cells(QUELL_FERIENZEILE, quellSpalten(i)), _
            zielZelle

        If Weekday(datumswerte(i), vbMonday) = 5 Then
            RahmenSetzen zielBlatt.Cells(ZIEL_KWZEILE, zielSpalten(i)), xlEdgeRight, xlMedium
            RahmenSetzen zielBlatt.Cells(ZIEL_TAGESZEILE, zielSpalten(i)), xlEdgeRight, xlMedium
        End If
    Next i

    ausgabeZeile = ZIEL_ERSTE_PERSONENZEILE

    For Each eintrag In ausgabeReihenfolge

        If CStr(eintrag) = TRENNZEILE_MARKER Then
            TrennzeileFormatieren _
                zielBlatt, _
                ausgabeZeile, _
                zielSpalten, _
                datumswerte, _
                datumsAnzahl
        Else
            personName = CStr(eintrag)
            r = CLng(namensZeilen(personName))

            With zielBlatt.Cells(ausgabeZeile, 1)
                .Value = personName
                .Font.Name = "Arial"
                .Font.Size = 8
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With

            RahmenSetzen zielBlatt.Cells(ausgabeZeile, 1), xlEdgeLeft, xlMedium
            RahmenSetzen zielBlatt.Cells(ausgabeZeile, 1), xlEdgeTop, xlMedium
            RahmenSetzen zielBlatt.Cells(ausgabeZeile, 1), xlEdgeBottom, xlMedium

            For i = 1 To datumsAnzahl
                Set quellZelle = quellBlatt.Cells(r, quellSpalten(i))
                Set zielZelle = zielBlatt.Cells(ausgabeZeile, zielSpalten(i))

                ZellinhaltKopieren quellZelle, zielZelle
                SichtbareFarbenKopieren quellZelle, zielZelle

                With zielZelle
                    .Font.Name = "Arial"
                    .Font.Size = 8
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                RahmenSetzen zielZelle, xlEdgeTop, xlMedium
                RahmenSetzen zielZelle, xlEdgeBottom, xlMedium

                If Weekday(datumswerte(i), vbMonday) = 5 Then
                    RahmenSetzen zielZelle, xlEdgeRight, xlMedium
                End If
            Next i
        End If

        ausgabeZeile = ausgabeZeile + 1
    Next eintrag

    ' Rechte Abschlusskante hinter dem letzten Datum.
    For r = ZIEL_KWZEILE To letzteDatenZeile
        RahmenSetzen zielBlatt.Cells(r, zielSpalten(datumsAnzahl)), xlEdgeRight, xlMedium
    Next r

    LegendeAnlegen _
        quellBlatt, _
        zielBlatt, _
        ersteLegendenZeile, _
        quellLegendenZeile, _
        letzteQuellZeile, _
        letztePersonenZeile:=quellLegendenZeile - 1, _
        letzteQuellSpalte:=quellSpalten(datumsAnzahl)

    If VERGANGENE_MONATE_AUSBLENDEN Then
        stichtag = DateSerial(Year(Date), Month(Date), 1)

        For i = 1 To datumsAnzahl
            If datumswerte(i) >= stichtag Then
                ersterSichtbarerDatumsspalte = zielSpalten(i)
                Exit For
            End If
        Next i

        ' Nur ausblenden, wenn die Quelldatei mindestens einen aktuellen
        ' oder zukünftigen Monat enthält. Dadurch bleibt eine rein historische
        ' Datei weiterhin vollständig sichtbar.
        If ersterSichtbarerDatumsspalte > 0 Then
            For i = 1 To datumsAnzahl
                If datumswerte(i) < stichtag Then
                    zielBlatt.Columns(zielSpalten(i)).Hidden = True
                End If
            Next i

            For i = 1 To jahresAnzahl
                allesVorStichtag = True

                For j = 1 To datumsAnzahl
                    If Year(datumswerte(j)) = jahresWerte(i) _
                       And datumswerte(j) >= stichtag Then
                        allesVorStichtag = False
                        Exit For
                    End If
                Next j

                If allesVorStichtag Then
                    zielBlatt.Columns(jahresSpalten(i)).Hidden = True
                End If
            Next i
        End If
    End If

    If ERSTE_JAHRESSPALTE_AUSBLENDEN And jahresAnzahl > 0 Then
        zielBlatt.Columns(jahresSpalten(1)).Hidden = True
    End If

End Sub

Private Sub TrennzeileFormatieren( _
    ByVal zielBlatt As Worksheet, _
    ByVal zeile As Long, _
    ByRef zielSpalten() As Long, _
    ByRef datumswerte() As Date, _
    ByVal datumsAnzahl As Long)

    Dim i As Long
    Dim zelle As Range

    Set zelle = zielBlatt.Cells(zeile, 1)

    With zelle
        .ClearContents
        .Interior.Color = RGB(89, 89, 89)
    End With

    RahmenSetzen zelle, xlEdgeLeft, xlMedium
    RahmenSetzen zelle, xlEdgeTop, xlMedium
    RahmenSetzen zelle, xlEdgeBottom, xlMedium

    For i = 1 To datumsAnzahl
        Set zelle = zielBlatt.Cells(zeile, zielSpalten(i))

        With zelle
            .ClearContents
            .Interior.Color = RGB(89, 89, 89)
        End With

        RahmenSetzen zelle, xlEdgeTop, xlMedium
        RahmenSetzen zelle, xlEdgeBottom, xlMedium

        If Weekday(datumswerte(i), vbMonday) = 5 Then
            RahmenSetzen zelle, xlEdgeRight, xlMedium
        End If
    Next i

End Sub

Private Sub LegendeAnlegen( _
    ByVal quellBlatt As Worksheet, _
    ByVal zielBlatt As Worksheet, _
    ByVal ersteLegendenZeile As Long, _
    ByVal quellLegendenZeile As Long, _
    ByVal letzteQuellZeile As Long, _
    ByVal letztePersonenZeile As Long, _
    ByVal letzteQuellSpalte As Long)

    Dim legendenTexte As Variant
    Dim legendenCodes As Variant
    Dim i As Long
    Dim zielZeile As Long
    Dim farbQuelle As Range
    Dim zielZelle As Range

    legendenTexte = Array( _
        "Geplant oder Beantragt", _
        "Genehmigt", _
        "Feiertag", _
        "Grundkurs", _
        "Berufsschule", _
        "Ausbildungsmesse", _
        "PV+BA+AP", _
        "Ausbildung", _
        "Weihnachtsputz", _
        "Kurzarbeit", _
        "Weiterbildung", _
        "Kein Arbeitstag" _
    )

    legendenCodes = Array( _
        "UB", _
        "UG", _
        "{FEIERTAG}", _
        "GK", _
        "BS", _
        "AM", _
        "PV|BA|AP", _
        "AZ", _
        "WP", _
        "KA", _
        "WB", _
        "KAT" _
    )

    For i = LBound(legendenTexte) To UBound(legendenTexte)
        zielZeile = ersteLegendenZeile + i * 2
        Set zielZelle = zielBlatt.Cells(zielZeile, 1)

        With zielZelle
            .Value = CStr(legendenTexte(i))
            .Font.Name = "Arial"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        Set farbQuelle = QuellzelleFuerLegendenfarbe( _
            quellBlatt, _
            CStr(legendenCodes(i)), _
            quellLegendenZeile, _
            letzteQuellZeile, _
            letztePersonenZeile, _
            letzteQuellSpalte _
        )

        If Not farbQuelle Is Nothing Then
            SichtbareFarbenKopieren farbQuelle, zielZelle
        Else
            LegendenFallbackFarbe zielZelle, CStr(legendenCodes(i))
        End If

        KontrastreicheSchriftSetzen zielZelle
    Next i

End Sub

Private Function QuellzelleFuerLegendenfarbe( _
    ByVal quellBlatt As Worksheet, _
    ByVal codeListe As String, _
    ByVal quellLegendenZeile As Long, _
    ByVal letzteQuellZeile As Long, _
    ByVal letztePersonenZeile As Long, _
    ByVal letzteQuellSpalte As Long) As Range

    Dim teile As Variant
    Dim code As Variant
    Dim c As Long
    Dim textWert As String
    Dim fundstelle As Range
    Dim suchBereich As Range

    If codeListe = "{FEIERTAG}" Then
        For c = 2 To letzteQuellSpalte
            textWert = Trim$(CStr(quellBlatt.Cells(QUELL_FERIENZEILE, c).Value2))

            If Len(textWert) > 0 _
               And StrComp(textWert, "Ferien", vbTextCompare) <> 0 Then
                Set QuellzelleFuerLegendenfarbe = quellBlatt.Cells(QUELL_FERIENZEILE, c)
                Exit Function
            End If
        Next c

        Exit Function
    End If

    teile = Split(codeListe, "|")

    ' Zuerst die Farbzelle der ursprünglichen Legende verwenden.
    For Each code In teile
        Set suchBereich = quellBlatt.Range( _
            quellBlatt.Cells(quellLegendenZeile + 1, 1), _
            quellBlatt.Cells(letzteQuellZeile, 1) _
        )

        Set fundstelle = suchBereich.Find( _
            What:=CStr(code), _
            After:=suchBereich.Cells(suchBereich.Cells.Count), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False _
        )

        If Not fundstelle Is Nothing Then
            Set QuellzelleFuerLegendenfarbe = quellBlatt.Cells(fundstelle.Row, 3)
            Exit Function
        End If
    Next code

    ' Falls die Legende geändert wurde, eine tatsächlich belegte Datenzelle suchen.
    Set suchBereich = quellBlatt.Range( _
        quellBlatt.Cells(QUELL_ERSTE_PERSONENZEILE, 2), _
        quellBlatt.Cells(letztePersonenZeile, letzteQuellSpalte) _
    )

    For Each code In teile
        Set fundstelle = suchBereich.Find( _
            What:=CStr(code), _
            After:=suchBereich.Cells(suchBereich.Cells.Count), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False _
        )

        If Not fundstelle Is Nothing Then
            Set QuellzelleFuerLegendenfarbe = fundstelle
            Exit Function
        End If
    Next code

End Function

Private Sub LegendenFallbackFarbe(ByVal zielZelle As Range, ByVal codeListe As String)

    Select Case codeListe
        Case "UB"
            zielZelle.Interior.Color = RGB(255, 255, 0)
        Case "UG"
            zielZelle.Interior.Color = RGB(144, 238, 144)
        Case "{FEIERTAG}"
            zielZelle.Interior.Color = RGB(0, 176, 240)
        Case "GK"
            zielZelle.Interior.Color = RGB(196, 164, 132)
        Case "BS"
            zielZelle.Interior.Color = RGB(64, 224, 208)
        Case "AM"
            zielZelle.Interior.Color = RGB(255, 105, 180)
        Case "PV|BA|AP"
            zielZelle.Interior.Color = RGB(176, 132, 204)
        Case "AZ"
            zielZelle.Interior.Color = RGB(0, 97, 0)
        Case "WP"
            zielZelle.Interior.Color = RGB(101, 67, 33)
        Case "KA"
            zielZelle.Interior.Color = RGB(255, 0, 0)
        Case "WB"
            zielZelle.Interior.Color = RGB(255, 165, 0)
        Case "KAT"
            zielZelle.Interior.Color = RGB(0, 0, 0)
    End Select

End Sub

Private Sub ZellinhaltKopieren(ByVal quellZelle As Range, ByVal zielZelle As Range)

    If quellZelle.HasFormula Then
        zielZelle.FormulaR1C1 = quellZelle.FormulaR1C1
    Else
        zielZelle.Value2 = quellZelle.Value2
    End If

End Sub

Private Sub SichtbareFarbenKopieren(ByVal quellZelle As Range, ByVal zielZelle As Range)

    On Error GoTo NormalesFormat

    If quellZelle.DisplayFormat.Interior.Pattern = xlNone Then
        zielZelle.Interior.Pattern = xlNone
    Else
        zielZelle.Interior.Pattern = xlSolid
        zielZelle.Interior.Color = quellZelle.DisplayFormat.Interior.Color
    End If

    zielZelle.Font.Color = quellZelle.DisplayFormat.Font.Color
    Exit Sub

NormalesFormat:
    On Error Resume Next

    If quellZelle.Interior.Pattern = xlNone Then
        zielZelle.Interior.Pattern = xlNone
    Else
        zielZelle.Interior.Pattern = xlSolid
        zielZelle.Interior.Color = quellZelle.Interior.Color
    End If

    zielZelle.Font.Color = quellZelle.Font.Color
    On Error GoTo 0

End Sub

Private Sub KontrastreicheSchriftSetzen(ByVal zelle As Range)

    Dim farbe As Long
    Dim rot As Long
    Dim gruen As Long
    Dim blau As Long
    Dim helligkeit As Double

    If zelle.Interior.Pattern = xlNone Then
        zelle.Font.Color = RGB(0, 0, 0)
        Exit Sub
    End If

    farbe = zelle.Interior.Color
    rot = farbe Mod 256
    gruen = (farbe \ 256) Mod 256
    blau = (farbe \ 65536) Mod 256

    helligkeit = rot * 0.299 + gruen * 0.587 + blau * 0.114

    If helligkeit < 115 Then
        zelle.Font.Color = RGB(255, 255, 255)
    Else
        zelle.Font.Color = RGB(0, 0, 0)
    End If

End Sub

Private Sub RahmenAussen(ByVal bereich As Range, ByVal gewicht As XlBorderWeight)

    RahmenSetzen bereich, xlEdgeLeft, gewicht
    RahmenSetzen bereich, xlEdgeTop, gewicht
    RahmenSetzen bereich, xlEdgeRight, gewicht
    RahmenSetzen bereich, xlEdgeBottom, gewicht

End Sub

Private Sub RahmenSetzen( _
    ByVal bereich As Range, _
    ByVal position As XlBordersIndex, _
    ByVal gewicht As XlBorderWeight)

    With bereich.Borders(position)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = gewicht
    End With

End Sub

Private Function DatumAusKopf(ByVal wert As Variant) As Date

    Dim textWert As String
    Dim teile As Variant
    Dim datumTeile As Variant
    Dim i As Long
    Dim tag As Long
    Dim monat As Long
    Dim jahr As Long
    Dim pruefDatum As Date

    On Error GoTo KeinDatum

    If IsDate(wert) And VarType(wert) <> vbString Then
        DatumAusKopf = CDate(wert)
        Exit Function
    End If

    textWert = Trim$(Replace(CStr(wert), ChrW(160), " "))
    If Len(textWert) = 0 Then Exit Function

    teile = Split(textWert, " ")

    For i = UBound(teile) To LBound(teile) Step -1
        If InStr(1, CStr(teile(i)), ".", vbBinaryCompare) > 0 Then
            datumTeile = Split(CStr(teile(i)), ".")

            If UBound(datumTeile) = 2 Then
                If IsNumeric(datumTeile(0)) _
                   And IsNumeric(datumTeile(1)) _
                   And IsNumeric(datumTeile(2)) Then

                    tag = CLng(datumTeile(0))
                    monat = CLng(datumTeile(1))
                    jahr = CLng(datumTeile(2))
                    pruefDatum = DateSerial(jahr, monat, tag)

                    If Day(pruefDatum) = tag _
                       And Month(pruefDatum) = monat _
                       And Year(pruefDatum) = jahr Then
                        DatumAusKopf = pruefDatum
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    If IsDate(textWert) Then
        DatumAusKopf = CDate(textWert)
    End If

    Exit Function

KeinDatum:
    DatumAusKopf = 0

End Function

Private Function ISOJahreswoche(ByVal datumWert As Date) As Long

    ISOJahreswoche = DatePart("ww", datumWert, vbMonday, vbFirstFourDays)

End Function

Private Function MonatsnameDeutsch(ByVal monat As Long) As String

    Dim namen As Variant

    namen = Array( _
        "", _
        "Januar", _
        "Februar", _
        "März", _
        "April", _
        "Mai", _
        "Juni", _
        "Juli", _
        "August", _
        "September", _
        "Oktober", _
        "November", _
        "Dezember" _
    )

    MonatsnameDeutsch = CStr(namen(monat))

End Function

Private Function ZeileMitExaktemText( _
    ByVal blatt As Worksheet, _
    ByVal spalte As Long, _
    ByVal suchText As String) As Long

    Dim suchBereich As Range
    Dim fundstelle As Range

    Set suchBereich = blatt.Columns(spalte)

    Set fundstelle = suchBereich.Find( _
        What:=suchText, _
        After:=suchBereich.Cells(suchBereich.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )

    If Not fundstelle Is Nothing Then
        ZeileMitExaktemText = fundstelle.Row
    End If

End Function

Private Function ErsteZelleMitTextInZeile( _
    ByVal blatt As Worksheet, _
    ByVal zeile As Long, _
    ByVal ersteSpalte As Long, _
    ByVal suchText As String) As Range

    Dim letzteSpalte As Long
    Dim c As Long
    Dim textWert As String

    letzteSpalte = blatt.Cells(zeile, blatt.Columns.Count).End(xlToLeft).Column

    For c = ersteSpalte To letzteSpalte
        textWert = Trim$(CStr(blatt.Cells(zeile, c).Value2))

        If StrComp(textWert, suchText, vbTextCompare) = 0 Then
            Set ErsteZelleMitTextInZeile = blatt.Cells(zeile, c)
            Exit Function
        End If
    Next c

End Function

Private Function WertIstInArray(ByVal suchWert As String, ByVal werte As Variant) As Boolean

    Dim eintrag As Variant

    For Each eintrag In werte
        If StrComp(suchWert, CStr(eintrag), vbTextCompare) = 0 Then
            WertIstInArray = True
            Exit Function
        End If
    Next eintrag

End Function

Private Function BlattExistiert(ByVal wb As Workbook, ByVal blattName As String) As Boolean

    Dim blatt As Worksheet

    On Error Resume Next
    Set blatt = wb.Worksheets(blattName)
    BlattExistiert = Not blatt Is Nothing
    On Error GoTo 0

End Function

Private Function EindeutigerBlattname(ByVal wb As Workbook, ByVal basisName As String) As String

    Dim kandidat As String
    Dim zaehler As Long
    Dim suffix As String

    kandidat = Left$(basisName, 31)

    Do While BlattExistiert(wb, kandidat)
        zaehler = zaehler + 1
        suffix = "_" & CStr(zaehler)
        kandidat = Left$(basisName, 31 - Len(suffix)) & suffix
    Loop

    EindeutigerBlattname = kandidat

End Function

Private Sub ZielAnsichtAufAktuellenMonat( _
    ByVal zielBlatt As Worksheet, _
    ByRef zielSpalten() As Long, _
    ByRef datumswerte() As Date, _
    ByVal datumsAnzahl As Long)

    Dim stichtag As Date
    Dim i As Long
    Dim sichtbareSpalte As Long

    If Not VERGANGENE_MONATE_AUSBLENDEN Then Exit Sub

    stichtag = DateSerial(Year(Date), Month(Date), 1)

    For i = 1 To datumsAnzahl
        If datumswerte(i) >= stichtag Then
            sichtbareSpalte = zielSpalten(i)
            Exit For
        End If
    Next i

    If sichtbareSpalte > 0 Then
        zielBlatt.Activate
        ActiveWindow.ScrollColumn = sichtbareSpalte
    End If

End Sub
