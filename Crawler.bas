Attribute VB_Name = "Modul2"
Option Explicit

Sub DatenAbrufen()
    Dim driver As Object
    Dim ws As Worksheet
    Dim zeile As Long
    Dim lastRow As Long
    Dim url As String
    Dim haendlerName As String
    Dim bestellnr As String
    Dim bestellnrRaw As String
    Dim masseinheit As String
    Dim produktTitel As String
    Dim produktPreisRaw As String
    Dim produktPreis As Double
    Dim verfuegbarkeit As String
    Dim lieferfristText As String
    Dim screenshotPath As String
    Dim folderPath As String
    Dim safeFileName As String
    Dim dateStamp As String
    Dim steuer As Double
    Dim stueckzahl As Double
    Dim summe As Double
    Dim imgObj As Object
    Dim staffelpreise As String

    Set ws = ThisWorkbook.Sheets(1)

    ' ChromeDriver starten (Late Binding)
    On Error GoTo ErrCreateDriver
    Set driver = CreateObject("Selenium.ChromeDriver")
    On Error GoTo 0

    driver.Start "chrome"
    driver.ExecuteScript "window.resizeTo(1300, 1000);"

    ' Letzte Zeile in Spalte D (URL)
    lastRow = ws.cells(ws.rows.Count, 4).End(xlUp).row

    For zeile = 13 To lastRow
        On Error GoTo RowError

        ' Staffelpreis/Fehler-Spalten leeren (X=24, Y=25)
        ws.cells(zeile, 24).Value = ""
        ws.cells(zeile, 25).Value = ""

        url = Trim(ws.cells(zeile, 4).Value)
        If url = "" Or Not IsValidURL(url) Then GoTo NextRow

        ' Händler bestimmen (falls Spalte C leer)
        haendlerName = Trim(ws.cells(zeile, 3).Value)
        If haendlerName = "" Then
            haendlerName = BestimmeHaendlerAusURL(url)
            ws.cells(zeile, 3).Value = haendlerName
        End If

        ' Seite laden
        driver.Get url
        driver.Wait 1500

        ' Variablen zurücksetzen
        produktTitel = ""
        produktPreisRaw = ""
        produktPreis = 0
        verfuegbarkeit = ""
        lieferfristText = ""
        bestellnrRaw = ""
        bestellnr = ""
        masseinheit = "n/a"
        staffelpreise = ""

        ' -----------------------------
        ' Daten je nach Händler holen
        ' -----------------------------
        If InStr(1, haendlerName, "Conrad", vbTextCompare) > 0 Then
            ' ----- CONRAD -----
            produktTitel = SafeGetTextByTag(driver, "h1")
            produktPreisRaw = SafeGetTextById(driver, "productPriceUnitPrice")
            verfuegbarkeit = SafeGetTextByCss(driver, "div#currentOfferAvailability span[data-prerenderer='availabilityText']")
            lieferfristText = SafeGetTextByCss(driver, "#fastTrackDelivery > div > span.message-date")
            bestellnrRaw = SafeGetTextByCss(driver, "#productCode")

        ElseIf InStr(1, haendlerName, "Reichelt", vbTextCompare) > 0 Then
            ' ----- REICHELT -----
            Dim rawAvailOrig As String, rawAvailLower As String
            Dim lfSource As String, lf As String

            ' Titel
            produktTitel = SafeGetTextByCss(driver, ".productBuy > h1:nth-child(1)")
            ' Artikelnummer
            bestellnrRaw = SafeGetTextByCss(driver, "small.copytoclipboard:nth-child(3) > span:nth-child(1) > b:nth-child(1)")

            ' Preis – zuerst meta[itemprop=price], dann Fallback sichtbarer Text
            produktPreisRaw = SafeGetAttrByCss(driver, "meta[itemprop='price']", "content")
            If Trim(produktPreisRaw) = "" Then
                produktPreisRaw = SafeGetTextByCss(driver, "p.productPrice")
                If Trim(produktPreisRaw) = "" Then
                    produktPreisRaw = SafeGetTextByCss(driver, ".productPrice")
                End If
            End If

            ' rohe Verfügbarkeit
            rawAvailOrig = SafeGetTextByCss(driver, ".availability")
            rawAvailLower = LCase$(rawAvailOrig)

            ' Flags für Verfügbarkeitslogik
            Dim isPreorder As Boolean, isNotAvail As Boolean, isInStock As Boolean

            isPreorder = (InStr(rawAvailLower, "voraussichtlich") > 0) Or _
                         (InStr(rawAvailLower, "vorbestellbar") > 0)

            isNotAvail = (InStr(rawAvailLower, "nicht lieferbar") > 0) Or _
                         (InStr(rawAvailLower, "nicht lagernd") > 0) Or _
                         (InStr(rawAvailLower, "aktuell nicht") > 0) Or _
                         (InStr(rawAvailLower, "derzeit nicht") > 0) Or _
                         (InStr(rawAvailLower, "ausverkauft") > 0) Or _
                         (InStr(rawAvailLower, "nachricht bei verf") > 0)

            isInStock = (InStr(rawAvailLower, "ab lager") > 0) Or _
                        (InStr(rawAvailLower, "sofort lieferbar") > 0) Or _
                        (InStr(rawAvailLower, "sofort verf") > 0) Or _
                        (InStr(rawAvailLower, "lagernd") > 0) Or _
                        (InStr(rawAvailLower, "auf lager") > 0) Or _
                        (InStr(rawAvailLower, "lieferzeit:") > 0)

            If isPreorder Or isNotAvail Then
                verfuegbarkeit = "Ausverkauft"
            ElseIf isInStock Then
                verfuegbarkeit = "Auf Lager"
            Else
                ' Im Zweifel konservativ: nicht auf Lager
                verfuegbarkeit = "Ausverkauft"
            End If

            ' Lieferfrist herausziehen – nur der Zeitanteil bzw. ab-Datum
            lfSource = rawAvailOrig
            lf = ExtractLieferfristFromText(lfSource)
            lieferfristText = lf

            ' Maßeinheit direkt aus Grundpreis-Angabe (€/100ml, €/kg, €/m, ...)
            masseinheit = ExtractUnitFromGrundpreis(driver)

            ' Staffelpreise (falls vorhanden)
            staffelpreise = ExtractReicheltStaffelpreise(driver)

        Else
            ' ----- Andere Shops (Fallback) -----
            produktTitel = SafeGetTextByTag(driver, "h1")
            produktPreisRaw = SafeGetAttrByCss(driver, "meta[itemprop='price']", "content")
            ' verfuegbarkeit/lieferfrist bleiben leer
        End If

        ' -----------------------------
        ' Bestellnummer bereinigen
        ' -----------------------------
        If InStr(1, haendlerName, "Reichelt", vbTextCompare) > 0 Then
            bestellnr = Trim(bestellnrRaw)
        Else
            bestellnr = FilterBestellnummer(bestellnrRaw)
            bestellnr = Replace(bestellnr, vbCrLf, "")
            bestellnr = Replace(bestellnr, vbCr, "")
            bestellnr = Replace(bestellnr, vbLf, "")
            bestellnr = Trim(bestellnr)
        End If

        ' Maßeinheit bei Conrad:
        ' 1) versuchen wie bei Reichelt aus Grundpreis (€/kg, €/m, €/100g, ...)
        ' 2) wenn nichts Sinnvolles: Fallback auf Preistabelle
        If InStr(1, haendlerName, "Conrad", vbTextCompare) > 0 Then
            masseinheit = ExtractUnitFromGrundpreis(driver)
            If masseinheit = "Stk" Or masseinheit = "n/a" Or masseinheit = "" Then
                masseinheit = ExtractUnitFromTable(driver)
            End If
        End If

        ' Conrad FastTrack-Fallback
        If InStr(1, haendlerName, "Conrad", vbTextCompare) > 0 Then
            If Trim(lieferfristText) = "" Then
                If SafeElementExistsById(driver, "fastTrackDelivery") Then
                    lieferfristText = "1 Tag (FastTrack)"
                End If
            End If
        End If

        ' Preis in Double wandeln
        produktPreis = ParsePreisToDouble(produktPreisRaw)

        ' Berechnungen
        stueckzahl = Val(ws.cells(zeile, 7).Value)
        If stueckzahl = 0 Then stueckzahl = 1

        steuer = Round(produktPreis * 0.19, 2)
        summe = Round(stueckzahl * produktPreis, 2)

        ' Screenshot-Dateiname
        dateStamp = Format(Date, "DD-MM-YY")
        safeFileName = bestellnr
        If safeFileName = "" Then safeFileName = "no_order_" & zeile
        safeFileName = SanitizeFileName(safeFileName & "_" & dateStamp & ".png")

        ' Screenshot-Ordner
        folderPath = ThisWorkbook.Path & Application.PathSeparator & "Preisnachweise"
        If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
        screenshotPath = folderPath & Application.PathSeparator & safeFileName

        ' Screenshot aufnehmen
        On Error Resume Next
        Set imgObj = driver.TakeScreenshot()
        If Not imgObj Is Nothing Then imgObj.SaveAs screenshotPath
        On Error GoTo RowError

        ' Ergebnisse in Tabelle schreiben
        ws.cells(zeile, 5).Value = produktTitel
        ws.cells(zeile, 6).Value = bestellnr
        ws.cells(zeile, 8).Value = masseinheit
        ws.cells(zeile, 9).Value = verfuegbarkeit
        ws.cells(zeile, 10).Value = lieferfristText
        ws.cells(zeile, 11).Value = produktPreis
        ws.cells(zeile, 12).Value = steuer
        ws.cells(zeile, 13).Value = summe
        ws.cells(zeile, 23).Value = screenshotPath
        ws.cells(zeile, 24).Value = staffelpreise

NextRow:
    Next zeile

Cleanup:
    On Error Resume Next
    driver.Quit
    Set driver = Nothing
    On Error GoTo 0

    MsgBox "Datenabfrage abgeschlossen!"
    Exit Sub

ErrCreateDriver:
    MsgBox "Fehler beim Initialisieren des ChromeDriver: " & Err.Description, vbExclamation
    Exit Sub

RowError:
    ws.cells(zeile, 25).Value = "Fehler " & Err.Number & ": " & Err.Description
    Err.Clear
    Resume NextRow
End Sub

' ========================================
' Händler aus URL bestimmen
' ========================================
Private Function BestimmeHaendlerAusURL(ByVal url As String) As String
    If InStr(1, url, "conrad", vbTextCompare) > 0 Then
        BestimmeHaendlerAusURL = "Conrad"
    ElseIf InStr(1, url, "reichelt", vbTextCompare) > 0 Then
        BestimmeHaendlerAusURL = "Reichelt"
    ElseIf InStr(1, url, "digikey", vbTextCompare) > 0 Then
        BestimmeHaendlerAusURL = "DigiKey"
    Else
        BestimmeHaendlerAusURL = "Unbekannt"
    End If
End Function

' ========================================
' Prüfen, ob String eine URL ist
' ========================================
Private Function IsValidURL(ByVal url As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^(https?://)[\w\d\.-]+\.[a-z]{2,}(/.*)?$"
    regex.IgnoreCase = True
    IsValidURL = regex.Test(Trim(url))
End Function

' ========================================
' Bestellnummer filtern (v.a. Conrad)
' ========================================
Private Function FilterBestellnummer(rawText As String) As String
    Dim result As String, pos As Long
    If Trim(rawText) <> "" Then
        pos = InStr(1, rawText, "Bestell-Nr.:", vbTextCompare)
        If pos > 0 Then
            result = Trim(Mid$(rawText, pos + Len("Bestell-Nr.:")))
        Else
            pos = InStr(1, rawText, "Bestell-Nr.", vbTextCompare)
            If pos > 0 Then
                result = Trim(Mid$(rawText, pos + Len("Bestell-Nr.")))
            Else
                result = Trim(rawText)
            End If
        End If
        Do While Left$(result, 1) Like "[: .]"
            result = Mid$(result, 2)
        Loop
    End If
    FilterBestellnummer = Trim(result)
End Function

' ========================================
' Maßeinheit aus Conrad-Preistabelle (Fallback)
' ========================================
Private Function ExtractUnitFromTable(driver As Object) As String
    Dim cell As Object, txt As String
    Dim regex As Object, matches As Object
    ExtractUnitFromTable = "n/a"
    On Error Resume Next
    Set cell = driver.FindElementByXPath("//table[contains(@class, 'totalPriceTable__table')]//tbody/tr[1]/td[1]")
    If cell Is Nothing Then Set cell = driver.FindElementByCss(".totalPriceTable__table tbody tr:first-child td:first-child")
    If Not cell Is Nothing Then
        txt = Trim(cell.Text)
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "\b(ml|l|L|g|kg|St\.|Stk|Stück|StŸck|m|cm|mm|Pack|Pkg|Liter|Meter)\b"
        regex.IgnoreCase = True
        If regex.Test(txt) Then
            Set matches = regex.Execute(txt)
            ExtractUnitFromTable = matches(0).Value
        Else
            ExtractUnitFromTable = txt
        End If
    End If
    On Error GoTo 0
End Function

' ========================================
' Einheiten aus Grundpreis (Reichelt & Conrad):
' Beispiele:
'  - "69,90 €/100ml" ? ml
'  - "27,90 €/kg"    ? kg
'  - "0,99 €/m"      ? m
' ========================================
Private Function ExtractUnitFromGrundpreis(driver As Object) As String
    Dim els As Object, el As Object
    Dim t As String, tLow As String
    Dim pos As Long
    Dim rest As String
    Dim i As Long
    Dim ch As String
    Dim unit As String
    Dim startedLetters As Boolean

    ' Default-Fallback
    ExtractUnitFromGrundpreis = "Stk"

    On Error Resume Next
    Set els = driver.FindElementsByTag("small")
    If els Is Nothing Then GoTo CleanExit

    For Each el In els
        t = Trim(el.Text)
        If t <> "" Then
            tLow = LCase$(t)
            pos = InStr(tLow, "€/")
            If pos > 0 Then
                ' Alles NACH "€/"
                rest = Mid$(tLow, pos + Len("€/"))
                rest = Trim$(rest)

                ' Zuerst alle Ziffern, Leerzeichen, %, etc. überspringen (für Fälle wie "100ml")
                i = 1
                Do While i <= Len(rest)
                    ch = Mid$(rest, i, 1)
                    If ch Like "[0-9 ,.%]" Then
                        i = i + 1
                    Else
                        Exit Do
                    End If
                Loop

                ' Jetzt Buchstaben einsammeln (ml, g, kg, m, cm, mm, ...)
                unit = ""
                startedLetters = False
                Do While i <= Len(rest)
                    ch = Mid$(rest, i, 1)
                    If ch Like "[a-z]" Then
                        unit = unit & ch
                        startedLetters = True
                    Else
                        If startedLetters Then Exit Do
                    End If
                    i = i + 1
                Loop

                unit = Trim$(unit)
                If unit <> "" Then
                    Select Case unit
                        Case "ml", "g", "kg", "l", "m", "cm", "mm"
                            ExtractUnitFromGrundpreis = unit
                            Exit For
                        Case Else
                            ' z.B. "stk" ? ignorieren, Standard bleibt "Stk"
                    End Select
                End If
            End If
        End If
    Next el

CleanExit:
    On Error GoTo 0
End Function

' ========================================
' Reichelt: Staffelpreise als Text extrahieren
' ========================================
Private Function ExtractReicheltStaffelpreise(driver As Object) As String
    Dim tbl As Object, rows As Object, row As Object
    Dim cells As Object
    Dim qText As String, pText As String
    Dim result As String

    On Error Resume Next
    Set tbl = driver.FindElementByCss(".price_table")
    If tbl Is Nothing Then Set tbl = driver.FindElementByCss(".priceTable")
    If tbl Is Nothing Then Set tbl = driver.FindElementByCss(".graduatedPrices")
    If tbl Is Nothing Then
        ExtractReicheltStaffelpreise = ""
        GoTo CleanExit
    End If

    Set rows = tbl.FindElementsByTag("tr")
    If rows Is Nothing Then
        ExtractReicheltStaffelpreise = ""
        GoTo CleanExit
    End If

    For Each row In rows
        Set cells = row.FindElementsByTag("td")
        If Not cells Is Nothing Then
            If cells.Count >= 2 Then
                qText = Trim(cells(0).Text)
                pText = Trim(cells(1).Text)
                If qText <> "" And pText <> "" Then
                    If result <> "" Then result = result & "; "
                    result = result & qText & ": " & pText
                End If
            End If
        End If
    Next row

    ExtractReicheltStaffelpreise = result

CleanExit:
    On Error GoTo 0
End Function

' ========================================
' Lieferfrist aus Reichelt-Text extrahieren
' ========================================
Private Function ExtractLieferfristFromText(ByVal src As String) As String
    Dim s As String
    Dim posLT As Long, posColon As Long
    Dim sLow As String
    Dim posAb As Long, posColon2 As Long

    s = Trim$(src)
    If s = "" Then
        ExtractLieferfristFromText = ""
        Exit Function
    End If

    sLow = LCase$(s)

    ' Fälle "Voraussichtlich lieferbar ab: 18.12.2025 ..."
    If InStr(sLow, "voraussichtlich") > 0 And InStr(sLow, "ab") > 0 Then
        posAb = InStr(sLow, "ab")
        posColon2 = InStr(posAb, s, ":")
        If posColon2 > 0 Then
            s = "ab " & Trim$(Mid$(s, posColon2 + 1))
        Else
            s = Mid$(s, posAb)
        End If
        ExtractLieferfristFromText = Trim$(s)
        Exit Function
    End If

    ' Normale "Lieferzeit: X" Fälle
    posLT = InStr(1, sLow, "lieferzeit")
    If posLT > 0 Then
        posColon = InStr(posLT, s, ":")
        If posColon > 0 Then
            s = Mid$(s, posColon + 1)
        Else
            s = Mid$(s, posLT + Len("Lieferzeit"))
        End If
    End If

    ' Leading-Kommas / Striche entfernen
    s = Trim$(s)
    Do While Left$(s, 1) = "," Or Left$(s, 1) = "-" Or Left$(s, 1) = " "
        s = Mid$(s, 2)
        s = Trim$(s)
    Loop

    ' "Ab Lager" entfernen, falls noch vorn
    If LCase$(Left$(s, 8)) = "ab lager" Then
        s = Mid$(s, 9)
        s = Trim$(s)
    End If

    ExtractLieferfristFromText = s
End Function

' ========================================
' Preisstring nach Double
' ========================================
Private Function ParsePreisToDouble(ByVal preisText As String) As Double
    Dim decSep As String
    Dim tmp As String

    ParsePreisToDouble = 0

    preisText = Trim(preisText)
    If preisText = "" Then Exit Function

    decSep = Application.International(xlDecimalSeparator)

    tmp = preisText
    tmp = Replace(tmp, "€", "")
    tmp = Replace(tmp, "Û", "")
    tmp = Replace(tmp, "EUR", "")
    tmp = Replace(tmp, Chr(160), "")
    tmp = Trim(tmp)
    tmp = Replace(tmp, " ", "")

    If InStr(tmp, ",") > 0 And InStr(tmp, ".") > 0 Then
        tmp = Replace(tmp, ".", "")
        tmp = Replace(tmp, ",", decSep)
    ElseIf InStr(tmp, ",") > 0 Then
        tmp = Replace(tmp, ",", decSep)
    ElseIf InStr(tmp, ".") > 0 Then
        tmp = Replace(tmp, ".", decSep)
    End If

    If IsNumeric(tmp) Then
        ParsePreisToDouble = CDbl(tmp)
    End If
End Function

' ========================================
' Dateiname säubern
' ========================================
Private Function SanitizeFileName(ByVal fName As String) As String
    fName = Replace(fName, " ", "_")
    fName = Replace(fName, "\", "-")
    fName = Replace(fName, "/", "-")
    fName = Replace(fName, ":", "-")
    fName = Replace(fName, "*", "")
    fName = Replace(fName, "?", "")
    fName = Replace(fName, """", "")
    fName = Replace(fName, "<", "")
    fName = Replace(fName, ">", "")
    fName = Replace(fName, "|", "")
    SanitizeFileName = Trim$(fName)
End Function

' ========================================
' Sichere DOM-Helfer (geben "" statt Fehler)
' ========================================
Private Function SafeGetTextByCss(driver As Object, ByVal selector As String) As String
    Dim el As Object
    On Error Resume Next
    Set el = driver.FindElementByCss(selector)
    If Not el Is Nothing Then SafeGetTextByCss = el.Text Else SafeGetTextByCss = ""
    Set el = Nothing
    On Error GoTo 0
End Function

Private Function SafeGetAttrByCss(driver As Object, ByVal selector As String, ByVal attrName As String) As String
    Dim el As Object
    On Error Resume Next
    Set el = driver.FindElementByCss(selector)
    If Not el Is Nothing Then SafeGetAttrByCss = el.GetAttribute(attrName) Else SafeGetAttrByCss = ""
    Set el = Nothing
    On Error GoTo 0
End Function

Private Function SafeGetTextById(driver As Object, ByVal id As String) As String
    Dim el As Object
    On Error Resume Next
    Set el = driver.FindElementById(id)
    If Not el Is Nothing Then SafeGetTextById = el.Text Else SafeGetTextById = ""
    Set el = Nothing
    On Error GoTo 0
End Function

Private Function SafeGetTextByTag(driver As Object, ByVal tagName As String) As String
    Dim el As Object
    On Error Resume Next
    Set el = driver.FindElementByTag(tagName)
    If Not el Is Nothing Then SafeGetTextByTag = el.Text Else SafeGetTextByTag = ""
    Set el = Nothing
    On Error GoTo 0
End Function

Private Function SafeElementExistsById(driver As Object, ByVal id As String) As Boolean
    Dim el As Object
    On Error Resume Next
    Set el = driver.FindElementById(id)
    SafeElementExistsById = Not (el Is Nothing)
    Set el = Nothing
    On Error GoTo 0
End Function


