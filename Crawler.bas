Attribute VB_Name = "Modul1"
Option Explicit

Sub DatenAbrufen()
    Dim driver As New Selenium.ChromeDriver
    Dim ws As Worksheet
    Dim zeile As Long
    Dim url As String
    Dim händlerName As String
    Dim produktTitel As String
    Dim produktPreisRaw As String
    Dim produktPreis As Double
    Dim verfuegbarkeit As String
    Dim lieferfrist As String
    Dim lieferfristText As String
    Dim screenshotPath As String
    Dim folderPath As String
    Dim safeFileName As String
    Dim steuer As Double
    Dim stueckzahl As Double
    Dim summe As Double
    Dim imgObj As Object

    Set ws = ThisWorkbook.Sheets(1)

    ' Browser starten
    driver.Start "chrome"
    driver.ExecuteScript "window.resizeTo(1300, 1000);"

    ' Erste Datenzeile
    zeile = 8

    ' Durchlaufe Zeilen bis Ende
    Do While ws.cells(zeile, 4).Value <> "" Or ws.cells(zeile, 3).Value <> "" Or ws.cells(zeile, 23).Value <> ""

        url = Trim(ws.cells(zeile, 4).Value)

        ' Leere Zeilen überspringen
        If url <> "" Then
            ' Händlername falls leer ? aus Domain extrahieren
            händlerName = ws.cells(zeile, 3).Value
            If händlerName = "" Then
                If InStr(url, "conrad") > 0 Then
                    händlerName = "Conrad"
                ElseIf InStr(url, "reichelt") > 0 Then
                    händlerName = "Reichelt"
                ElseIf InStr(url, "digikey") > 0 Then
                    händlerName = "DigiKey"
                Else
                    händlerName = "Unbekannt"
                End If
                ws.cells(zeile, 3).Value = händlerName
            End If

            ' Seite laden
            On Error Resume Next
            driver.Get url
            driver.Wait 1500
            On Error GoTo 0

            ' --- Produktinformationen abrufen ---
            On Error Resume Next
            produktTitel = driver.FindElementByTag("h1").Text
            produktPreisRaw = driver.FindElementById("productPriceUnitPrice").Text
            verfuegbarkeit = driver.FindElementByCss("div#currentOfferAvailability span[data-prerenderer='availabilityText']").Text
            lieferfristText = driver.FindElementByCss("span.currentOfferAvailability__additionalDeliveryText").Text
            On Error GoTo 0

            If lieferfristText = "" Then
                On Error Resume Next
                If Not driver.FindElementById("fastTrackDelivery") Is Nothing Then
                    lieferfristText = "1 Tag (FastTrack)"
                End If
                On Error GoTo 0
            End If
            lieferfrist = lieferfristText

            ' Preis in Zahl umwandeln (z. B. "12,99 €")
            produktPreis = 0
            If produktPreisRaw <> "" Then
                produktPreisRaw = Replace(produktPreisRaw, "€", "")
                produktPreisRaw = Replace(produktPreisRaw, "EUR", "")
                produktPreisRaw = Replace(produktPreisRaw, " ", "")
                produktPreisRaw = Replace(produktPreisRaw, ",", ".")
                If IsNumeric(produktPreisRaw) Then produktPreis = CDbl(produktPreisRaw)
            End If

            ' --- Berechnungen ---
            stueckzahl = Val(ws.cells(zeile, 7).Value)
            If stueckzahl = 0 Then stueckzahl = 1
            steuer = Round(produktPreis * 0.19, 2)
            summe = Round(stueckzahl * produktPreis, 2)

            ' --- Screenshot-Dateiname ---
            safeFileName = produktTitel
            safeFileName = Replace(safeFileName, "\", "-")
            safeFileName = Replace(safeFileName, "/", "-")
            safeFileName = Replace(safeFileName, ":", "-")
            safeFileName = Replace(safeFileName, "*", "")
            safeFileName = Replace(safeFileName, "?", "")
            safeFileName = Replace(safeFileName, """", "")
            safeFileName = Replace(safeFileName, "<", "")
            safeFileName = Replace(safeFileName, ">", "")
            safeFileName = Replace(safeFileName, "|", "")
            safeFileName = Trim(safeFileName)

            ' Zielordner aus Spalte W
            folderPath = ws.cells(zeile, 23).Value
            If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
            screenshotPath = folderPath & safeFileName & ".png"

            ' Screenshot aufnehmen
            On Error GoTo ScreenshotFehler
            Set imgObj = driver.TakeScreenshot()
            imgObj.SaveAs screenshotPath
            On Error GoTo 0

            ' --- Ergebnisse zurück in Excel ---
            ws.cells(zeile, 6).Value = produktTitel
            ws.cells(zeile, 9).Value = verfuegbarkeit
            ws.cells(zeile, 10).Value = lieferfrist
            ws.cells(zeile, 11).Value = produktPreis
            ws.cells(zeile, 12).Value = steuer
            ws.cells(zeile, 13).Value = summe
            ws.cells(zeile, 23).Value = screenshotPath
        End If

        ' Nächste Zeile
        zeile = zeile + 1
    Loop

    driver.Quit
    MsgBox "Datenabfrage erfolgreich abgeschlossen!"

    Exit Sub

ScreenshotFehler:
    ws.cells(zeile, 23).Value = "Screenshot-Fehler!"
    Resume Next
End Sub

