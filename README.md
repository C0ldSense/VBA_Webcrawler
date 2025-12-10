# VBA Webcrawler
Dieses Repository enthält den Quellcode einer Excel-VBA-Automatisierung,
die mithilfe von Selenium Produktdaten (z. B. Preis, Verfügbarkeit,
Lieferzeit) direkt aus Online-Webshops wie Conrad ausliest und automatisiert
in ein Excel-Bestellformular überträgt.

Der Hauptzweck dieses Dokuments ist die nachhaltige Wartbarkeit des Projekts.
Insbesondere wird hier detailliert beschrieben, wie das Skript bei Änderungen
der Webseitenstruktur (HTML/DOM) der Händler repariert werden kann – auch durch
Personen ohne Vorkenntnisse in VBA, Selenium oder Webentwicklung.


## 1. GRUNDPRINZIP DES SKRIPTS
Das Skript funktioniert nach folgendem Grundschema:

1. Excel enthält eine Liste von Produkt-URLs
2. Selenium öffnet jede URL automatisch im Browser
3. Bestimmte Informationen werden direkt aus der Webseite gelesen
4. Die Daten werden aufbereitet und nach Excel zurückgeschrieben
5. Zusätzlich wird ein Screenshot der Produktseite gespeichert

Das Skript greift dabei NICHT auf offizielle Shop-APIs zu, sondern liest
die Informationen direkt aus dem sichtbaren Seitenaufbau (HTML/DOM).
Genau daraus ergibt sich später der Wartungsaufwand.

## 2. WARUM WEBSEITENÄNDERUNGEN PROBLEME VERURSACHEN
Webshops werden regelmäßig überarbeitet:
- neue Designs
- neue CSS-Klassen
- andere HTML-Strukturen
- Umbenennung von IDs

Das Skript greift gezielt auf bestimmte HTML-Elemente zu.
Wenn diese Elemente umbenannt oder verschoben werden,
kann Selenium sie nicht mehr finden.

Technisch äußert sich dies so:
- Selenium sucht ein Element
- findet es nicht
- der gelesene Wert bleibt leer
- Excel-Zellen bleiben leer oder unvollständig

Anmerkung:
Das Skript stürzt in der Regel NICHT ab.
Stattdessen entstehen „leise Fehler“ (leere Felder),
die aktiv erkannt werden müssen.

## 3. DOM-ABHÄNGIGE CODESTELLEN

Die folgenden Codezeilen sind ALLE Stellen im Projekt, an denen direkt auf
die Webseitenstruktur zugegriffen wird. Nur diese Zeilen müssen bei Änderungen der Shop-Webseiten angepasst werden.
Alle anderen Codebereiche sind unabhängig vom DOM.

### 3. Produkttitel
--------------------------------------------------

    produktTitel = driver.FindElementByTag("h1").Text

Erwartung:
Der Produktname befindet sich in einem h1-HTML-Tag.

Typisches Problem:
- Es gibt kein h1 mehr
- Es gibt mehrere h1-Tags
- Der Titel wurde in ein anderes Element verschoben


--------------------------------------------------
3.2 Produktpreis
--------------------------------------------------

    produktPreisRaw = driver.FindElementById("productPriceUnitPrice").Text

Erwartung:
Der Preis besitzt eine eindeutige HTML-ID.

Typisches Problem:
- ID wurde umbenannt
- ID existiert nicht mehr
- Preis ist jetzt mehrfach vorhanden (z. B. Staffelpreise)


--------------------------------------------------
3.3 Verfügbarkeit
--------------------------------------------------

    verfuegbarkeit = driver.FindElementByCss("div#currentOfferAvailability span[data-prerenderer='availabilityText']").Text

Erwartung:
- Ein div mit ID "currentOfferAvailability"
- darin ein span mit spezifischem Attribut

Typische Probleme:
- data-Attribute entfernt
- Verschachtelung geändert
- Klassen statt IDs eingeführt


--------------------------------------------------
3.4 Lieferfrist
--------------------------------------------------

    lieferfristText = driver.FindElementByCss("span.currentOfferAvailability__additionalDeliveryText").Text

Erwartung:
- Lieferzeit befindet sich in einem span mit spezifischer Klasse

Typische Probleme:
- Klasse umbenannt
- Lieferzeit in Textblock integriert
- separate Lieferzeit existiert nicht mehr


--------------------------------------------------
3.5 Sonderfall: FastTrack-Lieferung
--------------------------------------------------

    If lieferfristText = "" Then
        On Error Resume Next
        If Not driver.FindElementById("fastTrackDelivery") Is Nothing Then
            lieferfristText = "1 Tag (FastTrack)"
        End If
        On Error GoTo 0
    End If

Erwartung:
- Es existiert ein eindeutiges Element mit ID "fastTrackDelivery"

Typische Probleme:
- Feature entfernt
- anderes HTML-Element
- neue Klasse statt ID



## 4. TYPOLOGIE DER FEHLERBILDER (DIAGNOSE)
Diese Tabelle hilft, Fehler schnell einzugrenzen:

Produkttitel fehlt
→ Problem in Abschnitt 3.1

Preis fehlt oder = 0
→ Problem in Abschnitt 3.2 ODER Preisformat

Verfügbarkeit fehlt
→ Problem in Abschnitt 3.3

Lieferzeit fehlt
→ Problem in Abschnitt 3.4 oder 3.5

Screenshot vorhanden, Daten leer
→ DOM-Zugriff fehlgeschlagen, Browser funktioniert


## 5. SCHRITT-FÜR-SCHRITT-REPARATUR (KEINE VORKENNTNISSE)

--------------------------------------------------
5.1 VBA-Editor öffnen
--------------------------------------------------

1. Excel öffnen
2. Taste ALT + F11 drücken
3. Links im Projekt-Explorer das entsprechende Modul öffnen
4. Die Prozedur "DatenAbrufen" suchen


--------------------------------------------------
5.2 Fehler reproduzieren
--------------------------------------------------

1. Makro per F5 ausführen
2. Excel beobachten:
   - Welche Spalten bleiben leer?
3. Entsprechenden Abschnitt aus Kapitel 3 identifizieren


--------------------------------------------------
5.3 Entwicklerwerkzeuge öffnen
--------------------------------------------------

1. Produkt-URL aus Excel kopieren
2. Im Browser öffnen
3. Taste F12 drücken
4. Im Browser:
   - Auf das Symbol "Element auswählen" klicken 


--------------------------------------------------
5.4 HTML-Element identifizieren
--------------------------------------------------

1. Auf sichtbaren Preis / Titel / Lieferzeit klicken
2. Im Dev-Fenster erscheint der zugehörige HTML-Code
3. Nach folgenden Dingen suchen:
   - id="..."
   - class="..."
   - verwendeter Tag (span, div, h1, ...)


--------------------------------------------------
5.5 Neuen Selektor formulieren
--------------------------------------------------

Beispiele:

ID vorhanden:
    driver.FindElementById("unit-price")

Klasse vorhanden:
    driver.FindElementByCss("span.unitPrice")

Verschachtelung:
    driver.FindElementByCss("div.price-container span.price")


--------------------------------------------------
5.6 Code ersetzen
--------------------------------------------------

ALTE Zeile:
    produktPreisRaw = driver.FindElementById("productPriceUnitPrice").Text

NEUE Zeile:
    produktPreisRaw = driver.FindElementByCss("span.unitPrice").Text


--------------------------------------------------
5.7 Preisformat reparieren
--------------------------------------------------

Zusätzliche Ersetzungen möglich:

    produktPreisRaw = Replace(produktPreisRaw, "*", "")
    produktPreisRaw = Replace(produktPreisRaw, "ab", "")
    produktPreisRaw = Replace(produktPreisRaw, Chr(160), "")

Ziel:
Eine reine Zahl mit Punkt:
    12.99



## 6. Unverändert lassen (es sei denn du findest eine Verbesserung)


- Excel-Schleifen
- Stückzahl-Logik
- Mehrwertsteuer-Berechnung
- Screenshot-Logik
- Fehlerbehandlung

Diese Teile sind vollkommen unabhängig vom Webseiten-DOM.
