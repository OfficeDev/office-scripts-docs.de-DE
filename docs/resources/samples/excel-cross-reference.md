---
title: Querverweisen Excel Dateien mit Power Automate
description: Erfahren Sie, wie Sie Office Skripts und Power Automate zum Querverweisen und Formatieren einer Excel Datei verwenden.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 13ba6c8ba6f9232554ea6cfd5f98c308ea981683
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585485"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Querverweisen Excel Dateien mit Power Automate

Diese Lösung zeigt, wie Daten in zwei Excel Dateien verglichen werden, um Abweichungen zu finden. Es verwendet Office Skripts, um Daten zu analysieren und Power Automate, um zwischen den Arbeitsmappen zu kommunizieren.

## <a name="example-scenario"></a>Beispielszenario

Sie sind ein Ereignis-Coordinator, der Referenten für anstehende Konferenzen plant. Sie behalten die Ereignisdaten in einer Kalkulationstabelle und die Lautsprecherregistrierungen in einer anderen. Um sicherzustellen, dass die beiden Arbeitsmappen synchronisiert werden, verwenden Sie einen Fluss mit Office Skripts, um potenzielle Probleme hervorzuheben.

## <a name="sample-excel-files"></a>Beispieldateien für Excel

Laden Sie die folgenden Dateien herunter, um einsatzbereite Arbeitsmappen für das Beispiel zu erhalten.

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

Fügen Sie die folgenden Skripts hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-get-event-data"></a>Beispielcode: Abrufen von Ereignisdaten

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a>Beispielcode: Überprüfen von Lautsprecherregistrierungen

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate Ablauf: Inkonsistenzen in den Arbeitsmappen überprüfen

Dieser Fluss extrahiert die Ereignisinformationen aus der ersten Arbeitsmappe und verwendet diese Daten, um die zweite Arbeitsmappe zu überprüfen.

1. Melden Sie sich bei [Power Automate an](https://flow.microsoft.com), und erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen Sie **"Manuell auslösen" aus,** und wählen Sie " **Erstellen**" aus.
1. Fügen Sie einen **neuen Schritt** hinzu, der den **Connector Excel Online (Business)** mit der **Skriptaktion ausführen** verwendet. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: event-data.xlsx ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Skript**: Ereignisdaten abrufen

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Der fertige Excel Online-Connector (Business) für das erste Skript in Power Automate.":::

1. Fügen Sie einen zweiten **neuen Schritt** hinzu, der den **connector Excel Online (Business)** mit der **Skriptaktion ausführen** verwendet. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: speaker-registration.xlsx ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Skript**: Überprüfen der Lautsprecherregistrierung

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Der fertige Excel Online(Business)-Connector für das zweite Skript in Power Automate.":::
1. In diesem Beispiel wird Outlook als E-Mail-Client verwendet. Sie können einen beliebigen E-Mail-Connector verwenden, Power Automate unterstützt. Fügen Sie einen **neuen Schritt** hinzu, der den **Office 365 Outlook** Connector und die Aktion **"Senden und E-Mail(V2)"** verwendet. Verwenden Sie die folgenden Werte für die Aktion.
    * **An**: Ihr Test-E-Mail-Konto (oder persönliche E-Mail)
    * **Betreff**: Ergebnisse der Ereignisüberprüfung
    * **Textkörper**: Ergebnis (_dynamischer Inhalt aus **Skript 2 ausführen**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Der fertige Office 365 Outlook-Connector in Power Automate.":::
1. Speichern Sie den Fluss. Verwenden Sie die Schaltfläche **"Test** " auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.
1. Sie sollten eine E-Mail mit der Meldung "Konflikt gefunden" erhalten. Daten erfordern Ihre Überprüfung." Dies weist darauf hin, dass es Unterschiede zwischen Zeilen in **speaker-registrations.xlsx** und Zeilen in **event-data.xlsx** gibt. Öffnen **Siespeaker-registrations.xlsx** , um mehrere hervorgehobene Zellen anzuzeigen, in denen potenzielle Probleme mit den Eintragen für die Lautsprecherregistrierung auftreten.
