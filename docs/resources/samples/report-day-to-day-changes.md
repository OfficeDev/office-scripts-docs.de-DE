---
title: Aufzeichnen von täglichen Änderungen in Excel und Berichterstellung mit einem Power Automate-Fluss
description: Erfahren Sie, wie Sie Mithilfe von Office-Skripts und Power Automate Wertänderungen in einer Arbeitsmappe nachverfolgen können
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572652"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Aufzeichnen von täglichen Änderungen in Excel und Berichterstellung mit einem Power Automate-Fluss

Power Automate- und Office-Skripts werden kombiniert, um sich wiederholende Aufgaben für Sie zu erledigen. In diesem Beispiel haben Sie die Aufgabe, jeden Tag einen einzelnen numerischen Lesevorgang in einer Arbeitsmappe aufzuzeichnen und die Änderung seit gestern zu melden. Sie erstellen einen Fluss, um diesen Lesevorgang abzurufen, ihn in der Arbeitsmappe zu protokollieren und die Änderung per E-Mail zu melden.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [daily-readings.xlsx](daily-readings.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-record-and-report-daily-readings"></a>Beispielcode: Aufzeichnen und Melden von täglichen Messwerten

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>Beispielfluss: Melden von täglichen Änderungen

Führen Sie die folgenden Schritte aus, um einen [Power Automate-Fluss](https://powerautomate.microsoft.com/) für das Beispiel zu erstellen.

1. Erstellen Sie einen neuen **geplanten Cloudfluss**.
1. Planen Sie den Ablauf so, dass er jeden **1 Tag** wiederholt wird.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="Der Schritt zum Erstellen des Flusses, in dem sie angezeigt wird, wird jeden Tag wiederholt.":::
1. Wählen Sie **Erstellen** aus.
1. In einem echten Fluss fügen Sie einen Schritt hinzu, der Ihre Daten abruft. Die Daten können aus einer anderen Arbeitsmappe, einer adaptiven Teams-Karte oder einer anderen Quelle stammen. Um das Beispiel zu testen, erstellen Sie eine Testnummer. Fügen Sie einen neuen Schritt mit der **Initialisierungsvariablenaktion** hinzu. Weisen Sie ihr die folgenden Werte zu.
    1. **Name**: Eingabe
    1. **Typ**: Integer
    1. **Wert**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="Die Initialisierungsvariablenaktion mit den angegebenen Werten.":::
1. Fügen Sie einen neuen Schritt mit dem **Excel Online (Business)** -Connector mit der **Skriptaktion "Ausführen" hinzu** . Verwenden Sie die folgenden Werte für die Aktion.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei**: daily-readings.xlsx *(Über den Dateibrowser ausgewählt)*
    1. **Skript**: Ihr Skriptname
    1. **newData**: Input *(dynamischer Inhalt)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="Die Skriptaktion ausführen mit den angegebenen Werten.":::
1. Das Skript gibt den täglichen Leseunterschied als dynamischen Inhalt mit dem Namen "Ergebnis" zurück. Für das Beispiel können Sie die Informationen per E-Mail an sich selbst senden. Erstellen Sie einen neuen Schritt, der den **Outlook-Connector** mit der Aktion " **E-Mail senden(V2)" (** oder einem beliebigen E-Mail-Client) verwendet. Verwenden Sie die folgenden Werte, um die Aktion abzuschließen.
    1. **An**: Ihre E-Mail-Adresse
    1. **Betreff**: Tägliche Leseänderung
    1. **Textkörper**: Ergebnis "Unterschied zu gestern" *(dynamischer Inhalt aus Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="Der abgeschlossene Outlook-Connector in Power Automate.":::
1. Speichern Sie den Fluss, und probieren Sie ihn aus. Verwenden Sie die Schaltfläche " **Testen** " auf der Fluss-Editor-Seite. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.
