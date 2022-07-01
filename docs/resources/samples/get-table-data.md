---
title: Excel-Daten als JSON ausgeben
description: Erfahren Sie, wie Sie Excel-Tabellendaten als JSON ausgeben, um sie in Power Automate zu verwenden.
ms.date: 06/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6453d9f0e92f9b3fcccc6e3ec9c1b6c9af49859c
ms.sourcegitcommit: 82fb78e6907b7c3b95c5c53cfc83af4ea1067a78
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/01/2022
ms.locfileid: "66572342"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a>Ausgabe von Excel-Tabellendaten als JSON für die Verwendung in Power Automate

Excel-Tabellendaten können als Array von Objekten in Form von [JSON](https://www.w3schools.com/whatis/whatis_json.asp) dargestellt werden. Jedes Objekt stellt eine Zeile in der Tabelle dar. Dies hilft beim Extrahieren der Daten aus Excel in einem konsistenten Format, das für den Benutzer sichtbar ist. Die Daten können dann über Power Automate-Flüsse an andere Systeme übergeben werden.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie die Datei <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> für eine sofort einsatzbereite Arbeitsmappe herunter.

:::image type="content" source="../../images/table-input.png" alt-text="Ein Arbeitsblatt mit Eingabetabellendaten.":::

Eine Variante dieses Beispiels enthält auch die Hyperlinks in einer der Tabellenspalten. Dadurch können zusätzliche Ebenen von Zelldaten im JSON-Code angezeigt werden.

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Ein Arbeitsblatt mit einer Spalte mit Tabellendaten, die als Hyperlinks formatiert sind.":::

## <a name="sample-code-return-table-data-as-json"></a>Beispielcode: Zurückgeben von Tabellendaten als JSON

Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

> [!NOTE]
> Sie können die `interface TableData` Struktur so ändern, dass sie ihren Tabellenspalten entspricht. Beachten Sie, dass Sie bei Spaltennamen mit Leerzeichen ihren Schlüssel in Anführungszeichen platzieren, z. B. mit `"Event ID"` im Beispiel. Weitere Informationen zum Arbeiten mit JSON finden Sie unter [Verwenden von JSON zum Übergeben von Daten an und von Office-Skripts](../../develop/use-json.md).

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));
  return returnObjects
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object: {[key: string]: string} = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object as unknown as TableData);
  }

  return objectArray;
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a>Beispielausgabe des Arbeitsblatts "PlainTable"

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a>Beispielcode: Zurückgeben von Tabellendaten als JSON mit Linktext

> [!NOTE]
> Das Skript extrahiert immer Links aus der 4. Spalte (0 Index) der Tabelle. Sie können diese Reihenfolge ändern oder mehrere Spalten als Linkdaten einschließen, indem Sie den Code unter dem Kommentar ändern. `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray : TableData[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object as TableData);
  }
  return objectArray;
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output-from-the-withhyperlink-worksheet"></a>Beispielausgabe des Arbeitsblatts "WithHyperLink"

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a>Verwenden in Power Automate

Informationen zur Verwendung eines solchen Skripts in Power Automate finden [Sie unter Erstellen eines automatisierten Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).
