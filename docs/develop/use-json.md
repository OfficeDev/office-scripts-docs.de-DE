---
title: Verwenden von JSON zum Übergeben von Daten an und von Office Skripts
description: Erfahren Sie, wie Sie Daten in JSON-Objekte für die Verwendung mit externen Aufrufen und Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088152"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>Verwenden von JSON zum Übergeben von Daten an und von Office Skripts

[JSON (JavaScript Object Notation)](https://www.w3schools.com/whatis/whatis_json.asp) ist ein Format zum Speichern und Übertragen von Daten. Jedes JSON-Objekt ist eine Auflistung von Name-Wert-Paaren, die beim Erstellen definiert werden können. JSON ist bei Office-Skripts nützlich, da es die beliebige Komplexität von Bereichen, Tabellen und anderen Datenmustern in Excel verarbeiten kann. Mit JSON können Sie eingehende Daten aus [Webdiensten](external-calls.md) analysieren und komplexe Objekte durch [Power Automate Flüsse](power-automate-integration.md) übergeben.

Dieser Artikel konzentriert sich auf die Verwendung von JSON mit Office Skripts. Wir empfehlen Ihnen, zuerst mehr über das Format aus Artikeln wie [JSON-Einführung](https://www.w3schools.com/js/js_json_intro.asp) von W3-Schulen zu erfahren.

## <a name="parse-json-data-into-a-range-or-table"></a>Analysieren von JSON-Daten in einen Bereich oder eine Tabelle

Arrays von JSON-Objekten bieten eine konsistente Möglichkeit, Zeilen mit Tabellendaten zwischen Anwendungen und Webdiensten zu übergeben. In diesen Fällen stellt jedes JSON-Objekt eine Zeile dar, während die Eigenschaften die Spalten darstellen. Ein Office-Skript kann ein JSON-Array durchlaufen und als 2D-Array neu zusammensetzen. Dieses Array wird dann als Werte eines Bereichs festgelegt und in einer Arbeitsmappe gespeichert. Die Eigenschaftennamen können auch als Kopfzeilen hinzugefügt werden, um eine Tabelle zu erstellen.

Das folgende Skript zeigt JSON-Daten, die in eine Tabelle konvertiert werden. Beachten Sie, dass die Daten nicht aus einer externen Quelle stammen. Dies wird weiter unten in diesem Artikel behandelt.

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> Wenn Sie die Struktur des JSON kennen, können Sie eine eigene Schnittstelle erstellen, um das Abrufen bestimmter Eigenschaften zu vereinfachen. Sie können die Json-zu-Array-Konvertierungsschritte durch typsichere Verweise ersetzen. Der folgende Codeausschnitt zeigt diese (jetzt auskommentierten) Schritte, die durch Aufrufe ersetzt werden, die eine neue `ActionRow` Schnittstelle verwenden. Beachten Sie, dass die `convertJsonToRow` Funktion dadurch nicht mehr benötigt wird.
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>Abrufen von JSON-Daten aus externen Quellen

Es gibt zwei Möglichkeiten, JSON-Daten über ein Office Script in Ihre Arbeitsmappe zu importieren.

- Als [Parameter](power-automate-integration.md#main-parameters-pass-data-to-a-script) mit einem Power Automate-Fluss.
- Mit einem `fetch` Aufruf eines [externen Webdiensts](external-calls.md).

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Ändern des Beispiels für die Arbeit mit Power Automate

JSON-Daten in Power Automate können als generisches Objektarray übergeben werden. Fügen Sie dem Skript eine `object[]` Eigenschaft hinzu, um diese Daten zu akzeptieren.

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

Anschließend wird im Power Automate Connector eine Option angezeigt, die der **Skriptaktion "Ausführen**" hinzugefügt `jsonData` werden soll.

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Ein Excel Online(Business)-Connector, der eine Skriptaktion ausführen mit dem Parameter &quot;jsonData&quot; anzeigt.":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>Ändern des Beispiels zur Verwendung eines Anrufs `fetch`

Webdienste können auf `fetch` Aufrufe mit JSON-Daten antworten. Dadurch erhält Ihr Skript die benötigten Daten, während Sie sich in Excel halten. Weitere Informationen zu `fetch` und externen Aufrufen finden Sie in [der Unterstützung externer API-Aufrufe in Office-Skripts](external-calls.md).

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>Erstellen von JSON aus einem Bereich

Die Zeilen und Spalten eines Arbeitsblatts implizieren häufig Beziehungen zwischen ihren Datenwerten. Eine Zeile einer Tabelle wird konzeptuell einem Programmierobjekt zugeordnet, wobei jede Spalte eine Eigenschaft dieses Objekts ist. Betrachten Sie die folgende Datentabelle. Jede Zeile stellt eine Transaktion dar, die in der Kalkulationstabelle aufgezeichnet wird.

|ID |Datum     |Betrag |Anbieter                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |$43.54 |Am besten für Sie Organics Company |
|2  |6/3/2022 |$67.23 |Liberty Bakery und Cafe       |
|3  |6/3/2022 |$37.12 |Am besten für Sie Organics Company |
|4  |6/6/2022 |$86.95 |Coho Vineyard                 |
|5  |6/7/2022 |$13.64 |Liberty Bakery und Cafe       |

Jeder Transaktion (jeder Zeile) ist ein Satz von Eigenschaften zugeordnet: "ID", "Date", "Amount" und "Vendor". Dies kann in einem Office Script als Objekt modelliert werden.

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

Die Zeilen in der Beispieltabelle entsprechen den Eigenschaften in der Schnittstelle, sodass ein Skript jede Zeile problemlos in ein `Transaction` Objekt konvertieren kann. Dies ist nützlich beim Ausgeben der Daten für Power Automate. Das folgende Skript durchreitet jede Zeile in der Tabelle und fügt sie einem `Transaction[]`hinzu.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="Die Konsolenausgabe des vorherigen Skripts, in der die Eigenschaftswerte des Objekts angezeigt werden.":::

### <a name="use-a-generic-object"></a>Verwenden eines generischen Objekts

Im vorherigen Beispiel wird davon ausgegangen, dass die Tabellenkopfzeilenwerte konsistent sind. Wenn Ihre Tabelle variable Spalten enthält, müssen Sie ein generisches JSON-Objekt erstellen. Das folgende Skript zeigt ein Skript, das eine beliebige Tabelle als JSON protokolliert.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>Siehe auch

- [Externe API-Anruf Unterstützung in Office-Skripts](external-calls.md)
- [Beispiel: Verwenden externer Abrufaufrufe in Office Skripts](../resources/samples/external-fetch-calls.md)
- [Ausführen Office Skripts mit Power Automate](power-automate-integration.md)