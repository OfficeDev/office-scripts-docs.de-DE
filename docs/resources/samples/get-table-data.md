---
title: Ausgabe Excel Daten als JSON
description: Erfahren Sie, wie Excel tabellendaten als JSON ausgegeben werden, die in der Power Automate.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 9b8c0c48b969cfd05750ca4a6703a5ecbb9d18d2
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285815"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="b26f6-103">Ausgabe Excel Tabellendaten als JSON für die Verwendung in Power Automate</span><span class="sxs-lookup"><span data-stu-id="b26f6-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="b26f6-104">Excel Tabellendaten können als Array von Objekten in Form von JSON dargestellt werden.</span><span class="sxs-lookup"><span data-stu-id="b26f6-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="b26f6-105">Jedes Objekt stellt eine Zeile in der Tabelle dar.</span><span class="sxs-lookup"><span data-stu-id="b26f6-105">Each object represents a row in the table.</span></span> <span data-ttu-id="b26f6-106">Dadurch werden die Daten aus Excel in einem konsistenten Format extrahiert, das für den Benutzer sichtbar ist.</span><span class="sxs-lookup"><span data-stu-id="b26f6-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="b26f6-107">Die Daten können dann an andere Systeme über Power Automate werden.</span><span class="sxs-lookup"><span data-stu-id="b26f6-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="b26f6-108">_Eingabetabelle_</span><span class="sxs-lookup"><span data-stu-id="b26f6-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="Ein Arbeitsblatt mit Eingabetabellesdaten":::

<span data-ttu-id="b26f6-110">Eine Variante dieses Beispiels enthält auch die Hyperlinks in einer der Tabellenspalten.</span><span class="sxs-lookup"><span data-stu-id="b26f6-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="b26f6-111">Dadurch können zusätzliche Ebenen von Zelldaten im JSON angezeigt werden.</span><span class="sxs-lookup"><span data-stu-id="b26f6-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="b26f6-112">_Eingabetabelle mit Hyperlinks_</span><span class="sxs-lookup"><span data-stu-id="b26f6-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Ein Arbeitsblatt mit einer Spalte mit Tabellendaten, die als Hyperlinks formatiert sind":::

<span data-ttu-id="b26f6-114">_Dialogfeld zum Bearbeiten von Hyperlinks_</span><span class="sxs-lookup"><span data-stu-id="b26f6-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="Im Dialogfeld Hyperlink bearbeiten werden Optionen zum Ändern von Hyperlinks angezeigt":::

## <a name="sample-excel-file"></a><span data-ttu-id="b26f6-116">Beispieldatei Excel Datei</span><span class="sxs-lookup"><span data-stu-id="b26f6-116">Sample Excel file</span></span>

<span data-ttu-id="b26f6-117">Laden Sie die Datei <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx, </a> die in diesen Beispielen verwendet wird, herunter, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="b26f6-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="b26f6-118">Beispielcode: Zurückgeben von Tabellendaten als JSON</span><span class="sxs-lookup"><span data-stu-id="b26f6-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="b26f6-119">Sie können die Struktur `interface TableData` so ändern, dass sie ihren Tabellenspalten entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b26f6-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="b26f6-120">Beachten Sie, dass Sie bei Spaltennamen mit Leerzeichen den Schlüssel in Anführungszeichen setzen, z. B. `"Event ID"` mit im Beispiel.</span><span class="sxs-lookup"><span data-stu-id="b26f6-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="b26f6-121">Beispielausgabe aus dem Arbeitsblatt "PlainTable"</span><span class="sxs-lookup"><span data-stu-id="b26f6-121">Sample output from the "PlainTable" worksheet</span></span>

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

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="b26f6-122">Beispielcode: Zurückgeben von Tabellendaten als JSON mit Hyperlinktext</span><span class="sxs-lookup"><span data-stu-id="b26f6-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="b26f6-123">Das Skript extrahiert immer Hyperlinks aus der 4. Spalte (0 Index) der Tabelle.</span><span class="sxs-lookup"><span data-stu-id="b26f6-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="b26f6-124">Sie können diese Reihenfolge ändern oder mehrere Spalten als Hyperlinkdaten hinzufügen, indem Sie den Code unter dem Kommentar ändern. `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="b26f6-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
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

    objectArray.push(object);
  }
  return objectArray as TableData[];
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

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="b26f6-125">Beispielausgabe aus dem Arbeitsblatt "WithHyperLink"</span><span class="sxs-lookup"><span data-stu-id="b26f6-125">Sample output from the "WithHyperLink" worksheet</span></span>

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

## <a name="use-in-power-automate"></a><span data-ttu-id="b26f6-126">Verwenden in Power Automate</span><span class="sxs-lookup"><span data-stu-id="b26f6-126">Use in Power Automate</span></span>

<span data-ttu-id="b26f6-127">Informationen zur Verwendung eines solchen Skripts in Power Automate finden Sie unter [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="b26f6-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
