---
title: Ausgabe von Excel-Daten als JSON
description: Erfahren Sie, wie Sie Excel-Tabellendaten als JSON zur Verwendung in Power Automate aus.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 678506fee0b6a41ede8245fb360d485d635e2d64
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571430"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="bb915-103">Ausgabe von Excel-Tabellendaten als JSON für die Verwendung in Power Automate</span><span class="sxs-lookup"><span data-stu-id="bb915-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="bb915-104">Excel-Tabellendaten können als Array von Objekten in Form von JSON dargestellt werden.</span><span class="sxs-lookup"><span data-stu-id="bb915-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="bb915-105">Jedes Objekt stellt eine Zeile in der Tabelle dar.</span><span class="sxs-lookup"><span data-stu-id="bb915-105">Each object represents a row in the table.</span></span> <span data-ttu-id="bb915-106">Dadurch werden die Daten aus Excel in einem konsistenten Format extrahiert, das für den Benutzer sichtbar ist.</span><span class="sxs-lookup"><span data-stu-id="bb915-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="bb915-107">Die Daten können dann über Power Automate-Flüsse an andere Systeme gegeben werden.</span><span class="sxs-lookup"><span data-stu-id="bb915-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="bb915-108">_Eingabetabelle_</span><span class="sxs-lookup"><span data-stu-id="bb915-108">_Input table data_</span></span>

![Screenshot zum Anzeigen von Eingabetabellesdaten](../../images/table-input.png)

<span data-ttu-id="bb915-110">Eine Variante dieses Beispiels enthält auch die Hyperlinks in einer der Tabellenspalten.</span><span class="sxs-lookup"><span data-stu-id="bb915-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="bb915-111">Dadurch können zusätzliche Ebenen von Zelldaten im JSON angezeigt werden.</span><span class="sxs-lookup"><span data-stu-id="bb915-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="bb915-112">_Eingabetabelle mit Hyperlinks_</span><span class="sxs-lookup"><span data-stu-id="bb915-112">_Input table data that includes hyperlinks_</span></span>

![Screenshot mit Tabellendaten, die Hyperlinks enthalten](../../images/table-hyperlink-view.png)

<span data-ttu-id="bb915-114">_Dialogfeld zum Bearbeiten von Hyperlinks_</span><span class="sxs-lookup"><span data-stu-id="bb915-114">_Dialog to edit hyperlink_</span></span>

![Screenshot des Dialogfelds zum Bearbeiten des Hyperlinks](../../images/table-hyperlink-edit.png)

## <a name="sample-excel-file"></a><span data-ttu-id="bb915-116">Beispiel-Excel-Datei</span><span class="sxs-lookup"><span data-stu-id="bb915-116">Sample Excel file</span></span>

<span data-ttu-id="bb915-117">Laden Sie die Datei <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx, </a> die in diesen Beispielen verwendet wird, herunter, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="bb915-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="bb915-118">Beispielcode: Zurückgeben von Tabellendaten als JSON</span><span class="sxs-lookup"><span data-stu-id="bb915-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="bb915-119">Sie können die Struktur `interface TableData` so ändern, dass sie ihren Tabellenspalten entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bb915-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="bb915-120">Beachten Sie, dass Sie bei Spaltennamen mit Leerzeichen den Schlüssel in Anführungszeichen setzen, z. B. `"Event ID"` mit im Beispiel.</span><span class="sxs-lookup"><span data-stu-id="bb915-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="bb915-121">Beispielausgabe</span><span class="sxs-lookup"><span data-stu-id="bb915-121">Sample output</span></span>

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

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="bb915-122">Beispielcode: Zurückgeben von Tabellendaten als JSON mit Hyperlinktext</span><span class="sxs-lookup"><span data-stu-id="bb915-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="bb915-123">Das Skript extrahiert immer Hyperlinks aus der 4. Spalte (0 Index) der Tabelle.</span><span class="sxs-lookup"><span data-stu-id="bb915-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="bb915-124">Sie können diese Reihenfolge ändern oder mehrere Spalten als Hyperlinkdaten hinzufügen, indem Sie den Code unter dem Kommentar ändern. `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="bb915-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
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

### <a name="sample-output"></a><span data-ttu-id="bb915-125">Beispielausgabe</span><span class="sxs-lookup"><span data-stu-id="bb915-125">Sample output</span></span>

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

## <a name="use-in-power-automate"></a><span data-ttu-id="bb915-126">Verwenden in Power Automate</span><span class="sxs-lookup"><span data-stu-id="bb915-126">Use in Power Automate</span></span>

<span data-ttu-id="bb915-127">Informationen zur Verwendung eines solchen Skripts in Power Automate finden Sie unter [Erstellen eines automatisierten Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="bb915-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
