---
title: Filtern Excel Tabelle und Erhalten des sichtbaren Bereichs
description: Erfahren Sie, Office Skripts verwenden, um eine Excel zu filtern und den sichtbaren Bereich als Array von Objekten zu erhalten.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 196e39ffdfb7e6ff2d0898802665d3c2eccc7dbe
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285794"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="b995e-103">Filtern Excel Tabelle und Erhalten des sichtbaren Bereichs als JSON-Objekt</span><span class="sxs-lookup"><span data-stu-id="b995e-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="b995e-104">In diesem Beispiel wird eine Excel filtert und der sichtbare Bereich als JSON-Objekt zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="b995e-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="b995e-105">Diese JSON könnte einem Power Automate als Teil einer größeren Lösung bereitgestellt werden.</span><span class="sxs-lookup"><span data-stu-id="b995e-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="b995e-106">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="b995e-106">Example scenario</span></span>

* <span data-ttu-id="b995e-107">Wenden Sie einen Filter auf eine Tabellenspalte an.</span><span class="sxs-lookup"><span data-stu-id="b995e-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="b995e-108">Extrahieren Sie den sichtbaren Bereich nach dem Filtern.</span><span class="sxs-lookup"><span data-stu-id="b995e-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="b995e-109">Zusammenstellen und Zurückgeben eines Objekts mit einer [bestimmten JSON-Struktur](#sample-json).</span><span class="sxs-lookup"><span data-stu-id="b995e-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="b995e-110">Beispielcode: Filtern einer Tabelle und Erhalten eines sichtbaren Bereichs</span><span class="sxs-lookup"><span data-stu-id="b995e-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="b995e-111">Das folgende Skript filtert eine Tabelle und ruft den sichtbaren Bereich ab.</span><span class="sxs-lookup"><span data-stu-id="b995e-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="b995e-112">Laden Sie die Beispieldatei <a href="table-filter.xlsx">table-filter.xlsx</a> und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="b995e-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  // Get the "Station" column to use as key values in the filter.
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(value => value[0] as string);

  // Filter out repeated keys. This call to `filter` only returns the first instance of every unique element in the array.
  const uniqueKeys = keyColumnValues.filter((value, index, array) => array.indexOf(value) === index);
  console.log(uniqueKeys);

  const stationData: ReturnTemplate = {};

  // Filter the table to show only rows corresponding to each key.
  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    
    // Get the visible view when a single filter is active.
    const rangeView = table1.getRange().getVisibleView();

    // Create a JSON object with every visible row.
    stationData[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  });

  // Remove the filters.
  table1.getColumnByName('Station').getFilter().clear();

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(stationData));
  return stationData;
}

// This function converts a 2D-array of values into a generic JSON object.
function returnObjectFromValues(values: string[][]): BasicObject[] {
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

  return objectArray;
}

interface BasicObject {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObject[]
}
```

### <a name="sample-json"></a><span data-ttu-id="b995e-113">Beispiel-JSON</span><span class="sxs-lookup"><span data-stu-id="b995e-113">Sample JSON</span></span>

<span data-ttu-id="b995e-114">Jeder Schlüssel stellt einen eindeutigen Wert einer Tabelle dar.</span><span class="sxs-lookup"><span data-stu-id="b995e-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="b995e-115">Jede Arrayinstanz stellt die Zeile dar, die angezeigt wird, wenn der entsprechende Filter angewendet wird.</span><span class="sxs-lookup"><span data-stu-id="b995e-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="b995e-116">Schulungsvideo: Filtern einer Excel Tabelle und Erhalten des sichtbaren Bereichs</span><span class="sxs-lookup"><span data-stu-id="b995e-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="b995e-117">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/Mv7BrvPq84A)</span><span class="sxs-lookup"><span data-stu-id="b995e-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
