---
title: Filtern Excel Tabelle und Abrufen des sichtbaren Bereichs
description: Erfahren Sie, wie Sie Office Skripts verwenden, um eine Excel Tabelle zu filtern und den sichtbaren Bereich als Array von Objekten abzurufen.
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 103ec97111720ab872c0be843aa0573781d98c44
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088085"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtern Excel Tabelle und Abrufen des sichtbaren Bereichs als JSON-Objekt

In diesem Beispiel wird eine Excel Tabelle gefiltert und der sichtbare Bereich als [JSON-Objekt](https://www.w3schools.com/whatis/whatis_json.asp) zurückgegeben. Dieser JSON-Code kann einem Power Automate-Fluss als Teil einer größeren Lösung bereitgestellt werden.

## <a name="example-scenario"></a>Beispielszenario

* Wenden Sie einen Filter auf eine Tabellenspalte an.
* Extrahieren Sie den sichtbaren Bereich nach dem Filtern.
* Ein Objekt mit einer [bestimmten JSON-Struktur](#sample-json) zusammenstellen und zurückgeben.

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie <a href="table-filter.xlsx">table-filter.xlsx</a> für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Beispielcode: Filtern einer Tabelle und Abrufen eines sichtbaren Bereichs

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
  let objectArray: BasicObject[] = [];
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

### <a name="sample-json"></a>JSON-Beispiel

Jeder Schlüssel stellt einen eindeutigen Wert einer Tabelle dar. Jede Arrayinstanz stellt die Zeile dar, die sichtbar ist, wenn der entsprechende Filter angewendet wird. Weitere Informationen zum Arbeiten mit JSON finden Sie unter [Verwenden von JSON zum Übergeben von Daten an und von Office Skripts](../../develop/use-json.md).

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Schulungsvideo: Filtern einer Excel Tabelle und Abrufen des sichtbaren Bereichs

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/Mv7BrvPq84A).
