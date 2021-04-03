---
title: Filtern der Excel-Tabelle und Erhalten des sichtbaren Bereichs
description: Erfahren Sie, wie Sie Office-Skripts verwenden, um eine Excel-Tabelle zu filtern und den sichtbaren Bereich als Array von Objekten zu erhalten.
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571467"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtern der Excel-Tabelle und Erhalten des sichtbaren Bereichs als JSON-Objekt

In diesem Beispiel wird eine Excel-Tabelle filtert und der sichtbare Bereich als JSON-Objekt zurückgegeben. Diese JSON könnte einem Power Automate-Fluss als Teil einer größeren Lösung bereitgestellt werden.

## <a name="example-scenario"></a>Beispielszenario

* Wenden Sie einen Filter auf eine Tabellenspalte an.
* Extrahieren Sie den sichtbaren Bereich nach dem Filtern.
* Zusammenstellen und Zurückgeben eines Objekts mit einer [bestimmten JSON-Struktur](#sample-json).

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Beispielcode: Filtern einer Tabelle und Erhalten eines sichtbaren Bereichs

Das folgende Skript filtert eine Tabelle und ruft den sichtbaren Bereich ab.

Laden Sie die Beispieldatei <a href="table-filter.xlsx">table-filter.xlsx</a> und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a>Beispiel-JSON

Jeder Schlüssel stellt einen eindeutigen Wert einer Tabelle dar. Jede Arrayinstanz stellt die Zeile dar, die angezeigt wird, wenn der entsprechende Filter angewendet wird.

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Schulungsvideo: Filtern einer Excel-Tabelle und Erhalten des sichtbaren Bereichs

[![Schritt-für-Schritt-Video zum Filtern einer Excel-Tabelle und zum Erhalten des sichtbaren Bereichs ansehen](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Schritt-für-Schritt-Video zum Filtern einer Excel-Tabelle und zum Erhalten des sichtbaren Bereichs")
