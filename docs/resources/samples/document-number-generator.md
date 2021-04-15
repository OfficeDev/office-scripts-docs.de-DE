---
title: Generieren eines eindeutigen Bezeichners in einer Arbeitsmappe
description: Erfahren Sie, wie Sie Office-Skripts verwenden, um einen eindeutigen Bezeichner zu generieren und einer Tabelle und einem Bereich eine Zeile hinzuzufügen.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: c75eba551cc1c05426db613c78d41431b5177842
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754838"
---
# <a name="generate-a-unique-identifier-in-a-workbook"></a>Generieren eines eindeutigen Bezeichners in einer Arbeitsmappe

Dieses Szenario hilft einem Benutzer, eine eindeutige Dokumentnummer mit einem bestimmten Format zu generieren und sie als Eintrag zu einem Bereich oder einer Tabelle hinzuzufügen. Der neue Eintrag oder die neue Hinzugefügte Zeile enthält die neu generierte eindeutige Dokumentnummer und einige andere Attribute, die an das Skript übergeben werden.

Es gibt zwei Versionen des Beispiels für dieses Szenario.

* [Version 1: Lesen und Hinzufügen einer Zeile zu einem Arbeitsblatt mit einem einfachen Bereich](#sample-code-generate-key-and-add-row-to-range)

    _Bevor die neue Zeile hinzugefügt wird_

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="Ein Arbeitsblatt mit einem Datenbereich, bevor eine Zeile hinzugefügt wird.":::

    _Nachdem die neue Zeile hinzugefügt wurde_

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="Ein Arbeitsblatt mit einem Datenbereich nach dem Hinzufügen der Zeile.":::

* [Version 2: Lesen und Hinzufügen einer Zeile zu einer Tabelle](#sample-code-generate-key-and-add-row-to-table)

    _Bevor die neue Zeile hinzugefügt wird_

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="Ein Arbeitsblatt mit einer Tabelle, bevor eine Zeile hinzugefügt wird.":::

    _Nachdem die neue Zeile hinzugefügt wurde_

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="Ein Arbeitsblatt mit einer Tabelle, nachdem eine Zeile hinzugefügt wurde.":::

## <a name="sample-excel-file"></a>Beispiel-Excel-Datei

Laden Sie die Datei <a href="document-number-generator.xlsx">document-number-generator.xlsx, </a> die in dieser Lösung verwendet wird, um sie selbst auszuprobieren!

## <a name="sample-code-generate-key-and-add-row-to-range"></a>Beispielcode: Generieren des Schlüssels und Hinzufügen einer Zeile zum Bereich

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX  = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input:RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' && 
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('PlainSheet'); /* plain range sheet */
    const range = sheet.getUsedRange();

    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length-1];

    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;
    
    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);
    
    // Get last row and compute next row address.
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);

    // Add a row with incoming data plus the computed key value.
    target.setValues([
      [
        nextKey, 
        /* Capitalize the document type. */
        input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
        input.documentName
      ]
    ])
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```

## <a name="sample-code-generate-key-and-add-row-to-table"></a>Beispielcode: Generieren des Schlüssels und Hinzufügen einer Zeile zur Tabelle

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input: RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('TableSheet'); /* table sheet */
    const table = sheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it.
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length - 1];


    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;

    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);

    // Add a row with incoming data plus the computed key value.
    table.addRow(-1, [
            nextKey,
            /* Capitalize the document type. */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```
