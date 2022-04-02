---
title: Kombinieren von Arbeitsmappen in einer einzigen Arbeitsmappe
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um Zusammenführungsarbeitsblätter aus anderen Arbeitsmappen in einer einzelnen Arbeitsmappe zu erstellen.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: f90980f2e2d1f125f4ca2ffb80822f13ecdeed0e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585884"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>Kombinieren von Arbeitsblättern in einer einzelnen Arbeitsmappe

In diesem Beispiel wird gezeigt, wie Sie Daten aus mehreren Arbeitsmappen in eine einzelne, zentrale Arbeitsmappe ziehen. Es verwendet zwei Skripts: eins zum Abrufen von Informationen aus einer Arbeitsmappe und ein weiteres zum Erstellen neuer Arbeitsblätter mit diesen Informationen. Sie kombiniert die Skripts in einem Power Automate Fluss, der für einen gesamten OneDrive Ordner fungiert.

> [!IMPORTANT]
> In diesem Beispiel werden nur die Werte aus den anderen Arbeitsmappen kopiert. Formatierungen, Diagramme, Tabellen oder andere Objekte werden nicht beibehalten.

## <a name="scenario"></a>Szenario

1. Erstellen Sie eine neue Excel-Datei in Ihrer OneDrive, und fügen Sie ihr zwei Skripts aus diesem Beispiel hinzu.
1. Erstellen Sie einen Ordner in Ihrem OneDrive, und fügen Sie eine oder mehrere Arbeitsmappen mit Daten hinzu.
1. Erstellen Sie einen Flow, um alle Dateien dieses Ordners abzurufen.
1. Verwenden Sie das **Skript zum Zurückgeben von Arbeitsblattdaten** , um die Daten aus jedem Arbeitsblatt in jeder Arbeitsmappe abzurufen.
1. Verwenden Sie das Skript " **Arbeitsblätter hinzufügen** ", um ein neues Arbeitsblatt in einer einzelnen Arbeitsmappe für jedes Arbeitsblatt in allen anderen Dateien zu erstellen.

## <a name="sample-code-return-worksheet-data"></a>Beispielcode: Zurückgeben von Arbeitsblattdaten

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>Beispielcode: Hinzufügen von Arbeitsblättern

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate Ablauf: Kombinieren von Arbeitsblättern in einer einzelnen Arbeitsmappe

1. Melden Sie sich bei [Power Automate an](https://flow.microsoft.com), und erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen Sie **"Manuell auslösen" aus,** und wählen Sie " **Erstellen**" aus.
1. Rufen Sie alle Dateien im Ordner ab. In diesem Beispiel verwenden wir einen Ordner mit dem Namen "output". Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Listendateien in Ordneraktion** verwendet. Geben Sie den Ordnerpfad an, der die .csv Dateien enthält.
    * **Ordner**: /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="Der fertige OneDrive for Business-Connector in Power Automate.":::
1. Führen Sie das **Skript zum Zurückgeben von Arbeitsblattdaten** aus, um alle Daten aus den einzelnen Arbeitsmappen abzurufen. Fügen Sie den **connector Excel Online (Business)** mit der **Skriptaktion ausführen** hinzu. Verwenden Sie die folgenden Werte für die Aktion. Beachten Sie, dass beim Hinzufügen der *ID* für die Datei Power Automate die Aktion in ein **"Übernehmen" für jedes** Steuerelement umschließen, sodass die Aktion für jede Datei ausgeführt wird.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: *ID* (dynamischer Inhalt aus **Listendateien im Ordner**)
    * **Skript**: Zurückgeben von Arbeitsblattdaten
1. Führen Sie das Skript **"Arbeitsblätter hinzufügen**" in der neuen Excel datei aus, die Sie erstellt haben. Dadurch werden die Daten aus allen anderen Arbeitsmappen hinzugefügt. Fügen Sie nach der vorherigen **Skriptausführungsaktion** und innerhalb des Steuerelements **"Auf jedes Steuerelement anwenden****" einen Excel Online-Connector (Business)** mit der **Skriptaktion ausführen** hinzu. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: Ihre Datei
    * **Skript**: Hinzufügen von Arbeitsblättern
    * **workbookName**: *Name* (dynamischer Inhalt aus **Listendateien im Ordner**)
    * **worksheetInformation** (nachdem Sie die **Schaltfläche "Wechseln" ausgewählt haben, um die gesamte Arrayschaltfläche einzugeben** , sehen Sie sich die Notiz nach dem nächsten Bild an): *Ergebnis* (dynamischer Inhalt aus **Dem Skript ausführen**)

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="Die beiden Ausführen-Skriptaktionen innerhalb des Steuerelements &quot;Auf jedes Steuerelement anwenden&quot;.":::
    > [!NOTE]
    > Wählen Sie die **Schaltfläche "Switch" aus, um die gesamte Arrayschaltfläche** einzugeben, um das Arrayobjekt direkt anstelle einzelner Elemente für das Array hinzuzufügen.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="Die Schaltfläche, um zu wechseln, um ein gesamtes Array in ein Eingabefeld eines Steuerelementfelds einzugeben.":::
1. Speichern Sie den Fluss. Verwenden Sie die Schaltfläche **"Test** " auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.
1. Ihre Excel-Datei sollte nun neue Arbeitsblätter haben.
