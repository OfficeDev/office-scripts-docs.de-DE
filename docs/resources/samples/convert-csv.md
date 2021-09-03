---
title: Konvertieren von CSV-Dateien in Excel Arbeitsmappen
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um .xlsx Dateien aus .csv Dateien zu erstellen.
ms.date: 07/19/2021
localization_priority: Normal
ms.openlocfilehash: d67be06dc038fc22215426e5f7143e0af9ba9f0c
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862195"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Konvertieren von CSV-Dateien in Excel Arbeitsmappen

Viele Dienste exportieren Daten als CSV-Dateien (Kommastrennte Werte). Diese Lösung automatisiert das Konvertieren dieser CSV-Dateien in Excel Arbeitsmappen im .xlsx Dateiformat. Es wird ein [Power Automate](https://flow.microsoft.com) Ablauf verwendet, um Dateien mit der .csv Erweiterung in einem OneDrive Ordner und ein Office Skript zu suchen, um die Daten aus der .csv-Datei in eine neue Excel Arbeitsmappe zu kopieren.

## <a name="solution"></a>Lösung

1. Store die .csv dateien und eine leere "Template"-.xlsx-Datei in einem OneDrive Ordner.
1. Erstellen Sie ein Office Skript, um die CSV-Daten in einem Bereich zu analysieren.
1. Erstellen Sie einen Power Automate Fluss, um die .csv-Dateien zu lesen und deren Inhalte an das Skript zu übergeben.

## <a name="sample-files"></a>Beispieldateien

Laden Sie <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> herunter, um die Template.xlsx-Datei und zwei Beispieldateien .csv abzurufen. Extrahieren Sie die Dateien in einen Ordner in Ihrem OneDrive. In diesem Beispiel wird davon ausgegangen, dass der Ordner den Namen "output" hat.

Fügen Sie das folgende Skript hinzu, und erstellen Sie einen Fluss mit den schritten, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Beispielcode: Einfügen von kommagetrennten Werten in eine Arbeitsmappe

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate Ablauf: Erstellen neuer .xlsx Dateien

1. Melden Sie sich bei [Power Automate an,](https://flow.microsoft.com) und erstellen Sie einen neuen **geplanten Cloudfluss.**
1. Legen Sie den Fluss so fest, dass jeder "1" "Tag" **wiederholt** wird, und wählen Sie **"Erstellen"** aus.
1. Dient zum Abrufen der Vorlage Excel Datei. Dies ist die Basis für alle konvertierten .csv Dateien. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business-Connector** und die Aktion zum Abrufen von **Dateiinhalten** verwendet. Geben Sie den Dateipfad zur Datei "Template.xlsx" an.
    * **Datei:**/output/Template.xlsx
1. Benennen Sie den **Schritt "Dateiinhalt abrufen"** um, indem Sie das **Menü ...** dieses Schritts (in der oberen rechten Ecke des Connectors) aufrufen und die Option **"Umbenennen"** auswählen. Ändern Sie den Schrittnamen in "Abrufen Excel Vorlage".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Der fertige OneDrive for Business Connector in Power Automate, umbenannt in &quot;Get Excel vorlage&quot;.":::
1. Rufen Sie alle Dateien im Ordner "Ausgabe" ab. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Listendateien in Ordneraktion** verwendet. Geben Sie den Ordnerpfad an, der die .csv Dateien enthält.
    * **Ordner:**/output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Der fertige OneDrive for Business-Connector in Power Automate.":::
1. Fügen Sie eine Bedingung hinzu, damit der Fluss nur für .csv Dateien ausgeführt wird. Fügen Sie einen **neuen Schritt** hinzu, bei dem es sich um das **Bedingungssteuerelement** handelt. Verwenden Sie die folgenden Werte für die **Bedingung**.
    * **Wählen Sie einen Wert** aus: *Name* (dynamischer Inhalt aus **Listendateien im Ordner).** Beachten Sie, dass dieser dynamische Inhalt mehrere Ergebnisse hat, sodass ein **Apply-Steuerelement für jedes** *Wertsteuerelement* die **Bedingung** umgibt.
    * **endet mit** (aus der Dropdownliste)
    * **Wählen Sie einen Wert** aus: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Das abgeschlossene Bedingungssteuerelement mit dem Steuerelement &quot;Anwenden&quot; auf jedes Steuerelement, das es umgibt.":::
1. Der Rest des Flusses befindet sich unter dem Abschnitt **"Wenn ja",** da wir nur auf .csv Dateien reagieren möchten. Rufen Sie eine einzelne .csv Datei ab, indem Sie einen **neuen Schritt** hinzufügen, der den **OneDrive for Business** Connector und die **Aktion "Dateiinhalt abrufen"** verwendet. Verwenden Sie die **ID** aus dem dynamischen Inhalt aus **Listendateien im Ordner**.
    * **Datei:** *ID* (dynamischer Inhalt aus den **Listendateien im Ordnerschritt)**
1. Benennen Sie den neuen Inhaltsschritt **"Datei abrufen"** in "Get .csv file" um. Dadurch wird diese Datei von der Excel Vorlage unterschieden.
1. Erstellen Sie die neue .xlsx-Datei mithilfe der Excel-Vorlage als Basisinhalt. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Aktion "Datei erstellen"** verwendet. Verwenden Sie die folgenden Werte.
    * **Ordnerpfad:**/output
    * **Dateiname:** *Name ohne Erweiterung*.xlsx (wählen Sie den dynamischen Inhalt *"Name ohne Erweiterung"* aus den **Listendateien im Ordner** aus, und geben Sie danach manuell ".xlsx" ein)
    * **Dateiinhalt:** *Dateiinhalt* (dynamischer Inhalt aus **der Vorlage "Get Excel")**

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Die Schritte &quot;Datei abrufen .csv&quot; und &quot;Datei erstellen&quot; des Power Automate-Flusses.":::
1. Führen Sie das Skript aus, um Daten in die neue Arbeitsmappe zu kopieren. Fügen Sie den **connector Excel Online (Business)** mit der **Skriptaktion ausführen** hinzu. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei:** *ID* (dynamischer Inhalt aus **Datei erstellen)**
    * **Skript:** Csv konvertieren
    * **csv**: *File content* (dynamic content from Get **.csv file**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Der fertige Excel Online (Business)-Connector in Power Automate.":::
1. Speichern Sie den Fluss. Verwenden Sie die Schaltfläche **"Test"** auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse"** aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.
1. Sie sollten neben den ursprünglichen .csv Dateien neue .xlsx Dateien im Ordner "Ausgabe" finden. Die neuen Arbeitsmappen enthalten dieselben Daten wie die CSV-Dateien.
