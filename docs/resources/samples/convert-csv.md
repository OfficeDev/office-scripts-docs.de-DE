---
title: Konvertieren von CSV-Dateien in Excel Arbeitsmappen
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um .xlsx Dateien aus .csv Dateien zu erstellen.
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52619c1867b654fae3fce1a383a612f81f80d868
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585590"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Konvertieren von CSV-Dateien in Excel Arbeitsmappen

Viele Dienste exportieren Daten als CSV-Dateien (Kommastrennte Werte). Diese Lösung automatisiert das Konvertieren dieser CSV-Dateien in Excel Arbeitsmappen im .xlsx Dateiformat. Es wird ein [Power Automate](https://flow.microsoft.com) Ablauf verwendet, um Dateien mit der Erweiterung .csv in einem OneDrive Ordner zu suchen, und ein Office Skript, um die Daten aus der .csv Datei in eine neue Excel Arbeitsmappe zu kopieren.

## <a name="solution"></a>Lösung

1. Store die .csv dateien und eine leere "Template"-.xlsx-Datei in einem OneDrive Ordner.
1. Erstellen Sie ein Office Skript, um die CSV-Daten in einem Bereich zu analysieren.
1. Erstellen Sie einen Power Automate Ablauf, um die .csv Dateien zu lesen und deren Inhalte an das Skript zu übergeben.

## <a name="sample-files"></a>Beispieldateien

Laden Sie <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> herunter, um die Template.xlsx-Datei und zwei Beispieldateien .csv abzurufen. Extrahieren Sie die Dateien in einen Ordner in Ihrem OneDrive. In diesem Beispiel wird davon ausgegangen, dass der Ordner den Namen "output" hat.

Fügen Sie das folgende Skript hinzu, und erstellen Sie einen Flow mit den schritten, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Beispielcode: Einfügen von kommagetrennten Werten in eine Arbeitsmappe

```TypeScript
/**
 * Convert incoming CSV data into a range and add it to the workbook.
 */
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  // Remove any Windows \r characters.
  csv = csv.replace(/\r/g, "");

  // Split each line into a row.
  let rows = csv.split("\n");
  /*
   * For each row, match the comma-separated sections.
   * For more information on how to use regular expressions to parse CSV files,
   * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
   */
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g
  rows.forEach((value, index) => {
    if (value.length > 0) {
        let row = value.match(csvMatchRegex);
    
        // Check for blanks at the start of the row.
        if (row[0].charAt(0) === ',') {
          row.unshift("");
        }
    
        // Remove the preceding comma.
        row.forEach((cell, index) => {
          row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
        });
    
        // Create a 2D array with one row.
        let data: string[][] = [];
        data.push(row);
    
        // Put the data in the worksheet.
        let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
        range.setValues(data);
    }
  });

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate Ablauf: Erstellen neuer .xlsx Dateien

1. Melden Sie sich bei [Power Automate an](https://flow.microsoft.com), und erstellen Sie einen neuen **geplanten Cloudfluss**.
1. Legen Sie den Fluss auf "Jeden "1" "Tag" **wiederholen** fest, und wählen Sie " **Erstellen**" aus.
1. Dient zum Abrufen der Vorlage Excel Datei. Dies ist die Basis für alle konvertierten .csv Dateien. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Aktion "Dateiinhalt abrufen**" verwendet. Geben Sie den Dateipfad zur Datei "Template.xlsx" an.
    * **Datei**: /output/Template.xlsx
1. Benennen Sie den Schritt " **Dateiinhalt abrufen** " um, indem Sie zum Menü **"Dateiinhalt abrufen(...)"** dieses Schritts (in der oberen rechten Ecke des Connectors) wechseln und die Option **"Umbenennen** " auswählen. Ändern Sie den Schrittnamen in "Abrufen Excel Vorlage".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Der fertige OneDrive for Business-Connector in Power Automate, umbenannt in Vorlage &quot;Get Excel&quot;.":::
1. Rufen Sie alle Dateien im Ordner "Ausgabe" ab. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Listendateien in Ordneraktion** verwendet. Geben Sie den Ordnerpfad an, der die .csv Dateien enthält.
    * **Ordner**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Der fertige OneDrive for Business-Connector in Power Automate.":::
1. Fügen Sie eine Bedingung hinzu, sodass der Fluss nur für .csv Dateien ausgeführt wird. Fügen Sie einen **neuen Schritt** hinzu, bei dem es sich um das **Bedingungssteuerelement** handelt. Verwenden Sie die folgenden Werte für die **Bedingung**.
    * **Wählen Sie einen Wert** aus: *Name* (dynamischer Inhalt aus **Listendateien im Ordner**). Beachten Sie, dass dieser dynamische Inhalt mehrere Ergebnisse hat, sodass ein **Apply to each** *value* control die **Condition** umgibt.
    * **endet mit** (aus der Dropdownliste)
    * **Wählen Sie einen Wert** aus: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Das abgeschlossene Bedingungssteuerelement mit dem Steuerelement &quot;Anwenden&quot; auf jedes Steuerelement, das es umgibt.":::
1. Der Rest des Flusses befindet sich unter dem Abschnitt **"Wenn ja** ", da wir nur auf .csv Dateien reagieren möchten. Rufen Sie eine einzelne .csv Datei ab, indem Sie einen **neuen Schritt** hinzufügen, der den **connector OneDrive for Business** und die **Aktion zum Abrufen von Dateiinhalten** verwendet. Verwenden Sie die **ID** aus dem dynamischen Inhalt aus **Listendateien im Ordner**.
    * **Datei**: *ID* (dynamischer Inhalt aus den **Listendateien im Ordnerschritt** )
1. Benennen Sie den neuen Inhaltsschritt " **Datei abrufen** " in "Get .csv file" um. Dadurch wird diese Datei von der Excel Vorlage unterschieden.
1. Erstellen Sie die neue .xlsx-Datei mithilfe der Excel-Vorlage als Basisinhalt. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Aktion "Datei erstellen"** verwendet. Verwenden Sie die folgenden Werte.
    * **Ordnerpfad**: /output
    * **Dateiname**: *Name ohne Erweiterung*.xlsx (wählen Sie den dynamischen Inhalt *"Name ohne Erweiterung* " aus den **Listendateien im Ordner** aus, und geben Sie danach manuell ".xlsx" ein)
    * **Dateiinhalt**: *Dateiinhalt* (dynamischer Inhalt aus **der Vorlage "Get Excel**")

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Die Schritte &quot;Datei abrufen .csv&quot; und &quot;Datei erstellen&quot; des Power Automate-Flusses.":::
1. Führen Sie das Skript aus, um Daten in die neue Arbeitsmappe zu kopieren. Fügen Sie den **connector Excel Online (Business)** mit der **Skriptaktion ausführen** hinzu. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: *ID* (dynamischer Inhalt aus **Datei erstellen**)
    * **Skript**: Csv konvertieren
    * **csv**: *File content* (dynamic content from **Get .csv file**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Der fertige Excel Online(Business)-Connector in Power Automate.":::
1. Speichern Sie den Fluss. Verwenden Sie die Schaltfläche **"Test** " auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.
1. Sie sollten neben den ursprünglichen .csv Dateien neue .xlsx Dateien im Ordner "Ausgabe" finden. Die neuen Arbeitsmappen enthalten dieselben Daten wie die CSV-Dateien.

## <a name="troubleshooting"></a>Problembehandlung

### <a name="script-testing"></a>Skripttests

Um das Skript zu testen, ohne Power Automate zu verwenden, weisen Sie einen Wert zu`csv`, bevor Sie es verwenden. Versuchen Sie, den folgenden Code als erste Zeile der `main` Funktion hinzuzufügen und **"Ausführen" zu** drücken.

```TypeScript
  csv = `1, 2, 3
         4, 5, 6
         7, 8, 9`;
```

### <a name="semicolon-separated-files-and-other-alternative-separators"></a>Durch Semikolons getrennte Dateien und andere alternative Trennzeichen

Einige Regionen verwenden Semikolons (';'), um Zellwerte anstelle von Kommas zu trennen. In diesem Fall müssen Sie die folgenden Zeilen im Skript ändern.

1. Ersetzen Sie die Kommas in der Regulären Ausdrucksanweisung durch Semikolons. Dies beginnt mit `let row = value.match`.

    ```TypeScript
    let row = value.match(/(?:;|\n|^)("(?:(?:"")*[^"]*)*"|[^";\n]*|(?:\n|$))/g);
    ```

1. Ersetzen Sie das Komma durch ein Semikolon in der Prüfung für die leere erste Zelle. Dies beginnt mit `if (row[0].charAt(0)`.

    ```TypeScript
    if (row[0].charAt(0) === ';') {
    ```

1. Ersetzen Sie das Komma durch ein Semikolon in der Zeile, das das Trennzeichen aus dem angezeigten Text entfernt. Dies beginnt mit `row[index] = cell.indexOf`.

   ```TypeScript
      row[index] = cell.indexOf(";") === 0 ? cell.substr(1) : cell;
    ```

> [!NOTE]
> Wenn Ihre Datei Registerkarten oder ein anderes Zeichen verwendet, um die Werte zu trennen, ersetzen Sie die `;` in den obigen Ersetzungen aufgeführten Zeichen durch `\t` ein beliebiges zeichen, das verwendet wird.

### <a name="large-csv-files"></a>Große CSV-Dateien

Wenn Ihre Datei Hunderte von Tausenden von Zellen aufweist, können Sie das [Excel Datenübertragungslimit](../../testing/platform-limits.md#excel) erreichen. Sie müssen erzwingen, dass das Skript regelmäßig mit Excel synchronisiert wird. Die einfachste Möglichkeit besteht darin, einen Aufruf `console.log` auszuführen, nachdem ein Zeilenbatch verarbeitet wurde. Fügen Sie die folgenden Codezeilen hinzu, um dies zu ermöglichen.

1. Fügen Sie zuvor `rows.forEach((value, index) => {`die folgende Zeile hinzu.

    ```TypeScript
      let rowCount = 0;
    ```

1. Fügen Sie danach `range.setValues(data);`den folgenden Code hinzu. Beachten Sie, dass Sie je nach Anzahl der Spalten möglicherweise auf eine niedrigere Zahl reduzieren `5000` müssen.

    ```TypeScript
      rowCount++;
      if (rowCount % 5000 === 0) {
        console.log("Syncing 5000 rows.");
      }
    ```

> [!WARNING]
> Wenn Ihre CSV-Datei sehr groß ist, können Probleme [beim Timingout in Power Automate](../../testing/platform-limits.md#power-automate) auftreten. Sie müssen die CSV-Daten in mehrere Dateien aufteilen, bevor Sie sie in Excel Arbeitsmappen konvertieren.
