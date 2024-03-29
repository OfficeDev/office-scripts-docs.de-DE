---
title: Ausführen eines Scripts für alle Excel-Dateien in einem Ordner
description: Erfahren Sie, wie Sie ein Skript für alle Excel Dateien in einem Ordner auf OneDrive for Business ausführen.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 924fd1b0ea1d50e18877b5a7808feb906959215e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585877"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Ausführen eines Scripts für alle Excel-Dateien in einem Ordner

Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden. Es kann auch für einen SharePoint Ordner verwendet werden.
Es führt Berechnungen für die Excel Dateien durch, fügt Formatierungen hinzu und fügt einen Kommentar ein, [der](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) @mentions einen Kollegen @mentions.

## <a name="sample-excel-files"></a>Beispieldateien für Excel

Laden Sie <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> für alle Arbeitsmappen herunter, die Sie für dieses Beispiel benötigen. Extrahieren Sie diese Dateien in einen Ordner mit dem Titel " **Sales**". Fügen Sie der Skriptsammlung das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Beispielcode: Hinzufügen von Formatierungen und Einfügen eines Kommentars

Dies ist das Skript, das für jede einzelne Arbeitsmappe ausgeführt wird.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate Ablauf: Ausführen des Skripts für jede Arbeitsmappe im Ordner

Dieser Fluss führt das Skript für jede Arbeitsmappe im Ordner "Vertrieb" aus.

1. Erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen Sie **"Manuell auslösen" aus,** und wählen Sie " **Erstellen**" aus.
1. Fügen Sie einen **neuen Schritt** hinzu, der den **OneDrive for Business** Connector und die **Listendateien in Ordneraktion** verwendet.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Der fertige OneDrive for Business-Connector in Power Automate.":::
1. Wählen Sie den Ordner "Vertrieb" mit den extrahierten Arbeitsmappen aus.
1. Um sicherzustellen, dass nur Arbeitsmappen ausgewählt sind, wählen Sie **"Neuer Schritt"** und dann " **Bedingung**" aus. Verwenden Sie die folgenden Werte für die Bedingung.
    1. **Name** (der Wert OneDrive Dateinamens)
    1. "Endet mit"
    1. "xlsx"

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Der Power Automate Bedingungsblock, der nachfolgende Aktionen auf jede Datei anwendet.":::
1. Fügen Sie unter der Verzweigung **"Wenn ja**" den **Connector Excel Online (Business)** mit der **Skriptaktion ausführen hinzu**. Verwenden Sie die folgenden Werte für die Aktion.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei**: **ID** (der OneDrive Datei-ID-Wert)
    1. **Skript**: Ihr Skriptname

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Der fertige Excel Online(Business)-Connector in Power Automate.":::
1. Speichern Sie den Flow, und testen Sie ihn. Verwenden Sie die Schaltfläche **"Test** " auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Schulungsvideo: Ausführen eines Skripts für alle Excel Dateien in einem Ordner

[Sehen Sie sich an, wie Sie dieses Beispiel auf YouTube durchlaufen](https://youtu.be/xMg711o7k6w).
