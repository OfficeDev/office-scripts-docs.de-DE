---
title: Ausführen eines Scripts für alle Excel-Dateien in einem Ordner
description: Erfahren Sie, wie Sie ein Skript für alle Excel in einem Ordner auf einem OneDrive for Business.
ms.date: 04/02/2021
localization_priority: Normal
ms.openlocfilehash: 6376dcac0eb36c04c2b60b2717d18cd730a0a8ee
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026859"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Ausführen eines Scripts für alle Excel-Dateien in einem Ordner

Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf einem OneDrive for Business. Es kann auch für einen Ordner SharePoint werden.
Es führt Berechnungen für die Excel aus, fügt Formatierungen hinzu und fügt einen Kommentar ein, der [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) kollegen.

Laden Sie die Datei <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extrahieren Sie die Dateien in einen Ordner mit dem Titel **Sales,** der in diesem Beispiel verwendet wird, und testen Sie sie selbst!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Beispielcode: Hinzufügen von Formatierung und Einfügen eines Kommentars

Dies ist das Skript, das für jede einzelne Arbeitsmappe ausgeführt wird.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate: Ausführen des Skripts für jede Arbeitsmappe im Ordner

In diesem Fluss wird das Skript für jede Arbeitsmappe im Ordner "Sales" ausgeführt.

1. Erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**
1. Fügen Sie einen **neuen Schritt hinzu,** der den **OneDrive for Business** und die Aktion Dateien in Ordner **auflisten** verwendet.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Der abgeschlossene OneDrive for Business connector in Power Automate.":::
1. Wählen Sie den Ordner "Vertrieb" mit den extrahierten Arbeitsmappen aus.
1. Um sicherzustellen, dass nur Arbeitsmappen ausgewählt sind, wählen Sie **Neuer Schritt** aus, wählen Sie **dann Bedingung** aus, und legen Sie die folgenden Werte sicher:
    1. **Name** (der OneDrive Dateinamewert)
    1. "endet mit"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Der Power Automate Bedingungsblock, der nachfolgende Aktionen auf jede Datei angewendet.":::
1. Fügen Sie **unter If yes** branch den Excel Online **(Business)** mit der **Aktion Skript ausführen (Vorschau)** hinzu. Verwenden Sie die folgenden Werte für die Aktion:
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **Datei**: **ID** (der OneDrive Datei-ID-Wert)
    1. **Skript**: Ihr Skriptname

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Der abgeschlossene Excel Online (Business)-Connector in Power Automate.":::
1. Speichern Sie den Fluss, und testen Sie ihn.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Schulungsvideo: Ausführen eines Skripts für alle Excel in einem Ordner

[Sehen Sie sich schritt-für-Schritt-Video](https://youtu.be/xMg711o7k6w) an, wie Sie ein Skript für alle Excel in einem ordner OneDrive for Business oder SharePoint ausführen.
