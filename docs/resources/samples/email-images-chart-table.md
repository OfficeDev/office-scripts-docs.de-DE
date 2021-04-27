---
title: E-Mail der Bilder eines Excel Und Tabelle
description: Erfahren Sie, wie Sie Office Skripts und Power Automate verwenden, um die Bilder eines Diagramms und einer Tabelle Excel zu extrahieren und per E-Mail zu senden.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: 0265250f7fd885cb4899d0b9493b4285496965ff
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026869"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Verwenden Office Skripts und Power Automate, um Bilder eines Diagramms und einer Tabelle per E-Mail zu senden

In diesem Beispiel werden Office Skripts und Power Automate zum Erstellen eines Diagramms verwendet. Anschließend werden Bilder des Diagramms und seiner Basistabelle per E-Mail gesendet.

## <a name="example-scenario"></a>Beispielszenario

* Berechnen Sie, um die neuesten Ergebnisse zu erhalten.
* Diagramm erstellen.
* Get chart and table images.
* Senden Sie die Bilder per E-Mail Power Automate.

_Eingabedaten_

:::image type="content" source="../../images/input-data.png" alt-text="Ein Arbeitsblatt mit einer Tabelle mit Eingabedaten.":::

_Ausgabediagramm_

:::image type="content" source="../../images/chart-created.png" alt-text="Das vom Kunden fällige Spaltendiagramm.":::

_E-Mails, die über den Power Automate empfangen wurden_

:::image type="content" source="../../images/email-received.png" alt-text="Die vom Fluss gesendete E-Mail mit Excel im Text eingebetteten Diagramm.":::

## <a name="solution"></a>Lösung

Diese Lösung besteht aus zwei Teilen:

1. [Ein Office Skript zum Berechnen und Extrahieren Excel Diagramms und der Tabelle](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Ein Power Automate zum Aufrufen des Skripts und zum Senden der Ergebnisse per E-Mail. Ein Beispiel dazu finden Sie unter Erstellen eines automatisierten [Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Beispielcode: Berechnen und Extrahieren Excel Diagramm und Tabelle

Das folgende Skript berechnet und extrahiert ein Excel Und Tabelle.

Laden Sie die Beispieldatei <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> und verwenden Sie sie mit diesem Skript, um sie selbst auszuprobieren!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate: E-Mail-E-Mail des Diagramms und der Tabellenbilder

In diesem Fluss wird das Skript ausgeführt und die zurückgegebenen Bilder per E-Mail gesendet.

1. Erstellen Sie einen neuen **Instant Cloud Flow**.
1. Wählen **Sie Manuellen Fluss auslösen aus,** und drücken Sie **die Create -Taste.**
1. Fügen Sie einen **neuen Schritt** hinzu, der den Excel **Online (Business)** mit der **Aktion Skript ausführen (Vorschau)** verwendet. Verwenden Sie die folgenden Werte für die Aktion:
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: Ihre Arbeitsmappe ([ausgewählt mit der Datei-Auswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Skript**: Ihr Skriptname

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Der abgeschlossene Excel Online (Business)-Connector in Power Automate.":::
1. In diesem Beispiel wird Outlook als E-Mail-Client verwendet. Sie können beliebige E-Mail-Power Automate verwenden, aber bei den restlichen Schritten wird davon ausgegangen, dass Sie Outlook. Fügen Sie einen **Neuen Schritt hinzu,** der den **Office 365 Outlook** und die Aktion Senden **und E-Mail (V2)** verwendet. Verwenden Sie die folgenden Werte für die Aktion:
    * **An**: Ihr Test-E-Mail-Konto (oder persönliche E-Mail)
    * **Betreff**: Please Review Report Data
    * Wählen Sie **für das Feld Text** die Option "Codeansicht" ( ) `</>` aus, und geben Sie Folgendes ein:

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Der abgeschlossene Office 365 Outlook in Power Automate.":::
1. Speichern Sie den Fluss, und testen Sie ihn.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Schulungsvideo: Extrahieren und E-Mail-Bilder von Diagramm und Tabelle

[![Sehen Sie sich schritt-für-Schritt-Video zum Extrahieren und E-Mail-Bilder von Diagramm und Tabelle an](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Schritt-für-Schritt-Video zum Extrahieren und E-Mail-Bilder von Diagramm und Tabelle")
