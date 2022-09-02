---
title: Email der Bilder eines Excel-Diagramms und einer Excel-Tabelle
description: Erfahren Sie, wie Sie office-Skripts und Power Automate verwenden, um die Bilder eines Excel-Diagramms und einer Excel-Tabelle zu extrahieren und per E-Mail zu senden.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572465"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Verwenden von Office-Skripts und Power Automate zum Senden von E-Mail-Bildern eines Diagramms und einer Tabelle

In diesem Beispiel werden Office-Skripts und Power Automate zum Erstellen eines Diagramms verwendet. Anschließend werden Bilder des Diagramms und seiner Basistabelle per E-Mail gesendet.

## <a name="example-scenario"></a>Beispielszenario

* Berechnen, um die neuesten Ergebnisse zu erhalten.
* Diagramm erstellen.
* Dient zum Abrufen von Diagramm- und Tabellenbildern.
* Email die Bilder mit Power Automate.

_Eingabedaten_

:::image type="content" source="../../images/input-data.png" alt-text="Ein Arbeitsblatt mit einer Tabelle mit Eingabedaten.":::

_Ausgabediagramm_

:::image type="content" source="../../images/chart-created.png" alt-text="Das erstellte Säulendiagramm mit dem vom Kunden fälligen Betrag.":::

_Email, die über den Power Automate-Fluss empfangen wurden_

:::image type="content" source="../../images/email-received.png" alt-text="Die vom Fluss gesendete E-Mail mit dem im Textkörper eingebetteten Excel-Diagramm.":::

## <a name="solution"></a>Lösung

Diese Lösung umfasst zwei Teile:

1. [Ein Office-Skript zum Berechnen und Extrahieren eines Excel-Diagramms und einer Excel-Tabelle](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Ein Power Automate-Fluss, um das Skript aufzurufen und die Ergebnisse per E-Mail zu senden. Ein Beispiel dazu finden [Sie unter Erstellen eines automatisierten Workflows mit Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [email-chart-table.xlsx](email-chart-table.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Beispielcode: Berechnen und Extrahieren eines Excel-Diagramms und einer Excel-Tabelle

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate-Fluss: Email der Diagramm- und Tabellenbilder

In diesem Fluss wird das Skript ausgeführt und die zurückgegebenen Bilder per E-Mail gesendet.

1. Erstellen Sie einen neuen **Instant Cloud-Fluss**.
1. Wählen Sie **"Manuell auslösen" und** dann " **Erstellen**" aus.
1. Fügen Sie einen **neuen Schritt** hinzu, der den **Excel Online (Business)** -Connector mit der **Skriptaktion ausführen** verwendet. Verwenden Sie die folgenden Werte für die Aktion.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **Datei**: Ihre Arbeitsmappe ([ausgewählt mit der Dateiauswahl](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Skript**: Ihr Skriptname

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Der fertige Excel Online (Business)-Connector in Power Automate.":::
1. In diesem Beispiel wird Outlook als E-Mail-Client verwendet. Sie können jeden beliebigen E-Mail-Connector verwenden, der von Power Automate unterstützt wird, aber in den restlichen Schritten wird davon ausgegangen, dass Sie Outlook ausgewählt haben. Fügen Sie einen **neuen Schritt** hinzu, der den **Office 365 Outlook-Connector** und die Aktion **"Senden und E-Mail( V2)"** verwendet. Verwenden Sie die folgenden Werte für die Aktion.
    * **An**: Ihr Test-E-Mail-Konto (oder persönliche E-Mail)
    * **Betreff**: Bitte überprüfen Sie die Berichtsdaten
    * Wählen Sie für das Feld " **Textkörper** " "Codeansicht" (`</>`) aus, und geben Sie Folgendes ein:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Der fertige Office 365 Outlook-Connector in Power Automate.":::
1. Speichern Sie den Fluss, und probieren Sie ihn aus. Verwenden Sie die Schaltfläche " **Testen** " auf der Fluss-Editor-Seite, oder führen Sie den Fluss über die Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Schulungsvideo: Extrahieren und E-Mail-Bilder von Diagramm und Tabelle

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/152GJyqc-Kw).
