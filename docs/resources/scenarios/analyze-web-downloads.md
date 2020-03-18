---
title: 'Beispielszenario für Office-Skripts: Analysieren von Webdownloads'
description: Ein Beispiel, das Rohdaten im Internet Datenverkehr in einer Excel-Arbeitsmappe verwendet und den Ursprungsort bestimmt, bevor diese Informationen in einer Tabelle organisiert werden.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700239"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Beispielszenario für Office-Skripts: Analysieren von Webdownloads

In diesem Szenario haben Sie die Aufgabe, Download Berichte von der Website Ihres Unternehmens zu analysieren. Das Ziel dieser Analyse besteht darin zu ermitteln, ob der Webdatenverkehr aus den USA oder anderswo in der Welt stammt.

Ihre Kollegen laden die Rohdaten in Ihre Arbeitsmappe hoch. Die Datenmenge jeder Woche verfügt über ein eigenes Arbeitsblatt. Es gibt auch das **Zusammenfassungs** Arbeitsblatt mit einer Tabelle und einem Diagramm, in denen Wochen überwochen Trends angezeigt werden.

Sie entwickeln ein Skript, mit dem wöchentliche Downloads von Daten im aktiven Arbeitsblatt analysiert werden. Er analysiert die IP-Adresse, die jedem Download zugeordnet ist, und ermittelt, ob er aus den USA stammt oder nicht. Die Antwort wird als boolescher Wert in das Arbeitsblatt eingefügt ("true" oder "false"), und die bedingte Formatierung wird auf diese Zellen angewendet. Die Ergebnisse der IP-Adressposition werden auf dem Arbeitsblatt summiert und in die Zusammenfassungstabelle kopiert.

## <a name="scripting-skills-covered"></a>Abgedeckte Skript Fertigkeiten

- Text Analyse
- Unter Funktionen in Skripts
- Bedingte Formatierung
- Tabellen

## <a name="demo-video"></a>Demo Video

Dieses Beispiel wurde als Teil des Office-Add-ins Entwickler-Community-Aufrufs für den 2020. Februar verdemot.

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a>Setup Anweisungen

1. Laden Sie <a href="analyze-web-downloads.xlsx">Analyze-Web-Downloads. xlsx</a> in Ihre OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe mit Excel für das Internet.

3. Öffnen Sie auf der Registerkarte **automatisieren** den **Code-Editor**.

4. Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. Benennen Sie das Skript zum **Analysieren von Webdownloads** um, und speichern Sie es.

## <a name="running-the-script"></a>Ausführen des Skripts

Navigieren Sie zu einem der **Wochen\* ** Arbeitsblätter, und führen Sie das Skript zum **Analysieren von Webdownloads** aus. Das Skript wendet die bedingte Formatierung und die Speicherort Kennzeichnung auf dem aktuellen Blatt an. Außerdem wird das **Zusammenfassungs** Arbeitsblatt aktualisiert.

### <a name="before-running-the-script"></a>Vor dem Ausführen des Skripts

![Ein Arbeitsblatt, das Rohdaten von Webdatenverkehr anzeigt.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>Nach dem Ausführen des Skripts

![Ein Arbeitsblatt, in dem formatierte IP-Standortinformationen mit den vorherigen Zeilen des Webdatenverkehrs angezeigt werden.](../../images/scenario-analyze-web-downloads-after.png)

![Die Zusammenfassungstabelle und das Diagramm, in dem die Arbeitsblätter zusammengefasst werden, in denen das Skript ausgeführt wurde.](../../images/scenario-analyze-web-downloads-table.png)
