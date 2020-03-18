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
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="4447b-103">Beispielszenario für Office-Skripts: Analysieren von Webdownloads</span><span class="sxs-lookup"><span data-stu-id="4447b-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="4447b-104">In diesem Szenario haben Sie die Aufgabe, Download Berichte von der Website Ihres Unternehmens zu analysieren.</span><span class="sxs-lookup"><span data-stu-id="4447b-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="4447b-105">Das Ziel dieser Analyse besteht darin zu ermitteln, ob der Webdatenverkehr aus den USA oder anderswo in der Welt stammt.</span><span class="sxs-lookup"><span data-stu-id="4447b-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="4447b-106">Ihre Kollegen laden die Rohdaten in Ihre Arbeitsmappe hoch.</span><span class="sxs-lookup"><span data-stu-id="4447b-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="4447b-107">Die Datenmenge jeder Woche verfügt über ein eigenes Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="4447b-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="4447b-108">Es gibt auch das **Zusammenfassungs** Arbeitsblatt mit einer Tabelle und einem Diagramm, in denen Wochen überwochen Trends angezeigt werden.</span><span class="sxs-lookup"><span data-stu-id="4447b-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="4447b-109">Sie entwickeln ein Skript, mit dem wöchentliche Downloads von Daten im aktiven Arbeitsblatt analysiert werden.</span><span class="sxs-lookup"><span data-stu-id="4447b-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="4447b-110">Er analysiert die IP-Adresse, die jedem Download zugeordnet ist, und ermittelt, ob er aus den USA stammt oder nicht.</span><span class="sxs-lookup"><span data-stu-id="4447b-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="4447b-111">Die Antwort wird als boolescher Wert in das Arbeitsblatt eingefügt ("true" oder "false"), und die bedingte Formatierung wird auf diese Zellen angewendet.</span><span class="sxs-lookup"><span data-stu-id="4447b-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="4447b-112">Die Ergebnisse der IP-Adressposition werden auf dem Arbeitsblatt summiert und in die Zusammenfassungstabelle kopiert.</span><span class="sxs-lookup"><span data-stu-id="4447b-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="4447b-113">Abgedeckte Skript Fertigkeiten</span><span class="sxs-lookup"><span data-stu-id="4447b-113">Scripting skills covered</span></span>

- <span data-ttu-id="4447b-114">Text Analyse</span><span class="sxs-lookup"><span data-stu-id="4447b-114">Text parsing</span></span>
- <span data-ttu-id="4447b-115">Unter Funktionen in Skripts</span><span class="sxs-lookup"><span data-stu-id="4447b-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="4447b-116">Bedingte Formatierung</span><span class="sxs-lookup"><span data-stu-id="4447b-116">Conditional formatting</span></span>
- <span data-ttu-id="4447b-117">Tabellen</span><span class="sxs-lookup"><span data-stu-id="4447b-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="4447b-118">Demo Video</span><span class="sxs-lookup"><span data-stu-id="4447b-118">Demo video</span></span>

<span data-ttu-id="4447b-119">Dieses Beispiel wurde als Teil des Office-Add-ins Entwickler-Community-Aufrufs für den 2020. Februar verdemot.</span><span class="sxs-lookup"><span data-stu-id="4447b-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a><span data-ttu-id="4447b-120">Setup Anweisungen</span><span class="sxs-lookup"><span data-stu-id="4447b-120">Setup instructions</span></span>

1. <span data-ttu-id="4447b-121">Laden Sie <a href="analyze-web-downloads.xlsx">Analyze-Web-Downloads. xlsx</a> in Ihre OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="4447b-121">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="4447b-122">Öffnen Sie die Arbeitsmappe mit Excel für das Internet.</span><span class="sxs-lookup"><span data-stu-id="4447b-122">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="4447b-123">Öffnen Sie auf der Registerkarte **automatisieren** den **Code-Editor**.</span><span class="sxs-lookup"><span data-stu-id="4447b-123">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="4447b-124">Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="4447b-124">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="4447b-125">Benennen Sie das Skript zum **Analysieren von Webdownloads** um, und speichern Sie es.</span><span class="sxs-lookup"><span data-stu-id="4447b-125">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="4447b-126">Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="4447b-126">Running the script</span></span>

<span data-ttu-id="4447b-127">Navigieren Sie zu einem der \*\*Wochen\* \*\* Arbeitsblätter, und führen Sie das Skript zum **Analysieren von Webdownloads** aus.</span><span class="sxs-lookup"><span data-stu-id="4447b-127">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="4447b-128">Das Skript wendet die bedingte Formatierung und die Speicherort Kennzeichnung auf dem aktuellen Blatt an.</span><span class="sxs-lookup"><span data-stu-id="4447b-128">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="4447b-129">Außerdem wird das **Zusammenfassungs** Arbeitsblatt aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="4447b-129">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="4447b-130">Vor dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="4447b-130">Before running the script</span></span>

![Ein Arbeitsblatt, das Rohdaten von Webdatenverkehr anzeigt.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="4447b-132">Nach dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="4447b-132">After running the script</span></span>

![Ein Arbeitsblatt, in dem formatierte IP-Standortinformationen mit den vorherigen Zeilen des Webdatenverkehrs angezeigt werden.](../../images/scenario-analyze-web-downloads-after.png)

![Die Zusammenfassungstabelle und das Diagramm, in dem die Arbeitsblätter zusammengefasst werden, in denen das Skript ausgeführt wurde.](../../images/scenario-analyze-web-downloads-table.png)
