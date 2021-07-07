---
title: 'Office Skript-Beispielszenario: Analysieren von Webdownloads'
description: Ein Beispiel, das unformatierte Internetdatenverkehrsdaten in einer Excel Arbeitsmappe annimmt und den Ursprungsspeicherort ermittelt, bevor diese Informationen in einer Tabelle organisiert werden.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: ec7ccd2a4f534be825ee4aa358f5d81becc53937
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313932"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="59ef1-103">Office Skript-Beispielszenario: Analysieren von Webdownloads</span><span class="sxs-lookup"><span data-stu-id="59ef1-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="59ef1-104">In diesem Szenario haben Sie die Aufgabe, Downloadberichte von der Website Ihres Unternehmens zu analysieren.</span><span class="sxs-lookup"><span data-stu-id="59ef1-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="59ef1-105">Das Ziel dieser Analyse besteht darin, festzustellen, ob der Webdatenverkehr aus den USA oder aus anderen Ländern der Welt stammt.</span><span class="sxs-lookup"><span data-stu-id="59ef1-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="59ef1-106">Ihre Kollegen laden die Rohdaten in Ihre Arbeitsmappe hoch.</span><span class="sxs-lookup"><span data-stu-id="59ef1-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="59ef1-107">Die Datengruppe jeder Woche verfügt über ein eigenes Arbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="59ef1-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="59ef1-108">Es gibt auch das Arbeitsblatt **Zusammenfassung** mit einer Tabelle und einem Diagramm, in dem Die Wochen-über-Woche-Trends angezeigt werden.</span><span class="sxs-lookup"><span data-stu-id="59ef1-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="59ef1-109">Sie entwickeln ein Skript, das wöchentliche Downloads von Daten im aktiven Arbeitsblatt analysiert.</span><span class="sxs-lookup"><span data-stu-id="59ef1-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="59ef1-110">Sie analysiert die ip-Adresse, die jedem Download zugeordnet ist, und bestimmt, ob sie aus den USA stammt.</span><span class="sxs-lookup"><span data-stu-id="59ef1-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="59ef1-111">Die Antwort wird als boolescher Wert ("TRUE" oder "FALSE") in das Arbeitsblatt eingefügt, und die bedingte Formatierung wird auf diese Zellen angewendet.</span><span class="sxs-lookup"><span data-stu-id="59ef1-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="59ef1-112">Die Ergebnisse des IP-Adressspeicherorts werden auf dem Arbeitsblatt summiert und in die Zusammenfassungstabelle kopiert.</span><span class="sxs-lookup"><span data-stu-id="59ef1-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="59ef1-113">Abgedeckte Skriptfähigkeiten</span><span class="sxs-lookup"><span data-stu-id="59ef1-113">Scripting skills covered</span></span>

- <span data-ttu-id="59ef1-114">Textparsing</span><span class="sxs-lookup"><span data-stu-id="59ef1-114">Text parsing</span></span>
- <span data-ttu-id="59ef1-115">Unterfunktionen in Skripts</span><span class="sxs-lookup"><span data-stu-id="59ef1-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="59ef1-116">Bedingte Formatierung</span><span class="sxs-lookup"><span data-stu-id="59ef1-116">Conditional formatting</span></span>
- <span data-ttu-id="59ef1-117">Tabellen</span><span class="sxs-lookup"><span data-stu-id="59ef1-117">Tables</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="59ef1-118">Setup-Anweisungen</span><span class="sxs-lookup"><span data-stu-id="59ef1-118">Setup instructions</span></span>

1. <span data-ttu-id="59ef1-119">Laden Sie <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> auf Ihre OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="59ef1-119">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="59ef1-120">Öffnen Sie die Arbeitsmappe mit Excel für das Web.</span><span class="sxs-lookup"><span data-stu-id="59ef1-120">Open the workbook with Excel for the web.</span></span>

1. <span data-ttu-id="59ef1-121">Wählen Sie auf der Registerkarte **"Automatisieren"** die Option **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="59ef1-121">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
        * End the script early if either object is not in the workbook.
        */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }
  
      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();
  
      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
            "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1);
  
      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);
  
      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];
  
      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }
  
      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];
  
      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);
  
      // Call the local function to add summary data to the worksheet.
      addSummaryData();
  
      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);
  
      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();
  
      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();
  
      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;
  
      /**
       * A function to add summary data on the worksheet.
        */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
          formula1: "=TRUE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
          formula1: "=FALSE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
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
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

1. <span data-ttu-id="59ef1-122">Benennen Sie das Skript um, um **Webdownloads** zu analysieren und zu speichern.</span><span class="sxs-lookup"><span data-stu-id="59ef1-122">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="59ef1-123">Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="59ef1-123">Running the script</span></span>

<span data-ttu-id="59ef1-124">Navigieren Sie zu einem der **Arbeitsblätter \* \* "Woche",** und führen Sie das Skript **"Webdownloads analysieren" aus.**</span><span class="sxs-lookup"><span data-stu-id="59ef1-124">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="59ef1-125">Das Skript wendet die bedingte Formatierung und standortbedingte Kennzeichnung auf das aktuelle Blatt an.</span><span class="sxs-lookup"><span data-stu-id="59ef1-125">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="59ef1-126">Außerdem wird das **Arbeitsblatt Zusammenfassung** aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="59ef1-126">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="59ef1-127">Vor dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="59ef1-127">Before running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Ein Arbeitsblatt mit unformatierten Webdatenverkehrsdaten.":::

### <a name="after-running-the-script"></a><span data-ttu-id="59ef1-129">Nach dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="59ef1-129">After running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Ein Arbeitsblatt mit formatierten IP-Standortinformationen mit den vorherigen Webdatenverkehrzeilen.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="Die Zusammenfassungstabelle und das Diagramm, in denen die Arbeitsblätter zusammengefasst sind, auf denen das Skript ausgeführt wurde.":::
