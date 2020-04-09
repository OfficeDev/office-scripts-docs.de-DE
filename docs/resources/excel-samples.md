---
title: Beispielskripts für Office-Skripts in Excel im Internet
description: Eine Sammlung von Codebeispielen, die mit Office-Skripts in Excel im Internet verwendet werden sollen.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191006"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="5b05f-103">Beispielskripts für Office-Skripts in Excel im Internet (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="5b05f-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="5b05f-104">Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können.</span><span class="sxs-lookup"><span data-stu-id="5b05f-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="5b05f-105">So verwenden Sie Sie in Excel im Internet:</span><span class="sxs-lookup"><span data-stu-id="5b05f-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="5b05f-106">Öffnen Sie die Registerkarte **Automatisieren**.</span><span class="sxs-lookup"><span data-stu-id="5b05f-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="5b05f-107">Drücken Sie **Code-Editor**.</span><span class="sxs-lookup"><span data-stu-id="5b05f-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="5b05f-108">Klicken Sie im Aufgabenbereich des Code-Editors auf **Neues Skript** .</span><span class="sxs-lookup"><span data-stu-id="5b05f-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="5b05f-109">Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.</span><span class="sxs-lookup"><span data-stu-id="5b05f-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="5b05f-110">Klicken Sie im Aufgabenbereich des Code-Editors auf **Ausführen** .</span><span class="sxs-lookup"><span data-stu-id="5b05f-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="5b05f-111">Grundlagen der Skripterstellung</span><span class="sxs-lookup"><span data-stu-id="5b05f-111">Scripting basics</span></span>

<span data-ttu-id="5b05f-112">In diesen Beispielen werden grundlegende Bausteine für Office-Skripts veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="5b05f-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="5b05f-113">Fügen Sie diese zu Ihren Skripts hinzu, um Ihre Lösung zu erweitern und häufige Probleme zu lösen.</span><span class="sxs-lookup"><span data-stu-id="5b05f-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="5b05f-114">Lesen und Protokollieren einer Zelle</span><span class="sxs-lookup"><span data-stu-id="5b05f-114">Read and log one cell</span></span>

<span data-ttu-id="5b05f-115">In diesem Beispiel wird der Wert von **a1** gelesen und in der Konsole gedruckt.</span><span class="sxs-lookup"><span data-stu-id="5b05f-115">This sample reads the value of **A1** and prints it to the console.</span></span>

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a><span data-ttu-id="5b05f-116">Arbeiten mit Datumsangaben</span><span class="sxs-lookup"><span data-stu-id="5b05f-116">Work with dates</span></span>

<span data-ttu-id="5b05f-117">In den Beispielen in diesem Abschnitt wird gezeigt, wie das JavaScript- [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) -Objekt verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="5b05f-117">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="5b05f-118">Im folgenden Beispiel werden das aktuelle Datum und die Uhrzeit abgerufen, und anschließend werden diese Werte in zwei Zellen im aktiven Arbeitsblatt geschrieben.</span><span class="sxs-lookup"><span data-stu-id="5b05f-118">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

<span data-ttu-id="5b05f-119">Im nächsten Beispiel wird ein Datum gelesen, das in Excel gespeichert und in ein JavaScript-Date-Objekt übersetzt wird.</span><span class="sxs-lookup"><span data-stu-id="5b05f-119">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="5b05f-120">Es verwendet die [numerische Seriennummer des Datums](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) als Eingabe für das JavaScript-Datum.</span><span class="sxs-lookup"><span data-stu-id="5b05f-120">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="5b05f-121">Anzeigen von Daten</span><span class="sxs-lookup"><span data-stu-id="5b05f-121">Display data</span></span>

<span data-ttu-id="5b05f-122">In diesen Beispielen wird gezeigt, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.</span><span class="sxs-lookup"><span data-stu-id="5b05f-122">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="5b05f-123">Anwenden bedingter Formatierung</span><span class="sxs-lookup"><span data-stu-id="5b05f-123">Apply conditional formatting</span></span>

<span data-ttu-id="5b05f-124">In diesem Beispiel wird die bedingte Formatierung auf den aktuell verwendeten Bereich im Arbeitsblatt angewendet.</span><span class="sxs-lookup"><span data-stu-id="5b05f-124">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="5b05f-125">Die bedingte Formatierung ist eine grüne Füllung für die oberen 10% der Werte.</span><span class="sxs-lookup"><span data-stu-id="5b05f-125">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="5b05f-126">Erstellen einer sortierten Tabelle</span><span class="sxs-lookup"><span data-stu-id="5b05f-126">Create a sorted table</span></span>

<span data-ttu-id="5b05f-127">In diesem Beispiel wird eine Tabelle aus dem verwendeten Bereich des aktuellen Arbeitsblatts erstellt und dann basierend auf der ersten Spalte sortiert.</span><span class="sxs-lookup"><span data-stu-id="5b05f-127">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a><span data-ttu-id="5b05f-128">Zusammenarbeit</span><span class="sxs-lookup"><span data-stu-id="5b05f-128">Collaboration</span></span>

<span data-ttu-id="5b05f-129">In diesen Beispielen wird gezeigt, wie Sie mit Zusammenarbeits bezogenen Features von Excel wie Kommentaren arbeiten.</span><span class="sxs-lookup"><span data-stu-id="5b05f-129">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="5b05f-130">Aufgelöste Kommentare löschen</span><span class="sxs-lookup"><span data-stu-id="5b05f-130">Delete resolved comments</span></span>

<span data-ttu-id="5b05f-131">In diesem Beispiel werden alle aufgelösten Kommentare aus dem aktuellen Arbeitsblatt gelöscht.</span><span class="sxs-lookup"><span data-stu-id="5b05f-131">This sample deletes all resolved comments from the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="5b05f-132">Szenario-Beispiele</span><span class="sxs-lookup"><span data-stu-id="5b05f-132">Scenario samples</span></span>

<span data-ttu-id="5b05f-133">Beispiele für größere Lösungen aus der realen Welt finden Sie unter [Beispielszenarien für Office-Skripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="5b05f-133">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="5b05f-134">Neue Beispiele vorschlagen</span><span class="sxs-lookup"><span data-stu-id="5b05f-134">Suggest new samples</span></span>

<span data-ttu-id="5b05f-135">Wir begrüßen Vorschläge für neue Beispiele.</span><span class="sxs-lookup"><span data-stu-id="5b05f-135">We welcome suggestions for new samples.</span></span> <span data-ttu-id="5b05f-136">Wenn es ein gängiges Szenario gibt, das anderen Skript Entwicklern helfen würde, teilen Sie uns dies im Abschnitt Feedback unten.</span><span class="sxs-lookup"><span data-stu-id="5b05f-136">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
