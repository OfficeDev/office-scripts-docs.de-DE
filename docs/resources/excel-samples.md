---
title: Beispielskripts für Office-Skripts in Excel im Internet
description: Eine Sammlung von Codebeispielen, die mit Office-Skripts in Excel im Internet verwendet werden sollen.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700243"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="25400-103">Beispielskripts für Office-Skripts in Excel im Internet (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="25400-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="25400-104">Die folgenden Beispiele sind einfache Skripts, mit denen Sie Ihre eigenen Arbeitsmappen ausprobieren können.</span><span class="sxs-lookup"><span data-stu-id="25400-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="25400-105">So verwenden Sie Sie in Excel im Internet:</span><span class="sxs-lookup"><span data-stu-id="25400-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="25400-106">Öffnen Sie die Registerkarte **automatisieren** .</span><span class="sxs-lookup"><span data-stu-id="25400-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="25400-107">Drücken Sie **Code-Editor**.</span><span class="sxs-lookup"><span data-stu-id="25400-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="25400-108">Klicken Sie im Aufgabenbereich des Code-Editors auf **Neues Skript** .</span><span class="sxs-lookup"><span data-stu-id="25400-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="25400-109">Ersetzen Sie das gesamte Skript durch das Beispiel Ihrer Wahl.</span><span class="sxs-lookup"><span data-stu-id="25400-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="25400-110">Klicken Sie im Aufgabenbereich des Code-Editors auf **Ausführen** .</span><span class="sxs-lookup"><span data-stu-id="25400-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="25400-111">Grundlagen der Skripterstellung</span><span class="sxs-lookup"><span data-stu-id="25400-111">Scripting basics</span></span>

<span data-ttu-id="25400-112">In diesen Beispielen werden grundlegende Bausteine für Office-Skripts veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="25400-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="25400-113">Fügen Sie diese zu Ihren Skripts hinzu, um Ihre Lösung zu erweitern und häufige Probleme zu lösen.</span><span class="sxs-lookup"><span data-stu-id="25400-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="25400-114">Lesen und Protokollieren einer Zelle</span><span class="sxs-lookup"><span data-stu-id="25400-114">Read and log one cell</span></span>

<span data-ttu-id="25400-115">In diesem Beispiel wird der Wert von **a1** gelesen und in der Konsole gedruckt.</span><span class="sxs-lookup"><span data-stu-id="25400-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="25400-116">Arbeiten mit Datumsangaben</span><span class="sxs-lookup"><span data-stu-id="25400-116">Work with dates</span></span>

<span data-ttu-id="25400-117">In diesem Beispiel wird das [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) -Objekt von JavaScript verwendet, um das aktuelle Datum und die Uhrzeit abzurufen, und dann werden diese Werte in zwei Zellen im aktiven Arbeitsblatt geschrieben.</span><span class="sxs-lookup"><span data-stu-id="25400-117">This sample uses the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object to get the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="25400-118">Anzeigen von Daten</span><span class="sxs-lookup"><span data-stu-id="25400-118">Display data</span></span>

<span data-ttu-id="25400-119">In diesen Beispielen wird gezeigt, wie Sie mit Arbeitsblattdaten arbeiten und Benutzern eine bessere Ansicht oder Organisation bieten.</span><span class="sxs-lookup"><span data-stu-id="25400-119">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="25400-120">Anwenden bedingter Formatierung</span><span class="sxs-lookup"><span data-stu-id="25400-120">Apply conditional formatting</span></span>

<span data-ttu-id="25400-121">In diesem Beispiel wird die bedingte Formatierung auf den aktuell verwendeten Bereich im Arbeitsblatt angewendet.</span><span class="sxs-lookup"><span data-stu-id="25400-121">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="25400-122">Die bedingte Formatierung ist eine grüne Füllung für die oberen 10% der Werte.</span><span class="sxs-lookup"><span data-stu-id="25400-122">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="25400-123">Erstellen einer sortierten Tabelle</span><span class="sxs-lookup"><span data-stu-id="25400-123">Create a sorted table</span></span>

<span data-ttu-id="25400-124">In diesem Beispiel wird eine Tabelle aus dem verwendeten Bereich des aktuellen Arbeitsblatts erstellt und dann basierend auf der ersten Spalte sortiert.</span><span class="sxs-lookup"><span data-stu-id="25400-124">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

## <a name="collaboration"></a><span data-ttu-id="25400-125">Zusammenarbeit</span><span class="sxs-lookup"><span data-stu-id="25400-125">Collaboration</span></span>

<span data-ttu-id="25400-126">In diesen Beispielen wird gezeigt, wie Sie mit Zusammenarbeits bezogenen Features von Excel wie Kommentaren arbeiten.</span><span class="sxs-lookup"><span data-stu-id="25400-126">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="25400-127">Aufgelöste Kommentare löschen</span><span class="sxs-lookup"><span data-stu-id="25400-127">Delete resolved comments</span></span>

<span data-ttu-id="25400-128">In diesem Beispiel werden alle aufgelösten Kommentare aus dem aktuellen Arbeitsblatt gelöscht.</span><span class="sxs-lookup"><span data-stu-id="25400-128">This sample deletes all resolved comments from the current worksheet.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="25400-129">Szenario-Beispiele</span><span class="sxs-lookup"><span data-stu-id="25400-129">Scenario samples</span></span>

<span data-ttu-id="25400-130">Beispiele für größere Lösungen aus der realen Welt finden Sie unter [Beispielszenarien für Office-Skripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="25400-130">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="25400-131">Neue Beispiele vorschlagen</span><span class="sxs-lookup"><span data-stu-id="25400-131">Suggest new samples</span></span>

<span data-ttu-id="25400-132">Wir begrüßen Vorschläge für neue Beispiele.</span><span class="sxs-lookup"><span data-stu-id="25400-132">We welcome suggestions for new samples.</span></span> <span data-ttu-id="25400-133">Wenn es ein gängiges Szenario gibt, das anderen Skript Entwicklern helfen würde, teilen Sie uns dies im Abschnitt Feedback unten.</span><span class="sxs-lookup"><span data-stu-id="25400-133">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
