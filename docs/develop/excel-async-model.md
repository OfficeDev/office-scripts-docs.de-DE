---
title: Unterstützung älterer Office Skripts, die die asynchronen APIs verwenden
description: Ein Primer auf den Office Scripts Async-APIs und wie das Lade-/Synchronisierungsmuster für ältere Skripts verwendet wird.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545075"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="ffe2b-103">Unterstützung älterer Office Skripts, die die asynchronen APIs verwenden</span><span class="sxs-lookup"><span data-stu-id="ffe2b-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="ffe2b-104">In diesem Artikel erfahren Sie, wie Sie Skripts verwalten und aktualisieren, die die asynchronen APIs des älteren Modells verwenden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-104">This article teaches you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="ffe2b-105">Diese APIs verfügen über die gleiche Kernfunktionalität wie die jetzt standardmäßigen, synchronen Office Scripts-APIs, erfordern jedoch, dass Ihr Skript die Datensynchronisierung zwischen dem Skript und der Arbeitsmappe steuert.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ffe2b-106">Das async-Modell kann nur mit Skripts verwendet werden, die vor der Implementierung des aktuellen [API-Modells](scripting-fundamentals.md)erstellt wurden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md).</span></span> <span data-ttu-id="ffe2b-107">Skripts sind bei der Erstellung dauerhaft an das API-Modell gesperrt, das sie haben.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="ffe2b-108">Dies bedeutet auch, dass Sie ein brandneues Skript erstellen müssen, wenn Sie ein altes Skript in das neue Modell konvertieren möchten.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="ffe2b-109">Es wird empfohlen, dass Sie Ihre alten Skripts auf das neue Modell aktualisieren, wenn Sie Änderungen vornehmen, da das aktuelle Modell einfacher zu verwenden ist.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="ffe2b-110">Der Abschnitt [Konvertieren von Asynchronskripten in den aktuellen Modellabschnitt](#convert-async-scripts-to-the-current-model) enthält Hinweise zum Umwandeln dieses Übergangs.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-110">The [Converting async scripts to the current model](#convert-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="older-main-function-signature"></a><span data-ttu-id="ffe2b-111">Ältere `main` Funktionssignatur</span><span class="sxs-lookup"><span data-stu-id="ffe2b-111">Older `main` function signature</span></span>

<span data-ttu-id="ffe2b-112">Skripts, die die async-APIs verwenden, haben eine andere `main` Funktion.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="ffe2b-113">Es ist eine `async` Funktion, die einen `Excel.RequestContext` als ersten Parameter hat.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="ffe2b-114">Context</span><span class="sxs-lookup"><span data-stu-id="ffe2b-114">Context</span></span>

<span data-ttu-id="ffe2b-115">Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="ffe2b-116">Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="ffe2b-117">Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="ffe2b-118">Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="ffe2b-119">Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="ffe2b-120">Das `context`-Objekt verwaltet diese Transaktionen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="ffe2b-121">Synchronisieren und Laden</span><span class="sxs-lookup"><span data-stu-id="ffe2b-121">Sync and load</span></span>

<span data-ttu-id="ffe2b-122">Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="ffe2b-123">In der async-API werden Befehle in die Warteschlange gestellt, bis das Skript den `sync` Vorgang zum Synchronisieren des Skripts und der Arbeitsmappe explizit aufruft.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="ffe2b-124">Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:</span><span class="sxs-lookup"><span data-stu-id="ffe2b-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="ffe2b-125">Daten aus der Arbeitsmappe lesen (nach einem `load`-Vorgang oder einer Methode, die ein[ClientResultat](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) zurückgibt).</span><span class="sxs-lookup"><span data-stu-id="ffe2b-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span></span>
- <span data-ttu-id="ffe2b-126">Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).</span><span class="sxs-lookup"><span data-stu-id="ffe2b-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="ffe2b-127">In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:</span><span class="sxs-lookup"><span data-stu-id="ffe2b-127">The following image shows an example control flow between the script and workbook:</span></span>

:::image type="content" source="../images/load-sync.png" alt-text="Ein Diagramm, in dem Lese- und Schreibvorgänge angezeigt werden, die aus dem Skript an die Arbeitsmappe":::

### <a name="sync"></a><span data-ttu-id="ffe2b-129">Synchronisierung</span><span class="sxs-lookup"><span data-stu-id="ffe2b-129">Sync</span></span>

<span data-ttu-id="ffe2b-130">Wenn Ihr async-Skript Daten aus der Arbeitsmappe lesen oder in die Arbeitsmappe schreiben muss, rufen Sie die Methode auf, `RequestContext.sync` wie im folgenden Codeausschnitt gezeigt:</span><span class="sxs-lookup"><span data-stu-id="ffe2b-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown in the following code snippet:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="ffe2b-131">`context.sync()` wird implizit aufgerufen, wenn ein Skript endet.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="ffe2b-132">Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="ffe2b-133">Ein Schreibvorgang setzt eine beliebige Eigenschaft für ein Excel Objekt (z. B. ) oder eine Methode auf, die `range.format.fill.color = "red"` eine Eigenschaft ändert (z. B. `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="ffe2b-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="ffe2b-134">Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).</span><span class="sxs-lookup"><span data-stu-id="ffe2b-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="ffe2b-135">Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="ffe2b-136">Minimieren Sie die Anzahl der `sync` Aufrufe, damit Ihr Skript schnell ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="ffe2b-137">Andernfalls sind die asynchronen APIs nicht schneller als die standardmäßigen, synchronen APIs.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="ffe2b-138">Laden</span><span class="sxs-lookup"><span data-stu-id="ffe2b-138">Load</span></span>

<span data-ttu-id="ffe2b-139">Ein async-Skript muss Daten aus der Arbeitsmappe laden, bevor es gelesen wird.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="ffe2b-140">Das Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="ffe2b-141">Mit der `load` Methode kann das Skript genau angeben, welche Daten aus der Arbeitsmappe abgerufen werden sollen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="ffe2b-142">Die `load`-Methode ist für jedes Excel-Objekt verfügbar.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="ffe2b-143">Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="ffe2b-144">Wenn Sie dies nicht tun, führt dies zu einem Fehler.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-144">Not doing so results in an error.</span></span>

<span data-ttu-id="ffe2b-145">In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="ffe2b-146">Absicht</span><span class="sxs-lookup"><span data-stu-id="ffe2b-146">Intent</span></span> |<span data-ttu-id="ffe2b-147">Beispielbefehl</span><span class="sxs-lookup"><span data-stu-id="ffe2b-147">Example Command</span></span> | <span data-ttu-id="ffe2b-148">Auswirkung</span><span class="sxs-lookup"><span data-stu-id="ffe2b-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="ffe2b-149">Laden einer Eigenschaft</span><span class="sxs-lookup"><span data-stu-id="ffe2b-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="ffe2b-150">Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="ffe2b-151">Laden mehrerer Eigenschaften</span><span class="sxs-lookup"><span data-stu-id="ffe2b-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="ffe2b-152">Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="ffe2b-153">Alles laden</span><span class="sxs-lookup"><span data-stu-id="ffe2b-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="ffe2b-154">Lädt alle Eigenschaften des Zellbereichs.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-154">Loads all the properties on the range.</span></span> <span data-ttu-id="ffe2b-155">Dies ist keine empfohlene Lösung, da es Ihr Skript verlangsamen wird, indem unnötige Daten erhalten.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="ffe2b-156">Verwenden Sie dies nur, während Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="ffe2b-157">Ihr Skript muss `context.sync()` aufrufen, bevor es geladene Werte ausliest.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

<span data-ttu-id="ffe2b-158">Sie können auch Eigenschaften aus einer ganzen Sammlung laden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="ffe2b-159">Jedes Auflistungsobjekt in der async-API verfügt über eine `items` Eigenschaft, die ein Array ist, das die Objekte in dieser Auflistung enthält.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="ffe2b-160">Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="ffe2b-161">Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a><span data-ttu-id="ffe2b-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="ffe2b-162">ClientResult</span></span>

<span data-ttu-id="ffe2b-163">Methoden in der asynchronen API, die Informationen aus der Arbeitsmappe zurückgeben, haben ein ähnliches Muster wie das `load` / `sync` Paradigma.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="ffe2b-164">`TableCollection.getCount` ruft zum Beispiel die Anzahl von Tabellen in der Auflistung ab.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="ffe2b-165">`getCount` gibt ein `ClientResult<number>` zurück, was bedeutet, dass die `value`-Eigenschaft in der zurückgegebenen [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) eine Zahl ist.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number.</span></span> <span data-ttu-id="ffe2b-166">Ihr Skript kann erst auf diesen Wert zugreifen, wenn `context.sync()` aufgerufen wird.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="ffe2b-167">Ähnlich wie beim Laden einer Eigenschaft ist der `value` bis zu diesem `sync`-Aufruf ein lokaler "leerer" Wert.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="ffe2b-168">Das folgende Skript ruft die Gesamtanzahl der Tabellen in der Arbeitsmappe ab und protokolliert diese Anzahl in der Konsole.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="convert-async-scripts-to-the-current-model"></a><span data-ttu-id="ffe2b-169">Konvertieren von Async-Skripts in das aktuelle Modell</span><span class="sxs-lookup"><span data-stu-id="ffe2b-169">Convert async scripts to the current model</span></span>

<span data-ttu-id="ffe2b-170">Das aktuelle API-Modell verwendet keine `load` , `sync` oder eine `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="ffe2b-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="ffe2b-171">Dies macht die Skripte viel einfacher zu schreiben und zu verwalten.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="ffe2b-172">Ihre beste Ressource zum Konvertieren alter Skripts ist [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span><span class="sxs-lookup"><span data-stu-id="ffe2b-172">Your best resource for converting old scripts is [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="ffe2b-173">Dort können Sie die Community um Hilfe bei bestimmten Szenarien bitten.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="ffe2b-174">In den folgenden Anleitungen sollten Sie die allgemeinen Schritte erläutern, die Sie ausführen müssen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="ffe2b-175">Erstellen Sie ein neues Skript, und kopieren Sie den alten asynchronen Code in dieses.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="ffe2b-176">Achten Sie darauf, die alte Methodensignatur nicht einzuschließen, `main` sondern stattdessen die aktuelle. `function main(workbook: ExcelScript.Workbook)`</span><span class="sxs-lookup"><span data-stu-id="ffe2b-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="ffe2b-177">Entfernen Sie alle `load` und `sync` Anrufe.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="ffe2b-178">Sie sind nicht mehr notwendig.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="ffe2b-179">Alle Eigenschaften wurden entfernt.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-179">All properties have been removed.</span></span> <span data-ttu-id="ffe2b-180">Sie greifen nun über diese Objekte und Methoden auf diese Objekte `get` `set` zu, daher müssen Sie diese Eigenschaftsverweise auf Methodenaufrufe umstellen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="ffe2b-181">Anstatt beispielsweise die Füllfarbe einer Zelle durch den Eigenschaftenzugriff wie folgt `mySheet.getRange("A2:C2").format.fill.color = "blue";` festzulegen: verwenden Sie jetzt Methoden wie die folgende: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="ffe2b-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="ffe2b-182">Auflistungsklassen wurden durch Arrays ersetzt.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="ffe2b-183">Die `add` und Methoden dieser `get` Auflistungsklassen wurden in das Objekt verschoben, das die Auflistung besaß, daher müssen Ihre Verweise entsprechend aktualisiert werden.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="ffe2b-184">Verwenden Sie beispielsweise den folgenden Code, um ein Diagramm mit dem Namen "MyChart" aus dem ersten Arbeitsblatt in der Arbeitsmappe abzubekommen: `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="ffe2b-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="ffe2b-185">Beachten Sie `[0]` die, um auf den ersten Wert des `Worksheet[]` zurückgegebenen von `getWorksheets()` zuzugreifen.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="ffe2b-186">Einige Methoden wurden aus Gründen der Übersichtlichkeit umbenannt und aus Gründen der Benutzerfreundlichkeit hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="ffe2b-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="ffe2b-187">Weitere Informationen finden Sie in der [API-Referenz Office Scripts.](/javascript/api/office-scripts/overview)</span><span class="sxs-lookup"><span data-stu-id="ffe2b-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="ffe2b-188">Office Skripts async API Referenzdokumentation</span><span class="sxs-lookup"><span data-stu-id="ffe2b-188">Office Scripts async API reference documentation</span></span>

<span data-ttu-id="ffe2b-189">Die asynchronen APIs entsprechen denen, die in Office Add-Ins verwendet werden. Die Referenzdokumentation finden Sie im [Excel Abschnitt der Office JavaScript-API-Referenz .](/javascript/api/excel?view=excel-js-online&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="ffe2b-189">The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>
