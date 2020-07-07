---
title: Ausführen von Office-Skripts mit Power Automation
description: Vorgehensweise Abrufen von Office-Skripts für Excel im Internet arbeiten mit einem Power automatisieren Workflow.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 0ea58324998d23020e04cb37dfeea065791757f5
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: de-DE
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043384"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="a6fbe-103">Ausführen von Office-Skripts mit Power Automation</span><span class="sxs-lookup"><span data-stu-id="a6fbe-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="a6fbe-104">[Power Automation](https://flow.microsoft.com) ermöglicht Ihnen das Hinzufügen von Office-Skripts zu einem größeren, automatisierten Workflow.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="a6fbe-105">Sie können Power automatisieren verwenden Sie Dinge wie das Hinzufügen des Inhalts einer e-Mail zur Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in ihren Projektverwaltungstools basierend auf Arbeitsmappen-Kommentaren.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="a6fbe-106">Wenn Sie neu bei Power Automation sind, empfehlen wir Ihnen, die [Erste Schritte mit Power automatisieren](/power-automate/getting-started)zu besuchen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="a6fbe-107">Hier erfahren Sie mehr über die Automatisierung Ihrer Workflows in mehreren Diensten.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a6fbe-108">Derzeit können Sie Office-Skripts nicht über einen [freigegebenen Fluss](/power-automate/share-buttons)ausführen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="a6fbe-109">Nur der Benutzer, der ein Skript erstellt hat, kann es auch mithilfe von Power Automation ausführen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="a6fbe-110">Erste Schritte</span><span class="sxs-lookup"><span data-stu-id="a6fbe-110">Getting started</span></span>

<span data-ttu-id="a6fbe-111">Um mit der Kombination von Power Automation und Office-Skripts zu beginnen, führen Sie das Lernprogramm [Start using scripts with Power Automation aus](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="a6fbe-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="a6fbe-112">In diesem Artikel erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="a6fbe-113">Wenn Sie das Lernprogramm und die [automatisch ausgeführten Skripts mit Power Automation](../tutorials/excel-power-automate-trigger.md) Tutorial abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office-Skripts mit Power Automation Flows ein.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="a6fbe-114">Excel Online-Connector (Business)</span><span class="sxs-lookup"><span data-stu-id="a6fbe-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="a6fbe-115">[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automation und Anwendungen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="a6fbe-116">Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gibt Ihrem Fluss Zugriff auf Excel-Arbeitsmappen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="a6fbe-117">Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, das über die ausgewählte Arbeitsmappe zugänglich ist.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="a6fbe-118">Sie können nicht nur Skripts über einen Fluss ausführen, sondern auch Daten an und von der Arbeitsmappe mit dem Fluss durch die Skripts übergeben.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a6fbe-119">Die Aktion "Skript ausführen" gibt Benutzern, die den Excel Connector verwenden, wichtigen Zugriff auf Ihre Arbeitsmappe und deren Daten.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="a6fbe-120">Darüber hinaus gibt es Sicherheitsrisiken mit Skripts, die externe API-Aufrufe durchführen, wie in [externe Aufrufe von Power Automation](external-calls.md)erläutert.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="a6fbe-121">Wenn Ihr Administrator mit der Exposition hoch vertraulicher Daten befasst ist, können Sie entweder den Excel Online Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administrator Steuerelemente](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)einschränken.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="a6fbe-122">Übergeben von Daten aus Power Automation in ein Skript</span><span class="sxs-lookup"><span data-stu-id="a6fbe-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="a6fbe-123">Alle Skript Eingaben werden als zusätzliche Parameter für die `main` Funktion angegeben.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="a6fbe-124">Wenn Sie beispielsweise möchten, dass ein Skript einen akzeptiert, `string` das einen Namen als Eingabe darstellt, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="a6fbe-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="a6fbe-125">Wenn Sie einen Fluss in Power Automation konfigurieren, können Sie Skript Eingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamische Inhalte angeben.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="a6fbe-126">Details zu den Connectors eines einzelnen Diensts finden Sie in der [Power Automation Connector-Dokumentation](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="a6fbe-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="a6fbe-127">Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Zulagen und Einschränkungen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="a6fbe-128">Der erste Parameter muss vom Typ sein `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="a6fbe-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="a6fbe-129">Der Name des Parameters spielt keine Rolle.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="a6fbe-130">Jeder Parameter muss einen Typ aufweisen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="a6fbe-131">Die Grundtypen `string` , `number` ,,, `boolean` `any` `unknown` , `object` und `undefined` werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="a6fbe-132">Arrays der zuvor aufgelisteten Grundtypen werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="a6fbe-133">Geschachtelte Arrays werden als Parameter unterstützt (jedoch nicht als Rückgabetypen).</span><span class="sxs-lookup"><span data-stu-id="a6fbe-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="a6fbe-134">Union-Typen sind zulässig, wenn es sich um eine Vereinigung von literalen handelt, die zu einem einzelnen Typ ( `string` , `number` oder `boolean` ) gehören.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="a6fbe-135">Gewerkschaften eines unterstützten Typs mit undefined werden ebenfalls unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="a6fbe-136">Objekttypen sind zulässig, wenn Sie Eigenschaften vom Typ `string` , `number` ,, `boolean` unterstützte Arrays oder andere unterstützte Objekte enthalten.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="a6fbe-137">Im folgenden Beispiel werden geschachtelte Objekte gezeigt, die als Parametertypen unterstützt werden:</span><span class="sxs-lookup"><span data-stu-id="a6fbe-137">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="a6fbe-138">Objekte müssen Ihre Schnittstelle oder Klassendefinition im Skript definiert haben.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="a6fbe-139">Ein Objekt kann auch anonym Inline definiert werden, wie im folgenden Beispiel dargestellt:</span><span class="sxs-lookup"><span data-stu-id="a6fbe-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="a6fbe-140">Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer als solche gekennzeichnet werden `?` (beispielsweise `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="a6fbe-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="a6fbe-141">Standardparameterwerte sind zulässig (beispielsweise `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="a6fbe-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="a6fbe-142">Zurückgeben von Daten aus einem Skript an Power Automation</span><span class="sxs-lookup"><span data-stu-id="a6fbe-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="a6fbe-143">Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power-Automatisierungs Fluss verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="a6fbe-144">Wie bei Eingabeparametern stellt Power Automation einige Einschränkungen für den Rückgabetyp dar.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="a6fbe-145">Die Grundtypen `string` , `number` , `boolean` , `void` und `undefined` werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="a6fbe-146">Union-Typen, die als Rückgabetypen verwendet werden, verfolgen dieselben Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="a6fbe-147">Array Typen sind zulässig, wenn Sie vom Typ `string` , `number` oder sind `boolean` .</span><span class="sxs-lookup"><span data-stu-id="a6fbe-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="a6fbe-148">Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="a6fbe-149">Als Rückgabetypen verwendete Objekttypen entsprechen denselben Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="a6fbe-150">Die implizite Typisierung wird unterstützt, muss aber denselben Regeln wie ein definierter Typ entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="a6fbe-151">Vermeiden der Verwendung relativer Verweise</span><span class="sxs-lookup"><span data-stu-id="a6fbe-151">Avoid using relative references</span></span>

<span data-ttu-id="a6fbe-152">Power Automation führt Ihr Skript in der ausgewählten Excel-Arbeitsmappe in Ihrem Auftrag aus.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="a6fbe-153">Die Arbeitsmappe wird möglicherweise geschlossen, wenn dies geschieht.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="a6fbe-154">Jede API, die den aktuellen Status des Benutzers verwendet, beispielsweise `Workbook.getActiveWorksheet` , schlägt fehl, wenn die Leistung automatisiert ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="a6fbe-155">Achten Sie beim Entwerfen Ihrer Skripts darauf, absolute Bezüge für Arbeitsblätter und Bereiche zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="a6fbe-156">Die folgenden Funktionen lösen einen Fehler aus und schlagen fehl, wenn Sie aus einem Skript in einem Power automatisieren-Flow aufgerufen werden.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a><span data-ttu-id="a6fbe-157">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a6fbe-157">Example</span></span>

<span data-ttu-id="a6fbe-158">Der folgende Screenshot zeigt einen Power-Automatisierungs Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) -Problem zugewiesen ist.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="a6fbe-159">Der Fluss führt ein Skript aus, mit dem das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzugefügt wird.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="a6fbe-160">Wenn es fünf oder mehr Probleme in dieser Tabelle gibt, sendet der Fluss eine e-Mail-Erinnerung.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Das Beispiel fließt wie im Power automatisieren-Fluss-Editor dargestellt.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="a6fbe-162">Die `main` Funktion des Skripts gibt die Problem-ID und den Titel des Problems als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Ausgabetabelle zurück.</span><span class="sxs-lookup"><span data-stu-id="a6fbe-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="a6fbe-163">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="a6fbe-163">See also</span></span>

- [<span data-ttu-id="a6fbe-164">Ausführen von Office-Skripts in Excel im Internet mit Power Automation</span><span class="sxs-lookup"><span data-stu-id="a6fbe-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="a6fbe-165">Automatisches Ausführen von Skripts mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="a6fbe-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="a6fbe-166">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="a6fbe-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="a6fbe-167">Erste Schritte mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="a6fbe-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="a6fbe-168">Referenzdokumentation zu Excel Online (Business) Connector</span><span class="sxs-lookup"><span data-stu-id="a6fbe-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
