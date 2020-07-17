---
title: Ausführen von Office-Skripts mit Power Automation
description: Vorgehensweise Abrufen von Office-Skripts für Excel im Internet arbeiten mit einem Power automatisieren Workflow.
ms.date: 07/01/2020
localization_priority: Normal
ms.openlocfilehash: 40a67f3d0e8f049a8ec5516c0af54c5fc6fb9319
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081593"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="518a5-103">Ausführen von Office-Skripts mit Power Automation</span><span class="sxs-lookup"><span data-stu-id="518a5-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="518a5-104">[Power Automation](https://flow.microsoft.com) ermöglicht Ihnen das Hinzufügen von Office-Skripts zu einem größeren, automatisierten Workflow.</span><span class="sxs-lookup"><span data-stu-id="518a5-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="518a5-105">Sie können Power automatisieren verwenden Sie Dinge wie das Hinzufügen des Inhalts einer e-Mail zur Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in ihren Projektverwaltungstools basierend auf Arbeitsmappen-Kommentaren.</span><span class="sxs-lookup"><span data-stu-id="518a5-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="518a5-106">Wenn Sie neu bei Power Automation sind, empfehlen wir Ihnen, die [Erste Schritte mit Power automatisieren](/power-automate/getting-started)zu besuchen.</span><span class="sxs-lookup"><span data-stu-id="518a5-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="518a5-107">Hier erfahren Sie mehr über die Automatisierung Ihrer Workflows in mehreren Diensten.</span><span class="sxs-lookup"><span data-stu-id="518a5-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="518a5-108">Derzeit können Sie Office-Skripts nicht über einen [freigegebenen Fluss](/power-automate/share-buttons)ausführen.</span><span class="sxs-lookup"><span data-stu-id="518a5-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="518a5-109">Nur der Benutzer, der ein Skript erstellt hat, kann es auch mithilfe von Power Automation ausführen.</span><span class="sxs-lookup"><span data-stu-id="518a5-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="518a5-110">Erste Schritte</span><span class="sxs-lookup"><span data-stu-id="518a5-110">Getting started</span></span>

<span data-ttu-id="518a5-111">Um mit der Kombination von Power Automation und Office-Skripts zu beginnen, führen Sie das Lernprogramm [Start using scripts with Power Automation aus](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="518a5-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="518a5-112">In diesem Artikel erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft.</span><span class="sxs-lookup"><span data-stu-id="518a5-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="518a5-113">Wenn Sie das Lernprogramm und das Lernprogramm [automatisch ausgeführte Skripts mit automatischer Power-Automatisierung ausführen](../tutorials/excel-power-automate-trigger.md) abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office-Skripts mit Power Automation Flows ein.</span><span class="sxs-lookup"><span data-stu-id="518a5-113">After you've completed that tutorial and the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="518a5-114">Excel Online-Connector (Business)</span><span class="sxs-lookup"><span data-stu-id="518a5-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="518a5-115">[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automation und Anwendungen.</span><span class="sxs-lookup"><span data-stu-id="518a5-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="518a5-116">Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gibt Ihrem Fluss Zugriff auf Excel-Arbeitsmappen.</span><span class="sxs-lookup"><span data-stu-id="518a5-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="518a5-117">Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, das über die ausgewählte Arbeitsmappe zugänglich ist.</span><span class="sxs-lookup"><span data-stu-id="518a5-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="518a5-118">Sie können nicht nur Skripts über einen Fluss ausführen, sondern auch Daten an und von der Arbeitsmappe mit dem Fluss durch die Skripts übergeben.</span><span class="sxs-lookup"><span data-stu-id="518a5-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="518a5-119">Die Aktion "Skript ausführen" gibt Benutzern, die den Excel Connector verwenden, wichtigen Zugriff auf Ihre Arbeitsmappe und deren Daten.</span><span class="sxs-lookup"><span data-stu-id="518a5-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="518a5-120">Darüber hinaus gibt es Sicherheitsrisiken mit Skripts, die externe API-Aufrufe durchführen, wie in [externe Aufrufe von Power Automation](external-calls.md)erläutert.</span><span class="sxs-lookup"><span data-stu-id="518a5-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="518a5-121">Wenn Ihr Administrator mit der Exposition hoch vertraulicher Daten befasst ist, können Sie entweder den Excel Online Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administrator Steuerelemente](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)einschränken.</span><span class="sxs-lookup"><span data-stu-id="518a5-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="518a5-122">Datenübertragung in Flows für Skripts</span><span class="sxs-lookup"><span data-stu-id="518a5-122">Data transfer in flows for scripts</span></span>

<span data-ttu-id="518a5-123">Mit Power Automation können Sie Datenteile zwischen den einzelnen Schritten Ihres Flows übergeben.</span><span class="sxs-lookup"><span data-stu-id="518a5-123">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="518a5-124">Skripts können so konfiguriert werden, dass alle Arten von Informationen akzeptiert werden, die Sie benötigen, und Sie geben alles aus Ihrer Arbeitsmappe zurück, die Sie in Ihrem Flow wünschen.</span><span class="sxs-lookup"><span data-stu-id="518a5-124">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="518a5-125">Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben.</span><span class="sxs-lookup"><span data-stu-id="518a5-125">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="518a5-126">Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu deklariert `main` .</span><span class="sxs-lookup"><span data-stu-id="518a5-126">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="518a5-127">Wenn Sie einen Block "Skript ausführen" im Flow erstellen, werden die akzeptierten Parameter und die zurückgegebenen Typen aufgefüllt.</span><span class="sxs-lookup"><span data-stu-id="518a5-127">When you create a "Run Script" block in you flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="518a5-128">Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den Block "Skript ausführen" des Flusses wiederholen.</span><span class="sxs-lookup"><span data-stu-id="518a5-128">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="518a5-129">Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.</span><span class="sxs-lookup"><span data-stu-id="518a5-129">This ensure the data is being parsed correctly.</span></span>

<span data-ttu-id="518a5-130">In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts behandelt, die in Power Automation verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="518a5-130">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="518a5-131">Wenn Sie eine praktische Herangehensweise zum Erlernen dieses Themas wünschen, probieren Sie das Lernprogramm [automatisch ausgeführte Skripts mit automatischer Power-Automatisierung](../tutorials/excel-power-automate-trigger.md) aus, oder erkunden Sie das Beispielszenario für [automatisierte Aufgaben Erinnerungen](../resources/scenarios/task-reminders.md) .</span><span class="sxs-lookup"><span data-stu-id="518a5-131">If you'd like a hands-on approach to learning this topic, try out the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="518a5-132">`main`Parameter: übergeben von Daten an ein Skript</span><span class="sxs-lookup"><span data-stu-id="518a5-132">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="518a5-133">Alle Skript Eingaben werden als zusätzliche Parameter für die `main` Funktion angegeben.</span><span class="sxs-lookup"><span data-stu-id="518a5-133">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="518a5-134">Wenn Sie beispielsweise möchten, dass ein Skript einen akzeptiert, `string` das einen Namen als Eingabe darstellt, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="518a5-134">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="518a5-135">Wenn Sie einen Fluss in Power Automation konfigurieren, können Sie Skript Eingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamische Inhalte angeben.</span><span class="sxs-lookup"><span data-stu-id="518a5-135">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="518a5-136">Details zu den Connectors eines einzelnen Diensts finden Sie in der [Power Automation Connector-Dokumentation](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="518a5-136">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="518a5-137">Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Zulagen und Einschränkungen.</span><span class="sxs-lookup"><span data-stu-id="518a5-137">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="518a5-138">Der erste Parameter muss vom Typ sein `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="518a5-138">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="518a5-139">Der Name des Parameters spielt keine Rolle.</span><span class="sxs-lookup"><span data-stu-id="518a5-139">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="518a5-140">Jeder Parameter muss einen Typ aufweisen.</span><span class="sxs-lookup"><span data-stu-id="518a5-140">Every parameter must have a type.</span></span>

3. <span data-ttu-id="518a5-141">Die Grundtypen `string` , `number` ,,, `boolean` `any` `unknown` , `object` und `undefined` werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="518a5-141">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="518a5-142">Arrays der zuvor aufgelisteten Grundtypen werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="518a5-142">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="518a5-143">Geschachtelte Arrays werden als Parameter unterstützt (jedoch nicht als Rückgabetypen).</span><span class="sxs-lookup"><span data-stu-id="518a5-143">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="518a5-144">Union-Typen sind zulässig, wenn es sich um eine Vereinigung von literalen handelt, die zu einem einzelnen Typ ( `string` , `number` oder `boolean` ) gehören.</span><span class="sxs-lookup"><span data-stu-id="518a5-144">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="518a5-145">Gewerkschaften eines unterstützten Typs mit undefined werden ebenfalls unterstützt.</span><span class="sxs-lookup"><span data-stu-id="518a5-145">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="518a5-146">Objekttypen sind zulässig, wenn Sie Eigenschaften vom Typ `string` , `number` ,, `boolean` unterstützte Arrays oder andere unterstützte Objekte enthalten.</span><span class="sxs-lookup"><span data-stu-id="518a5-146">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="518a5-147">Im folgenden Beispiel werden geschachtelte Objekte gezeigt, die als Parametertypen unterstützt werden:</span><span class="sxs-lookup"><span data-stu-id="518a5-147">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="518a5-148">Objekte müssen Ihre Schnittstelle oder Klassendefinition im Skript definiert haben.</span><span class="sxs-lookup"><span data-stu-id="518a5-148">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="518a5-149">Ein Objekt kann auch anonym Inline definiert werden, wie im folgenden Beispiel dargestellt:</span><span class="sxs-lookup"><span data-stu-id="518a5-149">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="518a5-150">Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer als solche gekennzeichnet werden `?` (beispielsweise `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="518a5-150">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="518a5-151">Standardparameterwerte sind zulässig (beispielsweise `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="518a5-151">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script"></a><span data-ttu-id="518a5-152">Zurückgeben von Daten aus einem Skript</span><span class="sxs-lookup"><span data-stu-id="518a5-152">Returning data from a script</span></span>

<span data-ttu-id="518a5-153">Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power-Automatisierungs Fluss verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="518a5-153">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="518a5-154">Wie bei Eingabeparametern stellt Power Automation einige Einschränkungen für den Rückgabetyp dar.</span><span class="sxs-lookup"><span data-stu-id="518a5-154">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="518a5-155">Die Grundtypen `string` , `number` , `boolean` , `void` und `undefined` werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="518a5-155">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="518a5-156">Union-Typen, die als Rückgabetypen verwendet werden, verfolgen dieselben Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="518a5-156">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="518a5-157">Array Typen sind zulässig, wenn Sie vom Typ `string` , `number` oder sind `boolean` .</span><span class="sxs-lookup"><span data-stu-id="518a5-157">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="518a5-158">Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.</span><span class="sxs-lookup"><span data-stu-id="518a5-158">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="518a5-159">Als Rückgabetypen verwendete Objekttypen entsprechen denselben Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="518a5-159">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="518a5-160">Die implizite Typisierung wird unterstützt, muss aber denselben Regeln wie ein definierter Typ entsprechen.</span><span class="sxs-lookup"><span data-stu-id="518a5-160">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="518a5-161">Vermeiden der Verwendung relativer Verweise</span><span class="sxs-lookup"><span data-stu-id="518a5-161">Avoid using relative references</span></span>

<span data-ttu-id="518a5-162">Power Automation führt Ihr Skript in der ausgewählten Excel-Arbeitsmappe in Ihrem Auftrag aus.</span><span class="sxs-lookup"><span data-stu-id="518a5-162">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="518a5-163">Die Arbeitsmappe wird möglicherweise geschlossen, wenn dies geschieht.</span><span class="sxs-lookup"><span data-stu-id="518a5-163">The workbook might be closed when this happens.</span></span> <span data-ttu-id="518a5-164">Jede API, die den aktuellen Status des Benutzers verwendet, beispielsweise `Workbook.getActiveWorksheet` , schlägt fehl, wenn die Leistung automatisiert ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="518a5-164">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="518a5-165">Achten Sie beim Entwerfen Ihrer Skripts darauf, absolute Bezüge für Arbeitsblätter und Bereiche zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="518a5-165">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="518a5-166">Die folgenden Funktionen lösen einen Fehler aus und schlagen fehl, wenn Sie aus einem Skript in einem Power automatisieren-Flow aufgerufen werden.</span><span class="sxs-lookup"><span data-stu-id="518a5-166">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

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

## <a name="example"></a><span data-ttu-id="518a5-167">Beispiel</span><span class="sxs-lookup"><span data-stu-id="518a5-167">Example</span></span>

<span data-ttu-id="518a5-168">Der folgende Screenshot zeigt einen Power-Automatisierungs Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) -Problem zugewiesen ist.</span><span class="sxs-lookup"><span data-stu-id="518a5-168">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="518a5-169">Der Fluss führt ein Skript aus, mit dem das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzugefügt wird.</span><span class="sxs-lookup"><span data-stu-id="518a5-169">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="518a5-170">Wenn es fünf oder mehr Probleme in dieser Tabelle gibt, sendet der Fluss eine e-Mail-Erinnerung.</span><span class="sxs-lookup"><span data-stu-id="518a5-170">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Das Beispiel fließt wie im Power automatisieren-Fluss-Editor dargestellt.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="518a5-172">Die `main` Funktion des Skripts gibt die Problem-ID und den Titel des Problems als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Ausgabetabelle zurück.</span><span class="sxs-lookup"><span data-stu-id="518a5-172">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="518a5-173">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="518a5-173">See also</span></span>

- [<span data-ttu-id="518a5-174">Ausführen von Office-Skripts in Excel im Internet mit Power Automation</span><span class="sxs-lookup"><span data-stu-id="518a5-174">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="518a5-175">Automatisches Ausführen von Skripts mit automatisiertem Power-Automatisierungs Fluss</span><span class="sxs-lookup"><span data-stu-id="518a5-175">Automatically run scripts with automated Power Automate flows</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="518a5-176">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="518a5-176">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="518a5-177">Erste Schritte mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="518a5-177">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="518a5-178">Referenzdokumentation zu Excel Online (Business) Connector</span><span class="sxs-lookup"><span data-stu-id="518a5-178">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
