---
title: Ausführen Office Skripts mit Power Automate
description: So erhalten Sie Office Skripts für Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 96b07501e07383ace5ff88a8bc6b64ef145ebd5e
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074424"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="bbc35-103">Ausführen Office Skripts mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="bbc35-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="bbc35-104">[Power Automate](https://flow.microsoft.com) können Sie einem größeren, automatisierten Workflow Office Skripts hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="bbc35-105">Sie können Power Automate aktionen wie das Hinzufügen des Inhalts einer E-Mail zu einer Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren verwenden.</span><span class="sxs-lookup"><span data-stu-id="bbc35-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="bbc35-106">Erste Schritte</span><span class="sxs-lookup"><span data-stu-id="bbc35-106">Get started</span></span>

<span data-ttu-id="bbc35-107">Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, ["Erste Schritte mit Power Automate"](/power-automate/getting-started)zu besuchen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="bbc35-108">Dort können Sie mehr über alle Automatisierungsmöglichkeiten erfahren, die Ihnen zur Verfügung stehen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="bbc35-109">Die dokumente hier konzentrieren sich darauf, wie Office Skripts mit Power Automate arbeiten und wie dies dazu beitragen kann, Ihre Excel Erfahrung zu verbessern.</span><span class="sxs-lookup"><span data-stu-id="bbc35-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="bbc35-110">Um mit der Kombination von Power Automate und Office Skripts zu beginnen, führen Sie das Lernprogramm mit der [Verwendung von Skripts mit Power Automate aus.](../tutorials/excel-power-automate-manual.md)</span><span class="sxs-lookup"><span data-stu-id="bbc35-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="bbc35-111">Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft.</span><span class="sxs-lookup"><span data-stu-id="bbc35-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="bbc35-112">Nachdem Sie dieses Lernprogramm abgeschlossen haben und die [Daten an Skripts in einem automatisch ausgeführten Power Automate](../tutorials/excel-power-automate-trigger.md) Fluss-Lernprogramm übergeben haben, kehren Sie hier zurück, um ausführliche Informationen zum Verbinden von Office Skripts mit Power Automate Flüssen zu finden.</span><span class="sxs-lookup"><span data-stu-id="bbc35-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="bbc35-113">Excel Online-Connector (Business)</span><span class="sxs-lookup"><span data-stu-id="bbc35-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="bbc35-114">[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="bbc35-115">Der [connector Excel Online (Business)](/connectors/excelonlinebusiness) gewährt Ihren Flüssen Zugriff auf Excel Arbeitsmappen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="bbc35-116">Mit der Aktion "Skript ausführen" können Sie jedes Office Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann.</span><span class="sxs-lookup"><span data-stu-id="bbc35-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="bbc35-117">Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Sie können festlegen, dass Ihr Skript Informationen für spätere Schritte im Fluss zurückgibt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bbc35-118">Die Aktion "Skript ausführen" bietet Personen, die den Excel Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten.</span><span class="sxs-lookup"><span data-stu-id="bbc35-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="bbc35-119">Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie in [externen Aufrufen von Power Automate](external-calls.md)erläutert.</span><span class="sxs-lookup"><span data-stu-id="bbc35-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="bbc35-120">Wenn Sich Ihr Administrator mit der Offenlegung streng vertraulicher Daten befasst, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office Skripts über die [Administratorsteuerelemente Office Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings)einschränken.</span><span class="sxs-lookup"><span data-stu-id="bbc35-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="bbc35-121">Datenübertragung in Flüssen für Skripts</span><span class="sxs-lookup"><span data-stu-id="bbc35-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="bbc35-122">mit Power Automate können Sie Datenteile zwischen den Schritten des Flusses übergeben.</span><span class="sxs-lookup"><span data-stu-id="bbc35-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="bbc35-123">Skripts können so konfiguriert werden, dass sie alle benötigten Informationstypen akzeptieren und alles aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Flow wünschen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="bbc35-124">Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben.</span><span class="sxs-lookup"><span data-stu-id="bbc35-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="bbc35-125">Die Ausgabe des Skripts wird deklariert, indem ein Rückgabetyp zu `main` hinzugefügt wird.</span><span class="sxs-lookup"><span data-stu-id="bbc35-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="bbc35-126">Wenn Sie einen "Skript ausführen"-Block in Ihrem Flow erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="bbc35-127">Wenn Sie die Parameter ändern oder Typen ihres Skripts zurückgeben, müssen Sie den "Skript ausführen"-Block des Flusses wiederholen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="bbc35-128">Dadurch wird sichergestellt, dass die Daten richtig analysiert werden.</span><span class="sxs-lookup"><span data-stu-id="bbc35-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="bbc35-129">In den folgenden Abschnitten werden die Eingabe- und Ausgabedetails für Skripts behandelt, die in Power Automate verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="bbc35-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="bbc35-130">Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas wünschen, testen Sie die [Pass-Daten an Skripts in einem automatisch ausgeführten lernprogramm für Power Automate Fluss](../tutorials/excel-power-automate-trigger.md) oder erkunden Sie das Beispielszenario für [automatisierte Aufgabenerinnerungen.](../resources/scenarios/task-reminders.md)</span><span class="sxs-lookup"><span data-stu-id="bbc35-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="bbc35-131">`main` Parameter: Übergeben von Daten an ein Skript</span><span class="sxs-lookup"><span data-stu-id="bbc35-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="bbc35-132">Alle Skripteingaben werden als zusätzliche Parameter für die `main` Funktion angegeben.</span><span class="sxs-lookup"><span data-stu-id="bbc35-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="bbc35-133">Wenn sie beispielsweise möchten, dass ein Skript einen Namen akzeptiert, der `string` einen Namen als Eingabe darstellt, ändern Sie die Signatur in `main` `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="bbc35-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="bbc35-134">Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamischen Inhalt angeben.</span><span class="sxs-lookup"><span data-stu-id="bbc35-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="bbc35-135">Details zum Connector eines einzelnen Diensts finden Sie in der [Dokumentation zu Power Automate Connector.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="bbc35-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="bbc35-136">Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Einschränkungen und Einschränkungen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="bbc35-137">Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein.</span><span class="sxs-lookup"><span data-stu-id="bbc35-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="bbc35-138">Der Parametername spielt keine Rolle.</span><span class="sxs-lookup"><span data-stu-id="bbc35-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="bbc35-139">Jeder Parameter muss einen Typ aufweisen (z. B. `string` oder `number` ).</span><span class="sxs-lookup"><span data-stu-id="bbc35-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="bbc35-140">Die grundlegenden Typen `string` , , , , , und werden `number` `boolean` `unknown` `object` `undefined` unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="bbc35-141">Arrays der zuvor aufgeführten Basistypen werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="bbc35-142">Geschachtelte Arrays werden als Parameter unterstützt (aber nicht als Rückgabetypen).</span><span class="sxs-lookup"><span data-stu-id="bbc35-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="bbc35-143">Union-Typen sind zulässig, wenn es sich um eine Vereinigung von Literalen handelt, die zu einem einzelnen Typ gehören (z. B. `"Left" | "Right"` ).</span><span class="sxs-lookup"><span data-stu-id="bbc35-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="bbc35-144">Außerdem werden Verbindungen eines unterstützten Typs mit nicht definierten Typen unterstützt (z. B. `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="bbc35-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="bbc35-145">Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string` , `number` , `boolean` unterstützten Arrays oder anderen unterstützten Objekten enthalten.</span><span class="sxs-lookup"><span data-stu-id="bbc35-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="bbc35-146">Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:</span><span class="sxs-lookup"><span data-stu-id="bbc35-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="bbc35-147">Für Objekte muss ihre Schnittstelle oder Klassendefinition im Skript definiert sein.</span><span class="sxs-lookup"><span data-stu-id="bbc35-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="bbc35-148">Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel gezeigt:</span><span class="sxs-lookup"><span data-stu-id="bbc35-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="bbc35-149">Optionale Parameter sind zulässig und können mithilfe des optionalen Modifizierers (z. B. ) als solche gekennzeichnet `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` werden.</span><span class="sxs-lookup"><span data-stu-id="bbc35-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="bbc35-150">Standardwerte sind zulässig `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` (z. B. .</span><span class="sxs-lookup"><span data-stu-id="bbc35-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="bbc35-151">Zurückgeben von Daten aus einem Skript</span><span class="sxs-lookup"><span data-stu-id="bbc35-151">Return data from a script</span></span>

<span data-ttu-id="bbc35-152">Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate-Fluss verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="bbc35-153">Wie bei Eingabeparametern werden Power Automate einige Einschränkungen für den Rückgabetyp festlegen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="bbc35-154">Die grundlegenden Typen `string` , , , , und werden `number` `boolean` `void` `undefined` unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="bbc35-155">Union-Typen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="bbc35-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="bbc35-156">Arraytypen sind zulässig, wenn sie vom Typ `string` `number` sind, oder `boolean` .</span><span class="sxs-lookup"><span data-stu-id="bbc35-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="bbc35-157">Sie sind auch zulässig, wenn es sich bei dem Typ um eine unterstützte Vereinigung oder einen unterstützten Literaltyp handelt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="bbc35-158">Objekttypen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei der Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="bbc35-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="bbc35-159">Implizite Eingaben werden unterstützt, sie müssen jedoch denselben Regeln wie ein definierter Typ entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bbc35-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="bbc35-160">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bbc35-160">Example</span></span>

<span data-ttu-id="bbc35-161">Der folgende Screenshot zeigt einen Power Automate Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) Problem zugewiesen wird.</span><span class="sxs-lookup"><span data-stu-id="bbc35-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="bbc35-162">Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel Arbeitsmappe hinzufügt.</span><span class="sxs-lookup"><span data-stu-id="bbc35-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="bbc35-163">Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.</span><span class="sxs-lookup"><span data-stu-id="bbc35-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate Fluss-Editor mit dem Beispielfluss.":::

<span data-ttu-id="bbc35-165">Die `main` Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Problemtabelle zurück.</span><span class="sxs-lookup"><span data-stu-id="bbc35-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bbc35-166">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="bbc35-166">See also</span></span>

- [<span data-ttu-id="bbc35-167">Ausführen Office Skripts in Excel im Web mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="bbc35-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="bbc35-168">Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss</span><span class="sxs-lookup"><span data-stu-id="bbc35-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="bbc35-169">Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow</span><span class="sxs-lookup"><span data-stu-id="bbc35-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="bbc35-170">Problembehandlungsinformationen für Power Automate mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="bbc35-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="bbc35-171">Erste Schritte mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="bbc35-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="bbc35-172">Excel Referenzdokumentation zu Online-Connectors (Business)</span><span class="sxs-lookup"><span data-stu-id="bbc35-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
