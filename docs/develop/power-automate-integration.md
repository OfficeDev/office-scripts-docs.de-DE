---
title: Ausführen von Office-Skripts mit Power Automate
description: So erhalten Sie Office Scripts for Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755070"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="ca420-103">Ausführen von Office-Skripts mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="ca420-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="ca420-104">[Mit Power Automate](https://flow.microsoft.com) können Sie Office-Skripts zu einem größeren, automatisierten Workflow hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="ca420-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="ca420-105">Sie können Power Automate do-Aktionen verwenden, z. B. den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.</span><span class="sxs-lookup"><span data-stu-id="ca420-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="ca420-106">Erste Schritte</span><span class="sxs-lookup"><span data-stu-id="ca420-106">Getting started</span></span>

<span data-ttu-id="ca420-107">Wenn Sie noch keine Power Automate-Version haben, empfehlen wir, [erste Schritte mit Power Automate zu besuchen.](/power-automate/getting-started)</span><span class="sxs-lookup"><span data-stu-id="ca420-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="ca420-108">Dort erfahren Sie mehr über alle Automatisierungsmöglichkeiten, die Ihnen zur Verfügung stehen.</span><span class="sxs-lookup"><span data-stu-id="ca420-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="ca420-109">Die Dokumente hier konzentrieren sich auf die Funktionsweise von Office-Skripts mit Power Automate und die Verbesserung Ihrer Excel-Erfahrung.</span><span class="sxs-lookup"><span data-stu-id="ca420-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="ca420-110">Um mit der Kombination von Power Automate- und Office-Skripts zu beginnen, folgen Sie dem Lernprogramm [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="ca420-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="ca420-111">Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft.</span><span class="sxs-lookup"><span data-stu-id="ca420-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="ca420-112">Nachdem Sie dieses Lernprogramm und das Übergeben von Daten an Skripts in einem automatisch ausgeführten Power Automate-Fluss-Lernprogramm abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office-Skripts mit Power [Automate-Flüssen](../tutorials/excel-power-automate-trigger.md) zurück.</span><span class="sxs-lookup"><span data-stu-id="ca420-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="ca420-113">Excel Online (Business)-Connector</span><span class="sxs-lookup"><span data-stu-id="ca420-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="ca420-114">[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen.</span><span class="sxs-lookup"><span data-stu-id="ca420-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="ca420-115">Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gewährt Ihren Flüssen Zugriff auf Excel-Arbeitsmappen.</span><span class="sxs-lookup"><span data-stu-id="ca420-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="ca420-116">Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann.</span><span class="sxs-lookup"><span data-stu-id="ca420-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="ca420-117">Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Ihr Skript kann Informationen zu späteren Schritten im Fluss zurückgeben.</span><span class="sxs-lookup"><span data-stu-id="ca420-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ca420-118">Die Aktion "Skript ausführen" ermöglicht Personen, die den Excel-Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten.</span><span class="sxs-lookup"><span data-stu-id="ca420-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="ca420-119">Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie unter [Externe Aufrufe von Power Automate erläutert.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="ca420-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="ca420-120">Wenn Ihr Administrator mit der Belichtung besonders vertraulicher Daten zurecht kommt, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administratorsteuerelemente einschränken.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="ca420-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="ca420-121">Datenübertragung in Flüssen für Skripts</span><span class="sxs-lookup"><span data-stu-id="ca420-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="ca420-122">Mit Power Automate können Sie Datenteile zwischen den Schritten Ihres Datenflusses übergeben.</span><span class="sxs-lookup"><span data-stu-id="ca420-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="ca420-123">Skripts können so konfiguriert werden, dass sie alle arten von Informationen akzeptieren, die Sie benötigen, und alle Informationen aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Fluss wünschen.</span><span class="sxs-lookup"><span data-stu-id="ca420-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="ca420-124">Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben.</span><span class="sxs-lookup"><span data-stu-id="ca420-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="ca420-125">Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu `main` deklariert.</span><span class="sxs-lookup"><span data-stu-id="ca420-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="ca420-126">Wenn Sie einen "Skript ausführen"-Block in Ihrem Fluss erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt.</span><span class="sxs-lookup"><span data-stu-id="ca420-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="ca420-127">Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den Block "Skript ausführen" des Ablaufs wiederholen.</span><span class="sxs-lookup"><span data-stu-id="ca420-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="ca420-128">Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.</span><span class="sxs-lookup"><span data-stu-id="ca420-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="ca420-129">In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts beschrieben, die in Power Automate verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="ca420-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="ca420-130">Wenn Sie einen praktischen Ansatz zum Erlernen dieses [Themas](../resources/scenarios/task-reminders.md) verwenden möchten, testen Sie das Beispielszenario Daten an Skripts übergeben in einem automatisch ausgeführten [Power Automate-Fluss-Lernprogramm](../tutorials/excel-power-automate-trigger.md) oder erkunden Sie das Beispielszenario Automatisierte Aufgabenerinnerungen.</span><span class="sxs-lookup"><span data-stu-id="ca420-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="ca420-131">`main` Parameter: Übergeben von Daten an ein Skript</span><span class="sxs-lookup"><span data-stu-id="ca420-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="ca420-132">Alle Skripteingaben werden als zusätzliche Parameter für die Funktion `main` angegeben.</span><span class="sxs-lookup"><span data-stu-id="ca420-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="ca420-133">Wenn Sie z. B. möchten, dass ein Skript einen Namen als Eingabe akzeptiert, ändern Sie `string` die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="ca420-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="ca420-134">Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische [Werte,](/power-automate/use-expressions-in-conditions)Ausdrücke oder dynamische Inhalte angeben.</span><span class="sxs-lookup"><span data-stu-id="ca420-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="ca420-135">Details zum Connector eines einzelnen Diensts finden Sie in der [Power Automate Connector-Dokumentation](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="ca420-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="ca420-136">Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur `main` Skriptfunktion die folgenden Freibeträge und Einschränkungen.</span><span class="sxs-lookup"><span data-stu-id="ca420-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="ca420-137">Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein.</span><span class="sxs-lookup"><span data-stu-id="ca420-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="ca420-138">Der Parametername spielt keine Rolle.</span><span class="sxs-lookup"><span data-stu-id="ca420-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="ca420-139">Jeder Parameter muss einen Typ (z. B. `string` oder `number` ) haben.</span><span class="sxs-lookup"><span data-stu-id="ca420-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="ca420-140">Die grundlegenden `string` Typen , , , , , und werden `number` `boolean` `any` `unknown` `object` `undefined` unterstützt.</span><span class="sxs-lookup"><span data-stu-id="ca420-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="ca420-141">Arrays der zuvor aufgeführten Basistypen werden unterstützt.</span><span class="sxs-lookup"><span data-stu-id="ca420-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="ca420-142">Geschachtelte Arrays werden als Parameter (jedoch nicht als Rückgabetypen) unterstützt.</span><span class="sxs-lookup"><span data-stu-id="ca420-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="ca420-143">Union-Typen sind zulässig, wenn sie eine Vereinigung von Literalen sind, die zu einem einzelnen Typ gehören (z. B. `"Left" | "Right"` ).</span><span class="sxs-lookup"><span data-stu-id="ca420-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="ca420-144">Auch Unionen eines unterstützten Typs mit nicht definierten Werden unterstützt (z. B. `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="ca420-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="ca420-145">Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string` , , , unterstützte Arrays oder andere unterstützte Objekte `number` `boolean` enthalten.</span><span class="sxs-lookup"><span data-stu-id="ca420-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="ca420-146">Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:</span><span class="sxs-lookup"><span data-stu-id="ca420-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="ca420-147">Objekte müssen ihre Schnittstelle oder Klassendefinition im Skript definiert haben.</span><span class="sxs-lookup"><span data-stu-id="ca420-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="ca420-148">Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel:</span><span class="sxs-lookup"><span data-stu-id="ca420-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="ca420-149">Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer (z. B. ) als solche bezeichnet `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` werden.</span><span class="sxs-lookup"><span data-stu-id="ca420-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="ca420-150">Standardparameterwerte sind zulässig (z. `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` B. .</span><span class="sxs-lookup"><span data-stu-id="ca420-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="ca420-151">Zurückgeben von Daten aus einem Skript</span><span class="sxs-lookup"><span data-stu-id="ca420-151">Returning data from a script</span></span>

<span data-ttu-id="ca420-152">Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate-Fluss verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="ca420-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="ca420-153">Wie bei Eingabeparametern legt Power Automate einige Einschränkungen für den Rückgabetyp fest.</span><span class="sxs-lookup"><span data-stu-id="ca420-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="ca420-154">Die grundlegenden `string` Typen , , , und werden `number` `boolean` `void` `undefined` unterstützt.</span><span class="sxs-lookup"><span data-stu-id="ca420-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="ca420-155">Union-Typen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="ca420-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="ca420-156">Arraytypen sind zulässig, wenn sie vom Typ `string` , `number` oder `boolean` sind.</span><span class="sxs-lookup"><span data-stu-id="ca420-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="ca420-157">Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.</span><span class="sxs-lookup"><span data-stu-id="ca420-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="ca420-158">Objekttypen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.</span><span class="sxs-lookup"><span data-stu-id="ca420-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="ca420-159">Die implizite Eingabe wird unterstützt, muss jedoch dieselben Regeln wie ein definierter Typ befolgen.</span><span class="sxs-lookup"><span data-stu-id="ca420-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="ca420-160">Beispiel</span><span class="sxs-lookup"><span data-stu-id="ca420-160">Example</span></span>

<span data-ttu-id="ca420-161">Der folgende Screenshot zeigt einen Power Automate-Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub-Problem](https://github.com/) zugewiesen wird.</span><span class="sxs-lookup"><span data-stu-id="ca420-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="ca420-162">Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzufügt.</span><span class="sxs-lookup"><span data-stu-id="ca420-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="ca420-163">Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.</span><span class="sxs-lookup"><span data-stu-id="ca420-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate-Fluss-Editor, der den Beispielfluss zeigt.":::

<span data-ttu-id="ca420-165">Die Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der `main` Zeilen in der Problemtabelle zurück.</span><span class="sxs-lookup"><span data-stu-id="ca420-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="ca420-166">Weitere Informationen</span><span class="sxs-lookup"><span data-stu-id="ca420-166">See also</span></span>

- [<span data-ttu-id="ca420-167">Ausführen von Office-Skripts in Excel im Web mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="ca420-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="ca420-168">Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss</span><span class="sxs-lookup"><span data-stu-id="ca420-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="ca420-169">Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow</span><span class="sxs-lookup"><span data-stu-id="ca420-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="ca420-170">Problembehandlungsinformationen für Power Automate mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="ca420-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="ca420-171">Erste Schritte mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="ca420-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="ca420-172">Referenzdokumentation zu Excel Online (Business)-Connectors</span><span class="sxs-lookup"><span data-stu-id="ca420-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
