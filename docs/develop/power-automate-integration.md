---
title: Ausführen Office Skripts mit Power Automate
description: So erhalten Sie Office Skripts für Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545040"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen Office Skripts mit Power Automate

[mit Power Automate](https://flow.microsoft.com) können Sie Office Skripts zu einem größeren, automatisierten Workflow hinzufügen. Sie können Power Automate z. B. den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.

## <a name="get-started"></a>Erste Schritte

Wenn Sie noch Power Automate haben, empfehlen wir Ihnen, erste Schritte [mit Power Automate](/power-automate/getting-started)zu besuchen. Dort erfahren Sie mehr über alle Automatisierungsmöglichkeiten, die Ihnen zur Verfügung stehen. Die Dokumente hier konzentrieren sich darauf, wie Office Skripts mit Power Automate arbeiten und wie dies dazu beitragen kann, Ihre Excel Erfahrung zu verbessern.

Um mit der Kombination Power Automate und Office Skripts zu beginnen, folgen Sie dem Tutorial [Starten Sie mit Skripts mit Power Automate](../tutorials/excel-power-automate-manual.md). Dadurch erfahren Sie, wie Sie einen Flow erstellen, der ein einfaches Skript aufruft. Nachdem Sie dieses Tutorial und die [Daten an Skripts in einem automatisch ausgeführten Power Automate-Flow-Tutorial](../tutorials/excel-power-automate-trigger.md) übergeben haben, kehren Sie hier zurück, um detaillierte Informationen zum Verbinden Office Skripts mit Power Automate-Flows zu erhalten.

## <a name="excel-online-business-connector"></a>Excel Online-Connector (Geschäft)

[Steckverbinder](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [Excel Online-Connector (Business)](/connectors/excelonlinebusiness) ermöglicht Ihren Flows den Zugriff auf Excel Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie jedes Office Skriptaufrufen aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Flow bereitgestellt werden können, oder Sie können Informationen für spätere Schritte im Flow zurückgeben.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" bietet Personen, die den Excel Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten. Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe tätigen, wie in [externen Aufrufen von Power Automate](external-calls.md)erläutert. Wenn Ihr Administrator mit der Offenlegung hochsensibler Daten befasst ist, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office Skripts über die [Office-Skripts-Administratorsteuerelemente](/microsoft-365/admin/manage/manage-office-scripts-settings)einschränken.

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flows für Skripte

mit Power Automate können Sie Daten zwischen den Schritten Ihres Flows übergeben. Skripts können so konfiguriert werden, dass sie alle Arten von Informationen akzeptieren, die Sie benötigen, und alles aus Ihrer Arbeitsmappe zurückgeben, das Sie in Ihrem Flow benötigen. Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben. Die Ausgabe aus dem Skript wird deklariert, indem ein Rückgabetyp zu hinzugefügt `main` wird.

> [!NOTE]
> Wenn Sie einen "Run Script"-Block in Ihrem Flow erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den "Skriptausführen"-Block Ihres Flows wiederholen. Dadurch wird sichergestellt, dass die Daten korrekt analysiert werden.

Die folgenden Abschnitte behandeln die Details der Eingabe und Ausgabe von Skripten, die in Power Automate verwendet werden. Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas wünschen, probieren Sie die [Daten an Skripts übergeben in einem automatisch ausgeführten Power Automate-Flow-Lernprogramm](../tutorials/excel-power-automate-trigger.md) aus, oder erkunden Sie das Beispielszenario für [automatisierte Aufgabenerinnerungen.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die `main` Funktion angegeben. Wenn Sie z. B. möchten, dass ein Skript einen `string` Namen als Eingabe akzeptiert, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .

Wenn Sie einen Flow in Power Automate konfigurieren, können Sie Skripteingabe als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamischen Inhalt angeben. Details zum Connector eines einzelnen Dienstes finden Sie in der [dokumentation Power Automate Connector](/connectors/).

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines `main` Skripts die folgenden Einschränkungen und Einschränkungen.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein. Sein Parametername spielt keine Rolle.

2. Jeder Parameter muss einen Typ haben (z. B. `string` oder `number` ).

3. Die Basistypen `string` , , , , und werden `number` `boolean` `unknown` `object` `undefined` unterstützt.

4. Arrays der zuvor aufgeführten Basistypen werden unterstützt.

5. Geschachtelte Arrays werden als Parameter (jedoch nicht als Rückgabetypen) unterstützt.

6. Union-Typen sind zulässig, wenn sie eine Vereinigung von Literalen sind, die zu einem einzelnen Typ gehören (z. `"Left" | "Right"` B. ). Auch Unions eines unterstützten Typs mit nicht definiertem Typ werden unterstützt (z. `string | undefined` B. ).

7. Objekttypen sind zulässig, wenn sie Eigenschaften des `string` Typs , `number` , `boolean` unterstützten Arrays oder anderen unterstützten Objekten enthalten. Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:

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

8. Für Objekte muss ihre Schnittstellen- oder Klassendefinition im Skript definiert sein. Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optionale Parameter sind zulässig und können als solche bezeichnet werden, indem der optionale Modifikator verwendet wird `?` (z. `function main(workbook: ExcelScript.Workbook, Name?: string)` B. ).

10. Standardparameterwerte sind zulässig (z. `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` B. .

### <a name="return-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate-Flow verwendet werden sollen. Wie bei Eingabeparametern stellt Power Automate einige Einschränkungen für den Rückgabetyp bereit.

1. Die Basistypen `string` , , , und werden `number` `boolean` `void` `undefined` unterstützt.

2. Union-Typen, die als Rückgabetypen verwendet werden, folgen denselben Einschränkungen wie bei der Verwendung als Skriptparameter.

3. Arraytypen sind zulässig, wenn sie vom Typ `string` , `number` oder `boolean` sind. Sie sind auch zulässig, wenn es sich bei dem Typ um einen unterstützten Union- oder unterstützten Literaltyp handelt.

4. Objekttypen, die als Rückgabetypen verwendet werden, folgen denselben Einschränkungen wie bei der Verwendung als Skriptparameter.

5. Implizite Typisierung wird unterstützt, obwohl sie denselben Regeln wie ein definierter Typ folgen muss.

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) Problem zugewiesen wird. Der Flow führt ein Skript aus, das das Problem einer Tabelle in einer Excel Arbeitsmappe hinzufügt. Wenn diese Tabelle fünf oder mehr Probleme enthält, sendet der Flow eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate-Flow-Editor mit dem Beispielfluss":::

Die `main` Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Issue-Tabelle zurück.

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

## <a name="see-also"></a>Siehe auch

- [Führen Sie Office Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Fehlerbehebungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Excel Online -Connector-Referenzdokumentation (Business)](/connectors/excelonlinebusiness/)
