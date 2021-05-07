---
title: Ausführen Office Skripts mit Power Automate
description: So erhalten Sie Office Skripts für Excel im Web arbeiten mit einem Power Automate Workflow.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: fd2622880f08c253f4333e642d1ebb0410bce681
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232417"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen Office Skripts mit Power Automate

[Power Automate](https://flow.microsoft.com) können Sie Office Skripts zu einem größeren, automatisierten Workflow hinzufügen. Sie können die Power Automate verwenden, z. B. den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.

## <a name="getting-started"></a>Erste Schritte

Wenn Sie noch nicht mit Power Automate, empfehlen wir Den Besuch [Erste Schritte mit Power Automate](/power-automate/getting-started). Dort erfahren Sie mehr über alle Automatisierungsmöglichkeiten, die Ihnen zur Verfügung stehen. Die dokumente hier konzentrieren sich darauf, wie Office Skripts mit Power Automate arbeiten und wie dies zur Verbesserung Ihrer Excel beitragen kann.

Um mit der Kombination Power Automate und Office skripts zu beginnen, folgen Sie dem Lernprogramm [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md). Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Nachdem Sie dieses Lernprogramm und das Übergeben von Daten an Skripts in einem [automatisch](../tutorials/excel-power-automate-trigger.md) ausgeführten Power Automate-Fluss-Lernprogramm abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office Skripts mit Power Automate zurück.

## <a name="excel-online-business-connector"></a>Excel Online (Business)-Connector

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [Excel Online (Business) ermöglicht](/connectors/excelonlinebusiness) Ihren Flüssen zugriff auf Excel Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie beliebige Office Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Ihr Skript kann Informationen zu späteren Schritten im Fluss zurückgeben.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" bietet Personen, die den Excel verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten. Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie unter [Externe Aufrufe von Power Automate.](external-calls.md) Wenn Ihr Administrator mit der Belichtung besonders vertraulicher Daten zurecht kommt, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office-Skripts über die Office Scripts-Administratorsteuerelemente [einschränken.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flüssen für Skripts

Power Automate können Sie Datenteile zwischen Denkschritten des Datenflusses übergeben. Skripts können so konfiguriert werden, dass sie alle arten von Informationen akzeptieren, die Sie benötigen, und alle Informationen aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Fluss wünschen. Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben. Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu `main` deklariert.

> [!NOTE]
> Wenn Sie einen "Skript ausführen"-Block in Ihrem Fluss erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den Block "Skript ausführen" des Ablaufs wiederholen. Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.

In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts beschrieben, die in der Power Automate. Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas haben möchten, testen Sie das Beispielszenario Daten an [Skripts](../resources/scenarios/task-reminders.md) übergeben in einem automatisch ausgeführten Power Automate-Fluss-Lernprogramm [oder](../tutorials/excel-power-automate-trigger.md) erkunden Sie das Beispielszenario Automatisierte Aufgabenerinnerungen.

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die Funktion `main` angegeben. Wenn Sie z. B. möchten, dass ein Skript einen Namen als Eingabe akzeptiert, ändern Sie `string` die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .

Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische [Werte,](/power-automate/use-expressions-in-conditions)Ausdrücke oder dynamischen Inhalt angeben. Details zum Connector eines einzelnen Diensts finden Sie in der Dokumentation [Power Automate Connector](/connectors/).

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur `main` Skriptfunktion die folgenden Freibeträge und Einschränkungen.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein. Der Parametername spielt keine Rolle.

2. Jeder Parameter muss einen Typ (z. B. `string` oder `number` ) haben.

3. Die grundlegenden `string` Typen , , , , , und werden `number` `boolean` `any` `unknown` `object` `undefined` unterstützt.

4. Arrays der zuvor aufgeführten Basistypen werden unterstützt.

5. Geschachtelte Arrays werden als Parameter (jedoch nicht als Rückgabetypen) unterstützt.

6. Union-Typen sind zulässig, wenn sie eine Vereinigung von Literalen sind, die zu einem einzelnen Typ gehören (z. B. `"Left" | "Right"` ). Auch Unionen eines unterstützten Typs mit nicht definierten Werden unterstützt (z. B. `string | undefined` ).

7. Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string` , , , unterstützte Arrays oder andere unterstützte Objekte `number` `boolean` enthalten. Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:

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

8. Objekte müssen ihre Schnittstelle oder Klassendefinition im Skript definiert haben. Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer (z. B. ) als solche bezeichnet `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` werden.

10. Standardparameterwerte sind zulässig (z. `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` B. .

### <a name="returning-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate werden. Wie bei Eingabeparametern Power Automate einige Einschränkungen für den Rückgabetyp.

1. Die grundlegenden `string` Typen , , , und werden `number` `boolean` `void` `undefined` unterstützt.

2. Union-Typen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.

3. Arraytypen sind zulässig, wenn sie vom Typ `string` , `number` oder `boolean` sind. Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.

4. Objekttypen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.

5. Die implizite Eingabe wird unterstützt, muss jedoch dieselben Regeln wie ein definierter Typ befolgen.

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate, der ausgelöst wird, wenn [ihnen ein GitHub](https://github.com/) zugewiesen wird. Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Arbeitsmappe Excel hinzufügt. Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate-Fluss-Editor, der den Beispielfluss zeigt":::

Die Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der `main` Zeilen in der Problemtabelle zurück.

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

- [Ausführen Office Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Problembehandlungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Excel Referenzdokumentation zu Onlineconnector (Business)](/connectors/excelonlinebusiness/)
