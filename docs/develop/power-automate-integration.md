---
title: Ausführen Office Skripts mit Power Automate
description: So erhalten Sie Office Skripts für Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 61b43904cbc46b97a0102230c9c87c1051edd1516668f42fbded63c53c958de9
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846519"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen Office Skripts mit Power Automate

[Power Automate](https://flow.microsoft.com) können Sie einem größeren, automatisierten Workflow Office Skripts hinzufügen. Sie können Power Automate aktionen wie das Hinzufügen des Inhalts einer E-Mail zu einer Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren verwenden.

## <a name="get-started"></a>Erste Schritte

Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, ["Erste Schritte mit Power Automate"](/power-automate/getting-started)zu besuchen. Dort können Sie mehr über alle Automatisierungsmöglichkeiten erfahren, die Ihnen zur Verfügung stehen. Die dokumente hier konzentrieren sich darauf, wie Office Skripts mit Power Automate arbeiten und wie dies dazu beitragen kann, Ihre Excel Erfahrung zu verbessern.

Um mit der Kombination von Power Automate und Office Skripts zu beginnen, führen Sie das Lernprogramm mit der [Verwendung von Skripts mit Power Automate aus.](../tutorials/excel-power-automate-manual.md) Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Nachdem Sie dieses Lernprogramm abgeschlossen haben und die [Daten an Skripts in einem automatisch ausgeführten Power Automate](../tutorials/excel-power-automate-trigger.md) Fluss-Lernprogramm übergeben haben, kehren Sie hier zurück, um ausführliche Informationen zum Verbinden von Office Skripts mit Power Automate Flüssen zu finden.

## <a name="excel-online-business-connector"></a>Excel Online-Connector (Business)

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [connector Excel Online (Business)](/connectors/excelonlinebusiness) gewährt Ihren Flüssen Zugriff auf Excel Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie jedes Office Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Ihr Skript Informationen für spätere Schritte im Fluss zurückgeben lassen.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" bietet Personen, die den Excel Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten. Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie in [externen Aufrufen von Power Automate](external-calls.md)erläutert. Wenn Sich Ihr Administrator mit der Offenlegung streng vertraulicher Daten befasst, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office Skripts über die [Administratorsteuerelemente Office Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings)einschränken.

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flüssen für Skripts

mit Power Automate können Sie Datenteile zwischen den Schritten des Flusses übergeben. Skripts können so konfiguriert werden, dass sie alle arten von Informationen akzeptieren, die Sie benötigen, und alles aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Flow wünschen. Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben. Die Ausgabe des Skripts wird deklariert, indem ein Rückgabetyp zu `main` hinzugefügt wird.

> [!NOTE]
> Wenn Sie einen "Skript ausführen"-Block in Ihrem Flow erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter ändern oder Typen ihres Skripts zurückgeben, müssen Sie den "Skript ausführen"-Block des Flusses wiederholen. Dadurch wird sichergestellt, dass die Daten richtig analysiert werden.

In den folgenden Abschnitten werden die Eingabe- und Ausgabedetails für Skripts behandelt, die in Power Automate verwendet werden. Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas wünschen, testen Sie die [Pass-Daten an Skripts in einem automatisch ausgeführten Lernprogramm Power Automate Fluss oder](../tutorials/excel-power-automate-trigger.md) erkunden Sie das Beispielszenario für [automatisierte Aufgabenerinnerungen.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die `main` Funktion angegeben. Wenn sie beispielsweise möchten, dass ein Skript einen Wert akzeptiert, der `string` einen Namen als Eingabe darstellt, ändern Sie die Signatur in `main` `function main(workbook: ExcelScript.Workbook, name: string)` .

Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamischen Inhalt angeben. Details zum Connector eines einzelnen Diensts finden Sie in der [Dokumentation zu Power Automate Connector.](/connectors/)

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Einschränkungen und Einschränkungen.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein. Der Parametername spielt keine Rolle.

2. Jeder Parameter muss einen Typ aufweisen (z. B. `string` oder `number` ).

3. Die grundlegenden Typen `string` , , , , , und werden `number` `boolean` `unknown` `object` `undefined` unterstützt.

4. Arrays der zuvor aufgeführten Basistypen werden unterstützt.

5. Geschachtelte Arrays werden als Parameter unterstützt (aber nicht als Rückgabetypen).

6. Union-Typen sind zulässig, wenn es sich um eine Vereinigung von Literalen handelt, die zu einem einzelnen Typ gehören (z. B. `"Left" | "Right"` ). Außerdem werden Verbindungen eines unterstützten Typs mit nicht definierten Typen unterstützt (z. B. `string | undefined` ).

7. Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string` , `number` , `boolean` unterstützten Arrays oder anderen unterstützten Objekten enthalten. Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:

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

8. Für Objekte muss ihre Schnittstelle oder Klassendefinition im Skript definiert sein. Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel gezeigt:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optionale Parameter sind zulässig und können mithilfe des optionalen Modifizierers (z. B. ) als solche gekennzeichnet `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` werden.

10. Standardwerte sind zulässig `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` (z. B. .

### <a name="return-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate Fluss verwendet werden sollen. Wie bei Eingabeparametern legt Power Automate einige Einschränkungen für den Rückgabetyp.

1. Die grundlegenden Typen `string` , , , und werden `number` `boolean` `void` `undefined` unterstützt.

2. Union-Typen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei der Verwendung als Skriptparameter.

3. Arraytypen sind zulässig, wenn sie vom Typ `string` `number` sind, oder `boolean` . Sie sind auch zulässig, wenn es sich bei dem Typ um eine unterstützte Vereinigung oder einen unterstützten Literaltyp handelt.

4. Objekttypen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei der Verwendung als Skriptparameter.

5. Implizite Eingaben werden unterstützt, müssen jedoch denselben Regeln wie ein definierter Typ entsprechen.

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) Problem zugewiesen wird. Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel Arbeitsmappe hinzufügt. Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate Fluss-Editor mit dem Beispielfluss.":::

Die `main` Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Problemtabelle zurück.

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

## <a name="see-also"></a>Weitere Artikel

- [Ausführen Office Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Problembehandlungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Excel Referenzdokumentation zu Online-Connectors (Business)](/connectors/excelonlinebusiness/)
