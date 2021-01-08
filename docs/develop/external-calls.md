---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitungen für externe API-Aufrufe in einem Office-Skript.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784144"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skriptautoren sollten kein konsistentes Verhalten erwarten, wenn [sie externe APIs](https://developer.mozilla.org/docs/Web/API) während der Vorschauphase der Plattform verwenden. Verlassen Sie sich daher nicht auf externe APIs für kritische Skriptszenarien.

Aufrufe an externe APIs können nur über die Excel-Anwendung und nicht über Power Automate [unter normalen Umständen vorgenommen werden.](#external-calls-from-power-automate)

> [!CAUTION]
> Externe Anrufe können dazu führen, dass vertrauliche Daten für unerwünschte Endpunkte offengelegt werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.

## <a name="working-with-fetch"></a>Arbeiten mit `fetch`

Die [Fetch-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen von externen Diensten ab. Es handelt sich um eine API, daher müssen Sie `async` die Signatur Ihres `main` Skripts anpassen. Make the `main` function and have it return a `async` `Promise<void>` . Sie sollten auch den Anruf und den Abruf `await` `fetch` `json` sicherstellen. Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen werden, bevor das Skript beendet wird.

Das folgende Skript verwendet `fetch` JSON-Daten vom Testserver in der angegebenen URL abzurufen.

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

Das Beispielszenario für [Office-Skripts:](../resources/scenarios/noaa-data-fetch.md) Graph-Daten auf Wasserebene aus NOAA veranschaulicht den Abrufbefehl, der zum Abrufen von Datensätzen aus der Datenbank "National Oceanic and Demonstrates Administration's Scripts and Currents" verwendet wird.

## <a name="external-calls-from-power-automate"></a>Externe Aufrufe von Power Automate

Alle externen API-Aufrufe führen zu einem Fehler, wenn ein Skript mit Power Automate ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen der Ausführung eines Skripts über den Excel-Client und über Power Automate. Überprüfen Sie ihre Skripts auf solche Verweise, bevor Sie sie in einen Fluss erstellen.

> [!WARNING]
> Externe Aufrufe über den Power Automate [Excel Online Connector](/connectors/excelonlinebusiness) führen zu einem Fehler, um vorhandene Richtlinien zur Verhinderung von Datenverlust zu unterstützen. Skripts, die über Power Automate ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation ausgeführt. Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von Office-Skripts steuern. Ihr Administrator kann entweder den Excel Online Connector in Power Automate deaktivieren oder die Office-Skripts für Excel im Web über die Administratorsteuerelemente für [Office-Skripts deaktivieren.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Siehe auch

- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
- [Beispielszenario für Office-Skripts: Graph:Daten zum Wasserstand von NOAA](../resources/scenarios/noaa-data-fetch.md)
