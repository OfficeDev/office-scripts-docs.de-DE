---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitung zum Ausführen externer API-Aufrufe in einem Office-Skript.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570311"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skriptautoren sollten während der Vorschauphase der Plattform kein konsistentes Verhalten bei der Verwendung externer [APIs](https://developer.mozilla.org/docs/Web/API) erwarten. Verlassen Sie sich daher bei kritischen Skriptszenarien nicht auf externe APIs.

Aufrufe externer APIs können nur über die Excel-Anwendung und nicht über Power Automate [unter normalen Umständen vorgenommen werden.](#external-calls-from-power-automate)

> [!CAUTION]
> Externe Anrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.

## <a name="working-with-fetch"></a>Arbeiten mit `fetch`

Die [Fetch-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen von externen Diensten ab. Es handelt sich um eine API, daher müssen Sie `async` die Signatur Ihres `main` Skripts anpassen. Erstellen Sie `main` die `async` Funktion, und lassen Sie sie einen `Promise<void>` zurückgeben. Sie sollten auch den Anruf und den `await` `fetch` Abruf `json` sicherstellen. Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen sind, bevor das Skript endet.

Das folgende Skript `fetch` verwendet, um JSON-Daten vom Testserver in der angegebenen URL abzurufen.

```TypeScript
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

Das [Beispielszenario office scripts: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.

## <a name="external-calls-from-power-automate"></a>Externe Anrufe von Power Automate

Bei externen API-Aufrufen wird ein Fehler ausgeführt, wenn ein Skript mit Power Automate ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über den Excel-Client und power Automate. Überprüfen Sie ihre Skripts auf solche Verweise, bevor Sie sie in einen Fluss erstellen.

> [!WARNING]
> Externe Aufrufe über den Power Automate [Excel Online-Connector](/connectors/excelonlinebusiness) führen zu einem Fehler, um vorhandene Richtlinien zur Verhinderung von Datenverlust zu unterstützen. Skripts, die über Power Automate ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation ausgeführt. Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von Office-Skripts steuern. Ihr Administrator kann entweder den Excel Online-Connector in Power Automate deaktivieren oder Office-Skripts für Excel im Web über die [Office Scripts-Administratorsteuerelemente deaktivieren.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Siehe auch

- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
- [Beispielszenario für Office-Skripts: Graphen von Daten auf Wasserebene aus NOAA](../resources/scenarios/noaa-data-fetch.md)
