---
title: Unterstützung externer API-Anrufe in Office-Skripts
description: Unterstützung und Anleitungen für die Erstellung externer API-Aufrufe in einem Office-Skript.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878756"
---
# <a name="external-api-call-support-in-office-scripts"></a>Unterstützung externer API-Anrufe in Office-Skripts

Die Office-Skriptplattform unterstützt keine Aufrufe [externer APIs](https://developer.mozilla.org/docs/Web/API). Diese Aufrufe können jedoch unter den richtigen Umständen ausgeführt werden. Externe Anrufe können nur über den Excel-Client vorgenommen werden, nicht über Power Automation [unter normalen Umständen](#external-calls-from-power-automate).

Skriptautoren sollten bei der Verwendung externer APIs während der Preview-Phase der Plattform kein einheitliches Verhalten erwarten. Dies liegt daran, wie die JavaScript-Laufzeit die Interaktion mit der Arbeitsmappe verwaltet. Das Skript endet möglicherweise vor dem Abschluss des API-Aufrufs (oder `Promise` ist vollständig aufgelöst). Verlassen Sie sich daher nicht auf externe APIs für kritische Skript Szenarien.

> [!CAUTION]
> Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden. Ihr Administrator kann Firewallschutz gegen solche Anrufe einrichten.

## <a name="definition-files-for-external-apis"></a>Definitionsdateien für externe APIs

Die Definitionsdateien für externe APIs sind in Office-Skripts nicht enthalten. Durch die Verwendung solcher APIs werden Kompilierzeitfehler für fehlende Definitionen generiert. Die APIs werden weiterhin ausgeführt (allerdings nur, wenn Sie über den Excel-Client ausgeführt werden), wie im folgenden Skript dargestellt:

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a>Externe Anrufe von Power Automation

Bei externen API-aufrufen tritt ein Fehler auf, wenn ein Skript mit Power Automation ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über den Excel-Client und der Power Automation. Achten Sie darauf, Ihre Skripts auf solche Verweise zu überprüfen, bevor Sie Sie in einem Flow erstellen.

> [!WARNING]
> Der Ausfall externer Anrufe [Excel Online Connector](/connectors/excelonlinebusiness) in Power Automation dient zur Wahrung vorhandener Richtlinien zur Verhinderung von Datenverlust. Die Skripts, die über Power Automation ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation durchgeführt. Um zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung zu erhalten, kann Ihr Administrator die Verwendung von Office-Skripts steuern. Ihr Administrator kann entweder den Excel Online-Connector in Power automatisieren oder Office-Skripts für Excel im Internet über die [Office Scripts-Administrator Steuerelemente](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)deaktivieren.

## <a name="see-also"></a>Siehe auch

- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)