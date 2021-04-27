---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitung zum Ausführen externer API-Aufrufe in einem Office Skripts.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: 7e6054fc50723dfbd95ded2e6e83eea3d38d2660
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026814"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skriptautoren sollten während der Vorschauphase der Plattform kein konsistentes Verhalten bei der Verwendung externer [APIs](https://developer.mozilla.org/docs/Web/API) erwarten. Verlassen Sie sich daher bei kritischen Skriptszenarien nicht auf externe APIs.

Aufrufe an externe APIs können nur über die Excel und nicht über Power Automate unter normalen [Umständen vorgenommen werden.](#external-calls-from-power-automate)

> [!CAUTION]
> Externe Anrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.

## <a name="configure-your-script-for-external-calls"></a>Konfigurieren des Skripts für externe Anrufe

Externe Aufrufe [sind asynchron](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) und erfordern, dass Ihr Skript als gekennzeichnet `async` ist. Fügen Sie `async` das Präfix zu Ihrer Funktion `main` hinzu, und geben Sie ein `Promise` zurück, wie hier gezeigt:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Skripts, die andere Informationen zurückgeben, können einen `Promise` dieser Typen zurückgeben. Wenn Ihr Skript z. B. ein Objekt zurückgeben muss, würde die `Employee` Rückgabesignatur `: Promise <Employee>`

Sie müssen die Schnittstellen des externen Diensts erlernen, um Anrufe an diesen Dienst zu nehmen. Wenn Sie oder `fetch` [REST-APIs verwenden,](https://wikipedia.org/wiki/Representational_state_transfer)müssen Sie die JSON-Struktur der zurückgegebenen Daten ermitteln. Für Eingaben und Ausgaben ihres Skripts sollten Sie eine so erstellen, dass sie den `interface` erforderlichen JSON-Strukturen entsprechen. Dies gibt dem Skript mehr Typsicherheit. Ein Beispiel dafür finden Sie unter [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Einschränkungen bei externen Aufrufen von Office Skripts

* Es gibt keine Möglichkeit zum Anmelden oder Verwenden von OAuth2-Authentifizierungsflüssen. Alle Schlüssel und Anmeldeinformationen müssen hartcodiert sein (oder aus einer anderen Quelle gelesen werden).
* Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln. Dies muss vom Benutzer verwaltet werden.
* Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden oder externe Daten in interne Arbeitsmappen gebracht werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten. Überprüfen Sie unbedingt lokale Richtlinien, bevor Sie sich auf externe Anrufe verlassen.
* Achten Sie darauf, den Datendurchsatz zu überprüfen, bevor Sie eine Abhängigkeit verwenden. Beispielsweise ist das Herunterziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken zu erhalten.

### <a name="working-with-fetch"></a>Arbeiten mit `fetch`

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

Das [beispielszenario Office Scripts: Graph](../resources/scenarios/noaa-data-fetch.md) Daten auf Wasserebene aus NOAA veranschaulicht den Abrufbefehl, der zum Abrufen von Datensätzen aus der Flut- und Currents-Datenbank der National Oceanic and Atmospheric Administration verwendet wird.

## <a name="external-calls-from-power-automate"></a>Externe Anrufe von Power Automate

Bei externen API-Aufrufen wird ein Fehler ausgeführt, wenn ein Skript mit Power Automate. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über Excel client und Power Automate. Überprüfen Sie ihre Skripts auf solche Verweise, bevor Sie sie in einen Fluss erstellen.

Sie müssen HTTP mit [Azure AD](/connectors/webcontents/) oder andere gleichwertige Aktionen verwenden, um Daten von einem externen Dienst zu ziehen oder an einen externen Dienst zu pushen.

> [!WARNING]
> Externe Anrufe, die über den Power Automate [Excel Onlineconnector](/connectors/excelonlinebusiness) vorgenommen werden, führen zu einem Fehler, um vorhandene Richtlinien zur Verhinderung von Datenverlust zu unterstützen. Skripts, die durch Power Automate ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation ausgeführt. Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von skripts Office steuern. Ihr Administrator kann entweder den Excel Onlineconnector in Power Automate deaktivieren oder Office Skripts für Excel im Web über die Office Scripts-Administratorsteuerelemente [deaktivieren.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Siehe auch

* [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
* [Verwenden externer Abrufaufrufe in Office Skripts](../resources/samples/external-fetch-calls.md)
* [Office Beispielszenario für Skripts: Graph daten auf Wasserebene aus NOAA](../resources/scenarios/noaa-data-fetch.md)
