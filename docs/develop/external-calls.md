---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitung für externe API-Aufrufe in einem Office-Skript.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088043"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skripts unterstützen Anrufe an externe Dienste. Verwenden Sie diese Dienste, um Daten und andere Informationen an Ihre Arbeitsmappe zu übermitteln.

> [!CAUTION]
> Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten offengelegt werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.

> [!IMPORTANT]
> Aufrufe an externe APIs können nur über die Excel Anwendung erfolgen, nicht über Power Automate [unter normalen Umständen](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Konfigurieren Des Skripts für externe Anrufe

Externe [Aufrufe sind asynchron](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) und erfordern, dass Ihr Skript als `async`gekennzeichnet ist. Fügen Sie der Funktion das `async` Präfix `main` hinzu, und geben Sie ein `Promise`Präfix zurück, wie hier gezeigt:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Skripts, die andere Informationen zurückgeben, können einen `Promise` solchen Typ zurückgeben. Wenn ihr Skript z. B. ein `Employee` Objekt zurückgeben muss, lautet die Rückgabesignatur `: Promise <Employee>`

Sie müssen die Schnittstellen des externen Diensts kennen lernen, um Aufrufe an diesen Dienst zu tätigen. Wenn Sie [REST-APIs](https://wikipedia.org/wiki/Representational_state_transfer) verwenden`fetch`, müssen Sie die JSON-Struktur der zurückgegebenen Daten ermitteln. Für Eingaben und Ausgaben aus Ihrem Skript sollten Sie eine `interface` Übereinstimmung mit den erforderlichen JSON-Strukturen erstellen. Dadurch erhält das Skript mehr Typsicherheit. Ein Beispiel hierfür finden Sie unter [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Einschränkungen bei externen Aufrufen von Office-Skripts

* Es gibt keine Möglichkeit, sich anzumelden oder den OAuth2-Authentifizierungsflusstyp zu verwenden. Alle Schlüssel und Anmeldeinformationen müssen hartcodiert (oder aus einer anderen Quelle gelesen) werden.
* Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln. Dies muss vom Benutzer verwaltet werden.
* Dokumentcookies `localStorage`und `sessionStorage` Objekte werden nicht unterstützt.
* Externe Aufrufe können dazu führen, dass vertrauliche Daten für unerwünschte Endpunkte verfügbar gemacht werden oder externe Daten in interne Arbeitsmappen eingefügt werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten. Achten Sie darauf, lokale Richtlinien zu überprüfen, bevor Sie sich auf externe Anrufe verlassen.
* Achten Sie darauf, den Umfang des Datendurchsatzes zu überprüfen, bevor Sie eine Abhängigkeit erstellen. Beispielsweise ist das Herunterziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken abzurufen.

## <a name="retrieve-information-with-fetch"></a>Abrufen von Informationen mit `fetch`

Die [Abruf-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen aus externen Diensten ab. Es handelt sich um eine `async` API, daher müssen Sie die `main` Signatur Ihres Skripts anpassen. Erstellen Sie die `main` Funktion `async`. Sie sollten auch auf den Anruf und `json` den `fetch` Abruf achten`await`. Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen sind, bevor das Skript beendet wird.

Alle JSON-Daten, die von abgerufen werden `fetch` , müssen mit einer im Skript definierten Schnittstelle übereinstimmen. Der zurückgegebene Wert muss einem bestimmten Typ zugewiesen werden, da [Office Skripts den `any` Typ nicht unterstützen](typescript-restrictions.md#no-any-type-in-office-scripts). Sie sollten in der Dokumentation für Ihren Dienst nachsehen, wie die Namen und Typen der zurückgegebenen Eigenschaften sind. Fügen Sie dann die übereinstimmende Schnittstelle oder Schnittstellen zu Ihrem Skript hinzu.

Das folgende Skript verwendet `fetch` JSON-Daten vom Testserver in der angegebenen URL abzurufen. Beachten Sie die `JSONData` Schnittstelle zum Speichern der Daten als übereinstimmenden Typ.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a>Weitere `fetch` Beispiele

* Das Beispiel [zum Verwenden externer Abrufaufrufe in Office Skripts](../resources/samples/external-fetch-calls.md) zeigt, wie Sie grundlegende Informationen zu den repositorys GitHub eines Benutzers abrufen.
* Das [Beispielszenario Office Skripts: Graph Wasserstandsdaten aus NOAA](../resources/scenarios/noaa-data-fetch.md) veranschaulicht den Abrufbefehl, der zum Abrufen von Datensätzen aus der Datenbank "Tides and Currents" der National Oceanic and Atmospheric Administration verwendet wird.

## <a name="external-calls-from-power-automate"></a>Externe Anrufe von Power Automate

Ein externer API-Aufruf schlägt fehl, wenn ein Skript mit Power Automate ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über die Excel Anwendung und durch Power Automate. Achten Sie darauf, ihre Skripts auf solche Verweise zu überprüfen, bevor Sie sie in einem Fluss erstellen.

Sie müssen [HTTP mit Azure AD](/connectors/webcontents/) oder anderen entsprechenden Aktionen verwenden, um Daten abzuziehen oder an einen externen Dienst zu übertragen.

> [!WARNING]
> Externe Anrufe, die über den Power Automate [Excel Online Connector](/connectors/excelonlinebusiness) getätigt werden, schlagen fehl, um vorhandene Richtlinien zur Verhinderung von Datenverlust einzuhalten. Skripts, die über Power Automate ausgeführt werden, erfolgen jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation. Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von Office Skripts steuern. Ihr Administrator kann entweder den Excel Online-Connector in Power Automate deaktivieren oder Office Skripts für Excel im Web über die [Administratorsteuerelemente Office Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings) deaktivieren.

## <a name="see-also"></a>Siehe auch

* [Verwenden von JSON zum Übergeben von Daten an und von Office Skripts](use-json.md)
* [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
* [Verwenden von externen Abrufanrufen in Office-Skripts](../resources/samples/external-fetch-calls.md)
* [beispielszenario für Office-Skripts: Graph von Daten auf Wasserstandsebene aus NOAA](../resources/scenarios/noaa-data-fetch.md)
