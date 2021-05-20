---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitung für externe API-Aufrufe in einem Office Skript.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545082"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skriptautoren sollten während der Vorschauphase der Plattform kein konsistentes Verhalten erwarten, wenn [externe APIs](https://developer.mozilla.org/docs/Web/API) verwendet werden. Verlassen Sie sich daher nicht auf externe APIs für kritische Skriptszenarien.

Aufrufe externer APIs können nur über die Excel Anwendung erfolgen, nicht über Power Automate [unter normalen Umständen](#external-calls-from-power-automate).

> [!CAUTION]
> Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden. Ihr Administrator kann Einen Firewallschutz gegen solche Aufrufe einrichten.

## <a name="configure-your-script-for-external-calls"></a>Konfigurieren Ihres Skripts für externe Anrufe

Externe Aufrufe sind [asynchron](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) und erfordern, dass Ihr Skript als markiert `async` ist. Fügen Sie das `async` Präfix zu Ihrer Funktion hinzu, und lassen Sie `main` es eine `Promise` zurückgeben, wie hier gezeigt:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Skripts, die andere Informationen zurückgeben, können einen `Promise` dieses Typs zurückgeben. Wenn Ihr Skript z. B. ein Objekt zurückgeben `Employee` muss, wird die `: Promise <Employee>`

Sie müssen die Schnittstellen des externen Dienstes erlernen, um Anrufe an diesen Dienst zu tätigen. Wenn Sie `fetch` [REST-APIs](https://wikipedia.org/wiki/Representational_state_transfer)verwenden oder rest, müssen Sie die JSON-Struktur der zurückgegebenen Daten ermitteln. Für die Eingabe und Ausgabe aus Ihrem Skript sollten Sie eine erstellen, `interface` die den benötigten JSON-Strukturen entspricht. Dies gibt dem Skript mehr Typsicherheit. Ein Beispiel hierfür finden Sie unter [Verwenden von fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Einschränkungen bei externen Aufrufen von Office Skripts

* Es gibt keine Möglichkeit, sich anzumelden oder OAuth2-Authentifizierungsflüsse zu verwenden. Alle Schlüssel und Anmeldeinformationen müssen hartcodiert (oder aus einer anderen Quelle gelesen) werden.
* Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln. Dies muss vom Benutzer verwaltet werden.
* Dokumentcookies, `localStorage` und Objekte werden nicht `sessionStorage` unterstützt. 
* Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden oder externe Daten in interne Arbeitsmappen eingefügt werden. Ihr Administrator kann Einen Firewallschutz gegen solche Aufrufe einrichten. Achten Sie darauf, sich mit lokalen Richtlinien zu erkundigen, bevor Sie sich auf externe Anrufe verlassen.
* Achten Sie darauf, die Menge des Datendurchsatzes zu überprüfen, bevor Sie eine Abhängigkeit übernehmen. Beispielsweise ist das Herunterziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken abrufe.

## <a name="retrieve-information-with-fetch"></a>Abrufen von Informationen mit `fetch`

Die [Abruf-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen von externen Diensten ab. Es handelt sich um eine `async` API, daher müssen Sie die `main` Signatur Ihres Skripts anpassen. Machen Sie die `main` `async` Funktion, und lassen Sie sie eine `Promise<void>` zurückgeben. Sie sollten auch sicher sein, `await` den `fetch` Anruf und `json` abzurufen. Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen sind, bevor das Skript beendet wird.

Alle JSON-Daten, die von abgerufen `fetch` werden, müssen mit einer im Skript definierten Schnittstelle übereinstimmen. Der zurückgegebene Wert muss einem bestimmten Typ zugewiesen werden, da [Office Skripts den `any` Typ nicht unterstützen.](typescript-restrictions.md#no-any-type-in-office-scripts) Sie sollten in der Dokumentation für Ihren Dienst nach den Namen und Typen der zurückgegebenen Eigenschaften lesen. Fügen Sie dann die entsprechende Schnittstelle oder Schnittstellen zu Ihrem Skript hinzu.

Das folgende Skript `fetch` verwendet, um JSON-Daten vom Testserver in der angegebenen URL abzurufen. Beachten Sie die `JSONData` Schnittstelle, um die Daten als übereinstimmenden Typ zu speichern.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
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

### <a name="other-fetch-samples"></a>Andere `fetch` Proben

* Das Beispiel [externe Abrufaufrufe in Office Skripts](../resources/samples/external-fetch-calls.md) verwenden zeigt, wie Sie grundlegende Informationen über die GitHub Repositorys eines Benutzers abrufen.
* Das [Office Scripts-Beispielszenario: Graph Daten auf Wasserebene von NOAA](../resources/scenarios/noaa-data-fetch.md) zeigt den Fetch-Befehl, der zum Abrufen von Datensätzen aus der Tides and Currents-Datenbank der National Oceanic and Atmospheric Administration verwendet wird.

## <a name="external-calls-from-power-automate"></a>Externe Anrufe aus Power Automate

Jeder externe API-Aufruf schlägt fehl, wenn ein Skript mit Power Automate ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über die Excel Anwendung und durch Power Automate. Achten Sie darauf, Ihre Skripts auf solche Verweise zu überprüfen, bevor Sie sie in einen Flow erstellen.

Sie müssen HTTP [mit Azure AD](/connectors/webcontents/) oder andere gleichwertige Aktionen verwenden, um Daten aus einem externen Dienst abzuziehen oder an einen externen Dienst zu übertragen.

> [!WARNING]
> Externe Anrufe, die über den Power Automate [Excel Online-Connector](/connectors/excelonlinebusiness) getätigt werden, schlagen fehl, um die bestehenden Richtlinien zur Verhinderung von Datenverlust zu unterstützen. Skripts, die über Power Automate ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation ausgeführt. Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von Office Skripts steuern. Ihr Administrator kann den Excel Online-Connector in Power Automate deaktivieren oder Office Skripts für Excel im Web über die [Office Scripts-Administratorsteuerelemente](/microsoft-365/admin/manage/manage-office-scripts-settings)deaktivieren.

## <a name="see-also"></a>Siehe auch

* [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
* [Verwenden von externen Abrufanrufen in Office-Skripts](../resources/samples/external-fetch-calls.md)
* [Office Scripts-Beispielszenario: Graph Wasserstandsdaten von NOAA](../resources/scenarios/noaa-data-fetch.md)
