---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitung für externe API-Aufrufe in einem Office-Skript.
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: 14b98e49907ff989684eceb9509edf56a1a72d9e
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332052"
---
# <a name="external-api-call-support-in-office-scripts"></a>Externe API-Anruf Unterstützung in Office-Skripts

Skripts unterstützen Aufrufe externer Dienste. Verwenden Sie diese Dienste, um Daten und andere Informationen für Ihre Arbeitsmappe bereitzustellen.

> [!CAUTION]
> Externe Aufrufe können dazu führen, dass vertrauliche Daten für unerwünschte Endpunkte verfügbar gemacht werden. Ihr Administrator kann Firewallschutz gegen solche Anrufe einrichten.

> [!IMPORTANT]
> Aufrufe externer APIs können nur über die Excel Anwendung erfolgen, nicht über Power Automate [unter normalen Umständen.](#external-calls-from-power-automate)

## <a name="configure-your-script-for-external-calls"></a>Konfigurieren Ihres Skripts für externe Aufrufe

Externe Aufrufe sind [asynchron](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) und erfordern, dass Ihr Skript als gekennzeichnet `async` ist. Fügen Sie das `async` Präfix zu Ihrer `main` Funktion hinzu, und lassen Sie es ein `Promise` zurückgeben, wie hier gezeigt:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Skripts, die andere Informationen zurückgeben, können einen `Promise` dieser Typen zurückgeben. Wenn Ihr Skript beispielsweise ein Objekt zurückgeben `Employee` muss, lautet die Rückgabesignatur `: Promise <Employee>`

Sie müssen die Schnittstellen des externen Diensts kennen lernen, um Aufrufe an diesen Dienst zu tätigen. Wenn Sie `fetch` [REST-APIs](https://wikipedia.org/wiki/Representational_state_transfer)verwenden, müssen Sie die JSON-Struktur der zurückgegebenen Daten bestimmen. Für die Eingabe in und die Ausgabe von Ihrem Skript sollten Sie eine an `interface` die erforderlichen JSON-Strukturen anpassen. Dadurch erhält das Skript mehr Typsicherheit. Ein Beispiel hierfür finden Sie unter [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Einschränkungen bei externen Aufrufen von Office Skripts

* Es gibt keine Möglichkeit, sich anzumelden oder den OAuth2-Authentifizierungsflusstyp zu verwenden. Alle Schlüssel und Anmeldeinformationen müssen hartcodiert (oder aus einer anderen Quelle gelesen) werden.
* Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln. Dies muss vom Benutzer verwaltet werden.
* `localStorage`Dokumentcookies und `sessionStorage` -objekte werden nicht unterstützt.
* Externe Aufrufe können dazu führen, dass vertrauliche Daten für unerwünschte Endpunkte verfügbar gemacht werden oder externe Daten in interne Arbeitsmappen übertragen werden. Ihr Administrator kann Firewallschutz gegen solche Anrufe einrichten. Überprüfen Sie unbedingt die lokalen Richtlinien, bevor Sie sich auf externe Aufrufe verlassen.
* Überprüfen Sie unbedingt die Datendurchsatzmenge, bevor Sie eine Abhängigkeit aufnehmen. Beispielsweise ist das Ziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken abzurufen.

## <a name="retrieve-information-with-fetch"></a>Abrufen von Informationen mit `fetch`

Die [Abruf-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen von externen Diensten ab. Es handelt sich um eine `async` API, daher müssen Sie die `main` Signatur Ihres Skripts anpassen. Erstellen Sie die `main` `async` Funktion, und lassen Sie sie ein `Promise<void>` . Sie sollten auch sicherstellen, dass `await` der `fetch` Anruf und `json` abruft. Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen sind, bevor das Skript beendet wird.

Alle VON abgerufenen JSON-Daten `fetch` müssen mit einer im Skript definierten Schnittstelle übereinstimmen. Der zurückgegebene Wert muss einem bestimmten Typ zugewiesen werden, da [Office Skripts den `any` Typ nicht unterstützen.](typescript-restrictions.md#no-any-type-in-office-scripts) Die Namen und Typen der zurückgegebenen Eigenschaften finden Sie in der Dokumentation für Ihren Dienst. Fügen Sie dann die übereinstimmende Schnittstelle oder Schnittstellen zu Ihrem Skript hinzu.

Das folgende Skript `fetch` verwendet, um JSON-Daten vom Testserver in der angegebenen URL abzurufen. Beachten Sie die `JSONData` Schnittstelle zum Speichern der Daten als übereinstimmenden Typ.

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

### <a name="other-fetch-samples"></a>Andere `fetch` Beispiele

* Das Beispiel ["Verwenden externer Abrufaufrufe in Office Skripts"](../resources/samples/external-fetch-calls.md) zeigt, wie Sie grundlegende Informationen zu den repositorys für GitHub eines Benutzers abrufen.
* Beispielszenario [für Office Skripts: Graph Wasserstandsdaten von NOAA](../resources/scenarios/noaa-data-fetch.md) veranschaulicht den Abrufbefehl, der zum Abrufen von Datensätzen aus der Datenbank der National- und Verwaltungsebene (National Wateric and Retriev Administration) verwendet wird.

## <a name="external-calls-from-power-automate"></a>Externe Aufrufe von Power Automate

Bei einem externen API-Aufruf tritt ein Fehler auf, wenn ein Skript mit Power Automate ausgeführt wird. Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über die Excel Anwendung und Power Automate. Überprüfen Sie ihre Skripts unbedingt auf solche Verweise, bevor Sie sie in einen Fluss integrieren.

Sie müssen HTTP [mit Azure AD](/connectors/webcontents/) oder anderen gleichwertigen Aktionen verwenden, um Daten aus einem externen Dienst abzurufen oder an einen externen Dienst zu übertragen.

> [!WARNING]
> Externe Aufrufe, die über den Power Automate [Excel Online Connector](/connectors/excelonlinebusiness) getätigt werden, schlagen fehl, um vorhandene Richtlinien zur Verhinderung von Datenverlust zu beheben. Skripts, die über Power Automate ausgeführt werden, erfolgen jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation. Um zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung zu erhalten, kann Ihr Administrator die Verwendung von Office Skripts steuern. Ihr Administrator kann entweder den Excel Online-Connector in Power Automate deaktivieren oder Office Skripts für Excel im Web über die [Administratorsteuerelemente Office Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings)deaktivieren.

## <a name="see-also"></a>Siehe auch

* [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
* [Verwenden von externen Abrufanrufen in Office-Skripts](../resources/samples/external-fetch-calls.md)
* [Office Skripts-Beispielszenario: Graph Wasserstandsdaten von NOAA](../resources/scenarios/noaa-data-fetch.md)
