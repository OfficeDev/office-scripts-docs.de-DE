---
title: Ausführen externer API-Aufrufe in Office-Skripts
description: Erfahren Sie, wie Sie externe API-Aufrufe in Office-Skripts ausführen.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754803"
---
# <a name="external-api-calls-from-office-scripts"></a>Externe API-Aufrufe von Office-Skripts

Office Scripts ermöglicht [die eingeschränkte Unterstützung für externe API-Aufrufe.](../../develop/external-calls.md)

> [!IMPORTANT]
>
> * Es gibt keine Möglichkeit zum Anmelden oder Verwenden von OAuth2-Authentifizierungsflüssen. Alle Schlüssel und Anmeldeinformationen müssen hartcodiert sein (oder aus einer anderen Quelle gelesen werden).
> * Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln. Dies muss vom Benutzer verwaltet werden.
> * Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden oder externe Daten in interne Arbeitsmappen gebracht werden. Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten. Überprüfen Sie unbedingt lokale Richtlinien, bevor Sie sich auf externe Anrufe verlassen.
> * Wenn ein Skript einen API-Aufruf verwendet, funktioniert es in einem Power Automate-Szenario nicht. Sie müssen die POWER Automate-HTTP-Aktion oder gleichwertige Aktionen verwenden, um Daten von einem externen Dienst zu ziehen oder an einen externen Dienst zu pushen.
> * Ein externer #A0 umfasst eine asynchrone API-Syntax und erfordert geringfügig erweiterte Kenntnisse der asynchronen Kommunikation.
> * Achten Sie darauf, den Datendurchsatz zu überprüfen, bevor Sie eine Abhängigkeit verwenden. Beispielsweise ist das Herunterziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken zu erhalten.

## <a name="useful-knowledge-and-resources"></a>Nützliches Wissen und Ressourcen

* [REST-API:](https://en.wikipedia.org/wiki/Representational_state_transfer)Wahrscheinlich verwenden Sie den API-Aufruf.
* : Verstehen, wie dies funktioniert. [ `async` `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Verstehen, wie dies funktioniert.

## <a name="steps"></a>Schritte

1. Markieren Sie `main` Ihre Funktion als asynchrone Funktion, indem Sie präfix `async` hinzufügen. Beispiel: `async function main(workbook: ExcelScript.Workbook)`.
1. Welche Art von API-Aufruf wird von Ihnen gemacht? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Weitere Informationen finden Sie unter REST-API-Material.
1. Rufen Sie den Dienst-API-Endpunkt, Authentifizierungsanforderungen, Kopfzeilen usw. ab.
1. Definieren Sie die Eingabe oder `interface` Ausgabe, um die Codevollständigung und die Überprüfung der Entwicklungszeit zu unterstützen. Weitere [Informationen finden](#training-video-how-to-make-external-api-calls) Sie unter Video.
1. Code, Test, Optimierung. Sie können eine Funktion für Ihre API-Aufrufroutine erstellen, um sie aus anderen Teilen Ihres Skripts wiederverwendbar zu machen oder um sie in einem anderen Skript wiederzuverwenden (das Einfügen von Kopien wird auf diese Weise wesentlich einfacher).

## <a name="scenario"></a>Szenario

Dieses Skript ruft grundlegende Informationen zu den GitHub-Repositorys des Benutzers ab.

## <a name="resources-used-in-the-sample"></a>Ressourcen, die im Beispiel verwendet werden

1. [Abrufen von Repositorys Github-API-Referenz.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. API-Aufrufausgabe: Wechseln Sie zu einem Webbrowser oder einer beliebigen HTTP-Schnittstelle, und geben Sie ein, und ersetzen Sie den `https://api.github.com/users/{USERNAME}/repos` Platzhalter {USERNAME} durch Ihre Github-ID.
1. Abgerufene Informationen: repo.name, repo.size, repo.owner.id, repo.license?. name

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Beispielcode: Abrufen grundlegender Informationen zu den GitHub-Repositorys des Benutzers

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>Schulungsvideo: So nehmen Sie externe API-Aufrufe vor

[![Video zum Erstellen externer API-Aufrufe ansehen](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video zum Erstellen externer API-Aufrufe")
