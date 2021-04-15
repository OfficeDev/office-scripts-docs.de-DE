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
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="97fd3-103">Externe API-Aufrufe von Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="97fd3-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="97fd3-104">Office Scripts ermöglicht [die eingeschränkte Unterstützung für externe API-Aufrufe.](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="97fd3-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="97fd3-105">Es gibt keine Möglichkeit zum Anmelden oder Verwenden von OAuth2-Authentifizierungsflüssen.</span><span class="sxs-lookup"><span data-stu-id="97fd3-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="97fd3-106">Alle Schlüssel und Anmeldeinformationen müssen hartcodiert sein (oder aus einer anderen Quelle gelesen werden).</span><span class="sxs-lookup"><span data-stu-id="97fd3-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="97fd3-107">Es gibt keine Infrastruktur zum Speichern von API-Anmeldeinformationen und -Schlüsseln.</span><span class="sxs-lookup"><span data-stu-id="97fd3-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="97fd3-108">Dies muss vom Benutzer verwaltet werden.</span><span class="sxs-lookup"><span data-stu-id="97fd3-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="97fd3-109">Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden oder externe Daten in interne Arbeitsmappen gebracht werden.</span><span class="sxs-lookup"><span data-stu-id="97fd3-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="97fd3-110">Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.</span><span class="sxs-lookup"><span data-stu-id="97fd3-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="97fd3-111">Überprüfen Sie unbedingt lokale Richtlinien, bevor Sie sich auf externe Anrufe verlassen.</span><span class="sxs-lookup"><span data-stu-id="97fd3-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="97fd3-112">Wenn ein Skript einen API-Aufruf verwendet, funktioniert es in einem Power Automate-Szenario nicht.</span><span class="sxs-lookup"><span data-stu-id="97fd3-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="97fd3-113">Sie müssen die POWER Automate-HTTP-Aktion oder gleichwertige Aktionen verwenden, um Daten von einem externen Dienst zu ziehen oder an einen externen Dienst zu pushen.</span><span class="sxs-lookup"><span data-stu-id="97fd3-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="97fd3-114">Ein externer #A0 umfasst eine asynchrone API-Syntax und erfordert geringfügig erweiterte Kenntnisse der asynchronen Kommunikation.</span><span class="sxs-lookup"><span data-stu-id="97fd3-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="97fd3-115">Achten Sie darauf, den Datendurchsatz zu überprüfen, bevor Sie eine Abhängigkeit verwenden.</span><span class="sxs-lookup"><span data-stu-id="97fd3-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="97fd3-116">Beispielsweise ist das Herunterziehen des gesamten externen Datasets möglicherweise nicht die beste Option, und stattdessen sollte die Paginierung verwendet werden, um Daten in Blöcken zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="97fd3-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="97fd3-117">Nützliches Wissen und Ressourcen</span><span class="sxs-lookup"><span data-stu-id="97fd3-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="97fd3-118">[REST-API:](https://en.wikipedia.org/wiki/Representational_state_transfer)Wahrscheinlich verwenden Sie den API-Aufruf.</span><span class="sxs-lookup"><span data-stu-id="97fd3-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="97fd3-119">: Verstehen, wie dies funktioniert. [ `async` `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)</span><span class="sxs-lookup"><span data-stu-id="97fd3-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="97fd3-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Verstehen, wie dies funktioniert.</span><span class="sxs-lookup"><span data-stu-id="97fd3-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="97fd3-121">Schritte</span><span class="sxs-lookup"><span data-stu-id="97fd3-121">Steps</span></span>

1. <span data-ttu-id="97fd3-122">Markieren Sie `main` Ihre Funktion als asynchrone Funktion, indem Sie präfix `async` hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="97fd3-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="97fd3-123">Beispiel: `async function main(workbook: ExcelScript.Workbook)`.</span><span class="sxs-lookup"><span data-stu-id="97fd3-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="97fd3-124">Welche Art von API-Aufruf wird von Ihnen gemacht?</span><span class="sxs-lookup"><span data-stu-id="97fd3-124">Which type of API call are you making?</span></span> <span data-ttu-id="97fd3-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="97fd3-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="97fd3-126">Weitere Informationen finden Sie unter REST-API-Material.</span><span class="sxs-lookup"><span data-stu-id="97fd3-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="97fd3-127">Rufen Sie den Dienst-API-Endpunkt, Authentifizierungsanforderungen, Kopfzeilen usw. ab.</span><span class="sxs-lookup"><span data-stu-id="97fd3-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="97fd3-128">Definieren Sie die Eingabe oder `interface` Ausgabe, um die Codevollständigung und die Überprüfung der Entwicklungszeit zu unterstützen.</span><span class="sxs-lookup"><span data-stu-id="97fd3-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="97fd3-129">Weitere [Informationen finden](#training-video-how-to-make-external-api-calls) Sie unter Video.</span><span class="sxs-lookup"><span data-stu-id="97fd3-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="97fd3-130">Code, Test, Optimierung.</span><span class="sxs-lookup"><span data-stu-id="97fd3-130">Code, test, optimize.</span></span> <span data-ttu-id="97fd3-131">Sie können eine Funktion für Ihre API-Aufrufroutine erstellen, um sie aus anderen Teilen Ihres Skripts wiederverwendbar zu machen oder um sie in einem anderen Skript wiederzuverwenden (das Einfügen von Kopien wird auf diese Weise wesentlich einfacher).</span><span class="sxs-lookup"><span data-stu-id="97fd3-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="97fd3-132">Szenario</span><span class="sxs-lookup"><span data-stu-id="97fd3-132">Scenario</span></span>

<span data-ttu-id="97fd3-133">Dieses Skript ruft grundlegende Informationen zu den GitHub-Repositorys des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="97fd3-133">This script gets basic information about the user's GitHub repositories.</span></span>

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="97fd3-134">Ressourcen, die im Beispiel verwendet werden</span><span class="sxs-lookup"><span data-stu-id="97fd3-134">Resources used in the sample</span></span>

1. [<span data-ttu-id="97fd3-135">Abrufen von Repositorys Github-API-Referenz.</span><span class="sxs-lookup"><span data-stu-id="97fd3-135">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="97fd3-136">API-Aufrufausgabe: Wechseln Sie zu einem Webbrowser oder einer beliebigen HTTP-Schnittstelle, und geben Sie ein, und ersetzen Sie den `https://api.github.com/users/{USERNAME}/repos` Platzhalter {USERNAME} durch Ihre Github-ID.</span><span class="sxs-lookup"><span data-stu-id="97fd3-136">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="97fd3-137">Abgerufene Informationen: repo.name, repo.size, repo.owner.id, repo.license?. name</span><span class="sxs-lookup"><span data-stu-id="97fd3-137">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="97fd3-138">Beispielcode: Abrufen grundlegender Informationen zu den GitHub-Repositorys des Benutzers</span><span class="sxs-lookup"><span data-stu-id="97fd3-138">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="97fd3-139">Schulungsvideo: So nehmen Sie externe API-Aufrufe vor</span><span class="sxs-lookup"><span data-stu-id="97fd3-139">Training video: How to make external API calls</span></span>

<span data-ttu-id="97fd3-140">[![Video zum Erstellen externer API-Aufrufe ansehen](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video zum Erstellen externer API-Aufrufe")</span><span class="sxs-lookup"><span data-stu-id="97fd3-140">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
