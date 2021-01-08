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
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="903b1-103">Externe API-Anruf Unterstützung in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="903b1-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="903b1-104">Skriptautoren sollten kein konsistentes Verhalten erwarten, wenn [sie externe APIs](https://developer.mozilla.org/docs/Web/API) während der Vorschauphase der Plattform verwenden.</span><span class="sxs-lookup"><span data-stu-id="903b1-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="903b1-105">Verlassen Sie sich daher nicht auf externe APIs für kritische Skriptszenarien.</span><span class="sxs-lookup"><span data-stu-id="903b1-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="903b1-106">Aufrufe an externe APIs können nur über die Excel-Anwendung und nicht über Power Automate [unter normalen Umständen vorgenommen werden.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="903b1-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="903b1-107">Externe Anrufe können dazu führen, dass vertrauliche Daten für unerwünschte Endpunkte offengelegt werden.</span><span class="sxs-lookup"><span data-stu-id="903b1-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="903b1-108">Ihr Administrator kann einen Firewallschutz gegen solche Anrufe einrichten.</span><span class="sxs-lookup"><span data-stu-id="903b1-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="903b1-109">Arbeiten mit `fetch`</span><span class="sxs-lookup"><span data-stu-id="903b1-109">Working with `fetch`</span></span>

<span data-ttu-id="903b1-110">Die [Fetch-API](https://developer.mozilla.org/docs/Web/API/Fetch_API) ruft Informationen von externen Diensten ab.</span><span class="sxs-lookup"><span data-stu-id="903b1-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="903b1-111">Es handelt sich um eine API, daher müssen Sie `async` die Signatur Ihres `main` Skripts anpassen.</span><span class="sxs-lookup"><span data-stu-id="903b1-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="903b1-112">Make the `main` function and have it return a `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="903b1-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="903b1-113">Sie sollten auch den Anruf und den Abruf `await` `fetch` `json` sicherstellen.</span><span class="sxs-lookup"><span data-stu-id="903b1-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="903b1-114">Dadurch wird sichergestellt, dass diese Vorgänge abgeschlossen werden, bevor das Skript beendet wird.</span><span class="sxs-lookup"><span data-stu-id="903b1-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="903b1-115">Das folgende Skript verwendet `fetch` JSON-Daten vom Testserver in der angegebenen URL abzurufen.</span><span class="sxs-lookup"><span data-stu-id="903b1-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

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

<span data-ttu-id="903b1-116">Das Beispielszenario für [Office-Skripts:](../resources/scenarios/noaa-data-fetch.md) Graph-Daten auf Wasserebene aus NOAA veranschaulicht den Abrufbefehl, der zum Abrufen von Datensätzen aus der Datenbank "National Oceanic and Demonstrates Administration's Scripts and Currents" verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="903b1-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="903b1-117">Externe Aufrufe von Power Automate</span><span class="sxs-lookup"><span data-stu-id="903b1-117">External calls from Power Automate</span></span>

<span data-ttu-id="903b1-118">Alle externen API-Aufrufe führen zu einem Fehler, wenn ein Skript mit Power Automate ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="903b1-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="903b1-119">Dies ist ein Verhaltensunterschied zwischen der Ausführung eines Skripts über den Excel-Client und über Power Automate.</span><span class="sxs-lookup"><span data-stu-id="903b1-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="903b1-120">Überprüfen Sie ihre Skripts auf solche Verweise, bevor Sie sie in einen Fluss erstellen.</span><span class="sxs-lookup"><span data-stu-id="903b1-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="903b1-121">Externe Aufrufe über den Power Automate [Excel Online Connector](/connectors/excelonlinebusiness) führen zu einem Fehler, um vorhandene Richtlinien zur Verhinderung von Datenverlust zu unterstützen.</span><span class="sxs-lookup"><span data-stu-id="903b1-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="903b1-122">Skripts, die über Power Automate ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="903b1-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="903b1-123">Für zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung kann Ihr Administrator die Verwendung von Office-Skripts steuern.</span><span class="sxs-lookup"><span data-stu-id="903b1-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="903b1-124">Ihr Administrator kann entweder den Excel Online Connector in Power Automate deaktivieren oder die Office-Skripts für Excel im Web über die Administratorsteuerelemente für [Office-Skripts deaktivieren.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="903b1-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="903b1-125">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="903b1-125">See also</span></span>

- [<span data-ttu-id="903b1-126">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="903b1-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="903b1-127">Beispielszenario für Office-Skripts: Graph:Daten zum Wasserstand von NOAA</span><span class="sxs-lookup"><span data-stu-id="903b1-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
