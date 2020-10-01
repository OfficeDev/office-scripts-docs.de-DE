---
title: Externe API-Anruf Unterstützung in Office-Skripts
description: Unterstützung und Anleitungen für die Erstellung externer API-Aufrufe in einem Office-Skript.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: fa77e606e2b3ab90144507660d71561b278e82e5
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319630"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="bd15a-103">Externe API-Anruf Unterstützung in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="bd15a-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="bd15a-104">Die Office-Skriptplattform unterstützt keine Aufrufe [externer APIs](https://developer.mozilla.org/docs/Web/API).</span><span class="sxs-lookup"><span data-stu-id="bd15a-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="bd15a-105">Diese Aufrufe können jedoch unter den richtigen Umständen ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="bd15a-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="bd15a-106">Externe Anrufe können nur über den Excel-Client vorgenommen werden, nicht über Power Automation [unter normalen Umständen](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="bd15a-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="bd15a-107">Skriptautoren sollten bei der Verwendung externer APIs während der Preview-Phase der Plattform kein einheitliches Verhalten erwarten.</span><span class="sxs-lookup"><span data-stu-id="bd15a-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="bd15a-108">Dies liegt daran, wie die JavaScript-Laufzeit die Interaktion mit der Arbeitsmappe verwaltet.</span><span class="sxs-lookup"><span data-stu-id="bd15a-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="bd15a-109">Das Skript endet möglicherweise vor dem Abschluss des API-Aufrufs (oder `Promise` ist vollständig aufgelöst).</span><span class="sxs-lookup"><span data-stu-id="bd15a-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="bd15a-110">Verlassen Sie sich daher nicht auf externe APIs für kritische Skript Szenarien.</span><span class="sxs-lookup"><span data-stu-id="bd15a-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="bd15a-111">Externe Aufrufe können dazu führen, dass vertrauliche Daten unerwünschten Endpunkten ausgesetzt werden.</span><span class="sxs-lookup"><span data-stu-id="bd15a-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="bd15a-112">Ihr Administrator kann Firewallschutz gegen solche Anrufe einrichten.</span><span class="sxs-lookup"><span data-stu-id="bd15a-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="bd15a-113">Definitionsdateien für externe APIs</span><span class="sxs-lookup"><span data-stu-id="bd15a-113">Definition files for external APIs</span></span>

<span data-ttu-id="bd15a-114">Die Definitionsdateien für externe APIs sind in Office-Skripts nicht enthalten.</span><span class="sxs-lookup"><span data-stu-id="bd15a-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="bd15a-115">Durch die Verwendung solcher APIs werden Kompilierzeitfehler für fehlende Definitionen generiert.</span><span class="sxs-lookup"><span data-stu-id="bd15a-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="bd15a-116">Die APIs werden weiterhin ausgeführt (allerdings nur, wenn Sie über den Excel-Client ausgeführt werden), wie im folgenden Skript dargestellt:</span><span class="sxs-lookup"><span data-stu-id="bd15a-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

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

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="bd15a-117">Externe Anrufe von Power Automation</span><span class="sxs-lookup"><span data-stu-id="bd15a-117">External calls from Power Automate</span></span>

<span data-ttu-id="bd15a-118">Bei externen API-aufrufen tritt ein Fehler auf, wenn ein Skript mit Power Automation ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="bd15a-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="bd15a-119">Dies ist ein Verhaltensunterschied zwischen dem Ausführen eines Skripts über den Excel-Client und der Power Automation.</span><span class="sxs-lookup"><span data-stu-id="bd15a-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="bd15a-120">Achten Sie darauf, Ihre Skripts auf solche Verweise zu überprüfen, bevor Sie Sie in einem Flow erstellen.</span><span class="sxs-lookup"><span data-stu-id="bd15a-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="bd15a-121">Der Ausfall externer Anrufe [Excel Online Connector](/connectors/excelonlinebusiness) in Power Automation dient zur Wahrung vorhandener Richtlinien zur Verhinderung von Datenverlust.</span><span class="sxs-lookup"><span data-stu-id="bd15a-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="bd15a-122">Die Skripts, die über Power Automation ausgeführt werden, werden jedoch außerhalb Ihrer Organisation und außerhalb der Firewalls Ihrer Organisation durchgeführt.</span><span class="sxs-lookup"><span data-stu-id="bd15a-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="bd15a-123">Um zusätzlichen Schutz vor böswilligen Benutzern in dieser externen Umgebung zu erhalten, kann Ihr Administrator die Verwendung von Office-Skripts steuern.</span><span class="sxs-lookup"><span data-stu-id="bd15a-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="bd15a-124">Ihr Administrator kann entweder den Excel Online-Connector in Power automatisieren oder Office-Skripts für Excel im Internet über die [Office Scripts-Administrator Steuerelemente](/microsoft-365/admin/manage/manage-office-scripts-settings)deaktivieren.</span><span class="sxs-lookup"><span data-stu-id="bd15a-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="bd15a-125">Mehr dazu</span><span class="sxs-lookup"><span data-stu-id="bd15a-125">See also</span></span>

- [<span data-ttu-id="bd15a-126">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="bd15a-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)