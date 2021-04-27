---
title: Verwenden externer Abrufaufrufe in Office Skripts
description: Erfahren Sie, wie Sie externe API-Aufrufe in Office ausführen.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: a77ceb61c2ff46a7b6226b798462b7be2c8e1c54
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026994"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="a1650-103">Verwenden externer Abrufaufrufe in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="a1650-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="a1650-104">Dieses Skript ruft grundlegende Informationen zu den Repositorys eines Benutzers GitHub ab.</span><span class="sxs-lookup"><span data-stu-id="a1650-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="a1650-105">Es zeigt, wie sie `fetch` in einem einfachen Szenario verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a1650-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="a1650-106">Sie können mehr über die GItHub-APIs erfahren, die in der [GitHub-API-Referenz verwendet werden.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="a1650-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="a1650-107">Sie können auch die Ausgabe von unformatiertem API-Aufruf anzeigen, indem Sie einen Webbrowser aufrufen (stellen Sie sicher, dass Sie den `https://api.github.com/users/{USERNAME}/repos` Platzhalter {USERNAME} durch Ihre Github-ID ersetzen).</span><span class="sxs-lookup"><span data-stu-id="a1650-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![Get repositorys info example](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="a1650-109">Beispielcode: Grundlegende Informationen zum Benutzerrepository GitHub erhalten</span><span class="sxs-lookup"><span data-stu-id="a1650-109">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="a1650-110">Schulungsvideo: So nehmen Sie externe API-Aufrufe vor</span><span class="sxs-lookup"><span data-stu-id="a1650-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="a1650-111">[![Video zum Erstellen externer API-Aufrufe ansehen](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video zum Erstellen externer API-Aufrufe")</span><span class="sxs-lookup"><span data-stu-id="a1650-111">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
