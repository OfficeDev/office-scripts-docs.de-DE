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
# <a name="use-external-fetch-calls-in-office-scripts"></a>Verwenden externer Abrufaufrufe in Office Skripts

Dieses Skript ruft grundlegende Informationen zu den Repositorys eines Benutzers GitHub ab. Es zeigt, wie sie `fetch` in einem einfachen Szenario verwendet werden.

Sie können mehr über die GItHub-APIs erfahren, die in der [GitHub-API-Referenz verwendet werden.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Sie können auch die Ausgabe von unformatiertem API-Aufruf anzeigen, indem Sie einen Webbrowser aufrufen (stellen Sie sicher, dass Sie den `https://api.github.com/users/{USERNAME}/repos` Platzhalter {USERNAME} durch Ihre Github-ID ersetzen).

![Get repositorys info example](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Beispielcode: Grundlegende Informationen zum Benutzerrepository GitHub erhalten

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
