---
title: Verwenden von externen Abrufanrufen in Office-Skripts
description: Erfahren Sie, wie Sie externe API-Aufrufe in Office ausführen.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545752"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Verwenden von externen Abrufanrufen in Office-Skripts

Dieses Skript ruft grundlegende Informationen zu den Repositorys eines Benutzers GitHub ab. Es zeigt, wie sie `fetch` in einem einfachen Szenario verwendet werden. Weitere Informationen zur Verwendung oder zu anderen externen Aufrufen finden Sie unter Unterstützung für externe `fetch` [API-Aufrufe in Office Skripts.](../../develop/external-calls.md)

Sie können mehr über die GItHub-APIs erfahren, die in der [GitHub-API-Referenz verwendet werden.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Sie können auch die Ausgabe von unformatiertem API-Aufruf anzeigen, indem Sie in einem Webbrowser (stellen Sie sicher, dass Sie den Platzhalter {USERNAME} durch Ihre GitHub `https://api.github.com/users/{USERNAME}/repos` ersetzen).

![Get repositorys info example](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Beispielcode: Grundlegende Informationen zum Benutzerrepository GitHub erhalten

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License 
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>Schulungsvideo: So nehmen Sie externe API-Aufrufe vor

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/fulP29J418E)
