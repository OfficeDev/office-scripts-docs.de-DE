---
title: Verwenden von externen Abrufanrufen in Office-Skripts
description: Erfahren Sie, wie Sie externe API-Aufrufe in Office Skripts ausführen.
ms.date: 05/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: feff9d49f9f50f14fd83b1864568df8dab02d417
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585527"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Verwenden von externen Abrufanrufen in Office-Skripts

Dieses Skript ruft grundlegende Informationen zu den GitHub Repositorys eines Benutzers ab. Es zeigt, wie sie in einem einfachen Szenario verwendet `fetch` werden. Weitere Informationen zur Verwendung `fetch` oder anderen externen Aufrufen finden Sie [unter "Externe API-Aufrufunterstützung" in Office Skripts](../../develop/external-calls.md).

Weitere Informationen zu den GItHub-APIs, die in der [GitHub-API-Referenz](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) verwendet werden. Sie können die Ausgabe `https://api.github.com/users/{USERNAME}/repos` des unformatierten API-Aufrufs auch in einem Webbrowser anzeigen (ersetzen Sie unbedingt den {USERNAME}-Platzhalter durch Ihre GitHub-ID).

![Beispiel zum Abrufen von Repositorys-Informationen](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Beispielcode: Abrufen grundlegender Informationen zu GitHub Repositorys des Benutzers

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

## <a name="training-video-how-to-make-external-api-calls"></a>Schulungsvideo: So führen Sie externe API-Aufrufe durch

[Sehen Sie sich an, wie Sie dieses Beispiel auf YouTube durchlaufen](https://youtu.be/fulP29J418E).
