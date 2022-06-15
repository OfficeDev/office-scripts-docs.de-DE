---
title: Verwenden von externen Abrufanrufen in Office-Skripts
description: Erfahren Sie, wie Sie externe API-Aufrufe in Office Skripts ausführen.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088092"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Verwenden von externen Abrufanrufen in Office-Skripts

Dieses Skript ruft grundlegende Informationen zu den Repositorys eines Benutzers GitHub ab. Es zeigt die Verwendung `fetch` in einem einfachen Szenario. Weitere Informationen zur Verwendung `fetch` oder zu anderen externen Aufrufen finden Sie [in der Unterstützung externer API-Aufrufe in Office Skripts](../../develop/external-calls.md). Informationen zum Arbeiten mit [JSON]](https://www.w3schools.com/whatis/whatis_json.asp) Objekten, z. B. was von den GitHub-APIs zurückgegeben wird, finden [Sie unter Verwenden von JSON zum Übergeben von Daten an und von Office Skripts](../../develop/use-json.md).

Weitere Informationen zu den GItHub-APIs, die in der [GitHub-API-Referenz](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) verwendet werden. Sie können auch die unformatierte API-Aufrufausgabe anzeigen, indem Sie `https://api.github.com/users/{USERNAME}/repos` einen Besuch in einem Webbrowser ausführen (achten Sie darauf, den Platzhalter {USERNAME} durch Ihre GitHub-ID zu ersetzen).

![Beispiel zum Abrufen von Repositoryinformationen](../../images/git.png)

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
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
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

## <a name="training-video-how-to-make-external-api-calls"></a>Schulungsvideo: Durchführen externer API-Aufrufe

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/fulP29J418E).
