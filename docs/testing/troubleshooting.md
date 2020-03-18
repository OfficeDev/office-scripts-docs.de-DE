---
title: Problembehandlung bei Office-Skripts
description: Tipps und Techniken zum Debuggen von Office-Skripts sowie Hilferessourcen.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700203"
---
# <a name="troubleshooting-office-scripts"></a>Problembehandlung bei Office-Skripts

Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen. Das ist okay. Wir verfügen über Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen.

## <a name="console-logs"></a>Konsolen Protokolle

Manchmal möchten Sie bei der Problembehandlung Nachrichten auf dem Bildschirm drucken. Diese können Ihnen den aktuellen Wert der Variablen oder die Code Pfade zeigen, die ausgelöst werden. Protokollieren Sie dazu Text in der Konsole.

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> Vergessen Sie nicht `load` , Arbeitsblatt `sync` Daten und mit der Arbeitsmappe vor dem Protokollieren von Objekteigenschaften.

Zeichenfolgen,`console.log` die an übergeben werden, werden in der Protokollierungs Konsole des Code-Editors angezeigt. Klicken Sie zum Aktivieren der Konsole auf die Schaltfläche **Ellipsen** , und wählen Sie **Protokolle...**

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="error-messages"></a>Fehlermeldungen

Wenn Ihr Excel-Skript auf ein Problem stößt, wird ein Fehler ausgegeben. Es wird eine Aufforderung angezeigt, in der Sie gefragt werden, ob Sie **Protokolle anzeigen**möchten. Drücken Sie die-Taste, um die Konsole zu öffnen und Fehler anzuzeigen.

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungs Problemen zu helfen. Häufig können Sie die Lösung für Ihr Problem in einer schnell Stapel-Überlauf Suche finden. Wenn nicht, stellen Sie Ihre Frage und markieren Sie Sie mit dem Tag "Office-Scripts". Vergessen Sie nicht, dass Sie ein Office- *Skript*erstellen, kein Office *-Add-in*.

Wenn ein Problem mit der Office-JavaScript-API auftritt, erstellen Sie ein Problem im [officedev/Office-js-GitHub-](https://github.com/OfficeDev/office-js) Repository. Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung bereitstellen. Das Erstellen eines Problems im Repository **officedev/Office-js** zeigt, dass Sie einen Fehler in der Office-JavaScript-API-Bibliothek gefunden haben, die das Produktteam adressieren sollte.

Wenn ein Problem mit dem Aktions Recorder oder-Editor auftritt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Internet](../overview/excel.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet](../develop/scripting-fundamentals.md)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](undo.md)
