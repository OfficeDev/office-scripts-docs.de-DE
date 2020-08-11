---
title: Behandeln von Problemen mit Office-Skripts
description: Tipps und Techniken zum Debuggen von Office-Skripts sowie Hilferessourcen.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 00727b497d49a2d1d3f9c61e259b8d8d75028a59
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616682"
---
# <a name="troubleshooting-office-scripts"></a>Behandeln von Problemen mit Office-Skripts

Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen. Das ist okay. Wir verfügen über Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen.

## <a name="console-logs"></a>Konsolen Protokolle

Manchmal möchten Sie bei der Problembehandlung Nachrichten auf dem Bildschirm drucken. Diese können Ihnen den aktuellen Wert der Variablen oder die Code Pfade zeigen, die ausgelöst werden. Protokollieren Sie dazu Text in der Konsole.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Zeichenfolgen, die an übergeben werden, werden `console.log` in der Protokollierungs Konsole des Code-Editors angezeigt. Klicken Sie zum Aktivieren der Konsole auf die Schaltfläche **Ellipsen** , und wählen Sie **Protokolle...**

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="error-messages"></a>Fehlermeldungen

Wenn Ihr Excel-Skript auf ein Problem stößt, wird ein Fehler ausgegeben. Es wird eine Aufforderung angezeigt, in der Sie gefragt werden, ob Sie **Protokolle anzeigen**möchten. Drücken Sie die-Taste, um die Konsole zu öffnen und Fehler anzuzeigen.

## <a name="automate-tab-not-appearing"></a>Die Registerkarte "automatisieren" wird nicht angezeigt

Die folgenden Schritte sollten bei der Behandlung von Problemen im Zusammenhang mit der **automatischen** Registerkarte, die nicht in Excel für das Internet angezeigt wird, helfen.

1. [Stellen Sie sicher, dass Ihre Microsoft 365-Lizenz Office-Skripts enthält](../overview/excel.md#requirements).
1. [Lassen Sie den Administrator das Feature aktivieren](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf).
1. Stellen [Sie sicher, dass Ihr Browser unterstützt wird](platform-limits.md#browser-support).
1. [Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind](platform-limits.md#third-party-cookies).

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungs Problemen zu helfen. Häufig können Sie die Lösung für Ihr Problem in einer schnell Stapel-Überlauf Suche finden. Wenn nicht, stellen Sie Ihre Frage und markieren Sie Sie mit dem Tag "Office-Scripts". Vergessen Sie nicht, dass Sie ein Office- *Skript*erstellen, kein Office *-Add-in*.

Wenn ein Problem mit der Office-JavaScript-API auftritt, erstellen Sie ein Problem im [officedev/Office-js-GitHub-](https://github.com/OfficeDev/office-js) Repository. Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung bereitstellen. Das Erstellen eines Problems im Repository **officedev/Office-js** zeigt, dass Sie einen Fehler in der Office-JavaScript-API-Bibliothek gefunden haben, die das Produktteam adressieren sollte.

Wenn ein Problem mit dem Aktions Recorder oder-Editor auftritt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet](../develop/scripting-fundamentals.md)
- [Plattformbeschränkungen mit Office-Skripts](platform-limits.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](undo.md)
