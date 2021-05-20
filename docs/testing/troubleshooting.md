---
title: Fehlerbehebung Office Skripts
description: Debuggen von Tipps und Techniken für Office Skripts sowie Hilferessourcen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545556"
---
# <a name="troubleshoot-office-scripts"></a>Fehlerbehebung Office Skripts

Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen. Das ist okay. Sie haben die Tools, um die Probleme zu finden und Ihre Skripte perfekt funktionieren zu lassen.

## <a name="types-of-errors"></a>Arten von Fehlern

Office Skriptfehler lassen sich in eine von zwei Kategorien einteilen:

* Kompilierungszeitfehler oder Warnungen
* Laufzeitfehler

### <a name="compile-time-errors"></a>Kompilierungszeitfehler

Kompilierzeitfehler und Warnungen werden zunächst im Code-Editor angezeigt. Diese zeigen die welligen roten Unterstreichungen im Editor. Sie werden auch unter der Registerkarte **Probleme** am unteren Rand des Aufgabenbereichs Code-Editor angezeigt. Wenn Sie den Fehler auswählen, werden weitere Details zum Problem angezeigt und Lösungen vorgeschlagen. Kompilierungsfehler sollten vor dem Ausführen des Skripts behoben werden.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im Hovertext des Code-Editors angezeigt wird":::

Möglicherweise werden auch orangefarbene Warnhinweise und graue Informationsmeldungen angezeigt. Diese weisen auf Leistungsvorschläge oder andere Möglichkeiten hin, bei denen das Skript unbeabsichtigte Auswirkungen haben kann. Solche Warnungen sollten eingehend geprüft werden, bevor sie zurückgewiesen werden.

### <a name="runtime-errors"></a>Laufzeitfehler

Laufzeitfehler treten aufgrund von Logikproblemen im Skript auf. Dies kann daran liegen, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet, oder eine andere geringfügige Diskrepanz zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe besteht. Das folgende Skript generiert einen Fehler, wenn kein Arbeitsblatt mit dem Namen "TestSheet" vorhanden ist.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Konsolenmeldungen

Sowohl Kompilierungszeit- als auch Laufzeitfehler zeigen Fehlermeldungen in der Konsole an, wenn ein Skript ausgeführt wird. Sie geben eine Zeilennummer an, bei der das Problem aufgetreten ist. Beachten Sie, dass die Ursache eines Problems möglicherweise eine andere Codezeile als in der Konsole angegeben ist.

Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` ](../develop/typescript-restrictions.md) Compilerfehler. Beachten Sie den Text `[5, 16]` am Anfang der Fehlerzeichenfolge. Dies zeigt an, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Die Code-Editor-Konsole, die eine explizite &quot;any&quot;-Fehlermeldung anzeigt":::

Das folgende Bild zeigt die Konsolenausgabe für einen Laufzeitfehler. Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen. Beachten Sie erneut die "Zeile 2" vor dem Fehler, um anzuzeigen, welche Zeile untersucht werden soll.
:::image type="content" source="../images/runtime-error-console.png" alt-text="Die Code-Editor-Konsole, die einen Fehler aus dem &quot;addWorksheet&quot;-Aufruf anzeigt":::

## <a name="console-logs"></a>Konsolenprotokolle

Drucken Sie Nachrichten mit der Anweisung auf den `console.log` Bildschirm. Diese Protokolle können Ihnen den aktuellen Wert von Variablen anzeigen oder zeigen, welche Codepfade ausgelöst werden. Rufen Sie dazu `console.log` ein beliebiges Objekt als Parameter auf. In der Regel ist a der einfachste Typ, der in der Konsole gelesen werden `string` kann.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

An Strings, die übergeben `console.log` werden, werden in der Protokollierungskonsole des Code-Editors am unteren Rand des Aufgabenbereichs angezeigt. Protokolle werden auf der Registerkarte **Ausgabe** gefunden, obwohl die Registerkarte automatisch den Fokus erhält, wenn ein Protokoll geschrieben wird.

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Registerkarte automatisieren wird nicht angezeigt oder Office Skripts nicht verfügbar

Die folgenden Schritte sollten zur Behebung von Problemen im Zusammenhang mit der Registerkarte **Automatisieren** helfen, die nicht in Excel im Web angezeigt werden.

1. [Stellen Sie sicher, dass Ihre Microsoft 365 Lizenz Office Scripts enthält.](../overview/excel.md#requirements)
1. [Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)
1. [Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind](platform-limits.md#third-party-cookies).
1. [Stellen Sie sicher, dass Ihr Administrator Office Skripts im Microsoft 365 Admin Center nicht deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Fehlerbehebung von Skripts in Power Automate

Informationen zum Ausführen von Skripts durch Power Automate finden Sie unter [Problembehandlung Office Skripts, die in Power Automate ausgeführt werden.](power-automate-troubleshooting.md)

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungsproblemen zu helfen. Häufig können Sie die Lösung für Ihr Problem durch eine schnelle Stack Overflow-Suche finden. Wenn nicht, stellen Sie Ihre Frage und markieren Sie sie mit dem Tag "office-scripts". Erwähnen Sie, dass Sie ein Office *Skript* erstellen, nicht ein Office *Add-In*.

Wenn ein Problem mit der Office JavaScript-API auftritt, erstellen Sie ein Problem im [Repository OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub. Die Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung leisten. Das Erstellen eines Problems im **OfficeDev/office-js-Repository** zeigt an, dass Sie einen Fehler in der Office JavaScript-API-Bibliothek gefunden haben, den das Produktteam beheben sollte.

Wenn ein Problem mit dem Aktionsrekorder oder -editor vorliegt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.

## <a name="see-also"></a>Siehe auch

- [Bewährte Methoden in Office-Skripts](../develop/best-practices.md)
- [Plattformlimits mit Office Scripts](platform-limits.md)
- [Verbessern Sie die Leistung Ihrer Office Scripts](../develop/web-client-performance.md)
- [Beheben Office Skripts, die in PowerAutomate ausgeführt werden](power-automate-troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
