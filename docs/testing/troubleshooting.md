---
title: Problembehandlung Office Skripts
description: Debuggen von Tipps und Techniken für Office Skripts sowie Hilferessourcen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 04ea0ea5d49d40667d249a6f4f4b109e03362940
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631702"
---
# <a name="troubleshoot-office-scripts"></a>Problembehandlung Office Skripts

Wenn Sie skripts Office entwickeln, können Sie Fehler machen. Das ist okay. Sie verfügen über die Tools, um die Probleme zu finden und die Skripts optimal zu funktionieren.

## <a name="types-of-errors"></a>Fehlertypen

Office Skriptfehler fallen in eine von zwei Kategorien:

* Kompilierungszeitfehler oder -warnungen
* Laufzeitfehler

### <a name="compile-time-errors"></a>Kompilierungszeitfehler

Kompilierungsfehler und Warnungen werden zunächst im Code-Editor angezeigt. Diese werden durch die wellenförmigen roten Unterstreichungen im Editor angezeigt. Sie werden auch unten  im Aufgabenbereich des Code-Editors unter der Registerkarte Probleme angezeigt. Wenn Sie den Fehler auswählen, erhalten Sie weitere Details zum Problem und schlagen Lösungen vor. Kompilierungsfehler sollten vor dem Ausführen des Skripts behoben werden.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im #A0 angezeigt wird":::

Möglicherweise werden auch orangefarbene Warnmeldungen und graue Informationsmeldungen angezeigt. Diese zeigen Leistungsvorschläge oder andere Möglichkeiten an, bei denen das Skript unbeabsichtigte Auswirkungen haben kann. Solche Warnungen sollten vor dem Schließen genau geprüft werden.

### <a name="runtime-errors"></a>Laufzeitfehler

Laufzeitfehler können aufgrund von Logikproblemen im Skript auftreten. Dies kann daran liegen, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet oder eine andere geringfügige Diskrepanz zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe besteht. Das folgende Skript generiert einen Fehler, wenn ein Arbeitsblatt mit dem Namen "TestSheet" nicht vorhanden ist.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Konsolennachrichten

Sowohl Kompilierungszeit- als auch Laufzeitfehler zeigen Fehlermeldungen in der Konsole an, wenn ein Skript ausgeführt wird. Sie geben eine Zeilennummer, in der das Problem aufgetreten ist. Denken Sie daran, dass die Hauptursache eines Problems möglicherweise eine andere Codezeile als die in der Konsole angegebenen ist.

Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` Compilerfehler.](../develop/typescript-restrictions.md) Notieren Sie sich `[5, 16]` den Text am Anfang der Fehlerzeichenfolge. Dies gibt an, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Die Code-Editor-Konsole, in der eine explizite Fehlermeldung &quot;any&quot; angezeigt wird":::

In der folgenden Abbildung wird die Konsolenausgabe für einen Laufzeitfehler angezeigt. Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen. Notieren Sie sich erneut die Zeile 2 vor dem Fehler, um zu zeigen, welche Zeile untersucht werden soll.
:::image type="content" source="../images/runtime-error-console.png" alt-text="Die Code-Editor-Konsole, in der ein Fehler aus dem Aufruf &quot;addWorksheet&quot; angezeigt wird":::

## <a name="console-logs"></a>Konsolenprotokolle

Drucken von Nachrichten auf dem Bildschirm mit der `console.log` Anweisung. In diesen Protokollen können Sie den aktuellen Wert von Variablen oder die Codepfade anzeigen, die ausgelöst werden. Rufen Sie dazu ein `console.log` beliebiges Objekt als Parameter auf. In der Regel `string` ist a der einfachste Typ, der in der Konsole gelesen werden kann.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

An übergebene Zeichenfolgen werden in der Protokollierungskonsole des Code-Editors am unteren Rand `console.log` des Aufgabenbereichs angezeigt. Protokolle werden auf der Registerkarte **Ausgabe** gefunden, die Registerkarte erhält jedoch automatisch den Fokus, wenn ein Protokoll geschrieben wird.

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Automatisieren der Registerkarte nicht angezeigt oder Office Skripts nicht verfügbar

Die folgenden Schritte sollten bei der Problembehandlung bei Problemen im Zusammenhang mit der Registerkarte **Automatisieren** helfen, die nicht in der Excel im Web.

1. [Stellen Sie sicher, Microsoft 365 Ihre Lizenz skripts Office enthält.](../overview/excel.md#requirements)
1. [Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)
1. [Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind.](platform-limits.md#third-party-cookies)
1. [Stellen Sie sicher, dass Ihr Administrator die Skripts im Office Admin Center Microsoft 365 deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Behandeln von Skripts in Power Automate

Informationen zum Ausführen von Skripts über Power Automate finden Sie unter [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bei Codierungsproblemen helfen. Häufig können Sie die Lösung für Ihr Problem über eine schnelle Stack Overflow-Suche finden. Wenn nicht, stellen Sie Ihre Frage, und markieren Sie sie mit dem Tag "office-scripts". Achten Sie darauf, zu erwähnen, dass Sie ein *Office-Skript* erstellen, nicht ein *Office-Add-In .*

Wenn Sie eine Featureanforderung für Office Skripts übermitteln möchten, posten Sie Ihre Idee auf unserer [User Voice-Seite,](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)oder wenn die Featureanforderung bereits vorhanden ist, fügen Sie Ihre Stimme dafür hinzu. Stellen Sie sicher, dass Sie die Anforderung unter Excel für das Web in der Kategorie "Makros, Skripts und Add-Ins" speichern.

Wenn ein Problem mit der Aktionsaufzeichnung oder dem Editor vor liegt, teilen Sie uns dies bitte mit. Wählen Sie im Menü Code-Editor-Aufgabenbereich **...** die **Schaltfläche** Feedback senden aus, um Probleme zu teilen.

:::image type="content" source="../images/code-editor-feedback.png" alt-text="Das Überlaufmenü des Code-Editors mit der Schaltfläche &quot;Feedback senden&quot;":::

## <a name="see-also"></a>Sehen Sie ebenfalls

- [Bewährte Methoden in Office-Skripts](../develop/best-practices.md)
- [Plattformbeschränkungen mit Office Skripts](platform-limits.md)
- [Verbessern der Leistung Ihrer Office Skripts](../develop/web-client-performance.md)
- [Problembehandlung Office in PowerAutomate ausgeführten Skripts](power-automate-troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
