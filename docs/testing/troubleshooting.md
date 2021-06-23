---
title: Problembehandlung bei Office Skripts
description: Debugtipps und -techniken für Office Skripts sowie Hilferessourcen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074648"
---
# <a name="troubleshoot-office-scripts"></a>Problembehandlung bei Office Skripts

Wenn Sie Office Skripts entwickeln, machen Sie möglicherweise Fehler. Das ist okay. Sie verfügen über die Tools, mit denen Sie die Probleme finden und Ihre Skripts einwandfrei funktionieren können.

## <a name="types-of-errors"></a>Fehlertypen

Office Skriptfehler werden in eine von zwei Kategorien unterteilt:

* Kompilierungszeitfehler oder Warnungen
* Laufzeitfehler

### <a name="compile-time-errors"></a>Kompilierungszeitfehler

Fehler und Warnungen zur Kompilierungszeit werden zunächst im Code-Editor angezeigt. Diese werden von den wellenförmigen roten Unterstreichungen im Editor angezeigt. Sie werden auch unter der Registerkarte **"Probleme"** am unteren Rand des Code-Editor-Aufgabenbereichs angezeigt. Wenn Sie den Fehler auswählen, erhalten Sie weitere Details zu dem Problem und schlagen Lösungen vor. Kompilierungszeitfehler sollten behoben werden, bevor das Skript ausgeführt wird.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im Hovertext des Code-Editors angezeigt wird.":::

Möglicherweise werden auch orangefarbene Warnhinweise und graue Informationsmeldungen angezeigt. Diese deuten auf Leistungsvorschläge oder andere Möglichkeiten hin, bei denen das Skript unbeabsichtigte Effekte haben kann. Solche Warnungen sollten sorgfältig geprüft werden, bevor sie geschlossen werden.

### <a name="runtime-errors"></a>Laufzeitfehler

Laufzeitfehler treten aufgrund von Logikproblemen im Skript auf. Dies kann darauf zurückzuführen sein, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet, oder eine andere geringfügige Abweichung zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe. Das folgende Skript generiert einen Fehler, wenn kein Arbeitsblatt mit dem Namen "TestSheet" vorhanden ist.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Konsolenmeldungen

Bei Kompilierungs- und Laufzeitfehlern werden Fehlermeldungen in der Konsole angezeigt, wenn ein Skript ausgeführt wird. Sie geben eine Zeilennummer an, in der das Problem aufgetreten ist. Beachten Sie, dass die Ursache eines Problems möglicherweise eine andere Codezeile als die in der Konsole angegebene ist.

Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` ](../develop/typescript-restrictions.md) Compilerfehler. Notieren Sie sich den Text `[5, 16]` am Anfang der Fehlerzeichenfolge. Dies weist darauf hin, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="In der Code-Editor-Konsole wird eine explizite Any-Fehlermeldung angezeigt.":::

Die folgende Abbildung zeigt die Konsolenausgabe für einen Laufzeitfehler. Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen. Beachten Sie erneut die "Zeile 2" vor dem Fehler, um anzuzeigen, welche Zeile untersucht werden soll.
:::image type="content" source="../images/runtime-error-console.png" alt-text="In der Code-Editor-Konsole wird ein Fehler aus dem Aufruf von &quot;addWorksheet&quot; angezeigt.":::

## <a name="console-logs"></a>Konsolenprotokolle

Drucken Sie Nachrichten mit der Anweisung auf dem `console.log` Bildschirm. In diesen Protokollen können Sie den aktuellen Wert von Variablen anzeigen oder welche Codepfade ausgelöst werden. Rufen Sie dazu `console.log` ein beliebiges Objekt als Parameter auf. In der Regel ist a `string` der einfachste Typ, der in der Konsole gelesen werden kann.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Übergebene Zeichenfolgen `console.log` werden in der Protokollierungskonsole des Code-Editors am unteren Rand des Aufgabenbereichs angezeigt. Protokolle werden auf der Registerkarte **"Ausgabe"** gefunden, obwohl die Registerkarte automatisch den Fokus erhält, wenn ein Protokoll geschrieben wird.

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Registerkarte "Automatisieren" wird nicht angezeigt, oder Office Skripts nicht verfügbar

Die folgenden Schritte sollten ihnen helfen, Probleme im Zusammenhang mit der Registerkarte **"Automatisieren"** zu beheben, die nicht in Excel im Web angezeigt wird.

1. [Stellen Sie sicher, dass Ihre Microsoft 365-Lizenz Office Skripts enthält.](../overview/excel.md#requirements)
1. [Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)
1. [Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind.](platform-limits.md#third-party-cookies)
1. [Stellen Sie sicher, dass Ihr Administrator Office Skripts im Microsoft 365 Admin Center nicht deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Problembehandlung bei Skripts in Power Automate

Informationen zum Ausführen von Skripts über Power Automate finden Sie unter [Problembehandlung bei Office Skripts,](power-automate-troubleshooting.md)die in Power Automate ausgeführt werden.

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bei Codierungsproblemen helfen möchten. Häufig können Sie die Lösung für Ihr Problem mithilfe einer schnellen Stack Overflow-Suche finden. Wenn nicht, stellen Sie Ihre Frage, und markieren Sie sie mit dem Tag "office-scripts". Erwähnen Sie unbedingt, dass Sie ein Office *Skript* erstellen, nicht ein *Office-Add-In.*

Um eine Featureanforderung für Office Skripts zu übermitteln, veröffentlichen Sie Ihre Idee auf unserer [User Voice-Seite,](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)oder wenn die Featureanforderung bereits vorhanden ist, fügen Sie Ihre Stimme dafür hinzu. Stellen Sie sicher, dass Sie die Anforderung unter Excel für das Web in der Kategorie "Makros, Skripts und Add-Ins" ablegen.

Wenn ein Problem mit dem Action Recorder oder Editor vorliegt, teilen Sie uns dies bitte mit. Wählen Sie im Aufgabenbereich des Code-Editors **im Menü ...** die Schaltfläche **"Feedback senden"** aus, um Probleme zu teilen.

:::image type="content" source="../images/code-editor-feedback.png" alt-text="Das Code-Editor-Überlaufmenü mit der Schaltfläche &quot;Feedback senden&quot;.":::

## <a name="see-also"></a>Siehe auch

- [Bewährte Methoden in Office-Skripts](../develop/best-practices.md)
- [Plattformbeschränkungen mit Office Skripts](platform-limits.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Problembehandlung bei Office Skripts, die in PowerAutomate ausgeführt werden](power-automate-troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
