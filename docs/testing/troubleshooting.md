---
title: Problembehandlung bei Office-Skripts
description: Debuggen von Tipps und Techniken für Office-Skripts sowie Hilferessourcen.
ms.date: 10/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fe4a9b17d51d078403d1a46abed774d38eeaa80
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495467"
---
# <a name="troubleshoot-office-scripts"></a>Problembehandlung bei Office-Skripts

Beim Entwickeln von Office-Skripts können Sie Fehler machen. Das ist okay. Sie verfügen über die Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen können.

> [!NOTE]
> Informationen zur Problembehandlung speziell für Office-Skripts mit Power Automate finden Sie unter [Problembehandlung bei Office-Skripts, die in Power Automate ausgeführt werden](power-automate-troubleshooting.md).

## <a name="types-of-errors"></a>Arten von Fehlern

Office-Skriptfehler fallen in eine von zwei Kategorien:

* Fehler oder Warnungen zur Kompilierzeit
* Laufzeitfehler

### <a name="compile-time-errors"></a>Fehler bei der Kompilierungszeit

Fehler und Warnungen zur Kompilierungszeit werden zunächst im Code-Editor angezeigt. Diese werden durch die wellenförmigen roten Unterstreichungen im Editor angezeigt. Sie werden auch unter der Registerkarte **"Probleme** " am unteren Rand des Code-Editor-Aufgabenbereichs angezeigt. Wenn Sie den Fehler auswählen, werden weitere Details zum Problem angezeigt und Lösungen vorgeschlagen. Fehler bei der Kompilierungszeit sollten vor dem Ausführen des Skripts behoben werden.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im Hovertext des Code-Editors angezeigt wird.":::

Möglicherweise werden auch orangefarbene Warnungsunterstriche und graue Informationsmeldungen angezeigt. Diese zeigen Leistungsvorschläge oder andere Möglichkeiten an, bei denen das Skript unbeabsichtigte Effekte haben kann. Solche Warnungen sollten vor dem Verwerfen genau geprüft werden.

### <a name="runtime-errors"></a>Laufzeitfehler

Laufzeitfehler treten aufgrund von Logikproblemen im Skript auf. Dies kann darauf zurückzuführen sein, dass sich ein im Skript verwendete Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet oder eine andere geringfügige Abweichung zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe besteht. Das folgende Skript generiert einen Fehler, wenn kein Arbeitsblatt mit dem Namen "TestSheet" vorhanden ist.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Konsolennachrichten

Sowohl Kompilierzeitfehler als auch Laufzeitfehler zeigen Fehlermeldungen in der Konsole an, wenn ein Skript ausgeführt wird. Sie geben eine Zeilennummer an, in der das Problem aufgetreten ist. Denken Sie daran, dass die Ursache eines Problems möglicherweise eine andere Codezeile als die in der Konsole angegebene ist.

Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any`](../develop/typescript-restrictions.md) Compilerfehler. Notieren Sie sich den Text `[5, 16]` am Anfang der Fehlerzeichenfolge. Dies gibt an, dass sich der Fehler in Zeile 5 befindet, beginnend bei Zeichen 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Die Code-Editor-Konsole mit einer expliziten &quot;any&quot;-Fehlermeldung.":::

Die folgende Abbildung zeigt die Konsolenausgabe für einen Laufzeitfehler. Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen. Beachten Sie auch hier die "Zeile 2" vor dem Fehler, um anzuzeigen, welche Zeile untersucht werden soll.
:::image type="content" source="../images/runtime-error-console.png" alt-text="Die Code-Editor-Konsole mit einem Fehler aus dem Aufruf von &quot;addWorksheet&quot;.":::

## <a name="console-logs"></a>Konsolenprotokolle

Drucken Sie Nachrichten mit der Anweisung auf den `console.log` Bildschirm. In diesen Protokollen können Sie den aktuellen Wert von Variablen anzeigen oder welche Codepfade ausgelöst werden. Rufen Sie dazu `console.log` ein beliebiges Objekt als Parameter auf. Normalerweise ist a `string` der einfachste Typ, der in der Konsole gelesen werden kann.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

An übergebene `console.log` Zeichenfolgen werden in der Protokollierungskonsole des Code-Editors am unteren Rand des Aufgabenbereichs angezeigt. Protokolle finden Sie auf der Registerkarte " **Ausgabe** ", obwohl die Registerkarte automatisch den Fokus erhält, wenn ein Protokoll geschrieben wird.

Protokolle wirken sich nicht auf die Arbeitsmappe aus.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Die Registerkarte "Automatisieren" wird nicht angezeigt, oder Office-Skripts sind nicht verfügbar

Die folgenden Schritte sollten bei der Behandlung von Problemen im Zusammenhang mit der Registerkarte **"Automatisieren**" helfen, die in Excel im Web nicht angezeigt werden.

1. [Stellen Sie sicher, dass Ihre Microsoft 365-Lizenz Office-Skripts enthält](../overview/excel.md#requirements).
1. [Überprüfen Sie, ob Ihr Browser unterstützt wird](platform-limits.md#browser-support).
1. [Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind](platform-limits.md#third-party-cookies).
1. [Stellen Sie sicher, dass Ihr Administrator Office-Skripts in der Microsoft 365 Admin Center nicht deaktiviert hat](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. Stellen Sie sicher, dass Sie nicht als externer oder Gastbenutzer bei Ihrem Mandanten angemeldet sind.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

> [!NOTE]
> Es gibt ein bekanntes Problem, das verhindert, dass in SharePoint gespeicherte Skripts immer in der zuletzt verwendeten Liste angezeigt werden. Dies tritt auf, wenn Ihr Administrator Exchange-Webdienste (EWS) deaktiviert. Auf Ihre SharePoint-basierten Skripts kann weiterhin über das Dateidialogfeld zugegriffen und verwendet werden.

## <a name="help-resources"></a>Hilferessourcen

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungsproblemen zu helfen. Häufig können Sie die Lösung für Ihr Problem über eine schnelle Stack Overflow-Suche finden. Wenn nicht, stellen Sie Ihre Frage, und markieren Sie sie mit dem Tag "office-scripts". Erwähnen Sie unbedingt, dass Sie ein *Office-Skript* erstellen, nicht ein *Office-Add-In*.

## <a name="see-also"></a>Siehe auch

- [Bewährte Methoden in Office-Skripts](../develop/best-practices.md)
- [Plattformbeschränkungen mit Office-Skripts](platform-limits.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Problembehandlung bei Office-Skripts, die in PowerAutomate ausgeführt werden](power-automate-troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
