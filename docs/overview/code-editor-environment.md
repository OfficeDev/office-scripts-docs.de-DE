---
title: Office Skriptcode-Editor-Umgebung
description: Die Voraussetzungen und Umgebungsinformationen für Office Skripts in Excel im Web.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545822"
---
# <a name="office-scripts-code-editor-environment"></a>Office Skriptcode-Editor-Umgebung

Office Skripts werden entweder in TypeScript oder JavaScript geschrieben und verwenden die JavaScript-APIs Office Scripts, um mit einer Excel Arbeitsmappe zu interagieren. Der Code-Editor basiert auf Visual Studio Code, also, wenn Sie diese Umgebung bereits verwendet haben, werden Sie sich wie zu Hause fühlen.

## <a name="scripting-language-typescript-or-javascript"></a>Skriptsprache: TypeScript oder JavaScript

Office Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html)geschrieben, einem Übersetaur von [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Der Aktions-Recorder generiert Code in TypeScript, und die Office Scripts-Dokumentation verwendet TypeScript. Da TypeScript eine Obermenge von JavaScript ist, funktioniert jeder Skriptcode, den Sie in JavaScript schreiben, einwandfrei.

Office Skripte sind größtenteils eigenständige Codeteile. Es wird nur ein kleiner Teil der TypeScript-Funktionalität verwendet. Daher können Sie Skripts bearbeiten, ohne die Feinheiten von TypeScript erlernen zu müssen. Der Code-Editor behandelt auch die Installation, Kompilierung und Ausführung von Code, sodass Sie sich um nichts außer dem Skript selbst kümmern müssen. Es ist möglich, die Sprache zu lernen und Skripte ohne vorherige Programmierkenntnisse zu erstellen. Wenn Sie jedoch noch nicht mit der Programmierung zu arbeiten haben, empfehlen wir Ihnen, einige Grundlagen zu lernen, bevor Sie mit Office Skripts fortfahren:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Skripts JavaScript-API

Office Skripts verwenden eine spezielle Version der Office JavaScript-APIs für [Office Add-Ins](/office/dev/add-ins/overview/index). Obwohl die beiden APIs Ähnlichkeiten aufweisen, sollten Sie nicht davon ausgehen, dass Code zwischen den beiden Plattformen portiert werden kann. Die Unterschiede zwischen den beiden Plattformen werden im Artikel ["Unterschiede zwischen Office Skripts" und Office Add-Ins](../resources/add-ins-differences.md#apis) beschrieben. Sie können alle APIs anzeigen, die für Ihr Skript verfügbar sind, in der [Office-SKRIPT-API-Referenzdokumentation](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Externe Bibliotheksunterstützung

Office Skripts unterstützen nicht die Verwendung externer JavaScript-Bibliotheken von Drittanbietern. Derzeit können Sie keine andere Bibliothek als die Office Scripts-APIs aus einem Skript aufrufen. Sie haben weiterhin Zugriff auf alle [integrierten JavaScript-Objekte,](../develop/javascript-objects.md)z. B. [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense ist eine Code-Editor-Funktion, mit der Tippfehler und Syntaxfehler beim Bearbeiten des Skripts verhindert werden können. Es zeigt mögliche Objekt- und Feldnamen während der Eingabe sowie inline Dokumentation für jede API an.

Der Excel Code-Editor verwendet das gleiche IntelliSense-Modul wie Visual Studio Code. Weitere Informationen zu dieser Funktion finden Sie in [Visual Studio Code IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Tastenkombinationen

Die meisten Tastenkombinationen für Visual Studio Code funktionieren auch im code-Editor Office Scripts. Verwenden Sie die folgenden PDFs, um mehr über die verfügbaren Optionen zu erfahren und das Beste aus dem Code-Editor herauszuarbeiten:

- [Tastenkombinationen für macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Tastenkombinationen für Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Siehe auch

- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md)
