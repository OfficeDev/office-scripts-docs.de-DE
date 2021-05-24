---
title: Office Skripts-Code-Editor-Umgebung
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
# <a name="office-scripts-code-editor-environment"></a>Office Skripts-Code-Editor-Umgebung

Office Skripts werden entweder in TypeScript oder JavaScript geschrieben und verwenden die Office Skripts-JavaScript-APIs, um mit einer Excel interagieren. Der Code-Editor basiert auf Visual Studio Code. Wenn Sie diese Umgebung bereits verwendet haben, fühlen Sie sich wie zu Hause.

## <a name="scripting-language-typescript-or-javascript"></a>Skriptsprache: TypeScript oder JavaScript

Office Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html)geschrieben, das eine Übersatz von [JavaScript ist.](https://developer.mozilla.org/docs/Web/JavaScript) Der Aktionsrecorder generiert Code in TypeScript, Office Skriptdokumentation TypeScript verwendet. Da TypeScript eine Übersatz von JavaScript ist, funktioniert jeder Skriptcode, den Sie in JavaScript schreiben, einfach gut.

Office Skripts sind weitgehend eigenständige Codeteile. Es wird nur ein kleiner Teil der Funktionalität von TypeScript verwendet. Daher können Sie Skripts bearbeiten, ohne die Feinheiten von TypeScript erlernen zu müssen. Der Code-Editor behandelt auch die Installation, Kompilierung und Ausführung von Code, sodass Sie sich keine Gedanken über das Skript selbst machen müssen. Es ist möglich, die Sprache zu erlernen und Skripts ohne vorherige Programmierkenntnisse zu erstellen. Wenn Sie jedoch neu in der Programmierung sind, empfehlen wir, einige Grundlagen zu erlernen, bevor Sie mit den Office fortfahren:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Skripts-JavaScript-API

Office Skripts verwenden eine spezielle Version der Office JavaScript-APIs für [Office Add-Ins](/office/dev/add-ins/overview/index). Obwohl es Ähnlichkeiten in den beiden APIs gibt, sollten Sie nicht davon ausgehen, dass Code zwischen den beiden Plattformen portiert werden kann. Die Unterschiede zwischen den beiden Plattformen werden im Artikel Differences [between Office Scripts and Office Add-Ins](../resources/add-ins-differences.md#apis) beschrieben. Sie können alle APIs, die Für Ihr Skript verfügbar sind, in der Office [Skripts-API-Referenzdokumentation anzeigen.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Unterstützung für externe Bibliotheken

Office Skripts unterstützen die Verwendung externer JavaScript-Bibliotheken von Drittanbietern nicht. Derzeit können Sie keine andere Bibliothek als die Office Skript-APIs aus einem Skript aufrufen. Sie haben weiterhin Zugriff auf jedes integrierte [JavaScript-Objekt,](../develop/javascript-objects.md)z. B. [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense ist ein Code-Editor-Feature, das hilft, Tippfehler und Syntaxfehler beim Bearbeiten des Skripts zu verhindern. Bei der Eingabe werden mögliche Objekt- und Feldnamen sowie eine Inlinedokumentation für jede API angezeigt.

Der Excel-Code-Editor verwendet dasselbe IntelliSense-Modul wie Visual Studio Code. Weitere Informationen zum Feature finden Sie unter [Visual Studio Code IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Tastenkombinationen

Die meisten Tastenkombinationen für Visual Studio Code funktionieren auch im Office Skriptcode-Editor. Verwenden Sie die folgenden PDFs, um mehr über die verfügbaren Optionen zu erfahren und das Beste aus dem Code-Editor zu nutzen:

- [Tastenkombinationen für macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Tastenkombinationen für Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Siehe auch

- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md)
