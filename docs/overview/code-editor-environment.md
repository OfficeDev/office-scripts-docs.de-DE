---
title: Code-Editor-Umgebung für Office-Skripts
description: Die Voraussetzungen und Umgebungsinformationen für Office-Skripts in Excel im Web.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891253"
---
# <a name="office-scripts-code-editor-environment"></a>Code-Editor-Umgebung für Office-Skripts

Office-Skripts werden entweder in TypeScript oder JavaScript geschrieben und verwenden die JavaScript-APIs von Office-Skripts, um mit einer Excel-Arbeitsmappe zu interagieren. Der Code-Editor basiert auf Visual Studio Code. Wenn Sie diese Umgebung also bereits verwendet haben, werden Sie sich wie zu Hause fühlen.

> [!TIP]
> Wenn Sie mit Visual Studio Code vertraut sind, können Sie es jetzt zum Schreiben von Skripts verwenden. Besuchen Sie [Visual Studio Code für Office-Skripts (Vorschau),](../develop/vscode-for-scripts.md) um dieses Feature auszuprobieren.

## <a name="scripting-language-typescript-or-javascript"></a>Skriptsprache: TypeScript oder JavaScript

Office-Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html) geschrieben, einer Obermenge von [JavaScript-](https://developer.mozilla.org/docs/Web/JavaScript). Der Action Recorder generiert Code in TypeScript, und die Dokumentation zu Office-Skripts verwendet TypeScript. Da TypeScript eine Obermenge von JavaScript ist, funktioniert jeder Skriptcode, den Sie in JavaScript schreiben, problemlos.

Office-Skripts sind weitgehend eigenständige Codeteile. Nur ein kleiner Teil der TypeScript-Funktionalität wird verwendet. Daher können Sie Skripts bearbeiten, ohne die Feinheiten von TypeScript erlernen zu müssen. Der Code-Editor übernimmt auch die Installation, Kompilierung und Ausführung von Code, sodass Sie sich nur um das Skript selbst kümmern müssen. Es ist möglich, die Sprache zu erlernen und Skripts ohne vorherige Programmierkenntnisse zu erstellen. Wenn Sie jedoch noch nicht mit der Programmierung vertraut sind, empfehlen wir Ihnen, einige Grundlagen zu erlernen, bevor Sie mit Office-Skripts fortfahren:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>JavaScript-API für Office-Skripts

Office-Skripts verwenden eine spezielle Version der Office [JavaScript-APIs für Office-Add-Ins](/office/dev/add-ins/overview/index). Obwohl es Ähnlichkeiten in den beiden APIs gibt, sollten Sie nicht davon ausgehen, dass Code zwischen den beiden Plattformen portiert werden kann. Die Unterschiede zwischen den beiden Plattformen werden im Artikel [Unterschiede zwischen Office-Skripts und Office-Add-Ins](../resources/add-ins-differences.md#apis) beschrieben. Sie können alle für Ihr Skript verfügbaren APIs in der [Referenzdokumentation zur Office-Skript-API](/javascript/api/office-scripts/overview) anzeigen.

## <a name="external-library-support"></a>Unterstützung externer Bibliotheken

Office-Skripts unterstützen nicht die Verwendung externer JavaScript-Bibliotheken von Drittanbietern. Derzeit können Sie keine andere Bibliothek als die Office-Skript-APIs aus einem Skript aufrufen. Sie haben weiterhin Zugriff auf alle [integrierten JavaScript-Objekte](../develop/javascript-objects.md), z. B. [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

IntelliSense ist eine Reihe von Code-Editor-Features, die Ihnen beim Schreiben von Code helfen. Es bietet automatische Vervollständigung, Syntaxfehlermarkierung und Inline-API-Dokumentation.

IntelliSense gibt Während der Eingabe Vorschläge, ähnlich dem vorgeschlagenen Text in Excel. Wenn Sie die TAB- oder EINGABETASTE drücken, wird das vorgeschlagene Element eingefügt. Lösen Sie IntelliSense an der aktuellen Cursorposition aus, indem Sie STRG+LEERTASTEn drücken. Diese Vorschläge sind besonders nützlich, wenn eine Methode abgeschlossen wird. Die von IntelliSense angezeigte Methodensignatur enthält eine Liste der benötigten Argumente, den Typ jedes Arguments, unabhängig davon, ob ein bestimmtes Argument erforderlich oder optional ist, und den Rückgabetyp der Methode.

Zeigen Sie mit dem Cursor auf eine Methode, Klasse oder ein anderes Codeobjekt, um weitere Informationen anzuzeigen. Zeigen Sie auf einen Syntaxfehler oder Codevorschlag, dargestellt durch eine rote oder gelbe Wellenlinie, um Vorschläge zum Beheben des Problems anzuzeigen. IntelliSense bietet häufig die Option "Schnelle Fehlerbehebung", um den Code automatisch zu ändern.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Eine Fehlermeldung im Daraufzeigertext des Code-Editors mit einer Schaltfläche &quot;Schnellkorrektur&quot;.":::

Der Code-Editor für Office-Skripts verwendet dieselbe IntelliSense-Engine wie Visual Studio Code. Weitere Informationen zum Feature finden Sie unter [IntelliSense-Features von Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Tastenkombinationen

Die meisten Tastenkombinationen für Visual Studio Code funktionieren auch im Code-Editor für Office-Skripts. Verwenden Sie die folgenden PDF-Dateien, um mehr über die verfügbaren Optionen zu erfahren und den Code-Editor optimal zu nutzen:

- [Tastenkombinationen für macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Tastenkombinationen für Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Siehe auch

- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md)
- [Visual Studio Code für Office-Skripts (Vorschau)](../develop/vscode-for-scripts.md)
