---
title: Office Skript-Code-Editor-Umgebung
description: Die Voraussetzungen und Umgebungsinformationen für Office Skripts in Excel im Web.
ms.date: 05/27/2021
localization_priority: Normal
ms.openlocfilehash: 5b2f7afa193dc71e13a3d6763c9e8ff8344ee3e8be18e7e996f8431e03510509
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847093"
---
# <a name="office-scripts-code-editor-environment"></a>Office Skript-Code-Editor-Umgebung

Office Skripts werden entweder in TypeScript oder JavaScript geschrieben und verwenden die JavaScript-APIs Office Skripts, um mit einer Excel Arbeitsmappe zu interagieren. Der Code-Editor basiert auf Visual Studio Code. Wenn Sie diese Umgebung bereits verwendet haben, fühlen Sie sich also wie zu Hause.

## <a name="scripting-language-typescript-or-javascript"></a>Skriptsprache: TypeScript oder JavaScript

Office-Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html) geschrieben, einer Obermenge von [JavaScript-](https://developer.mozilla.org/docs/Web/JavaScript). Der Action Recorder generiert Code in TypeScript, und die Dokumentation Office Skripts verwendet TypeScript. Da TypeScript eine Obermenge von JavaScript ist, funktioniert jeder Skriptcode, den Sie in JavaScript schreiben, einwandfrei.

Office Skripts sind größtenteils eigenständige Codeelemente. Nur ein kleiner Teil der TypeScript-Funktionalität wird verwendet. Daher können Sie Skripts bearbeiten, ohne die Feinheiten von TypeScript erlernen zu müssen. Der Code-Editor übernimmt auch die Installation, Kompilierung und Ausführung von Code, sodass Sie sich keine Gedanken über etwas außer das Skript selbst machen müssen. Es ist möglich, die Sprache zu erlernen und Skripts ohne vorherige Programmierkenntnisse zu erstellen. Wenn Sie jedoch noch nicht mit der Programmierung vertraut sind, empfehlen wir Ihnen, einige Grundlagen zu erlernen, bevor Sie mit Office Skripts fortfahren:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office JavaScript-API für Skripts

Office Skripts verwenden eine spezielle Version der Office JavaScript-APIs für [Office-Add-Ins.](/office/dev/add-ins/overview/index) Obwohl es Ähnlichkeiten in den beiden APIs gibt, sollten Sie nicht davon ausgehen, dass Code zwischen den beiden Plattformen portiert werden kann. Die Unterschiede zwischen den beiden Plattformen werden im Artikel ["Unterschiede zwischen Office Skripts und Office-Add-Ins"](../resources/add-ins-differences.md#apis) beschrieben. Sie können alle APIs, die Ihrem Skript zur Verfügung stehen, in der Referenzdokumentation zur [Office Skripts-API](/javascript/api/office-scripts/overview)anzeigen.

## <a name="external-library-support"></a>Unterstützung externer Bibliotheken

Office Skripts unterstützen nicht die Verwendung externer JavaScript-Bibliotheken von Drittanbietern. Derzeit können Sie keine andere Bibliothek als die Office Skript-APIs aus einem Skript aufrufen. Sie haben weiterhin Zugriff auf ein [beliebiges integriertes JavaScript-Objekt,](../develop/javascript-objects.md)z. B. [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

IntelliSense ist eine Reihe von Code-Editor-Features, mit denen Sie Code schreiben können. Es bietet dokumentation zur automatischen Vervollständigung, zur Syntaxfehlerhervorhebung und zur Inline-API.

IntelliSense bietet während der Eingabe Vorschläge, ähnlich wie der vorgeschlagene Text in Excel. Durch Drücken der TAB- oder EINGABETASTE wird das vorgeschlagene Element eingefügt. Trigger IntelliSense at the current cursor location by pressing the CTrl+Space keys. Diese Vorschläge sind besonders hilfreich, wenn Sie eine Methode abschließen. Die von IntelliSense angezeigte Methodensignatur enthält eine Liste der benötigten Argumente, den Typ jedes Arguments, ob ein bestimmtes Argument erforderlich oder optional ist, und den Rückgabetyp der Methode.

Zeigen Sie mit dem Mauszeiger über eine Methode, Klasse oder ein anderes Codeobjekt, um weitere Informationen anzuzeigen. Zeigen Sie mit dem Mauszeiger über einen Syntaxfehler oder Codevorschlag, dargestellt durch eine rote oder gelbe Wellenlinie, um Vorschläge zur Behebung des Problems anzuzeigen. IntelliSense bietet häufig die Option "Schnellkorrektur", um den Code automatisch zu ändern.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Eine Fehlermeldung im Hovertext des Code-Editors mit der Schaltfläche &quot;Schnellkorrektur&quot;.":::

Der Code-Editor Office Skripts verwendet dasselbe IntelliSense-Modul wie Visual Studio Code. Weitere Informationen zu diesem Feature finden Sie in [den IntelliSense-Features von Visual Studio Code.](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>Tastenkombinationen

Die meisten Tastenkombinationen für Visual Studio Code auch im Code-Editor für Office Skripts funktionieren. Verwenden Sie die folgenden PDFs, um mehr über die verfügbaren Optionen zu erfahren und den Code-Editor optimal zu nutzen:

- [Tastenkombinationen für macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Tastenkombinationen für Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Weitere Artikel

- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md)
