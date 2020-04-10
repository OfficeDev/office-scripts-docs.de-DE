---
title: Code-Editor-Umgebung für Office-Skripts
description: Die Voraussetzungen und Umgebungsinformationen für Office-Skripts in Excel im Internet.
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: 6b26adf886172f085980bed0488b4aa7a6815991
ms.sourcegitcommit: b13dedb5ee2048f0a244aa2294bf2c38697cb62c
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215266"
---
# <a name="office-scripts-code-editor-environment"></a>Code-Editor-Umgebung für Office-Skripts

Office-Skripts werden entweder in Skript [oder JavaScript](#scripting-language-typescript-or-javascript) geschrieben und verwenden die [JavaScript-APIs für Office-Skripts](#office-scripts-javascript-api) , um mit einer Excel-Arbeitsmappe zu interagieren.

## <a name="scripting-language-typescript-or-javascript"></a>Skriptsprache: Manuskript oder JavaScript

Office-Skripts [sind in Skript](https://www.typescriptlang.org/docs/home.html) oder [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript)geschrieben. Die aktionsaufzeichnung generiert Code in der Datei "Scripting" (eine Obermenge von JavaScript). In der Dokumentation für Office-Skripts wird die Schreibweise verwendet, wenn Sie sich aber besser mit JavaScript vertraut machen, können Sie diese stattdessen verwenden.

Office-Skripts sind weitgehend eigenständige Codeabschnitte. Nur ein kleiner Teil der Funktionalität von "Manuskript" wird verwendet. Daher können Sie Skripts bearbeiten, ohne die Feinheiten von "Scripte" kennen zu müssen. Der Code-Editor behandelt auch die Installation, Kompilierung und Ausführung von Code, sodass Sie sich keine Gedanken über alles andere als das Skript selbst machen müssen. Es ist möglich, die Sprache zu erlernen und Skripts ohne vorherige Programmierkenntnisse zu erstellen. Wenn Sie jedoch noch nicht mit der Programmierung vertraut sind, empfehlen wir Ihnen, einige Grundlagen zu lernen, bevor Sie mit Office-Skripts fortfahren:

- Erfahren Sie mehr über die Grundlagen von JavaScript. Sie sollten sich mit Konzepten wie Variablen, Ablaufsteuerung, Funktionen und Datentypen wohl fühlen. [Mozilla bietet ein gutes, umfassendes Tutorial zu JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Erfahren Sie mehr über Typen in der Texteingabe. Scripting baut auf JavaScript auf, indem sichergestellt wird, dass zur Kompilierzeit die richtigen Typen für Methodenaufrufe und-Zuweisungen verwendet werden. Die Dokumentation zu [Schnittstellen](https://www.typescriptlang.org/docs/handbook/interfaces.html), [Klassen](https://www.typescriptlang.org/docs/handbook/classes.html), [Typrückschluss](https://www.typescriptlang.org/docs/handbook/type-inference.html)und [Typkompatibilität](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) ist die nützlichste.

## <a name="office-scripts-javascript-api"></a>JavaScript-API für Office-Skripts

Office-Skripts verwenden eine spezielle Version der Office-JavaScript-APIs, die von [Office-Add-ins](/office/dev/add-ins/overview/index)verwendet werden. Die Unterschiede zwischen den beiden Plattformen werden in den [unterschieden zwischen Office-Skripts und Office-Add-ins](../resources/add-ins-differences.md#apis) Artikel beschrieben. Sie können alle verfügbaren APIs für Ihr Skript in der [Office Scripts-API-Referenzdokumentation](/javascript/api/office-scripts/overview)anzeigen.

## <a name="intellisense"></a>IntelliSense

IntelliSense ist ein Code-Editor-Feature, das bei der Bearbeitung Ihres Skripts hilft, Tippfehler und Syntaxfehler zu vermeiden. Es werden mögliche Objekt-und Feldnamen während der Eingabe sowie eine Inline Dokumentation für jede API angezeigt.

Der Excel-Code-Editor verwendet das gleiche IntelliSense-Modul wie Visual Studio Code. Weitere Informationen zum Feature finden Sie in den [IntelliSense-Funktionen von Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="external-library-support"></a>Unterstützung für externe Bibliotheken

Office-Skripts unterstützen nicht die Verwendung externer JavaScript-Bibliotheken von Drittanbietern. Sie können derzeit keine andere Bibliothek als die Office Scripts-APIs aus einem Skript aufrufen. Sie haben weiterhin Zugriff auf ein [integriertes JavaScript-Objekt](../develop/javascript-objects.md), beispielsweise [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="see-also"></a>Siehe auch

- [Referenz der Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md)
