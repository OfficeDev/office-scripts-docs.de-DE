---
title: TypeScript-Einschränkungen in Office-Skripts
description: Die Spezifischen des TypeScript-Compilers und Linters, die vom Office Scripts Code Editor verwendet werden.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 8c9d1beafb236e7ba10dedf00fab944c40fb954d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570276"
---
# <a name="typescript-restrictions-in-office-scripts"></a>TypeScript-Einschränkungen in Office-Skripts

Office-Skripts verwenden die TypeScript-Sprache. In den meisten Beispielen funktioniert typeScript- oder JavaScript-Code in einem Office-Skript. Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie mit Ihrer Excel-Arbeitsmappe beabsichtigt funktioniert.

## <a name="no-any-type-in-office-scripts"></a>Kein "beliebiger" Typ in Office-Skripts

[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können. Office Script erfordert jedoch, dass eine Variable keine vom [Typ sein darf.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Sowohl explizit als auch implizit `any` sind in einem Office-Skript nicht zulässig. Diese Fälle werden als Fehler gemeldet.

### <a name="explicit-any"></a>Explicit `any`

Sie können eine Variable nicht explizit als Typ `any` in Office Scripts deklarieren (d. h. `let someVariable: any;` ). Der `any` Typ verursacht Probleme bei der Verarbeitung durch Excel. Beispielsweise muss ein wissen, dass `Range` ein Wert ein , oder `string` `number` `boolean` ist. Sie erhalten einen Kompilierungszeitfehler (ein Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ `any` im Skript definiert ist.

![Die explizite beliebige Nachricht im Hovertext des Code-Editors](../images/explicit-any-editor-message.png)

![Der explizite Fehler im Konsolenfenster](../images/explicit-any-error-message.png)

Im obigen Screenshot `[5, 16] Explicit Any is not allowed` wird angegeben, dass #5, #16 Typ definiert `any` wird. Auf diese Weise können Sie den Fehler ermitteln.

Um dieses Problem zu beheben, definieren Sie immer den Typ der Variablen. Wenn Sie unsicher sind, welche Art von Variable Sie haben, können Sie einen [Union-Typ verwenden.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Dies kann für Variablen nützlich sein, die Werte enthalten, die vom Typ , oder sein können (der Typ für Werte ist eine `Range` `string` Vereinigung der `number` `boolean` `Range` werte: `string | number | boolean` ).

### <a name="implicit-any"></a>Implizit `any`

TypeScript-Variablentypen können [implizit definiert](https://www.typescriptlang.org/docs/handbook/type-inference.html) werden. Wenn der TypeScript-Compiler den Typ einer Variablen nicht ermitteln kann (entweder, weil typ nicht explizit definiert ist oder typinferenz nicht möglich ist), handelt es sich um einen impliziten Fehler, und Sie erhalten einen Kompilierungszeitfehler. `any`

Der häufigste Fall für implizite `any` Deklarationen ist eine variable Deklaration, z. B. `let value;` . Es gibt zwei Möglichkeiten, dies zu vermeiden:

* Weisen Sie die Variable einem implizit identifizierbaren Typ zu ( `let value = 5;` oder `let value = workbook.getWorksheet();` ).
* Geben Sie die Variable explizit ein ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Keine erbenden Office Script-Klassen oder -Schnittstellen

Klassen und Schnittstellen, die in Ihrem Office Script erstellt werden, [können](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Keine Office Scripts-Klassen oder -Schnittstellen erweitern oder implementieren. Anders ausgedrückt: Nichts im `ExcelScript` Namespace kann Unterklassen oder Unterwebsites enthalten.

## <a name="incompatible-typescript-functions"></a>Inkompatible TypeScript-Funktionen

Office-Skript-APIs können nicht wie folgt verwendet werden:

* [Generatorfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` wird nicht unterstützt

Die JavaScript [eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.

## <a name="restricted-identifers"></a>Eingeschränkte Identiferen

Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden. Es handelt sich um reservierte Bedingungen.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Nur Pfeilfunktionen in Arrayrückrufen

Ihre Skripts können nur [Pfeilfunktionen verwenden,](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) wenn Sie Rückrufargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitstellen. Sie können keine Art von Bezeichner oder herkömmliche Funktion an diese Methoden übergeben.

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a>Leistungswarnungen

Der Linter des [Code-Editors](https://wikipedia.org/wiki/Lint_(software)) gibt Warnungen an, wenn das Skript Leistungsprobleme haben kann. Die Fälle und ihre Funktionsweise sind unter Verbessern der Leistung Ihrer [Office-Skripts dokumentiert.](web-client-performance.md)

## <a name="external-api-calls"></a>Externe API-Aufrufe

Weitere Informationen finden Sie [unter Unterstützung für externe API-Aufrufe in Office-Skripts.](external-calls.md)

## <a name="see-also"></a>Siehe auch

* [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
* [Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)
