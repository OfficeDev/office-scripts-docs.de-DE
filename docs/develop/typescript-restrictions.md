---
title: TypeScript-Einschränkungen in Office Skripts
description: 'Die Spezifischen des TypeScript-Compilers und Linters, die vom #A0 Office skripts-Code-Editor verwendet werden.'
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: 449a8abbcfdcfde53d0c9b96106f73259de368b1
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639871"
---
# <a name="typescript-restrictions-in-office-scripts"></a>TypeScript-Einschränkungen in Office Skripts

Office Skripts verwenden die TypeScript-Sprache. In den meisten Versionen funktioniert typeScript- oder JavaScript-Code in Office Skripts. Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie mit Ihrer Arbeitsmappe Excel ist.

## <a name="no-any-type-in-office-scripts"></a>Kein "beliebiger" Typ in Office Skripts

[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können. Für Office Skripts ist jedoch erforderlich, dass eine Variable keine vom [Typ sein darf.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Sowohl explizit als auch implizit `any` sind in skripts Office zulässig. Diese Fälle werden als Fehler gemeldet.

### <a name="explicit-any"></a>Explicit `any`

Sie können eine Variable nicht explizit als Typ in Office `any` (d. h. `let value: any;` ) deklarieren. Der `any` Typ verursacht Probleme bei der Verarbeitung durch Excel. Beispielsweise muss ein wissen, dass `Range` ein Wert ein , oder `string` `number` `boolean` ist. Sie erhalten einen Kompilierungszeitfehler (ein Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ `any` im Skript definiert ist.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite &quot;any&quot;-Fehler im Konsolenfenster":::

Im vorherigen Screenshot wird angegeben, dass #2, #14 `[2, 14] Explicit Any is not allowed` Typ `any` definiert wird. Auf diese Weise können Sie den Fehler ermitteln.

Um dieses Problem zu beheben, definieren Sie immer den Typ der Variablen. Wenn Sie unsicher sind, welche Art von Variable Sie haben, können Sie einen [Union-Typ verwenden.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Dies kann für Variablen nützlich sein, die Werte enthalten, die vom Typ , oder sein können (der Typ für Werte ist eine `Range` `string` Vereinigung der `number` `boolean` `Range` werte: `string | number | boolean` ).

### <a name="implicit-any"></a>Implizit `any`

TypeScript-Variablentypen können [implizit definiert](https://www.typescriptlang.org/docs/handbook/type-inference.html) werden. Wenn der TypeScript-Compiler den Typ einer Variablen nicht ermitteln kann (entweder, weil typ nicht explizit definiert ist oder typinferenz nicht möglich ist), handelt es sich um einen impliziten Fehler, und Sie erhalten einen Kompilierungszeitfehler. `any`

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Die implizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors":::

Der häufigste Fall für implizite `any` Deklarationen ist eine variable Deklaration, z. B. `let value;` . Es gibt zwei Möglichkeiten, dies zu vermeiden:

* Weisen Sie die Variable einem implizit identifizierbaren Typ zu ( `let value = 5;` oder `let value = workbook.getWorksheet();` ).
* Geben Sie die Variable explizit ein ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Keine Vererbung Office Skriptklassen oder -schnittstellen

Klassen und Schnittstellen, die in Ihrem Skript erstellt Office [Skripts](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) können keine Skriptklassen oder -schnittstellen Office erweitern oder implementieren. Anders ausgedrückt: Nichts im `ExcelScript` Namespace kann Unterklassen oder Unterwebsites enthalten.

## <a name="incompatible-typescript-functions"></a>Inkompatible TypeScript-Funktionen

Office Skript-APIs können nicht wie folgt verwendet werden:

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

Der Linter des [Code-Editors](https://wikipedia.org/wiki/Lint_(software)) gibt Warnungen an, wenn das Skript Leistungsprobleme haben kann. Die Fälle und deren Funktionsweise sind unter Verbessern der Leistung Ihrer Skripts [Office dokumentiert.](web-client-performance.md)

## <a name="external-api-calls"></a>Externe API-Aufrufe

Weitere Informationen finden Sie unter Unterstützung für externe [OFFICE in Skripts.](external-calls.md)

## <a name="see-also"></a>Siehe auch

* [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
* [Verbessern der Leistung Ihrer Office Skripts](web-client-performance.md)
