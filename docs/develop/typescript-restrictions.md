---
title: TypeScript-Einschränkungen in Office Skripts
description: Die Besonderheiten des TypeScript-Compilers und linter, die vom Office Scripts Code Editor verwendet werden.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545019"
---
# <a name="typescript-restrictions-in-office-scripts"></a>TypeScript-Einschränkungen in Office Skripts

Office Skripts verwenden die TypeScript-Sprache. In den meisten Teilen funktioniert jeder TypeScript- oder JavaScript-Code in Office Skripts. Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer arbeitsmappen arbeitsmappen Excel funktioniert.

## <a name="no-any-type-in-office-scripts"></a>Kein 'any'-Typ in Office Skripts

[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können. Office Skripts erfordert jedoch, dass eine Variable nicht [vom Typ "](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)sein kann. Sowohl explizit als auch implizit `any` sind in Office Skripts nicht zulässig. Diese Fälle werden als Fehler gemeldet.

### <a name="explicit-any"></a>explizit `any`

Sie können eine Variable nicht explizit als Typ `any` in Office Skripts (d. `let someVariable: any;` h. ) deklarieren. Der Typ verursacht Probleme, wenn er `any` von Excel verarbeitet wird. Ein Muss z. B. `Range` wissen, dass ein Wert eine `string` , `number` oder `boolean` ist. Sie erhalten einen Kompilierungsfehler (ein Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ im Skript definiert `any` ist.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Meldung im Hover-Text des Code-Editors":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite &quot;any&quot;-Fehler im Konsolenfenster":::

Im vorherigen Screenshot `[5, 16] Explicit Any is not allowed` wird angegeben, dass die Zeile #5, spalte #16 `any` den Typ definiert. Dies hilft Ihnen, den Fehler zu finden.

Um dieses Problem zu umgehen, definieren Sie immer den Typ der Variablen. Wenn Sie sich über den Typ einer Variablen nicht sicher sind, können Sie einen [Union-Typ](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)verwenden. Dies kann für Variablen nützlich sein, die `Range` Werte enthalten, die vom Typ `string` , oder `number` `boolean` (der Typ für Werte ist eine Vereinigung `Range` dieser: `string | number | boolean` ) sein.

### <a name="implicit-any"></a>Implizite `any`

TypeScript-Variablentypen können [implizit](https://www.typescriptlang.org/docs/handbook/type-inference.html) definiert werden. Wenn der TypeScript-Compiler nicht in der Lage ist, den Typ einer Variablen zu bestimmen (entweder weil der Typ nicht explizit definiert ist oder Typrückschluss nicht möglich ist), ist dies eine implizite `any` Und Sie erhalten einen Kompilierungsfehler.

Der häufigste Fall für `any` implizite ist in einer Variablendeklaration, z. `let value;` B. . Es gibt zwei Möglichkeiten, dies zu vermeiden:

* Weisen Sie die Variable einem implizit identifizierbaren Typ ( `let value = 5;` oder `let value = workbook.getWorksheet();` ) zu.
* Geben Sie explizit die Variable ein ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Kein Erb- Office Skriptklassen oder -schnittstellen

Klassen und Schnittstellen, die in Ihrem Office Skript erstellt werden, können Office Skriptklassen oder -schnittstellen nicht [erweitert oder implementiert](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) werden. Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterschnittstellen haben.

## <a name="incompatible-typescript-functions"></a>Inkompatible TypeScript-Funktionen

Office Skript-APIs können nicht in den folgenden Optionen verwendet werden:

* [Generatorfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` wird nicht unterstützt

Die [JavaScript-Eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.

## <a name="restricted-identifers"></a>Eingeschränkte Identifikatoren

Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden. Es handelt sich um reservierte Bedingungen.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Nur Pfeilfunktionen in Array-Rückrufen

Ihre Skripts können [nur Pfeilfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) verwenden, wenn Callbackargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitstellen. Sie können keine Art von Bezeichner oder "traditionelle" Funktion an diese Methoden übergeben.

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

Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Code-Editors gibt Warnungen aus, wenn das Skript Leistungsprobleme haben könnte. Die Fälle und die Umgehung der Fälle werden unter [Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)dokumentiert.

## <a name="external-api-calls"></a>Externe API-Aufrufe

Weitere Informationen finden Sie [unter Unterstützung für externe API-Aufrufe in Office Skripts.](external-calls.md)

## <a name="see-also"></a>Siehe auch

* [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
* [Verbessern Sie die Leistung Ihrer Office Scripts](web-client-performance.md)
