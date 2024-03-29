---
title: TypeScript-Einschränkungen in Office Skripts
description: Die Einzelheiten des TypeScript-Compilers und Linters, die vom Code-Editor für Office Skripts verwendet werden.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ba0dfe60081a0bb65dec4e694c7d534cb8df63
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585681"
---
# <a name="typescript-restrictions-in-office-scripts"></a>TypeScript-Einschränkungen in Office Skripts

Office Skripts verwenden die TypeScript-Sprache. In den meisten Fällen funktioniert jeder TypeScript- oder JavaScript-Code in Office Skripts. Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer Excel Arbeitsmappe funktioniert.

## <a name="no-any-type-in-office-scripts"></a>Kein "any"-Typ in Office Skripts

Das Schreiben [von Typen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) ist in TypeScript optional, da die Typen abgeleitet werden können. Office Skripts erfordert jedoch, dass eine Variable keinen [Typ aufweisen](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) kann. Explizite und implizite `any` Skripts sind in Office Skripts nicht zulässig. Diese Fälle werden als Fehler gemeldet.

### <a name="explicit-any"></a>Explizit `any`

Sie können eine Variable nicht explizit als Typ `any` in Office Skripts deklarieren (d. b`let value: any;`. ). Der `any` Typ verursacht Probleme, wenn er von Excel verarbeitet wird. Beispielsweise muss ein `Range` Wert ein `string`, `number`oder `boolean`wissen. Sie erhalten einen Kompilierungszeitfehler (einen Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als `any` Typ im Skript definiert ist.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite Fehler &quot;any&quot; im Konsolenfenster.":::

Im vorherigen Screenshot `[2, 14] Explicit Any is not allowed` wird der Typ in Zeile 2, Spalte 14, definiert `any` . Dies hilft Ihnen bei der Suche nach dem Fehler.

Um dieses Problem zu umgehen, definieren Sie immer den Typ der Variablen. Wenn Sie hinsichtlich des Typs einer Variablen unsicher sind, können Sie einen [Vereinigungstyp](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) verwenden. Dies kann für Variablen nützlich sein, die Werte enthalten `Range` , die vom Typ `string`sein können, `number`oder `boolean` (der Typ für `Range` Werte ist eine Vereinigung dieser Werte: `string | number | boolean`).

### <a name="implicit-any"></a>Implizite `any`

TypeScript-Variablentypen können [implizit](https://www.typescriptlang.org/docs/handbook/type-inference.html) definiert werden. Wenn der TypeScript-Compiler den Typ einer Variablen nicht ermitteln kann (entweder weil der Typ nicht explizit definiert ist oder keine Typrückleitung möglich ist), ist dies ein impliziter `any` Fehler, und Sie erhalten einen Kompilierungszeitfehler.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Die implizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors.":::

Der häufigste Fall für implizite `any` Variablen ist eine Variablendeklaration, z `let value;`. B. . Es gibt zwei Möglichkeiten, dies zu vermeiden:

* Weisen Sie die Variable einem implizit identifizierbaren Typ (`let value = 5;` oder `let value = workbook.getWorksheet();`) zu.
* Geben Sie die Variable explizit ein (`let value: number;`)

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Kein Erben Office Skriptklassen oder -schnittstellen

Klassen und Schnittstellen, die in Ihrem Office Skript erstellt werden, können Office Skripts-Klassen oder -Schnittstellen nicht [erweitern oder implementieren](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance). Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterinterfaces enthalten.

## <a name="incompatible-typescript-functions"></a>Inkompatible TypeScript-Funktionen

Office Skript-APIs können nicht in folgenden Komponenten verwendet werden:

* [Generatorfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` wird nicht unterstützt

Die [JavaScript-Eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.

## <a name="restricted-identifiers"></a>Eingeschränkte Bezeichner

Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden. Es handelt sich um reservierte Begriffe.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Nur Pfeilfunktionen in Arrayrückrufen

Ihre Skripts können [Pfeilfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) nur verwenden, wenn Rückrufargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitgestellt werden. Sie können keine Bezeichner- oder "herkömmlichen" Funktionen an diese Methoden übergeben.

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>Verbindungen von `ExcelScript` Typen und benutzerdefinierten Typen werden nicht unterstützt.

Office Skripts werden zur Laufzeit von synchronen in asynchrone Codeblöcke konvertiert. Die Kommunikation mit der Arbeitsmappe über [Zusagen](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) wird dem Skriptersteller verborgen. Diese Konvertierung unterstützt keine [Vereinigungstypen](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) , die Typen und benutzerdefinierte Typen enthalten `ExcelScript` . In diesem Fall wird das `Promise` Skript zurückgegeben, aber der Office Skriptcompiler erwartet es nicht, und der Ersteller des Skripts kann nicht mit dem `Promise`Skript interagieren.

Das folgende Codebeispiel zeigt eine nicht unterstützte Vereinigung zwischen `ExcelScript.Table` und einer benutzerdefinierten `MyTable` Schnittstelle.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>Konstruktoren unterstützen Office Skript-APIs und `console` -Anweisungen nicht

`console`Anweisungen und viele Office Skript-APIs erfordern eine Synchronisierung mit der Excel Arbeitsmappe. Diese Synchronisierungen verwenden `await` Anweisungen in kompilierter Laufzeitversion des Skripts. `await` wird in [Konstruktoren](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor) nicht unterstützt. Wenn Sie Klassen mit Konstruktoren benötigen, vermeiden Sie die Verwendung Office Skript-APIs oder `console` -Anweisungen in diesen Codeblöcken.

Im folgenden Codebeispiel wird dieses Szenario veranschaulicht. Es wird ein Fehler generiert, der besagt `failed to load [code] [library]`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>Leistungswarnungen

Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Code-Editors gibt Warnungen aus, wenn das Skript Leistungsprobleme haben kann. Die Fälle und wie Sie diese umgehen, sind in ["Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)" dokumentiert.

## <a name="external-api-calls"></a>Externe API-Aufrufe

Weitere Informationen finden Sie [unter Support für externe API-Aufrufe in Office Skripts](external-calls.md).

## <a name="see-also"></a>Siehe auch

* [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
* [Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)
