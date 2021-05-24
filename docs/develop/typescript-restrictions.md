---
title: TypeScript-Einschränkungen in Office Skripts
description: 'Die Spezifischen des TypeScript-Compilers und Linters, die vom #A0 Office skripts-Code-Editor verwendet werden.'
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545019"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="7cff9-103">TypeScript-Einschränkungen in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="7cff9-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="7cff9-104">Office Skripts verwenden die TypeScript-Sprache.</span><span class="sxs-lookup"><span data-stu-id="7cff9-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="7cff9-105">In den meisten Versionen funktioniert typeScript- oder JavaScript-Code in Office Skripts.</span><span class="sxs-lookup"><span data-stu-id="7cff9-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="7cff9-106">Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie mit Ihrer Arbeitsmappe Excel ist.</span><span class="sxs-lookup"><span data-stu-id="7cff9-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="7cff9-107">Kein "beliebiger" Typ in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="7cff9-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="7cff9-108">[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können.</span><span class="sxs-lookup"><span data-stu-id="7cff9-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="7cff9-109">Für Office Skripts ist jedoch erforderlich, dass eine Variable keine vom [Typ sein darf.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="7cff9-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="7cff9-110">Sowohl explizit als auch implizit `any` sind in skripts Office zulässig.</span><span class="sxs-lookup"><span data-stu-id="7cff9-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="7cff9-111">Diese Fälle werden als Fehler gemeldet.</span><span class="sxs-lookup"><span data-stu-id="7cff9-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="7cff9-112">Explicit `any`</span><span class="sxs-lookup"><span data-stu-id="7cff9-112">Explicit `any`</span></span>

<span data-ttu-id="7cff9-113">Sie können eine Variable nicht explizit als Typ in Office `any` (d. h. `let someVariable: any;` ) deklarieren.</span><span class="sxs-lookup"><span data-stu-id="7cff9-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="7cff9-114">Der `any` Typ verursacht Probleme bei der Verarbeitung durch Excel.</span><span class="sxs-lookup"><span data-stu-id="7cff9-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="7cff9-115">Beispielsweise muss ein wissen, dass `Range` ein Wert ein , oder `string` `number` `boolean` ist.</span><span class="sxs-lookup"><span data-stu-id="7cff9-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="7cff9-116">Sie erhalten einen Kompilierungszeitfehler (ein Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ `any` im Skript definiert ist.</span><span class="sxs-lookup"><span data-stu-id="7cff9-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite &quot;any&quot;-Fehler im Konsolenfenster":::

<span data-ttu-id="7cff9-119">Im vorherigen Screenshot `[5, 16] Explicit Any is not allowed` wird angegeben, dass #5, #16 Typ definiert `any` wird.</span><span class="sxs-lookup"><span data-stu-id="7cff9-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="7cff9-120">Auf diese Weise können Sie den Fehler ermitteln.</span><span class="sxs-lookup"><span data-stu-id="7cff9-120">This helps you locate the error.</span></span>

<span data-ttu-id="7cff9-121">Um dieses Problem zu beheben, definieren Sie immer den Typ der Variablen.</span><span class="sxs-lookup"><span data-stu-id="7cff9-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="7cff9-122">Wenn Sie unsicher sind, welche Art von Variable Sie haben, können Sie einen [Union-Typ verwenden.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="7cff9-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="7cff9-123">Dies kann für Variablen nützlich sein, die Werte enthalten, die vom Typ , oder sein können (der Typ für Werte ist eine `Range` `string` Vereinigung der `number` `boolean` `Range` werte: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="7cff9-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="7cff9-124">Implizit `any`</span><span class="sxs-lookup"><span data-stu-id="7cff9-124">Implicit `any`</span></span>

<span data-ttu-id="7cff9-125">TypeScript-Variablentypen können [implizit definiert](https://www.typescriptlang.org/docs/handbook/type-inference.html) werden.</span><span class="sxs-lookup"><span data-stu-id="7cff9-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="7cff9-126">Wenn der TypeScript-Compiler den Typ einer Variablen nicht ermitteln kann (entweder, weil typ nicht explizit definiert ist oder typinferenz nicht möglich ist), handelt es sich um einen impliziten Fehler, und Sie erhalten einen Kompilierungszeitfehler. `any`</span><span class="sxs-lookup"><span data-stu-id="7cff9-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="7cff9-127">Der häufigste Fall für implizite `any` Deklarationen ist eine variable Deklaration, z. B. `let value;` .</span><span class="sxs-lookup"><span data-stu-id="7cff9-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="7cff9-128">Es gibt zwei Möglichkeiten, dies zu vermeiden:</span><span class="sxs-lookup"><span data-stu-id="7cff9-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="7cff9-129">Weisen Sie die Variable einem implizit identifizierbaren Typ zu ( `let value = 5;` oder `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="7cff9-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="7cff9-130">Geben Sie die Variable explizit ein ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="7cff9-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="7cff9-131">Keine Vererbung Office Skriptklassen oder -schnittstellen</span><span class="sxs-lookup"><span data-stu-id="7cff9-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="7cff9-132">Klassen und Schnittstellen, die in Ihrem Skript erstellt Office [Skripts](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) können keine Skriptklassen oder -schnittstellen Office erweitern oder implementieren.</span><span class="sxs-lookup"><span data-stu-id="7cff9-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="7cff9-133">Anders ausgedrückt: Nichts im `ExcelScript` Namespace kann Unterklassen oder Unterwebsites enthalten.</span><span class="sxs-lookup"><span data-stu-id="7cff9-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="7cff9-134">Inkompatible TypeScript-Funktionen</span><span class="sxs-lookup"><span data-stu-id="7cff9-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="7cff9-135">Office Skript-APIs können nicht wie folgt verwendet werden:</span><span class="sxs-lookup"><span data-stu-id="7cff9-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="7cff9-136">Generatorfunktionen</span><span class="sxs-lookup"><span data-stu-id="7cff9-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="7cff9-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="7cff9-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="7cff9-138">`eval` wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="7cff9-138">`eval` is not supported</span></span>

<span data-ttu-id="7cff9-139">Die JavaScript [eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="7cff9-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="7cff9-140">Eingeschränkte Identiferen</span><span class="sxs-lookup"><span data-stu-id="7cff9-140">Restricted identifers</span></span>

<span data-ttu-id="7cff9-141">Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="7cff9-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="7cff9-142">Es handelt sich um reservierte Bedingungen.</span><span class="sxs-lookup"><span data-stu-id="7cff9-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="7cff9-143">Nur Pfeilfunktionen in Arrayrückrufen</span><span class="sxs-lookup"><span data-stu-id="7cff9-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="7cff9-144">Ihre Skripts können nur [Pfeilfunktionen verwenden,](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) wenn Sie Rückrufargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="7cff9-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="7cff9-145">Sie können keine Art von Bezeichner oder herkömmliche Funktion an diese Methoden übergeben.</span><span class="sxs-lookup"><span data-stu-id="7cff9-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="7cff9-146">Leistungswarnungen</span><span class="sxs-lookup"><span data-stu-id="7cff9-146">Performance warnings</span></span>

<span data-ttu-id="7cff9-147">Der Linter des [Code-Editors](https://wikipedia.org/wiki/Lint_(software)) gibt Warnungen an, wenn das Skript Leistungsprobleme haben kann.</span><span class="sxs-lookup"><span data-stu-id="7cff9-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="7cff9-148">Die Fälle und deren Funktionsweise sind unter Verbessern der Leistung Ihrer Skripts [Office dokumentiert.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="7cff9-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="7cff9-149">Externe API-Aufrufe</span><span class="sxs-lookup"><span data-stu-id="7cff9-149">External API calls</span></span>

<span data-ttu-id="7cff9-150">Weitere Informationen finden Sie unter Unterstützung für externe [OFFICE in Skripts.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="7cff9-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="7cff9-151">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="7cff9-151">See also</span></span>

* [<span data-ttu-id="7cff9-152">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="7cff9-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="7cff9-153">Verbessern der Leistung Ihrer Office Skripts</span><span class="sxs-lookup"><span data-stu-id="7cff9-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
