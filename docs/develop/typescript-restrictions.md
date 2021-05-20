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
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="c6c8c-103">TypeScript-Einschränkungen in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="c6c8c-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="c6c8c-104">Office Skripts verwenden die TypeScript-Sprache.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="c6c8c-105">In den meisten Teilen funktioniert jeder TypeScript- oder JavaScript-Code in Office Skripts.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="c6c8c-106">Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer arbeitsmappen arbeitsmappen Excel funktioniert.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="c6c8c-107">Kein 'any'-Typ in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="c6c8c-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="c6c8c-108">[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="c6c8c-109">Office Skripts erfordert jedoch, dass eine Variable nicht [vom Typ "](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)sein kann.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="c6c8c-110">Sowohl explizit als auch implizit `any` sind in Office Skripts nicht zulässig.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="c6c8c-111">Diese Fälle werden als Fehler gemeldet.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="c6c8c-112">explizit `any`</span><span class="sxs-lookup"><span data-stu-id="c6c8c-112">Explicit `any`</span></span>

<span data-ttu-id="c6c8c-113">Sie können eine Variable nicht explizit als Typ `any` in Office Skripts (d. `let someVariable: any;` h. ) deklarieren.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="c6c8c-114">Der Typ verursacht Probleme, wenn er `any` von Excel verarbeitet wird.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="c6c8c-115">Ein Muss z. B. `Range` wissen, dass ein Wert eine `string` , `number` oder `boolean` ist.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="c6c8c-116">Sie erhalten einen Kompilierungsfehler (ein Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ im Skript definiert `any` ist.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Meldung im Hover-Text des Code-Editors":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite &quot;any&quot;-Fehler im Konsolenfenster":::

<span data-ttu-id="c6c8c-119">Im vorherigen Screenshot `[5, 16] Explicit Any is not allowed` wird angegeben, dass die Zeile #5, spalte #16 `any` den Typ definiert.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="c6c8c-120">Dies hilft Ihnen, den Fehler zu finden.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-120">This helps you locate the error.</span></span>

<span data-ttu-id="c6c8c-121">Um dieses Problem zu umgehen, definieren Sie immer den Typ der Variablen.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="c6c8c-122">Wenn Sie sich über den Typ einer Variablen nicht sicher sind, können Sie einen [Union-Typ](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)verwenden.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="c6c8c-123">Dies kann für Variablen nützlich sein, die `Range` Werte enthalten, die vom Typ `string` , oder `number` `boolean` (der Typ für Werte ist eine Vereinigung `Range` dieser: `string | number | boolean` ) sein.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="c6c8c-124">Implizite `any`</span><span class="sxs-lookup"><span data-stu-id="c6c8c-124">Implicit `any`</span></span>

<span data-ttu-id="c6c8c-125">TypeScript-Variablentypen können [implizit](https://www.typescriptlang.org/docs/handbook/type-inference.html) definiert werden.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="c6c8c-126">Wenn der TypeScript-Compiler nicht in der Lage ist, den Typ einer Variablen zu bestimmen (entweder weil der Typ nicht explizit definiert ist oder Typrückschluss nicht möglich ist), ist dies eine implizite `any` Und Sie erhalten einen Kompilierungsfehler.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="c6c8c-127">Der häufigste Fall für `any` implizite ist in einer Variablendeklaration, z. `let value;` B. .</span><span class="sxs-lookup"><span data-stu-id="c6c8c-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="c6c8c-128">Es gibt zwei Möglichkeiten, dies zu vermeiden:</span><span class="sxs-lookup"><span data-stu-id="c6c8c-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="c6c8c-129">Weisen Sie die Variable einem implizit identifizierbaren Typ ( `let value = 5;` oder `let value = workbook.getWorksheet();` ) zu.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="c6c8c-130">Geben Sie explizit die Variable ein ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="c6c8c-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="c6c8c-131">Kein Erb- Office Skriptklassen oder -schnittstellen</span><span class="sxs-lookup"><span data-stu-id="c6c8c-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="c6c8c-132">Klassen und Schnittstellen, die in Ihrem Office Skript erstellt werden, können Office Skriptklassen oder -schnittstellen nicht [erweitert oder implementiert](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) werden.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="c6c8c-133">Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterschnittstellen haben.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="c6c8c-134">Inkompatible TypeScript-Funktionen</span><span class="sxs-lookup"><span data-stu-id="c6c8c-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="c6c8c-135">Office Skript-APIs können nicht in den folgenden Optionen verwendet werden:</span><span class="sxs-lookup"><span data-stu-id="c6c8c-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="c6c8c-136">Generatorfunktionen</span><span class="sxs-lookup"><span data-stu-id="c6c8c-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="c6c8c-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="c6c8c-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="c6c8c-138">`eval` wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="c6c8c-138">`eval` is not supported</span></span>

<span data-ttu-id="c6c8c-139">Die [JavaScript-Eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="c6c8c-140">Eingeschränkte Identifikatoren</span><span class="sxs-lookup"><span data-stu-id="c6c8c-140">Restricted identifers</span></span>

<span data-ttu-id="c6c8c-141">Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="c6c8c-142">Es handelt sich um reservierte Bedingungen.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="c6c8c-143">Nur Pfeilfunktionen in Array-Rückrufen</span><span class="sxs-lookup"><span data-stu-id="c6c8c-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="c6c8c-144">Ihre Skripts können [nur Pfeilfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) verwenden, wenn Callbackargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="c6c8c-145">Sie können keine Art von Bezeichner oder "traditionelle" Funktion an diese Methoden übergeben.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="c6c8c-146">Leistungswarnungen</span><span class="sxs-lookup"><span data-stu-id="c6c8c-146">Performance warnings</span></span>

<span data-ttu-id="c6c8c-147">Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Code-Editors gibt Warnungen aus, wenn das Skript Leistungsprobleme haben könnte.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="c6c8c-148">Die Fälle und die Umgehung der Fälle werden unter [Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)dokumentiert.</span><span class="sxs-lookup"><span data-stu-id="c6c8c-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="c6c8c-149">Externe API-Aufrufe</span><span class="sxs-lookup"><span data-stu-id="c6c8c-149">External API calls</span></span>

<span data-ttu-id="c6c8c-150">Weitere Informationen finden Sie [unter Unterstützung für externe API-Aufrufe in Office Skripts.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="c6c8c-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="c6c8c-151">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="c6c8c-151">See also</span></span>

* [<span data-ttu-id="c6c8c-152">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="c6c8c-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="c6c8c-153">Verbessern Sie die Leistung Ihrer Office Scripts</span><span class="sxs-lookup"><span data-stu-id="c6c8c-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
