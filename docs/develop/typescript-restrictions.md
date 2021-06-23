---
title: TypeScript-Einschränkungen in Office Skripts
description: Die Einzelheiten des TypeScript-Compilers und Linters, die vom Code-Editor für Office Skripts verwendet werden.
ms.date: 05/24/2021
localization_priority: Normal
ms.openlocfilehash: 0bc6b4c0acaf9bb42f8200a0850dd7254632f965
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074445"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="e32a2-103">TypeScript-Einschränkungen in Office Skripts</span><span class="sxs-lookup"><span data-stu-id="e32a2-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="e32a2-104">Office Skripts verwenden die TypeScript-Sprache.</span><span class="sxs-lookup"><span data-stu-id="e32a2-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="e32a2-105">In den meisten Fällen funktioniert jeder TypeScript- oder JavaScript-Code in Office Skripts.</span><span class="sxs-lookup"><span data-stu-id="e32a2-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="e32a2-106">Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer Excel Arbeitsmappe funktioniert.</span><span class="sxs-lookup"><span data-stu-id="e32a2-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="e32a2-107">Kein "any"-Typ in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="e32a2-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="e32a2-108">Das Schreiben [von Typen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) ist in TypeScript optional, da die Typen abgeleitet werden können.</span><span class="sxs-lookup"><span data-stu-id="e32a2-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="e32a2-109">Office Skripts erfordert jedoch, dass eine Variable keinen [Typ aufweisen](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)kann.</span><span class="sxs-lookup"><span data-stu-id="e32a2-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="e32a2-110">Explizite und implizite `any` Skripts sind in Office Skripts nicht zulässig.</span><span class="sxs-lookup"><span data-stu-id="e32a2-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="e32a2-111">Diese Fälle werden als Fehler gemeldet.</span><span class="sxs-lookup"><span data-stu-id="e32a2-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="e32a2-112">explizit `any`</span><span class="sxs-lookup"><span data-stu-id="e32a2-112">Explicit `any`</span></span>

<span data-ttu-id="e32a2-113">Sie können eine Variable nicht explizit als Typ `any` in Office Skripts deklarieren (d. b. `let value: any;` ).</span><span class="sxs-lookup"><span data-stu-id="e32a2-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let value: any;`).</span></span> <span data-ttu-id="e32a2-114">Der `any` Typ verursacht Probleme, wenn er von Excel verarbeitet wird.</span><span class="sxs-lookup"><span data-stu-id="e32a2-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="e32a2-115">Beispielsweise muss ein `Range` Wert ein , `string` oder `number` `boolean` wissen.</span><span class="sxs-lookup"><span data-stu-id="e32a2-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="e32a2-116">Sie erhalten einen Kompilierungszeitfehler (einen Fehler vor dem Ausführen des Skripts), wenn eine Variable explizit als Typ im Skript definiert `any` ist.</span><span class="sxs-lookup"><span data-stu-id="e32a2-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Die explizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Der explizite Fehler &quot;any&quot; im Konsolenfenster.":::

<span data-ttu-id="e32a2-119">Im vorherigen Screenshot `[2, 14] Explicit Any is not allowed` wird angegeben, dass zeile #2 Spalte #14 `any` Typ definiert.</span><span class="sxs-lookup"><span data-stu-id="e32a2-119">In the previous screenshot, `[2, 14] Explicit Any is not allowed` indicates that line #2, column #14 defines `any` type.</span></span> <span data-ttu-id="e32a2-120">Dies hilft Ihnen bei der Suche nach dem Fehler.</span><span class="sxs-lookup"><span data-stu-id="e32a2-120">This helps you locate the error.</span></span>

<span data-ttu-id="e32a2-121">Um dieses Problem zu umgehen, definieren Sie immer den Typ der Variablen.</span><span class="sxs-lookup"><span data-stu-id="e32a2-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="e32a2-122">Wenn Sie hinsichtlich des Typs einer Variablen unsicher sind, können Sie einen [Vereinigungstyp](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)verwenden.</span><span class="sxs-lookup"><span data-stu-id="e32a2-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="e32a2-123">Dies kann für Variablen nützlich `Range` sein, die Werte enthalten, die vom Typ sein `string` `number` können, oder `boolean` (der Typ für `Range` Werte ist eine Vereinigung dieser Werte: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="e32a2-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="e32a2-124">Implizite `any`</span><span class="sxs-lookup"><span data-stu-id="e32a2-124">Implicit `any`</span></span>

<span data-ttu-id="e32a2-125">TypeScript-Variablentypen können [implizit](https://www.typescriptlang.org/docs/handbook/type-inference.html) definiert werden.</span><span class="sxs-lookup"><span data-stu-id="e32a2-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="e32a2-126">Wenn der TypeScript-Compiler den Typ einer Variablen nicht ermitteln kann (entweder weil der Typ nicht explizit definiert ist oder keine Typrückleitung möglich ist), ist dies ein impliziter `any` Fehler, und Sie erhalten einen Kompilierungszeitfehler.</span><span class="sxs-lookup"><span data-stu-id="e32a2-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Die implizite &quot;any&quot;-Nachricht im Hovertext des Code-Editors.":::

<span data-ttu-id="e32a2-128">Der häufigste Fall für `any` implizite Variablen ist eine Variablendeklaration, z. `let value;` B. .</span><span class="sxs-lookup"><span data-stu-id="e32a2-128">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="e32a2-129">Es gibt zwei Möglichkeiten, dies zu vermeiden:</span><span class="sxs-lookup"><span data-stu-id="e32a2-129">There are two ways to avoid this:</span></span>

* <span data-ttu-id="e32a2-130">Weisen Sie die Variable einem implizit identifizierbaren Typ ( `let value = 5;` oder `let value = workbook.getWorksheet();` ) zu.</span><span class="sxs-lookup"><span data-stu-id="e32a2-130">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="e32a2-131">Geben Sie die Variable explizit ein ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="e32a2-131">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="e32a2-132">Kein Erben Office Skriptklassen oder -schnittstellen</span><span class="sxs-lookup"><span data-stu-id="e32a2-132">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="e32a2-133">Klassen und Schnittstellen, die in Ihrem Office Skript erstellt werden, können Office Skripts-Klassen oder -Schnittstellen nicht [erweitern oder implementieren.](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)</span><span class="sxs-lookup"><span data-stu-id="e32a2-133">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="e32a2-134">Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterinterfaces enthalten.</span><span class="sxs-lookup"><span data-stu-id="e32a2-134">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="e32a2-135">Inkompatible TypeScript-Funktionen</span><span class="sxs-lookup"><span data-stu-id="e32a2-135">Incompatible TypeScript functions</span></span>

<span data-ttu-id="e32a2-136">Office Skript-APIs können in Folgenden nicht verwendet werden:</span><span class="sxs-lookup"><span data-stu-id="e32a2-136">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="e32a2-137">Generatorfunktionen</span><span class="sxs-lookup"><span data-stu-id="e32a2-137">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="e32a2-138">Array.sort</span><span class="sxs-lookup"><span data-stu-id="e32a2-138">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="e32a2-139">`eval` wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="e32a2-139">`eval` is not supported</span></span>

<span data-ttu-id="e32a2-140">Die [JavaScript-Eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="e32a2-140">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="e32a2-141">Eingeschränkte Identifer</span><span class="sxs-lookup"><span data-stu-id="e32a2-141">Restricted identifers</span></span>

<span data-ttu-id="e32a2-142">Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="e32a2-142">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="e32a2-143">Es handelt sich um reservierte Begriffe.</span><span class="sxs-lookup"><span data-stu-id="e32a2-143">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="e32a2-144">Nur Pfeilfunktionen in Arrayrückrufen</span><span class="sxs-lookup"><span data-stu-id="e32a2-144">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="e32a2-145">Ihre Skripts können [Pfeilfunktionen](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) nur verwenden, wenn Rückrufargumente für [Array-Methoden](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) bereitgestellt werden.</span><span class="sxs-lookup"><span data-stu-id="e32a2-145">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="e32a2-146">Sie können keine Bezeichner- oder "herkömmlichen" Funktionen an diese Methoden übergeben.</span><span class="sxs-lookup"><span data-stu-id="e32a2-146">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="e32a2-147">Leistungswarnungen</span><span class="sxs-lookup"><span data-stu-id="e32a2-147">Performance warnings</span></span>

<span data-ttu-id="e32a2-148">Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Code-Editors gibt Warnungen aus, wenn das Skript Leistungsprobleme haben kann.</span><span class="sxs-lookup"><span data-stu-id="e32a2-148">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="e32a2-149">Die Fälle und wie Sie diese umgehen, sind in ["Verbessern der Leistung Ihrer Office Skripts"](web-client-performance.md)dokumentiert.</span><span class="sxs-lookup"><span data-stu-id="e32a2-149">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="e32a2-150">Externe API-Aufrufe</span><span class="sxs-lookup"><span data-stu-id="e32a2-150">External API calls</span></span>

<span data-ttu-id="e32a2-151">Weitere Informationen finden Sie [unter Support für externe API-Aufrufe in Office Skripts.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="e32a2-151">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="e32a2-152">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="e32a2-152">See also</span></span>

* [<span data-ttu-id="e32a2-153">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="e32a2-153">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="e32a2-154">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="e32a2-154">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
