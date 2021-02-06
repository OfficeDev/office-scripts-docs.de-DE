---
title: Einschränkungen für TypeScript in Office Scripts
description: 'Die Vom #A0 für #A1 verwendeten #A2 und Linter.'
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: d67e208561ce6ddd706d4c80cf29d2f013a32032
ms.sourcegitcommit: 98c7bc26f51dc8427669c571135c503d73bcee4c
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 02/06/2021
ms.locfileid: "50125934"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="90c6b-103">Einschränkungen für TypeScript in Office Scripts</span><span class="sxs-lookup"><span data-stu-id="90c6b-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="90c6b-104">Office Scripts verwenden die TypeScript-Sprache.</span><span class="sxs-lookup"><span data-stu-id="90c6b-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="90c6b-105">In den meisten Beispielen funktioniert jeder TypeScript- oder JavaScript-Code in einem Office-Skript.</span><span class="sxs-lookup"><span data-stu-id="90c6b-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="90c6b-106">Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer Excel-Arbeitsmappe funktioniert.</span><span class="sxs-lookup"><span data-stu-id="90c6b-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="90c6b-107">Kein "beliebiger" Typ in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="90c6b-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="90c6b-108">[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können.</span><span class="sxs-lookup"><span data-stu-id="90c6b-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="90c6b-109">Office Script erfordert jedoch, dass eine Variable keinen Typ [haben darf.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="90c6b-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="90c6b-110">Sowohl explizit als auch implizit `any` sind in einem Office-Skript nicht zulässig.</span><span class="sxs-lookup"><span data-stu-id="90c6b-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="90c6b-111">Diese Fälle werden als Fehler gemeldet.</span><span class="sxs-lookup"><span data-stu-id="90c6b-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="90c6b-112">Explizit `any`</span><span class="sxs-lookup"><span data-stu-id="90c6b-112">Explicit `any`</span></span>

<span data-ttu-id="90c6b-113">Sie können eine Variable nicht explizit als Typ `any` in office-Skripts deklarieren (d. h. `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="90c6b-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="90c6b-114">Der `any` Typ verursacht Probleme bei der Verarbeitung durch Excel.</span><span class="sxs-lookup"><span data-stu-id="90c6b-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="90c6b-115">Ein Muss z. `Range` B. wissen, dass ein Wert ein `string` , oder `number` `boolean` ist.</span><span class="sxs-lookup"><span data-stu-id="90c6b-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="90c6b-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span><span class="sxs-lookup"><span data-stu-id="90c6b-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![Die explizite Beliebige Nachricht im Hovertext des Code-Editors](../images/explicit-any-editor-message.png)

![Der explizite Fehler im Konsolenfenster](../images/explicit-any-error-message.png)

<span data-ttu-id="90c6b-119">Im obigen Screenshot wird angegeben, dass #5, #16 `[5, 16] Explicit Any is not allowed` Typ `any` definiert.</span><span class="sxs-lookup"><span data-stu-id="90c6b-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="90c6b-120">Auf diese Weise können Sie den Fehler ermitteln.</span><span class="sxs-lookup"><span data-stu-id="90c6b-120">This helps you locate the error.</span></span>

<span data-ttu-id="90c6b-121">Um dieses Problem zu beheben, definieren Sie immer den Typ der Variablen.</span><span class="sxs-lookup"><span data-stu-id="90c6b-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="90c6b-122">Wenn Sie unsicher sind, welche Art von Variable Sie haben, können Sie einen [Vereinigungstyp verwenden.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="90c6b-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="90c6b-123">Dies kann für Variablen nützlich sein, die Werte enthalten, die vom Typ , oder (der Typ für Werte ist eine Vereinigung der `Range` `string` `number` `boolean` `Range` werte: ) sein können. `string | number | boolean`</span><span class="sxs-lookup"><span data-stu-id="90c6b-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="90c6b-124">Implizit `any`</span><span class="sxs-lookup"><span data-stu-id="90c6b-124">Implicit `any`</span></span>

<span data-ttu-id="90c6b-125">TypeScript-Variablentypen können [implizit definiert](https://www.typescriptlang.org/docs/handbook/type-inference.html) werden.</span><span class="sxs-lookup"><span data-stu-id="90c6b-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="90c6b-126">Wenn der #A0 den Typ einer Variablen nicht ermitteln kann (entweder weil der Typ nicht explizit definiert ist oder der Typverweis nicht möglich ist), handelt es sich um einen impliziten Fehler, und Sie erhalten einen Kompilierungszeitfehler. `any`</span><span class="sxs-lookup"><span data-stu-id="90c6b-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="90c6b-127">Der häufigste Fall bei impliziten `any` Deklarationen ist eine Variablendeklaration, z. B. `let value;` .</span><span class="sxs-lookup"><span data-stu-id="90c6b-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="90c6b-128">Es gibt zwei Möglichkeiten, dies zu vermeiden:</span><span class="sxs-lookup"><span data-stu-id="90c6b-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="90c6b-129">Weisen Sie die Variable einem implizit identifizierbaren Typ ( `let value = 5;` oder `let value = workbook.getWorksheet();` ) zu.</span><span class="sxs-lookup"><span data-stu-id="90c6b-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="90c6b-130">Geben Sie die Variable explizit ein ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="90c6b-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="90c6b-131">Keine Vererbung von Office -Skriptklassen oder -Schnittstellen</span><span class="sxs-lookup"><span data-stu-id="90c6b-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="90c6b-132">Klassen und Schnittstellen, die in Ihrem Office Script erstellt [werden,](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) können keine Klassen oder Schnittstellen von Office Scripts erweitern oder implementieren.</span><span class="sxs-lookup"><span data-stu-id="90c6b-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="90c6b-133">Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterwebsites enthalten.</span><span class="sxs-lookup"><span data-stu-id="90c6b-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="90c6b-134">Inkompatible TypeScript-Funktionen</span><span class="sxs-lookup"><span data-stu-id="90c6b-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="90c6b-135">Office-Skript-APIs können nicht in den folgenden Beispielen verwendet werden:</span><span class="sxs-lookup"><span data-stu-id="90c6b-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="90c6b-136">Generatorfunktionen</span><span class="sxs-lookup"><span data-stu-id="90c6b-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="90c6b-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="90c6b-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="90c6b-138">`eval` wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="90c6b-138">`eval` is not supported</span></span>

<span data-ttu-id="90c6b-139">Die [JavaScript-eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="90c6b-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="90c6b-140">Eingeschränkte Identitäten</span><span class="sxs-lookup"><span data-stu-id="90c6b-140">Restricted identifers</span></span>

<span data-ttu-id="90c6b-141">Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="90c6b-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="90c6b-142">Es handelt sich um reservierte Begriffe.</span><span class="sxs-lookup"><span data-stu-id="90c6b-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="90c6b-143">Leistungswarnungen</span><span class="sxs-lookup"><span data-stu-id="90c6b-143">Performance warnings</span></span>

<span data-ttu-id="90c6b-144">Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Codeeditors gibt Warnungen aus, wenn beim Skript Leistungsprobleme auftreten können.</span><span class="sxs-lookup"><span data-stu-id="90c6b-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="90c6b-145">Die Fälle und deren Umarbeitung sind unter "Verbessern der Leistung Ihrer [Office-Skripts" dokumentiert.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="90c6b-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="90c6b-146">Externe API-Aufrufe</span><span class="sxs-lookup"><span data-stu-id="90c6b-146">External API calls</span></span>

<span data-ttu-id="90c6b-147">Weitere [Informationen finden Sie unter "Externe API-Anrufunterstützung in Office-Skripts".](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="90c6b-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="90c6b-148">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="90c6b-148">See also</span></span>

* [<span data-ttu-id="90c6b-149">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="90c6b-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="90c6b-150">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="90c6b-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
