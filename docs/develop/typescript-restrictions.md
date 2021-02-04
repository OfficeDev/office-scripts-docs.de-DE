---
title: Einschränkungen für TypeScript in Office Scripts
description: 'Die Vom Office Scripts Code Editor verwendeten #A0 und -Linter.'
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099902"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="7a8b9-103">Einschränkungen für TypeScript in Office Scripts</span><span class="sxs-lookup"><span data-stu-id="7a8b9-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="7a8b9-104">Office Scripts verwenden die TypeScript-Sprache.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="7a8b9-105">In den meisten Beispielen funktioniert jeder TypeScript- oder JavaScript-Code in einem Office-Skript.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="7a8b9-106">Es gibt jedoch einige Einschränkungen, die vom Code-Editor erzwungen werden, um sicherzustellen, dass Ihr Skript konsistent und wie beabsichtigt mit Ihrer Excel-Arbeitsmappe funktioniert.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="7a8b9-107">Kein "beliebiger" Typ in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7a8b9-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="7a8b9-108">[Schreibtypen](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sind in TypeScript optional, da die Typen abgeleitet werden können.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="7a8b9-109">Office Script erfordert jedoch, dass eine Variable keinen Typ [haben darf.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="7a8b9-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="7a8b9-110">Sowohl explizit als auch implizit `any` sind in einem Office Script nicht zulässig.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="7a8b9-111">Diese Fälle werden als Fehler gemeldet.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="7a8b9-112">Explizit `any`</span><span class="sxs-lookup"><span data-stu-id="7a8b9-112">Explicit `any`</span></span>

<span data-ttu-id="7a8b9-113">Sie können eine Variable nicht explizit als Typ `any` in office-Skripts deklarieren (d. h. `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="7a8b9-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="7a8b9-114">Der `any` Typ verursacht Probleme bei der Verarbeitung durch Excel.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="7a8b9-115">Ein Muss z. `Range` B. wissen, dass ein Wert ein `string` , oder `number` `boolean` ist.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="7a8b9-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![Die explizite Beliebige Nachricht im Hovertext des Code-Editors](../images/explicit-any-editor-message.png)

![Der explizite Fehler im Konsolenfenster](../images/explicit-any-error-message.png)

<span data-ttu-id="7a8b9-119">Im obigen Screenshot wird angegeben, dass #5, #16 `[5, 16] Explicit Any is not allowed` Typ `any` definiert.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="7a8b9-120">Auf diese Weise können Sie den Fehler ermitteln.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-120">This helps you locate the error.</span></span>

<span data-ttu-id="7a8b9-121">Um dieses Problem zu beheben, definieren Sie immer den Typ der Variablen.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="7a8b9-122">Wenn Sie unsicher sind, welche Art von Variable Sie haben, können Sie einen [Vereinigungstyp verwenden.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="7a8b9-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="7a8b9-123">Dies kann für Variablen nützlich sein, die Werte enthalten, die vom Typ , oder (der Typ für Werte ist eine Vereinigung `Range` `string` der `number` `boolean` `Range` werte: ) sein können. `string | number | boolean`</span><span class="sxs-lookup"><span data-stu-id="7a8b9-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="7a8b9-124">Implizit `any`</span><span class="sxs-lookup"><span data-stu-id="7a8b9-124">Implicit `any`</span></span>

<span data-ttu-id="7a8b9-125">TypeScript variable types can be [implicitly](( https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-125">TypeScript variable types can be [implicitly]((https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="7a8b9-126">Wenn der #A0 den Typ einer Variablen nicht ermitteln kann (entweder weil der Typ nicht explizit definiert ist oder der Typverweis nicht möglich ist), handelt es sich um einen impliziten Fehler, und Sie erhalten einen Kompilierungszeitfehler. `any`</span><span class="sxs-lookup"><span data-stu-id="7a8b9-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="7a8b9-127">Der häufigste Fall für implizite `any` Deklarationen ist eine Variablendeklaration, z. B. `let value;` .</span><span class="sxs-lookup"><span data-stu-id="7a8b9-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="7a8b9-128">Es gibt zwei Möglichkeiten, dies zu vermeiden:</span><span class="sxs-lookup"><span data-stu-id="7a8b9-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="7a8b9-129">Weisen Sie die Variable einem implizit identifizierbaren Typ ( `let value = 5;` oder `let value = workbook.getWorksheet();` ) zu.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="7a8b9-130">Geben Sie die Variable explizit ein ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="7a8b9-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="7a8b9-131">Keine Vererbung von Office -Skriptklassen oder -Schnittstellen</span><span class="sxs-lookup"><span data-stu-id="7a8b9-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="7a8b9-132">Klassen und Schnittstellen, die in Ihrem Office Script erstellt [werden,](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) können keine Klassen oder Schnittstellen von Office Scripts erweitern oder implementieren.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="7a8b9-133">Mit anderen Worten, nichts im `ExcelScript` Namespace kann Unterklassen oder Unterwebsites enthalten.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="7a8b9-134">Inkompatible TypeScript-Funktionen</span><span class="sxs-lookup"><span data-stu-id="7a8b9-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="7a8b9-135">Office-Skript-APIs können nicht in den folgenden Beispielen verwendet werden:</span><span class="sxs-lookup"><span data-stu-id="7a8b9-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="7a8b9-136">Generatorfunktionen</span><span class="sxs-lookup"><span data-stu-id="7a8b9-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="7a8b9-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="7a8b9-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="7a8b9-138">`eval` wird nicht unterstützt</span><span class="sxs-lookup"><span data-stu-id="7a8b9-138">`eval` is not supported</span></span>

<span data-ttu-id="7a8b9-139">Die [JavaScript-eval-Funktion](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) wird aus Sicherheitsgründen nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="7a8b9-140">Eingeschränkte Identitäten</span><span class="sxs-lookup"><span data-stu-id="7a8b9-140">Restricted identifers</span></span>

<span data-ttu-id="7a8b9-141">Die folgenden Wörter können nicht als Bezeichner in einem Skript verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="7a8b9-142">Es handelt sich um reservierte Begriffe.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="7a8b9-143">Leistungswarnungen</span><span class="sxs-lookup"><span data-stu-id="7a8b9-143">Performance warnings</span></span>

<span data-ttu-id="7a8b9-144">Der [Linter](https://wikipedia.org/wiki/Lint_(software)) des Codeeditors gibt Warnungen aus, wenn beim Skript Leistungsprobleme auftreten können.</span><span class="sxs-lookup"><span data-stu-id="7a8b9-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="7a8b9-145">Die Fälle und deren Umarbeitung sind unter "Verbessern der Leistung Ihrer [Office-Skripts" dokumentiert.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="7a8b9-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="7a8b9-146">Externe API-Aufrufe</span><span class="sxs-lookup"><span data-stu-id="7a8b9-146">External API calls</span></span>

<span data-ttu-id="7a8b9-147">Weitere [Informationen finden Sie unter "Externe API-Anrufunterstützung in Office-Skripts".](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="7a8b9-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="7a8b9-148">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="7a8b9-148">See also</span></span>

* [<span data-ttu-id="7a8b9-149">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="7a8b9-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="7a8b9-150">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7a8b9-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
