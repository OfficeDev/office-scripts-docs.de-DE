---
title: Verwalten des Berechnungsmodus in Excel
description: Erfahren Sie, wie Sie Office Skripts verwenden, um den Berechnungsmodus in Excel im Web.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285913"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="49ec4-103">Verwalten des Berechnungsmodus in Excel</span><span class="sxs-lookup"><span data-stu-id="49ec4-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="49ec4-104">In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) verwenden und Methoden in Excel im Web mithilfe Office berechnen.</span><span class="sxs-lookup"><span data-stu-id="49ec4-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="49ec4-105">Sie können das Skript in einer beliebigen Excel testen.</span><span class="sxs-lookup"><span data-stu-id="49ec4-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="49ec4-106">Szenario</span><span class="sxs-lookup"><span data-stu-id="49ec4-106">Scenario</span></span>

<span data-ttu-id="49ec4-107">Arbeitsmappen mit einer großen Anzahl von Formeln können eine Weile dauern, bis sie neu berechnet werden.</span><span class="sxs-lookup"><span data-stu-id="49ec4-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="49ec4-108">Anstatt Excel steuern zu lassen, wenn Berechnungen ausgeführt werden, können Sie sie als Teil Ihres Skripts verwalten.</span><span class="sxs-lookup"><span data-stu-id="49ec4-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="49ec4-109">Dies hilft bei der Leistung in bestimmten Szenarien.</span><span class="sxs-lookup"><span data-stu-id="49ec4-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="49ec4-110">Das Beispielskript legt den Berechnungsmodus auf manuell fest.</span><span class="sxs-lookup"><span data-stu-id="49ec4-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="49ec4-111">Dies bedeutet, dass die Arbeitsmappe formeln nur neu berechnet, wenn das Skript dies ansah (oder Sie manuell über die [Benutzeroberfläche berechnen).](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)</span><span class="sxs-lookup"><span data-stu-id="49ec4-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="49ec4-112">Das Skript zeigt dann den aktuellen Berechnungsmodus an und berechnet die gesamte Arbeitsmappe vollständig neu.</span><span class="sxs-lookup"><span data-stu-id="49ec4-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="49ec4-113">Beispielcode: Steuerelementberechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="49ec4-113">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="49ec4-114">Schulungsvideo: Verwalten des Berechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="49ec4-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="49ec4-115">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/iw6O8QH01CI)</span><span class="sxs-lookup"><span data-stu-id="49ec4-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
