---
title: Verwalten des Berechnungsmodus in Excel
description: Erfahren Sie, wie Sie Office Skripts verwenden, um den Berechnungsmodus in Excel im Web.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232452"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="4434c-103">Verwalten des Berechnungsmodus in Excel</span><span class="sxs-lookup"><span data-stu-id="4434c-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="4434c-104">In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) verwenden und Methoden in Excel im Web mithilfe Office berechnen.</span><span class="sxs-lookup"><span data-stu-id="4434c-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="4434c-105">Sie können das Skript in einer beliebigen Excel testen.</span><span class="sxs-lookup"><span data-stu-id="4434c-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="4434c-106">Szenario</span><span class="sxs-lookup"><span data-stu-id="4434c-106">Scenario</span></span>

<span data-ttu-id="4434c-107">In Excel im Web kann der Berechnungsmodus einer Datei mithilfe von APIs programmgesteuert gesteuert werden.</span><span class="sxs-lookup"><span data-stu-id="4434c-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="4434c-108">Die folgenden Aktionen sind mithilfe von skripts Office möglich.</span><span class="sxs-lookup"><span data-stu-id="4434c-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="4434c-109">Den Berechnungsmodus erhalten.</span><span class="sxs-lookup"><span data-stu-id="4434c-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="4434c-110">Legen Sie den Berechnungsmodus ein.</span><span class="sxs-lookup"><span data-stu-id="4434c-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="4434c-111">Berechnen Excel Formeln für Dateien, die auf den manuellen Modus festgelegt sind (auch als Neuberechnung bezeichnet).</span><span class="sxs-lookup"><span data-stu-id="4434c-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="4434c-112">Beispielcode: Steuerelementberechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="4434c-112">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="4434c-113">Schulungsvideo: Verwalten des Berechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="4434c-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="4434c-114">[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/iw6O8QH01CI)</span><span class="sxs-lookup"><span data-stu-id="4434c-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
