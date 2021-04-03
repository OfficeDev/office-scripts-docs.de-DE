---
title: Verwalten des Berechnungsmodus in Excel
description: Erfahren Sie, wie Sie Office-Skripts verwenden, um den Berechnungsmodus in Excel im Web zu verwalten.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571527"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="3e76b-103">Verwalten des Berechnungsmodus in Excel</span><span class="sxs-lookup"><span data-stu-id="3e76b-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="3e76b-104">In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus verwenden](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) und Methoden in Excel im Web mithilfe von Office-Skripts berechnen.</span><span class="sxs-lookup"><span data-stu-id="3e76b-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="3e76b-105">Sie können das Skript in einer beliebigen Excel-Datei ausprobieren.</span><span class="sxs-lookup"><span data-stu-id="3e76b-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="3e76b-106">Szenario</span><span class="sxs-lookup"><span data-stu-id="3e76b-106">Scenario</span></span>

<span data-ttu-id="3e76b-107">In Excel im Web kann der Berechnungsmodus einer Datei mithilfe von APIs programmgesteuert gesteuert werden.</span><span class="sxs-lookup"><span data-stu-id="3e76b-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="3e76b-108">Die folgenden Aktionen sind mit Office-Skripts möglich.</span><span class="sxs-lookup"><span data-stu-id="3e76b-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="3e76b-109">Den Berechnungsmodus erhalten.</span><span class="sxs-lookup"><span data-stu-id="3e76b-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="3e76b-110">Legen Sie den Berechnungsmodus ein.</span><span class="sxs-lookup"><span data-stu-id="3e76b-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="3e76b-111">Berechnen Sie Excel-Formeln für Dateien, die auf den manuellen Modus festgelegt sind (auch als Neuberechnung bezeichnet).</span><span class="sxs-lookup"><span data-stu-id="3e76b-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="3e76b-112">Beispielcode: Steuerelementberechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="3e76b-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="3e76b-113">Schulungsvideo: Verwalten des Berechnungsmodus</span><span class="sxs-lookup"><span data-stu-id="3e76b-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="3e76b-114">[![Schritt-für-Schritt-Video zur Verwaltung des Berechnungsmodus in Excel im Web ansehen](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Schritt-für-Schritt-Video zum Verwalten des Berechnungsmodus in Excel im Web")</span><span class="sxs-lookup"><span data-stu-id="3e76b-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
