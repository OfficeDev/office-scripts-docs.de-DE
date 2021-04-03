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
# <a name="manage-calculation-mode-in-excel"></a>Verwalten des Berechnungsmodus in Excel

In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus verwenden](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) und Methoden in Excel im Web mithilfe von Office-Skripts berechnen. Sie können das Skript in einer beliebigen Excel-Datei ausprobieren.

## <a name="scenario"></a>Szenario

In Excel im Web kann der Berechnungsmodus einer Datei mithilfe von APIs programmgesteuert gesteuert werden. Die folgenden Aktionen sind mit Office-Skripts möglich.

1. Den Berechnungsmodus erhalten.
1. Legen Sie den Berechnungsmodus ein.
1. Berechnen Sie Excel-Formeln für Dateien, die auf den manuellen Modus festgelegt sind (auch als Neuberechnung bezeichnet).

## <a name="sample-code-control-calculation-mode"></a>Beispielcode: Steuerelementberechnungsmodus

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

## <a name="training-video-manage-calculation-mode"></a>Schulungsvideo: Verwalten des Berechnungsmodus

[![Schritt-für-Schritt-Video zur Verwaltung des Berechnungsmodus in Excel im Web ansehen](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Schritt-für-Schritt-Video zum Verwalten des Berechnungsmodus in Excel im Web")
