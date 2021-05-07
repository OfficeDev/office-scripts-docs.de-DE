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
# <a name="manage-calculation-mode-in-excel"></a>Verwalten des Berechnungsmodus in Excel

In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) verwenden und Methoden in Excel im Web mithilfe Office berechnen. Sie können das Skript in einer beliebigen Excel testen.

## <a name="scenario"></a>Szenario

In Excel im Web kann der Berechnungsmodus einer Datei mithilfe von APIs programmgesteuert gesteuert werden. Die folgenden Aktionen sind mithilfe von skripts Office möglich.

1. Den Berechnungsmodus erhalten.
1. Legen Sie den Berechnungsmodus ein.
1. Berechnen Excel Formeln für Dateien, die auf den manuellen Modus festgelegt sind (auch als Neuberechnung bezeichnet).

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

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/iw6O8QH01CI)
