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
# <a name="manage-calculation-mode-in-excel"></a>Verwalten des Berechnungsmodus in Excel

In diesem Beispiel wird gezeigt, wie Sie den [Berechnungsmodus](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) verwenden und Methoden in Excel im Web mithilfe Office berechnen. Sie können das Skript in einer beliebigen Excel testen.

## <a name="scenario"></a>Szenario

Arbeitsmappen mit einer großen Anzahl von Formeln können eine Weile dauern, bis sie neu berechnet werden. Anstatt Excel steuern zu lassen, wenn Berechnungen ausgeführt werden, können Sie sie als Teil Ihres Skripts verwalten. Dies hilft bei der Leistung in bestimmten Szenarien.

Das Beispielskript legt den Berechnungsmodus auf manuell fest. Dies bedeutet, dass die Arbeitsmappe formeln nur neu berechnet, wenn das Skript dies ansah (oder Sie manuell über die [Benutzeroberfläche berechnen).](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4) Das Skript zeigt dann den aktuellen Berechnungsmodus an und berechnet die gesamte Arbeitsmappe vollständig neu.

## <a name="sample-code-control-calculation-mode"></a>Beispielcode: Steuerelementberechnungsmodus

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

## <a name="training-video-manage-calculation-mode"></a>Schulungsvideo: Verwalten des Berechnungsmodus

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/iw6O8QH01CI)
