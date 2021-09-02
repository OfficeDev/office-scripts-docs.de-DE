---
title: Verwalten des Berechnungsmodus in Excel
description: Erfahren Sie, wie Sie Office Skripts verwenden, um den Berechnungsmodus in Excel im Web zu verwalten.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: ee43c3c0477f0d70078cae271081bc5e1e008960
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862152"
---
# <a name="manage-calculation-mode-in-excel"></a>Verwalten des Berechnungsmodus in Excel

Dieses Beispiel zeigt, wie Sie den [Berechnungsmodus](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) verwenden und Methoden in Excel im Web mit Office Skripts berechnen. Sie können das Skript in einer beliebigen Excel-Datei testen.

## <a name="scenario"></a>Szenario

Arbeitsmappen mit einer großen Anzahl von Formeln können eine Weile dauern, bis die Neuberechnung erfolgt. Anstatt Excel steuern zu lassen, wann Berechnungen durchgeführt werden, können Sie diese als Teil Ihres Skripts verwalten. Dies hilft bei der Leistung in bestimmten Szenarien.

Das Beispielskript legt den Berechnungsmodus auf manuell fest. Dies bedeutet, dass die Arbeitsmappe Formeln nur neu berechnet, wenn das Skript sie anweist (oder Sie manuell [über die Benutzeroberfläche berechnen).](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4) Das Skript zeigt dann den aktuellen Berechnungsmodus an und berechnet die gesamte Arbeitsmappe vollständig neu.

## <a name="sample-code-control-calculation-mode"></a>Beispielcode: Steuern des Berechnungsmodus

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

[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/iw6O8QH01CI)
