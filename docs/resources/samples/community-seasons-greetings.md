---
title: Weihnachtsgrüße
description: Erfahren Sie, wie Sie office-Skripts verwenden, um eine singende Struktur in Excel im Web anzuzeigen.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: aee953bd3f92912b6b3bcf55c3a3da110ff38528
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572444"
---
# <a name="seasons-greetings"></a>Weihnachtsgrüße

Dieses Skript wurde von [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) im Geiste der Ferienzeit beigetragen! Es ist ein lustiges Skript, das eine singende Struktur in Excel im Web mitHilfe von Office-Skripts zeigt.

Genießen!

[Schauen Sie sich das Skript "Seasons Greetings" im YouTube-Kanal "Les es IT Blog" in Aktion an](https://youtu.be/HBiGEkzmkgo).

## <a name="script"></a>Skript

Laden Sie [happy-tree.xlsx](happy-tree.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

```TypeScript
/* Original version by Leslie Black.  */

function main(workbook: ExcelScript.Workbook) {
  let happyTree = workbook.getWorksheet('HappyTree');
  happyTree.activate();

  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setFlashingStarAndSmileRed(workbook) //red
  setFlashingStarAndSmileYellow(workbook) //yellow

  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  setFlashingStarAndSmileRed(workbook) //red
  setFlashingStarAndSmileYellow(workbook) //yellow
  blink(workbook)

  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow

  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow

  setOuterEdgeRed(workbook) //red
  setOuterEdgeYellow(workbook) //yellow
  unblink(workbook)

  console.log('Routine finished');

  function blink(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getWorksheet('HappyTree');
    // Set the eyes to brown.
    selectedSheet.getRanges("N16:Q17, G16: J17")
      .getFormat()
      .getFill()
      .setColor("C65911");
  }

  function unblink(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getWorksheet('HappyTree');
    // Set the eyes back to white (except the pupils).
    selectedSheet.getRanges("N16:N17, O16:Q16, G16:H17, I16:J16, P17:Q17, J17")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
  }

  function setFlashingStarAndSmileRed(workbook: ExcelScript.Workbook) {
    // Set the star to red.
    let selectedSheet = workbook.getWorksheet('HappyTree');
    selectedSheet.getRanges("L2:L6, K3:K5, M3:M5, N4, J4")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set the smile points to black.
    selectedSheet.getRanges("I26, O26")
      .getFormat()
      .getFill()
      .setColor("000000");
  }

  function setFlashingStarAndSmileYellow(workbook: ExcelScript.Workbook) {
    // Set the start to yellow.
    let selectedSheet = workbook.getWorksheet('HappyTree');
    selectedSheet.getRanges("L2:L6, K3:K5, M3:M5, N4, J4")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Clear the smile points.
    selectedSheet.getRanges("O26, I26")
      .getFormat()
      .getFill().clear();
  }
}

function setOuterEdgeYellow(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getWorksheet('HappyTree');
  // Set the outer edge to yellow.
  sheet.getRanges("Q11, G11, R12, F12, S13, E13, T14, D14, C15, U15, T16:T17, D16:D17, C18, U18, T19, D19, L2:L6, C21, U21, C23, U23, C25, U25, C27, U27, C29, U29, T30, D30, K3:K5, M3: M5, S31, E31, R32, F32, Q33, G33, P34, H34, O35, I35, N36:N37, J36:J37, K37:M37, N4, J4, K7, M7, N8, J8, O9, I9, P10, H10")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
}

function setOuterEdgeRed(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getWorksheet('HappyTree');
  // Set the outer edge to red.
  sheet.getRanges("Q11, G11, R12, F12, S13, E13, T14, D14, C15, U15, T16:T17, D16:D17, C18, U18, T19, D19, L2:L6, C21, U21, C23, U23, C25, U25, C27, U27, C29, U29, T30, D30, K3:K5, M3: M5, S31, E31, R32, F32, Q33, G33, P34, H34, O35, I35, N36:N37, J36:J37, K37:M37, N4, J4, K7, M7, N8, J8, O9, I9, P10, H10")
    .getFormat()
    .getFill()
    .setColor("FF0000");
}
```
