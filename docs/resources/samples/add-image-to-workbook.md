---
title: Hinzufügen von Bildern zu einer Arbeitsmappe
description: Erfahren Sie, wie Sie office-Skripts verwenden, um einer Arbeitsmappe ein Bild hinzuzufügen und es blätterübergreifend zu kopieren.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572605"
---
# <a name="add-images-to-a-workbook"></a>Hinzufügen von Bildern zu einer Arbeitsmappe

In diesem Beispiel wird gezeigt, wie Sie mit Bildern mithilfe eines Office-Skripts in Excel arbeiten.

## <a name="scenario"></a>Szenario

Bilder helfen bei Branding, visueller Identität und Vorlagen. Sie helfen dabei, eine Arbeitsmappe mehr als nur eine riesige Tabelle zu machen.

Im ersten Beispiel wird ein Bild von einem Arbeitsblatt in ein anderes kopiert. Dies kann verwendet werden, um das Logo Ihres Unternehmens an der gleichen Position auf jedem Blatt zu platzieren.

Im zweiten Beispiel wird ein Bild von einer URL kopiert. Dies kann verwendet werden, um Fotos, die ein Kollege in einem freigegebenen Ordner gespeichert hat, in eine verwandte Arbeitsmappe zu kopieren.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [add-images.xlsx](add-images.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie die folgenden Skripts hinzu, und probieren Sie das Beispiel selbst aus!

## <a name="sample-code-copy-an-image-across-worksheets"></a>Beispielcode: Kopieren eines Bilds auf Arbeitsblättern

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Beispielcode: Hinzufügen eines Bilds von einer URL zu einer Arbeitsmappe

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
