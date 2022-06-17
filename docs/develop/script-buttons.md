---
title: Ausführen Office Skripts in Excel mit Schaltflächen
description: Fügen Sie Arbeitsmappen Schaltflächen hinzu, die Office Skripts in Excel steuern.
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc19a13a97d4d11f73cb91bc46b70afff3eadf03
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128216"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Ausführen Office Skripts in Excel mit Schaltflächen

Helfen Sie Ihren Kollegen, Ihre Skripts zu finden und auszuführen, indem Sie einer Arbeitsmappe Skriptschaltflächen hinzufügen.

:::image type="content" source="../images/run-from-button.png" alt-text="Eine Schaltfläche im Arbeitsblatt, die ein Skript ausführt, wenn darauf geklickt wird.":::

## <a name="create-script-buttons"></a>Erstellen von Skriptschaltflächen

Wechseln Sie bei einem beliebigen Skript zum Menü **"Weitere Optionen (...)"** auf der Detailseite des Skripts oder im Aufgabenbereich des Code-Editors, und wählen Sie **die Schaltfläche "Hinzufügen"** aus. Dadurch wird in der Arbeitsmappe eine Schaltfläche erstellt, die bei Auswahl das zugehörige Skript ausführt. Außerdem wird das Skript für die Arbeitsmappe freigegeben, sodass jeder Benutzer mit Schreibberechtigungen für die Arbeitsmappe Ihre hilfreiche Automatisierung verwenden kann.

Der folgende Screenshot zeigt die Seite mit den Skriptdetails für ein Skript mit dem Titel **"PivotTable erstellen** " und die **Option "Hinzufügen"** im Menü **"Weitere Optionen (...)"** ist hervorgehoben.

:::image type="content" source="../images/add-button.png" alt-text="Die Option &quot;Schaltfläche hinzufügen&quot; im Menü der Skriptdetailseite.":::

## <a name="remove-script-buttons"></a>Entfernen von Skriptschaltflächen

Um die Freigabe eines Skripts über eine Schaltfläche zu beenden, wechseln Sie auf der Seite mit den Skriptdetails zum Menü **"Weitere Optionen (...)** ", und wählen Sie " **Freigabe beenden"** aus. Dadurch werden alle Schaltflächen, die das Skript ausführen, entfernt. Durch das Löschen einer einzelnen Schaltfläche wird das Skript von dieser einen Schaltfläche entfernt, selbst wenn der Vorgang rückgängig gemacht wird oder die Schaltfläche ausgeschnitten und eingefügt wird.

## <a name="script-buttons-with-excel-on-windows"></a>Skriptschaltflächen mit Excel auf Windows

Diese Skriptschaltflächen funktionieren auch unter Windows. Erstellen Sie die Schaltfläche in Excel im Web, und Benutzer auf Windows können Ihr Skript mit einem Klick auf eine Schaltfläche ausführen. Bitte beachten Sie, dass Sie skripts in Excel auf Windows nicht bearbeiten können. Skripts können nur in Excel im Web bearbeitet werden.

Einige Office Skript-APIs werden von Excel auf Windows möglicherweise nicht unterstützt, insbesondere ältere Builds. Dazu gehören neuere APIs und APIs für Nur-Web-Features. Wenn ein Skript nicht unterstützte APIs enthält, wird das Skript nicht ausgeführt, und stattdessen wird im Aufgabenbereich "**Skriptausführungsstatus**" eine Warnmeldung angezeigt, die besagt: "Dieses Skript muss derzeit auf Excel für das Web ausgeführt werden. Öffnen Sie die Arbeitsmappe im Browser, und versuchen Sie es dann erneut, oder wenden Sie sich an den Skriptbesitzer, um Hilfe zu bitten."  

> [!IMPORTANT]
> Skriptschaltflächen erfordern[, dass WebView2](/deployoffice/webview2-install) mit Excel auf Windows funktioniert. Dies ist standardmäßig mit den neuesten Versionen von Excel auf dem Desktop installiert. Wenn Sie jedoch nicht auf Skriptschaltflächen klicken können, besuchen [Sie "WebView2-Runtime herunterladen](https://developer.microsoft.com/microsoft-edge/webview2/#download-section)", und laden Sie das Browsermodul herunter.
