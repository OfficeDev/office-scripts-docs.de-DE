---
title: Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web
description: Dies ist ein Lernprogramm zu den Grundlagen von Office-Skripts, einschließlich dem Aufzeichnen von Skripts mithilfe der Aktionsaufzeichnung und dem Schreiben von Daten in eine Arbeitsmappe.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 1971ff2ffd80554beb6ac561677ee3384f87ca81
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700174"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web

In diesem Lernprogramm lernen Sie die Grundlagen zum Aufzeichnen, Bearbeiten und Schreiben eines Office-Skripts für Excel im Web kennen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Preview note](../includes/preview-note.md)]

Bevor Sie mit diesem Lernprogramm beginnen, benötigen Sie Zugriff auf Office-Skripts. Dies setzt Folgendes voraus:

- [Excel im Web](https://www.office.com/launch/excel).
- Bitten Sie Ihren Administrator, [Office-Skripts für Ihre Organisation zu aktivieren](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf). Dadurch wird die Registerkarte "**Automatisieren**" zum Menüband hinzugefügt.

> [!IMPORTANT]
> Dieses Lernprogramm richtet sich an Personen, die über mittlere JavaScript- oder TypeScript-Kenntnisse verfügen. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, sich das [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) durchzusehen. Weitere Informationen über die Skriptumgebung finden Sie unter [Office-Skripts in Excel im Web](../overview/excel.md).

## <a name="add-data-and-record-a-basic-script"></a>Hinzufügen von Daten und Aufzeichnen eines einfachen Skripts

Zuerst benötigen wir einige Daten und ein kleines Startskript.

1. Erstellen Sie eine neue Arbeitsmappe in Excel im Web.
2. Kopieren Sie die folgenden Obst-Umsatzdaten, und fügen Sie diese beginnend bei Zelle **A1** in das Arbeitsblatt ein.

    |Obstsorte |2018 |2019 |
    |:---|:---|:---|
    |Orangen |1000 |1200 |
    |Zitronen |800 |900 |
    |Limetten |600 |500 |
    |Grapefruits |900 |700 |

3. Öffnen Sie die Registerkarte **Automatisieren**. Falls die Registerkarte **Automatisieren** nicht angezeigt wird, überprüfen Sie den Menüband-Überlauf, indem Sie auf den Dropdownpfeil klicken.
4. Klicken Sie auf die Schaltfläche **Aktionen aufzeichnen**.
5. Wählen Sie Zellen **A2:C2** (die Zeile "Orangen") aus, und legen Sie die Füllfarbe auf Orange fest.
6. Beenden Sie die Aufzeichnung, indem Sie die Schaltfläche **Stopp** drücken.
7. Geben Sie in das Feld **Skriptnamen** einen einprägsamen Namen ein.
8. *Optional:* Füllen Sie das Feld **Beschreibung** mit einer aussagekräftigen Beschreibung aus. Diese wird verwendet, um den Kontext des Skripts bereitzustellen. Für dieses Lernprogramm können Sie "Farbcodierte Zeilen einer Tabelle" verwenden.

   > [!TIP]
   > Sie können die Beschreibung eines Skripts später im Bereich **Skriptdetails** bearbeiten, der sich im Menü **...** des Code-Editors befindet.

9. Klicken Sie auf die Schaltfläche **Speichern**, um das Skript zu speichern.

    Ihr Arbeitsblatt sollte wie folgt aussehen (machen Sie sich keine Sorgen, wenn die Farbe anders ist):

    ![Eine Zeile mit Obst-Umsatzdaten mit hervorgehobener orangefarbener Zeile "Orangen".](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a>Bearbeiten eines vorhandenen Skripts

Das vorherige Skript hat die Zeile "Orangen" orangefarben eingefärbt. Jetzt fügen wir eine gelbe Zeile für die "Zitronen" hinzu.

1. Öffnen Sie die Registerkarte **Automatisieren**.
2. Klicken Sie auf die Schaltfläche **Code-Editor**.
3. Öffnen Sie das im vorherigen Abschnitt aufgezeichnete Skript. Folgendes sollte nun auf dem Bildschirm angezeigt werden:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
    }
    ```

    Dieser Code ruft das aktuelle Arbeitsblatt ab, indem er zuerst auf die Arbeitsblattsammlung der Arbeitsmappe zugreift. Anschließend legt er die Füllfarbe des Bereichs **A2:C2** fest.

    Bereiche sind ein wesentliches Element von Office-Skripts in Excel im Web. Ein Bereich ist ein zusammenhängender, rechteckiger Block von Zellen, die Werte, Formeln und Formatierungen enthalten. Hierbei handelt es sich um die grundlegende Struktur von Zellen, durch die Sie die meisten Skript-Aufgaben ausführen werden.

4. Fügen Sie am Ende des Skripts die folgende Zeile ein (zwischen dem festgelegten `color` und der schließenden `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
    ```

5. Testen Sie das Skript, indem Sie **Ausführen** drücken. Ihre Arbeitsmappe sollte nun wie folgt aussehen:

    ![Eine Zeile mit Obst-Umsatzdaten mit orangefarben hervorgehobener Zeile "Orangen" und einer gelb hervorgehobenen Zeile "Zitronen".](../images/tutorial-2.png)

## <a name="create-a-table"></a>Erstellen einer Tabelle

Wandeln wir diese Obst-Umsatzdaten in eine Tabelle um. Wir verwenden unser Skript für den gesamten Prozess.

1. Fügen Sie die folgende Zeile am Ende des Skripts hinzu (vor der schließenden `}`):

    ```TypeScript
    let table = selectedSheet.tables.add("A1:C5", true);
    ```

2. Dieser Aufruf gibt ein `Table`-Objekt zurück. Verwenden wir diese Tabelle zum Sortieren der Daten. Wir werden die Daten basierend auf den Werten in der Spalte "Obstsorte" in aufsteigender Reihenfolge sortieren. Fügen Sie dann die folgende Zeile nach der Tabellenerstellung hinzu:

    ```TypeScript
    table.sort.apply([{ key: 0, ascending: true }]);
    ```

    Das Skript sollte wie folgt aussehen:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
      selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
      let table = selectedSheet.tables.add("A1:C5", true);
      table.sort.apply([{ key: 0, ascending: true }]);
    }
    ```

    Tabellen beinhalten ein `TableSort`-Objekt, auf das über die `Table.sort`-Eigenschaft zugegriffen wird. Sie können auf dieses Objekt Sortierkriterien anwenden. Die `apply`-Methode bezieht eine Reihe von `SortField`-Objekten ein. In diesem Fall gibt es nur ein Sortierkriterium, also verwenden wir nur ein `SortField`. `key: 0` legt für die Spalte mit den die Sortierung bestimmenden Werten "0" fest (dies ist die erste Spalte in der Tabelle, in diesem Fall **A**). `ascending: true` sortiert die Daten in aufsteigender Reihenfolge (statt in absteigender Reihenfolge).

3. Führen Sie das Skript aus. Es sollte eine Tabelle wie die folgende angezeigt werden:

    ![Eine Tabelle mit sortierten Obst-Umsatzdaten.](../images/tutorial-3.png)

    > [!NOTE]
    > Wenn Sie das Skript erneut ausführen, wird eine Fehlermeldung angezeigt. Der Grund dafür ist, dass Sie keine Tabelle über eine andere Tabelle erstellen können. Sie können das Skript jedoch auf ein anderes Arbeitsblatt oder eine andere Arbeitsmappe anwenden.

### <a name="re-run-the-script"></a>Das Skript erneut ausführen

1. Erstellen Sie ein neues Arbeitsblatt in der aktuellen Arbeitsmappe.
2. Kopieren Sie die Obstdaten am Anfang dieses Lernprogramms, und fügen Sie sie in das neue Arbeitsblatt ein, beginnend bei Zelle **A1**.
3. Führen Sie das Skript aus.

## <a name="next-steps"></a>Nächste Schritte

Führen Sie das Lernprogramm [Lesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web](excel-read-tutorial.md) aus. Hier erfahren Sie, wie Sie mithilfe eines Office-Skripts Daten aus einer Arbeitsmappe lesen können.
