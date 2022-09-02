---
title: 'Beispielszenario für Office-Skripts: Schaltfläche "Taktstempel"'
description: In diesem Beispiel wird eine Druckuhrtaste hinzugefügt, und ein Benutzer kann mit der aktuellen Zeit ein- und ausstempeln.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: ac128a33b653506b6168bd4acfe1713bf6d26759
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572682"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Beispielszenario für Office-Skripts: Schaltfläche "Taktstempel"

Die in diesem Beispiel verwendete Szenarioidee und das Skript wurden vom Office Scripts-Communitymitglied [Brian Gonzalez](https://github.com/b-gonzalez) beigetragen.

In diesem Szenario erstellen Sie eine Arbeitszeittabelle für einen Mitarbeiter, mit der er seine Start- und Endzeiten mit einem Klick auf eine [Schaltfläche](../../develop/script-buttons.md) aufzeichnen kann. Je nachdem, was zuvor aufgezeichnet wurde, startet das Drücken der Schaltfläche entweder ihren Tag (Ein-/Auschecken) oder beendet den Tag (auschecken). Das Beispiel funktioniert sowohl für Excel im Web als auch für Windows.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Eine Tabelle mit drei Spalten ('Clock In', 'Clock Out' und 'Duration') und einer Schaltfläche mit der Bezeichnung &quot;Punch clock&quot; in der Arbeitsmappe.":::

## <a name="setup-instructions"></a>Setupanweisungen

1. Laden Sie [punch-clock-sample.xlsx](punch-clock-sample.xlsx) auf OneDrive herunter.

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="Eine Tabelle mit drei Spalten: &quot;Clock In&quot;, &quot;Clock Out&quot; und &quot;Duration&quot;.":::

1. Öffnen Sie die Arbeitsmappe in Excel im Web.

1. Wählen Sie auf der Registerkarte **"Automatisieren** " die Option **"Neues Skript** " aus, und fügen Sie das folgende Skript in den Editor ein.

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. Benennen Sie das Skript in "Punch clock" um.

1. Speichern Sie das Skript.

1. Wählen Sie in der Arbeitsmappe Zelle **E2** aus.

1. Hinzufügen einer Skriptschaltfläche. Wechseln Sie auf der Seite "**Skriptdetails**" zum Menü "**Weitere Optionen (...) ",** und wählen Sie **die Schaltfläche "Hinzufügen"** aus.

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="Das Menü &quot;Weitere Optionen&quot; und die Schaltfläche &quot;Hinzufügen&quot;.":::

1. Speichern Sie die Arbeitsmappe.

## <a name="run-the-script"></a>Ausführen des Skripts

Drücken Sie die **Druckuhrtaste** , um das Skript auszuführen. Je nachdem, was zuvor eingegeben wurde, wird entweder die aktuelle Uhrzeit unter "Uhr ein" oder "Auschecken" protokolliert.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Die Tabelle und die Schaltfläche &quot;Punch clock&quot; in der Arbeitsmappe.":::

> [!NOTE]
> Die Dauer wird nur aufgezeichnet, wenn sie länger als eine Minute ist. Bearbeiten Sie manuell die Zeit "Einchecken", um größere Dauer zu testen.
