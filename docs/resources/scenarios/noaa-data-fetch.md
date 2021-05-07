---
title: 'Office Beispielszenario für Skripts: Graph daten auf Wasserebene aus NOAA'
description: Ein Beispiel, das JSON-Daten aus einer NOAA-Datenbank abruft und zum Erstellen eines Diagramms verwendet.
ms.date: 04/26/2021
localization_priority: Normal
ms.openlocfilehash: d35af59d9eed1abc9f3844834c92752ed80de80f
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232690"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office Skriptbeispielszenario: Abrufen und Graphen von Daten auf Wasserebene aus NOAA

In diesem Szenario müssen Sie den Wasserstand an der Station der [National Oceanic and Atmospheric Administration in Seattle plotten.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Sie verwenden externe Daten, um eine Tabellenkalkulation auffüllen und ein Diagramm zu erstellen.

Sie entwickeln ein Skript, das den Befehl zum Abfragen der `fetch` [NOAA-Flutungen- und Currents-Datenbank verwendet.](https://tidesandcurrents.noaa.gov/) Damit wird der Wasserstand über einen bestimmten Zeitraum erfasst. Die Informationen werden als JSON zurückgegeben, sodass ein Teil des Skripts dies in Bereichswerte übersetzt. Sobald sich die Daten in der Kalkulationstabelle befindet, werden sie zum Erstellen eines Diagramms verwendet.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptkenntnisse

- Externe API-Aufrufe ( `fetch` )
- JSON-Analyse
- Diagramme

## <a name="setup-instructions"></a>Setupanweisungen

1. Öffnen Sie die Arbeitsmappe mit Excel im Web.

1. Wählen Sie **auf der** Registerkarte Automatisieren **alle Skripts aus.**

1. Wählen Sie **im Aufgabenbereich Code-Editor** die Option **Neues Skript aus,** und fügen Sie das folgende Skript in den Editor ein.

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);
    
      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1",
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. Benennen Sie das Skript in **NOAA Water Level Chart um,** und speichern Sie es.

## <a name="running-the-script"></a>Ausführen des Skripts

Führen Sie auf einem beliebigen Arbeitsblatt das **NOAA Water Level Chart-Skript** aus. Das Skript ruft die Wasserstandsdaten vom 25. Dezember 2020 bis zum 27. Dezember 2020 ab. Die Variablen am Anfang des Skripts können geändert werden, um unterschiedliche `const` Datumsangaben zu verwenden oder andere Senderinformationen zu erhalten. In [der CO-OPS-API für den Datenabruf](https://api.tidesandcurrents.noaa.gov/api/prod/) wird beschrieben, wie sie alle diese Daten abrufen.

### <a name="after-running-the-script"></a>Nach dem Ausführen des Skripts

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="Das Arbeitsblatt nach dem Ausführen des Skripts zeigt einige Wasserstandsdaten und ein Diagramm an.":::
