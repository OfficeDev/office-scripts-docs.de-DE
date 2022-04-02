---
title: Verwendung von Power Query oder Office-Skripts
description: Die Szenarien, die am besten für die Plattformen Power Query und Office Skripts geeignet sind.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585905"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Verwendung von Power Query oder Office-Skripts

[Power Query](https://powerquery.microsoft.com) und Office Skripts sind leistungsstarke Automatisierungslösungen für Excel. Beide Lösungen ermöglichen Excel Benutzern das Bereinigen und Transformieren von Daten in Arbeitsmappen. Ein einzelnes Power Query- oder Office-Skript kann aktualisiert und erneut für neue Daten ausgeführt werden, um konsistente Ergebnisse zu erzielen, wodurch Sie Zeit sparen und schneller mit den resultierenden Informationen arbeiten können.

Dieser Artikel bietet eine allgemeine Übersicht darüber, wann Sie eine Plattform gegenüber der anderen bevorzugen. Im Allgemeinen eignet sich Power Query gut zum Abrufen und Transformieren von Daten aus großen, externen Datenquellen und Office Skripts eignen sich für schnelle, Excel-zentrierte Lösungen und [Power Automate Integrationen](../develop/power-automate-integration.md).

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Große Datenquellen und Datenabruf: Power Query

Wir empfehlen Power Query beim Umgang mit Datenquellen von unterstützten Plattformen.

Power Query verfügt [über integrierte Datenverbindungen mit Hunderten](https://powerquery.microsoft.com/connectors/) von Quellen. Power Query wurde speziell für das Abrufen, Transformieren und Kombinieren von Daten entwickelt. Wenn Sie Daten aus einer dieser Quellen benötigen, bietet ihnen Power Query eine codefreie Möglichkeit, diese Daten in Excel in der benötigten Form zu platzieren.

Diese Power Query Verbindungen sind für große Datasets ausgelegt. Sie haben nicht dieselben [Übertragungsgrenzwerte](../testing/platform-limits.md) wie Power Automate oder Excel für das Web.

Office Skripts bieten eine einfache Lösung für kleinere Datenquellen oder Datenquellen, die nicht von Power Query Connectors abgedeckt werden. Dazu gehören [die Verwendung `fetch` oder REST-APIs](../develop/external-calls.md) oder das Abrufen von Informationen aus Ad-hoc-Datenquellen, z. B. eine [Teams adaptive Karte](../resources/scenarios/task-reminders.md).

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Formatierung, Visualisierungen und programmgesteuerte Steuerung: Office Skripts

Wir empfehlen Office Skripts, wenn Ihre Anforderungen über das Importieren und Transformieren von Daten hinausgehen.

Nahezu alles, was Sie manuell über die Excel Benutzeroberfläche ausführen können, ist mit Office Skripts möglich. Sie  eignet sich hervorragend zum Anwenden einer konsistenten Formatierung auf Arbeitsmappen. Skripts erstellen Diagramme, PivotTables, Formen, Bilder und andere Arbeitsblattvisualisierungen. Skripts bieten Ihnen außerdem eine genaue Kontrolle über positionen, Größen, Farben und andere Attribute dieser Visualisierungen.

Durch die Einbindung von TypeScript-Code erhalten Sie ein hohes Maß an Anpassung. Programmgesteuerte Steuerungslogik wie `if...else` Anweisungen macht Ihr Skript robust. Auf diese Weise können Sie Beispielsweise Daten bedingt lesen, ohne sich auf komplexe Excel Formeln verlassen zu müssen, oder die Arbeitsmappe vor dem Ändern der Arbeitsmappe auf unerwartete Änderungen überprüfen.

Formatierungen können mit Power Query über Excel [Vorlagen](https://templates.office.com/power-query-tutorial-tm11414620) angewendet werden. Vorlagen werden jedoch auf Einzel- oder Organisationsebene aktualisiert, während Office Skripts eine präzisere Zugriffssteuerung bieten.

## <a name="power-automate-integrations"></a>Power Automate-Integrationen

Office Skripts bieten weitere Optionen für Power Automate Integration. Skripts sind auf Ihre Lösungen zugeschnitten. Sie definieren die [Eingabe und Ausgabe des Skripts](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), damit es mit einem beliebigen anderen Connector oder daten im Fluss funktioniert. Der folgende Screenshot zeigt ein Beispiel Power Automate Flusses, der Daten von einer Teams adaptiven Karte an ein Office-Skript übergibt.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Screenshot des Excel Online (Business)-Connectors im Flow-Designer. Der Connector verwendet die Skriptaktion ausführen, um Eingaben von einer Teams adaptiven Karte zu übernehmen und für ein Skript bereitzustellen.":::

Power Query wird in der [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate-Verbindung verwendet. Mit [der Transformation von Daten mit Power Query](/connectors/sql/#transform-data-using-power-query) Aktion können Sie eine Abfrage in Power Automate erstellen. Obwohl dies ein leistungsstarkes Tool für die Verwendung mit SQL Server ist, beschränkt es Power Query auf diese Eingabequelle, wie im folgenden Fluss-Screenshot dargestellt.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Screenshot der SQL Server-Verbindung im Flow-Designer. Der Connector verwendet die Transformationsdaten mithilfe Power Query Aktion.":::

## <a name="platform-dependencies"></a>Plattformabhängigkeiten

Office Skripts sind derzeit nur für Excel im Web verfügbar. Power Query ist derzeit nur für Excel auf dem Desktop verfügbar. Beide können über Power Automate verwendet werden, sodass der Fluss mit Excel Arbeitsmappen funktioniert, die in OneDrive gespeichert sind.

## <a name="see-also"></a>Weitere Informationen

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query mit Excel](https://powerquery.microsoft.com/excel/)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
