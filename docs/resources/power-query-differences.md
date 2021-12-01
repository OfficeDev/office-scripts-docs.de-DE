---
title: Verwendung von Power Query- oder Office-Skripts
description: Die Szenarien, die am besten für die Power Query- und Office Scripts-Plattformen geeignet sind.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1812b508b2cde4d304ecf228adfdd8f68de9808a
ms.sourcegitcommit: 383880e0dc0d09b8f76884675531e462a292d747
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 12/01/2021
ms.locfileid: "61245611"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Verwendung von Power Query- oder Office-Skripts

[Power Query-](https://powerquery.microsoft.com) und Office-Skripts sind leistungsstarke Automatisierungslösungen für Excel. Beide Lösungen ermöglichen Excel Benutzern das Bereinigen und Transformieren von Daten in Arbeitsmappen. Ein einzelnes Power Query- oder Office-Skript kann aktualisiert und erneut für neue Daten ausgeführt werden, um konsistente Ergebnisse zu erzielen, wodurch Sie Zeit sparen und schneller mit den resultierenden Informationen arbeiten können.

Dieser Artikel bietet eine allgemeine Übersicht darüber, wann Sie eine Plattform gegenüber der anderen bevorzugen. Im Allgemeinen eignet sich Power Query gut zum Abrufen und Transformieren von Daten aus großen, externen Datenquellen und Office Skripts eignen sich gut für schnelle, Excel orientierte Lösungen und [Power Automate Integrationen.](../develop/power-automate-integration.md)

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Große Datenquellen und Datenabruf: Power Query

Wir empfehlen Power Query beim Umgang mit Datenquellen von unterstützten Plattformen.

Power Query verfügt [über integrierte Datenverbindungen zu Hunderten](https://powerquery.microsoft.com/connectors/) von Quellen. Power Query wurde speziell für Datenabruf-, Transformations- und Kombinationsaufgaben entwickelt. Wenn Sie Daten aus einer dieser Quellen benötigen, bietet Ihnen Power Query eine codefreie Möglichkeit, diese Daten in Excel in dem shape zu bringen, das Sie benötigen.

Diese Power Query-Verbindungen sind für große Datasets konzipiert. Sie haben nicht dieselben [Übertragungsgrenzwerte](../testing/platform-limits.md) wie Power Automate oder Excel für das Web.

Office Skripts bieten eine einfache Lösung für kleinere Datenquellen oder Datenquellen, die nicht von Power Query-Connectors abgedeckt werden. Dazu gehören [die Verwendung `fetch` oder REST-APIs](../develop/external-calls.md) oder das Abrufen von Informationen aus Ad-hoc-Datenquellen, z. B. eine [Teams adaptive Karte.](../resources/scenarios/task-reminders.md)

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Formatierung, Visualisierungen und programmgesteuerte Steuerung: Office Skripts

Wir empfehlen Office Skripts, wenn Ihre Anforderungen über den Datenimport und die Transformation hinausgehen.

Nahezu alles, was Sie manuell über die Excel Benutzeroberfläche ausführen können, ist mit Office Skripts möglich. Sie eignet sich hervorragend zum Anwenden einer konsistenten Formatierung auf Arbeitsmappen. Skripts erstellen Diagramme, PivotTables, Formen, Bilder und andere Arbeitsblattvisualisierungen. Skripts bieten Ihnen außerdem eine genaue Kontrolle über positionen, Größen, Farben und andere Attribute dieser Visualisierungen.

Durch die Einbindung von TypeScript-Code erhalten Sie ein hohes Maß an Anpassung. Programmgesteuerte Steuerungslogik wie `if...else` Anweisungen macht Ihr Skript robust. Auf diese Weise können Sie Beispielsweise Daten bedingt lesen, ohne sich auf komplexe Excel Formeln verlassen zu müssen, oder die Arbeitsmappe auf unerwartete Änderungen überprüfen, bevor Sie die Arbeitsmappe ändern.

Die Formatierung kann mit Power Query über Excel [Vorlagen](https://templates.office.com/power-query-tutorial-tm11414620)angewendet werden. Vorlagen werden jedoch auf Einzel- oder Organisationsebene aktualisiert, während Office Skripts eine präzisere Zugriffssteuerung bieten.

## <a name="power-automate-integrations"></a>Power Automate-Integrationen

Office Skripts bieten weitere Optionen für Power Automate Integration. Skripts sind auf Ihre Lösungen zugeschnitten. Sie definieren die [Eingabe und Ausgabe des Skripts,](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)damit es mit jedem anderen Connector oder mit anderen Daten im Fluss funktioniert. Der folgende Screenshot zeigt ein Beispiel Power Automate Flusses, der Daten von einer Teams adaptiven Karte an ein Office Skript übergibt.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Screenshot des Excel Online (Business)-Connectors im Flow-Designer. Der Connector verwendet die Skriptaktion ausführen, um Eingaben von einer Teams adaptiven Karte zu übernehmen und für ein Skript bereitzustellen.":::

Power Query wird im [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate Connector verwendet. Mithilfe der Power [Query-Aktion "Daten transformieren"](/connectors/sql/#transform-data-using-power-query) können Sie eine Abfrage in Power Automate erstellen. Obwohl dies ein leistungsstarkes Tool für die Verwendung mit SQL Server ist, beschränkt es Power Query auf diese Eingabequelle, wie im folgenden Fluss-Screenshot dargestellt.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Screenshot der SQL Server-Verbindung im Flow-Designer. Der Connector verwendet die Transformationsdaten mithilfe der Power Query-Aktion.":::

## <a name="platform-dependencies"></a>Plattformabhängigkeiten

Office Skripts sind derzeit nur für Excel im Web verfügbar. Power Query ist derzeit nur für Excel auf dem Desktop verfügbar. Beide können über Power Automate verwendet werden, sodass der Fluss mit Excel Arbeitsmappen funktioniert, die in OneDrive gespeichert sind.

## <a name="see-also"></a>Siehe auch

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query mit Excel](https://powerquery.microsoft.com/excel/)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
