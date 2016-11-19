# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>Excel-Web-Add-In zum Bearbeiten der Tabellen- und Diagrammformatierung

Erfahren Sie, wie Sie in Excel-Web-Add-Ins Tabellen und Diagramme programmgesteuert formatieren und Daten in eine Tabelle importieren. Vergleichen Sie, wie diese Aufgaben im VSTO-Add-In [Tabellen und Diagramme](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3) ausgeführt werden. Dieses Excel-Web-Add-In veranschaulicht auch die Verwendung der Entwurfsbeispiele aus [Office Add-in UX Design Patterns Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 

## <a name="table-of-contents"></a>Inhalt
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [In diesem Add-In verwendete Entwurfsvorlagen](#design-templates-used-in-this-add-in)
* [Abrufen der PickADate-Bibliothek](get-the-pickadate-library)
* [Ausführen des Projekts](#run-the-project)
* [Vergleichen des Codes dieses Excel-Web-Add-Ins mit dem VSTO-Add-In-Beispiel](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

2. November 2016:

* Ursprüngliche Version

## <a name="prerequisites"></a>Voraussetzungen

* Excel 2016 für Windows (Build 16.0.6727.1000 oder höher), Excel Online oder Excel für Mac (Build 15.26 oder höher)
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>In diesem Add-In verwendete Entwurfsvorlagen

- Zielseite
- Markenleiste
- Registerkartenleiste
- Einstellungen

Weitere Informationen zu den Entwurfsmustern finden Sie unter [UX-Entwurfsmustervorlagen für Office-Add-Ins](https://dev.office.com/docs/add-ins/design/ux-design-patterns). Informationen zu Beispielimplementierungen finden Sie unter [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="get-the-pickadate-library"></a>Abrufen der PickADate-Bibliothek

Das Datumsauswahl-Steuerelement aus der Office-Fabric ist von der Bibliothek PickADate abhängig. Gehen Sie wie folgt vor, *nachdem* Sie das Beispiel heruntergeladen haben.

1. Laden Sie die Version 3.5.3 der Bibliothek aus [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3) herunter. 
2. Entpacken Sie das Paket, und navigieren Sie zum Ordner `\pickadate.js-3.5.3\lib`. 
3. Kopieren Sie alle Dateien und Ordner aus diesem Ordner *außer* den Ordnern `compressed` und `themes-source` in den Projektordner: `SalesTrackerWeb\Scripts\PickADate`.

## <a name="run-the-project"></a>Ausführen des Projekts

1. Öffnen Sie die Visual Studio-Projektmappe. 
2. Drücken Sie **F5**. 
3. Wenn Excel geöffnet wird, klicken Sie auf die Schaltfläche **Track sales** am rechten Ende des Menübands **Start**. Das Add-In wird in einem Aufgabenbereich geöffnet.

#### <a name="import-data"></a>Importieren von Daten

4. Geben Sie auf der Seite **Home** im Feld **Product name** einen der folgenden Produktnamen ein (Groß-/Kleinschreibung wird beachtet): **Keyboard**, **Mouse**, **Monitor**, **Laptop**.
5. Wählen Sie mit dem Datumsauswahl-Steuerelement ein Datum aus, das nicht später ist als der 16. September 2016, da die Beispieldaten keine Verkäufe nach diesem Datum enthalten.
6. Klicken Sie auf die Schaltfläche **Get sales data**. Nach ein paar Sekunden wechselt der Fokus in der Arbeitsmappe auf ein neues Arbeitsblatt **Sales**. 

#### <a name="change-table-settings"></a>Ändern von Tabelleneinstellungen

1. Wählen Sie **Tabelle** auf der Registerkartenleiste. 
2. Deaktivieren Sie Optionsfelder nach Bedarf, um die entsprechenden Spalten auszublenden.
3. Wählen Sie eine Farbe für die Tabelle aus.

#### <a name="change-chart-settings"></a>Ändern von Diagrammeinstellungen

1. Wählen Sie **Diagramm** auf der Registerkartenleiste. 
2. Wählen Sie eine Datenquelle für das Diagramm aus.
3. Wählen Sie einen Diagrammtyp aus.
4. Wählen Sie ein Diagrammfarbschema aus.

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>Vergleichen des Codes dieses Excel-Web-Add-Ins mit dem VSTO-Add-In-Beispiel

Der Code, der die JavaScript-APIs für Office und Word verwendet, ist in Home.js und Helpers.js enthalten. Sämtliche Formatierung erfolgt mit HTML5 und den Stylesheet-Dateien: settings.css, tab.bar.css und mehreren Office-Fabric-CSS-Dateien.

Vergleichen Sie diesen Code mit dem Code in [Tabellen und Diagramme](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Beachten Sie Folgendes:


- Excel-Web-Add-Ins werden auf mehreren Plattformen wie Windows, Mac und Office Online unterstützt. VSTO-Add-Ins werden nur unter Windows unterstützt.
- Das Ändern des Formats einer Tabelle erfolgt in VSTO- und Web-Add-Ins auf ähnliche Weise. In beiden Fällen weist der Code einer Eigenschaft einen Formatvorlagennamen wie z. B. TableStyleMedium3 zu. Im VSTO-Add-In ist dies ein Tabellenobjekt. (Siehe die verschiedenen `*_CheckedChanged`-Methoden in der Datei TableAndChartPane.cs.) In diesem Web-Add-In befindet sich die Eigenschaft in einem JavaScript-Objekt, das an die `setTableOptionsAsync`-Funktion übergeben wird. (Siehe die `setTableColor`-Funktion in der Datei helpers.js.)
- Das Umschalten der Sichtbarkeit von Spalten in einer Tabelle ist in den VSTO- und Web-Add-Ins sehr ähnlich. Vergleichen Sie die Logik der `ListObjectHeaders_Click`-Methode im VSTO-Beispiel mit der `toggleColumnVisibility`-Funktion im Web-Add-In.
- Zum Ändern der Formatvorlage eines Diagramms in einem VSTO-Add-In weisen Sie der `ChartStyle`-Eigenschaft des Diagrammobjekts eine ganze Zahl zu. Die ganze Zahl bezieht sich auf eine Sammlung von Formatvorlageneinstellungen. (Siehe die Datei TableAndChartPane.cs im VSTO-Add-In.) In einem Excel-Web-Add-In ersetzen Sie das Diagramm durch ein neues mit der gewünschten Formatvorlage. Sie können die aktuellen Formatvorlageneinstellungen für das Diagramm in einem JavaScript-Objekt aufzeichnen, wie es in diesem Beispiel in der Datei home.js geschieht. Zum Ändern einer einzelnen Formatvorlageneinstellung ändern Sie diese in Ihrem Code im settings-Objekt, das dann in helpers.js an die `changeChart`-Funktion übergeben wird.
- Das Ändern eines Diagramm*typs*, wie z. B. Line, Area oder ClusteredColumn, in im VSTO-Add-In und in diesem Web-Add-In sehr ähnlich. In beiden Fällen wird eine `switch-case`-Struktur verwendet, um einer Typeigenschaft einen Wert zuzuweisen. Vergleichen Sie die `ChartStyleComboBox_SelectedIndexChanged`-Methode im VSTO-Beispiel mit der `setChartType`-Methode in diesem Webbeispiel. 
- Wenn Sie in einem Excel-Web-Add-In wie diesem einen *einzelnen* Wert über einen Zeitraum (oder eine horizontale Achse) nachverfolgen möchten, muss das Diagramm aus einer Tabelle mit nur zwei sichtbaren Spalten erstellt werden: eine stellt die horizontale Achse bereit (in diesem Fall Datumsangaben) und eine zweite enthält den Wert, der im Diagramm angezeigt wird. Aus diesem Grund erstellt das Add-In ein *ausgeblendetes* Arbeitsblatt mit einer Kopie der Verkaufsdatentabelle. Diese Tabelle hat nur zwei sichtbare Spalten: die **Date**-Spalte und die Spalte mit der ausgewählten Datenquelle für das Diagramm. Obwohl das Diagramm auf dem Arbeitsblatt **Sales** neben der Tabelle angezeigt wird, werden die Daten aus der Tabelle auf dem ausgeblendeten (temporären) Arbeitsblatt abgerufen.
- Zum Ändern der Datenquelle für ein Diagramm im Excel-Web-Add-In wird im Code die Sichtbarkeit der Spalten in der ausgeblendeten Tabelle umgeschaltet. Siehe die `setChartDataSource`-Methode in der Datei helpers.js. In einem VSTO-Add-In gibt der Code an, welche Spalte in der Tabelle als Datenquelle verwendet wird, indem die `SetSourceData`-Funktion des Diagrammobjekts aufgerufen wird. Siehe die `chartDataSourceComboBox_SelectedIndexChanged`-Methode.
- In Office-Web-Add-Ins können Sie HTML5, JavaScript und CSS nutzen, um reichhaltige Benutzeroberflächen wie die Benutzeroberfläche in diesem Codebeispiel zu erstellen. 
- Da Office-Web-Add-Ins asynchrone Methodenaufrufe verwenden, wird die Benutzeroberfläche nie blockiert.
- Office-Web-Add-Ins verwenden AJAX-Aufrufe zum Abrufen von Daten von Onlinedienstanbietern. Dieses Beispiel ruft einfach JSON-Daten aus einer lokalen JSON-Datei ab. Siehe die `getSalesData`-Methode in der Datei Helpers.js. VSTO-Add-Ins verwenden einen WebClient in C#, um auf Onlineressourcen zuzugreifen. Siehe die `GetDataUpdatesFoOneDataSource`-Methode in TableAndChartPane.cs.   


## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich dieses Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Wenn Ihre Frage die Office JavaScript-APIs betrifft, sollte die Frage mit [office-js] und [API] kategorisiert sein.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* [Dokumentation zu Office-Add-Ins](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Office Dev Center](http://dev.office.com/)
* Weitere Office-Add-In-Beispiele unter [OfficeDev auf Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Alle Rechte vorbehalten.

