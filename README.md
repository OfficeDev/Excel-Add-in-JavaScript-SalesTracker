---
topic: sample
products:
- office-excel
- office-365
languages:
- javascript
- csharp
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 9/26/2016 2:05:38 PM
---
# Excel Web Add-in for Manipulating Table and Chart Formatting

Learn how to programmatically format tables and charts, and how to import data to a spreadsheet, in Excel Web Add-ins. Compare with how these tasks are done in the [Tables and Charts](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3) VSTO Add-in. This Excel web add-in also shows how to use the design samples from [Office Add-in UX Design Patterns Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [Design templates used in this add-in](#design-templates-used-in-this-add-in)
* [Get the PickADate library](get-the-pickadate-library)
* [Run the project](#run-the-project)
* [Compare this Excel web add-in code with the VSTO add-in sample](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

November 2, 2016:

* Updated code sample to use Fabric JS 1.2.0
* Initial version.

## Prerequisites

* Excel 2016 for Windows (build 16.0.6727.1000 or later), Excel Online, or Excel for Mac (build 15.26 or later).
* Visual Studio 2015 

## Design templates used in this add-in

- Landing page
- Brand bar
- Tab bar
- Settings

For more information about the design patterns, see [UX design pattern templates for Office Add-ins](https://dev.office.com/docs/add-ins/design/ux-design-patterns). And for sample implementations, see [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## Get the PickADate library

The Office Fabric date picker control has a dependency on the PickADate library. Take the following steps *after* you have downloaded this sample.

1. Download version 3.5.3 of the library from [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3). 
2. Unzip the package and navigate to the `\pickadate.js-3.5.3\lib` folder. 
3. Copy all the files and folders in that folder, *except* the `compressed` and `themes-source` folders, to the project folder: `SalesTrackerWeb\Scripts\PickADate`.

## Run the project

1. Open the Visual Studio solution file. 
2. Press **F5**. 
3. When Excel opens, click the **Track sales** button on the right end of the **Home** ribbon. The add-in opens in a task pane.

#### Import data

4. On the **Home** page, enter one of the following product names (case sensitive) in the **Product name** box: **Keyboard**, **Mouse**, **Monitor**, **Laptop**,
5. Use the date picker control to pick a date no later than September 16th, 2016, because there are no sales after this date in the sample data.
6. Select the **Get sales data** button. After a few seconds, the workbook will switch focus to a new **Sales** worksheet. 

#### Change table settings

1. Select **Table** on the tab bar. 
2. Deselect radio buttons as needed, to hide the corresponding columns.
3. Select a color for the table.

#### Change chart settings

1. Select **Chart** on the tab bar. 
2. Select a data source for the chart.
3. Select a chart type.
4. Select a chart color theme.

## Compare this Excel web add-in code with the VSTO add-in sample

The code that uses the Office and Word JavaScript APIs is in Home.js and Helpers.js. All of the styling is done with HTML5 and the stylesheet files: settings.css, tab.bar.css and several Office Fabric css files.

Compare this code with the code in [Tables and Charts](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Note the following:


- Excel Web add-ins are supported across several platforms including Windows, Mac, and Office Online. VSTO add-ins are only supported on Windows.
- Changing the style of a table is similar in both VSTO and web add-ins. In both cases, your code assigns a style name, such as TableStyleMedium3 to a property. In the VSTO add-in, this is a table object. (See the various `*_CheckedChanged` methods in the TableAndChartPane.cs file.) In this web add-in, the property is in a JavaScript object that is passed to the `setTableOptionsAsync` function. (See the `setTableColor` function in the helpers.js file.)
- Toggling the visibility of columns in a table is very similar in the VSTO and web add-ins. Compare the logic of the `ListObjectHeaders_Click` method in the VSTO sample with the `toggleColumnVisibility` function in the web add-in.
- To change the style of a chart in a VSTO add-in, you assign an integer to the chart object's `ChartStyle` property. The integer refers to a collection of style settings. See the TableAndChartPane.cs file in the VSTO add-in. In an Excel web add-in you replace the chart with a new one that has the desired style. You can record the current style settings for the chart in a JavaScript object as this sample does in the home.js file. To change a single style setting, your code changes it in the settings object which is then passed to the `changeChart` function in helpers.js.
- Changing a chart *type*, such as Line, Area, ClusteredColumn, is very similar in both the VSTO add-in and this web add-in. In both cases a `switch-case` structure is used to assign a value to a type property. Compare the `ChartStyleComboBox_SelectedIndexChanged` method in the VSTO sample with the `setChartType` method in this web sample. 
- In an Excel web add-in, like this one, when you want to track a *single* value over time (or any horizontal axis), the chart must be built off of a table with only two visible columns; one that provides the horizontal axis (dates, in this case), and a second that provides the value that is being displayed in the chart. For this reason, the add-in creates a *hidden* worksheet with a copy of the sales data table. This table has only two visible columns; the **Date** column, and the column with the chosen data source for the chart. Although the chart appears on the **Sales** worksheet beside the table, it is getting it's data from the table on the hidden (temp) worksheet.
- To change the data source for a chart in the Excel web add-in, your code toggles the visibility of the columns on the hidden table. See the `setChartDataSource` method in the helpers.js file. In a VSTO add-in, your code specifies which column in the table is to be used as the data source by calling the chart object's `SetSourceData` function. See the `chartDataSourceComboBox_SelectedIndexChanged` method.
- In Office web add-ins, you can leverage HTML5, JavaScript and CSS to make rich UIs like the UI in this code sample. 
- Because Office web add-ins make asynchronous method calls, the UI never blocks.
- Office Web add-ins make AJAX calls to retrieve data from online service providers. This sample simply fetches JSON data from a local JSON file. See the `getSalesData` method in Helpers.js. VSTO add-ins use a WebClient in C# to access online resources. See the `GetDataUpdatesFoOneDataSource` method in TableAndChartPane.cs.   


## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Office Dev Center](http://dev.office.com/)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

## Copyright
Copyright (c) 2016 Microsoft Corporation. All rights reserved.



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
