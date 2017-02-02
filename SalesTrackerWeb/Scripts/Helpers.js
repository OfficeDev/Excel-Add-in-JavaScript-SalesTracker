/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
    "use strict";

    // Create a namespace to hold application-wide settings with primitive data types.
    var SalesTrackerApp = window.SalesTrackerApp || {};

    SalesTrackerApp.deleteWorksheets = function (worksheetsBefore, strSalesSheet, strTempWorksheetName) {

        // Delete the sales or temp worksheet if they already exist.
        for (var i = 0; i < worksheetsBefore.items.length; i++) {
            if ((worksheetsBefore.items[i].name == strSalesSheet) || (worksheetsBefore.items[i].name == strTempWorksheetName)) {
                worksheetsBefore.items[i].delete();
            }
        }
    }

        // Ensures that all table data fields are selected in the UI. 
    SalesTrackerApp.reinitializeUI = function () {    
      
        $.each(SalesTrackerApp.CheckBoxElements, function () {
            this.check();  // Use the check method on the Office UI Fabric Checkbox component to check the checkbox.
        });

        var colorRadios = $('.colorChoiceRadio');  // Uses CSS class to identify the radio buttons.
        colorRadios.each(function () {
            var _this = $(this);
            _this.removeClass('is-checked');
        });
        $("#colorChoice-blue").siblings('label').addClass('is-checked');
        
    }

    SalesTrackerApp.setChartColorSettings = function (chartColor, fontColor, catFontColor, valFontColor, chartSettings) {

        // Saves chart colors.
        chartSettings.solidColor = chartColor;
        chartSettings.fontColor = fontColor;
        chartSettings.categoryAxisFontColor = catFontColor;
        chartSettings.valueAxisFontColor = valFontColor;
    }

    SalesTrackerApp.changeChartColor = function (chart, chartSettings) {

        // Changes the chart colors.
        chart.format.fill.setSolidColor(chartSettings.solidColor);
        chart.title.format.font.color = chartSettings.fontColor;
        chart.axes.categoryAxis.format.font.color = chartSettings.categoryAxisFontColor;
        chart.axes.valueAxis.format.font.color = chartSettings.valueAxisFontColor;
    }

    SalesTrackerApp.setChartColorTheme = function (event) {

        var themeChoice = event.data.chartColor;

        Excel.run(function (ctx) {

            // Get the active worksheet and load its charts.
            var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeWorksheet.load('charts');
            return ctx.sync()

            // Then change color theme of the chart.
            .then(function () {
                var chart = activeWorksheet.charts.getItem(event.data.chartSettings.name);

                switch (themeChoice) {
                    case 'white':
                        SalesTrackerApp.setChartColorSettings('White', "#41AEBD", '#000000', '#000000', event.data.chartSettings);
                        SalesTrackerApp.changeChartColor(chart, event.data.chartSettings);
                        break;
                    case 'gray':
                        SalesTrackerApp.setChartColorSettings('DarkSlateGray', "#FFFFFF", '#FFFFFF', '#FFFFFF', event.data.chartSettings);
                        SalesTrackerApp.changeChartColor(chart, event.data.chartSettings);
                        break;
                    case 'blue':
                        SalesTrackerApp.setChartColorSettings('RoyalBlue', "#FFFFFF", '#FFFFFF', '#FFFFFF', event.data.chartSettings);
                        SalesTrackerApp.changeChartColor(chart, event.data.chartSettings);
                        break;
                    default:
                        throw new Error(themeChoice + " is not a valid choice.");
                }
                return ctx.sync();
            });
        })
        .catch(SalesTrackerApp.errorHandler);
    }

    // This function is called to set up the chart on first data load, and to reinitialize the chart 
    // when users change chart options in the chart UI.
    SalesTrackerApp.changeChart = function (SalesHistory, tempTable, strDateColumn, chartSettings) {

        // Queue commands to add a new chart and format it.
        var dateColumnRange = tempTable.columns.getItem(strDateColumn).getDataBodyRange();
        var dataSourceColumnRange = tempTable.columns.getItem(chartSettings.dataSourceDisplayed).getDataBodyRange();

        var chartDataRange = dateColumnRange.getBoundingRect(dataSourceColumnRange);
        var chart = SalesHistory.charts.add(chartSettings.type, chartDataRange, Excel.ChartSeriesBy.auto);
        chart.setPosition(chartSettings.upperLeftCorner, chartSettings.lowerRightCorner);
        chart.title.text = chartSettings.title;
        SalesTrackerApp.changeChartColor(chart, chartSettings);
        chart.name = chartSettings.name;
    }

    SalesTrackerApp.setChartDataSource = function (event) {

        var sourceChoice = event.data.chartDataSource || 'totalProductSales';

        Excel.run(function (ctx) {

            // First get the active worksheet.
            var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeWorksheet.load('charts');
            var tempDataSheet = ctx.workbook.worksheets.getItem(event.data.strTempWorksheetName);

            return ctx.sync()
            .then(function () {

                var existingChart = activeWorksheet.charts.getItem(event.data.chartSettings.name);
                existingChart.delete();

                var tempTable = ctx.workbook.tables.getItem(event.data.strTempTableName);
                var SalesHistory = ctx.workbook.worksheets.getItem(event.data.strSalesSheet);

                // Change the chart's data source, and then show/hide columns in tempTable. The chart uses tempTable as its 
                // data source. If several columns are visible at once in tempTable, the chart shows several series.

                tempTable.columns.getItem(event.data.chartSettings.dataSourceDisplayed).getDataBodyRange().columnHidden = true;

                switch (sourceChoice) {

                    case 'totalNumberOfUnits':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strTotalNumberOfUnits;
                        break;

                    case 'averageUnitPrice':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strAverageUnitPriceColumn;
                        break;

                    case 'totalTax':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strTotalTaxColumn;
                        break;

                    case 'totalDiscount':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strTotalDiscountColumn;
                        break;

                    case 'totalProductSales':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strTotalProductSalesColumn;
                        break;

                    case 'totalServiceSales':
                        event.data.chartSettings.dataSourceDisplayed = event.data.strTotalServiceSalesColumn;
                        break;
                    default:
                        throw new Error(sourceChoice + " is not a valid choice.");
                }
                tempTable.columns.getItem(event.data.chartSettings.dataSourceDisplayed).getDataBodyRange().columnHidden = false;

                SalesTrackerApp.changeChart(SalesHistory, tempTable, event.data.strDateColumn, event.data.chartSettings);
                return ctx.sync();
            });
        })
        .catch(SalesTrackerApp.errorHandler);
    }

    SalesTrackerApp.setChartType = function (event) {

        var typeChoice  = event.data.ctype || 'line';

        Excel.run(function (ctx) {

            // Get the active worksheet.
            var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeWorksheet.load('charts');

            return ctx.sync()

                // Then delete the old chart and create new one of the chosen type.
                .then(function () {

                    var existingChart = activeWorksheet.charts.getItem(event.data.chartSettings.name);
                    existingChart.delete();

                    var tempTable = ctx.workbook.tables.getItem(event.data.strTempTableName);
                    var SalesHistory = ctx.workbook.worksheets.getItem(event.data.strSalesSheet);

                    // Queue commands to add a new chart and format it.
                    switch (typeChoice) {
                        case 'line':
                            event.data.chartSettings.type = "Line";
                            break;
                        case 'column':
                            event.data.chartSettings.type = "ColumnClustered";
                            break;
                        case 'area':
                            event.data.chartSettings.type = "Area";
                            break;
                        default:
                            throw new Error(typeChoice + " is not a valid choice.");
                    }
                    SalesTrackerApp.changeChart(SalesHistory, tempTable, event.data.strDateColumn, event.data.chartSettings);
                    return ctx.sync();
                }); 
        })
        .catch(SalesTrackerApp.errorHandler);
    }

    SalesTrackerApp.setSelectedDate = function () {

        // Convert date from date picker format to ISO8601 format.
        var date = new Date($(this).val());
        date = date.toISOString();
        date = date.split("T")[0];

        SalesTrackerApp.pStartDate = date;
    }

    SalesTrackerApp.setTableColor = function (event) {

        var colorChoice = event.data.color || 'blue';

        // Need to create a binding for the table before you can change its style properties.
        var fullyQualifiedTableName = event.data.strSalesSheet + "!" + event.data.tableName;

        Office.context.document.bindings.addFromNamedItemAsync(fullyQualifiedTableName, "table", { id: 'salesTable' }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log('Action failed. Error: ' + asyncResult.error.message);
            } else {
                console.log('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);

                // Map color choice to one of the built-in Excel table styles. These values come from
                // the Excel UI. The style picker on the ribbon has style names on the pattern 
                // "<color>, Table Style <weight> <integer>"; for example "White, Table Style Medium 1".
                // To get the programmatic equivalent, ignore the color and comma and remove the spaces,
                // so this example would be "TableStyleMedium1". (The color is built into the style
                // definition.)
                var styleNumber = '';
                switch (colorChoice) {
                    case 'black':
                        styleNumber = '1';
                        break;
                    case 'blue':
                        styleNumber = '2';
                        break;
                    case 'gray':
                        styleNumber = '4';
                        break;
                    case 'green':
                        styleNumber = '7';
                        break;
                    case 'orange':
                        styleNumber = '3';
                        break;
                    default:
                        throw new Error(colorChoice + " is not a valid choice.");
                }
                var styleName = "TableStyleMedium" + styleNumber;

                // Set a new table format.
                var newTableFormat = { style: styleName };
                Office.select("bindings#salesTable").setTableOptionsAsync(newTableFormat, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log('Action failed. Error: ' + asyncResult.error.message);
                    }
                });

            }
        })
    }

    SalesTrackerApp.toggleColumnVisibility = function (event) {

        // Get ID of column checkbox that was changed.
        var columnName = event.target.id;

        Excel.run(function (ctx) {

            // Change the column object to a range object that has a visibility property.
            var columnRange = ctx.workbook.tables.getItem(event.data.tableName).columns.getItem(columnName).getDataBodyRange();

            columnRange.load('columnHidden');
            return ctx.sync()
                .then(function () {
                    if (columnRange.columnHidden === true) {
                        columnRange.columnHidden = false;
                        columnRange.format.autofitColumns();
                    }
                    else {
                        columnRange.columnHidden = true;
                    }
                    return ctx.sync();
                });
        })
        .catch(SalesTrackerApp.errorHandler);
    }

    // Navigates to a different DIV when users click on the tab bar.
    SalesTrackerApp.newPage = function () {
        if (this.id === 'DateTab') {
            $('.DateSettings').show().siblings().hide();
        }
        else if (this.id === 'TableTab') {
            $('.TableSettings').show().siblings().hide();
        }
        else {
            $('.ChartSettings').show().siblings().hide();
        }
    };

    SalesTrackerApp.getTodayAsISO8601Date = function () {
        var today = new Date();
        var pYear = today.getFullYear();
        var pMonth = today.getMonth() + 1;
        var pDate = today.getDate();

        // Convert date to ISO8601 format.
        if (pDate < 10) {
            pDate = '0' + pDate;
        }
        if (pMonth < 10) {
            pMonth = '0' + pMonth;
        }

        var pEndDate = '' + pYear + '-' + pMonth + '-' + pDate;
        return pEndDate;
    }


    // Gets sales data, and displays the data in a table and chart.
    SalesTrackerApp.getSalesData = function (event) {    

        // Initially, the table and chart options are hidden. 
        document.getElementById("ChartTab").hidden = false;
        document.getElementById("TableTab").hidden = false;

        var pEndDate = SalesTrackerApp.getTodayAsISO8601Date();    

        $.ajax({
            url: 'SampleSalesData.json',
            dataType: "json"
        })
        // Process returned data.
        .then(function (data) {
            return Excel.run(function (ctx) {

                // Define variables that are global to the Excel.run,
                // so they are accessible to subsequent "then" calls.
                var SalesHistory;
                var tempDataSheet;
                var tempTable;
                var startRowNumber;
                var endRowNumber;
                var columnHeaderNames;
                var masterTableAddress;
                var masterTable;
                var tempTableAddress;

                // Load the worksheets.
                var worksheetsBefore = ctx.workbook.worksheets;
                worksheetsBefore.load();
                return ctx.sync()

                // Then, if data was previously loaded, first we need to delete those worksheets, and
                // reset the UI.
                .then(function () {                
                    SalesTrackerApp.deleteWorksheets(worksheetsBefore, event.data.strSalesSheet, event.data.strTempWorksheetName);
                    if (event.data.isDataLoaded) {
                        SalesTrackerApp.reinitializeUI();
                    }
                })
                .then(ctx.sync)

                // Then (re)create the worksheets.
                .then(function () {
                    SalesHistory = ctx.workbook.worksheets.add(event.data.strSalesSheet);

                    // tempDataSheet is a hidden sheet with data for the chart.
                    tempDataSheet = ctx.workbook.worksheets.add(event.data.strTempWorksheetName);
                })
                .then(ctx.sync)

                // Then, create and populate the main and temporary tables.
                .then(function () {
                    startRowNumber = 2;
                    endRowNumber = startRowNumber;
                    columnHeaderNames = [[event.data.strDateColumn,
                        event.data.strTotalNumberOfUnits, event.data.strAverageUnitPriceColumn, event.data.strTotalTaxColumn,
                        event.data.strTotalDiscountColumn, event.data.strTotalProductSalesColumn,
                        event.data.strTotalServiceSalesColumn]];

                    masterTableAddress = event.data.strSalesSheet + '!B' + startRowNumber + ':H' + startRowNumber;
                    masterTable = SalesTrackerApp.createTable(ctx, event.data.tableName, masterTableAddress, columnHeaderNames);

                    tempTableAddress = event.data.strTempWorksheetName + '!A1:G1';
                    tempTable = SalesTrackerApp.createTable(ctx, event.data.strTempTableName, tempTableAddress, columnHeaderNames)

                    // Add JSON data to an array, and then assign the array to the table.
                    var tableBodyData = [];

                    data.SalesTransactions.forEach(function (sale) {
                        if (sale.Date >= SalesTrackerApp.pStartDate) // Filter JSON data by date.
                        {
                            // If the user does NOT enter a product name,  add all values to the array.
                            // Otherwise, filter the JSON data by product name.  
                            if (document.getElementById("product-name").value.length == 0
                                ||
                                ((document.getElementById("product-name").value.length > 0) && (sale.Product == document.getElementById("product-name").value))) {
                                if (Office.context.requirements.isSetSupported('ExcelApi', 1.2)) {
                                    // If ExcelApi 1.2 is supported, then we can bulk add rows - first to an array, and then to the table.
                                    tableBodyData.push([sale.Date, sale.TotalNumberOfUnits, sale.AverageUnitPrice, sale.TotalTax, sale.TotalDiscount, sale.TotalProductSales, sale.TotalServiceSales]);
                                }
                                else {
                                    // Rows must be added to the table one at a time.
                                    masterTable.rows.add(null, [[sale.Date, sale.TotalNumberOfUnits, sale.AverageUnitPrice, sale.TotalTax, sale.TotalDiscount, sale.TotalProductSales, sale.TotalServiceSales]]);
                                    tempTable.rows.add(null, [[sale.Date, sale.TotalNumberOfUnits, sale.AverageUnitPrice, sale.TotalTax, sale.TotalDiscount, sale.TotalProductSales, sale.TotalServiceSales]]);
                                }
                            }
                            endRowNumber++;
                        }
                    });

                    if (tableBodyData.length == 0) {
                        $("#txtNoDataFound").text('No data matches the values entered. Please try again. For example, enter Keyboard.');
                    }
                    else {
                        $("#txtNoDataFound").text('Click on Table or Charts to change table or chart display options.');
                    }

                    if (Office.context.requirements.isSetSupported('ExcelApi', 1.2)) {
                        masterTable.rows.add(null, tableBodyData);
                        tempTable.rows.add(null, tableBodyData);
                    }
                        
                    // Queue a command to sort by date in descending order.  
                    var sortRange = masterTable.getDataBodyRange();
                    sortRange.sort.apply([
                        {
                            key: 0,
                            ascending: false,
                        },
                    ]);
                })
                .then(ctx.sync)
                        
                // Then create the chart.
                .then(function () {

                    // Queue command to add a new chart and format it.
                    SalesTrackerApp.changeChart(SalesHistory, tempTable, event.data.strDateColumn, event.data.chartSettings);

                    // Activate the worksheet so the table on it can be accessed.
                    SalesHistory.activate();
                })
                .then(ctx.sync)

                // Finally, format main table columns, suppress non-data-source columns in temp table, 
                // and hide the temp worksheet.
                .then(function () {

                    var rangeName = "D3:H" + (endRowNumber);
                    var currencyFormat = "$#,##0.00";
                    SalesTrackerApp.formatRangeColumns(SalesHistory, rangeName, currencyFormat);

                    // To display just one data trend on a chart, the chart is based on a temporary table
                    // that has all other columns hidden. 
                    var columnsToHide = [
                        event.data.strTotalNumberOfUnits,
                        event.data.strTotalProductSalesColumn, event.data.strAverageUnitPriceColumn,
                        event.data.strTotalTaxColumn, event.data.strTotalDiscountColumn,
                        event.data.strTotalServiceSalesColumn
                    ];
                    SalesTrackerApp.suppressNonDataSourceColumns(tempTable, event.data.chartSettings.dataSourceDisplayed, columnsToHide);
                    tempDataSheet.visibility = 'hidden';
                })
                .then(function (data) {
                    console.log("Data has been loaded into worksheet");
                    event.data.isDataLoaded = true;
                });
            }) // end Excel.run
            .catch(SalesTrackerApp.errorHandler);
        })
        .fail(function (xhr, textStatus, errorThrown) {
            $('#product-data').append("Unable to get data: " + errorThrown);
        });
    }

    SalesTrackerApp.createTable = function (excelContext, tableName, tableAddress, columnHeaderNames) {
        var table = excelContext.workbook.tables.add(tableAddress, true);
        table.name = tableName;
        table.getHeaderRowRange().values = columnHeaderNames;
        return table;
    }

    SalesTrackerApp.formatRangeColumns = function (worksheet, rangeName, currencyFormat) {

        // Set the currency formats.
        var objRange = worksheet.getRange(rangeName);
        objRange.numberFormat = currencyFormat;

        // Set the columns to fit the data in them.
        worksheet.getUsedRange().getEntireColumn().format.autofitColumns();
        worksheet.getUsedRange().getEntireRow().format.autofitRows();
    }

    SalesTrackerApp.suppressNonDataSourceColumns = function (table, dataSourceDisplayed, columnsToHide) {

        // Hide all columns except the data source. "true" means hide.
        for (var i = 0; i < columnsToHide.length; i++) {        
            // Change the visibility on the temp table to control what gets shown on the chart.
            table.columns.getItem(columnsToHide[i]).getDataBodyRange().columnHidden = true;
        }
        // Change the visibility on the temp table to control what gets shown on the chart.
        table.columns.getItem(dataSourceDisplayed).getDataBodyRange().columnHidden = false;
    }

    SalesTrackerApp.errorHandler = function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    window.SalesTrackerApp = SalesTrackerApp;
})();
