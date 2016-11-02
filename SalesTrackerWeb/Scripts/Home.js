/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
    "use strict";

    // Create a namespace to hold application-wide settings with primitive data types.
    var SalesTrackerApp = window.SalesTrackerApp || {};

    // Need a global reference to the beginning of the date range for which data is needed. Set a default.
    SalesTrackerApp.pStartDate = '2015-09-01';

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {

        var tableName = 'SalesHistoryData';
        var strSalesSheet = "Sales";  // Worksheet that is displayed to the user with the table and chart.
        var strTempTableName = "TempChartData";  // Name of the table used to store temp data that the chart uses.
        var strTempWorksheetName = "temp";  // Name of the worksheet used to store temp data that the chart uses.
        var isDataLoaded = false;

        // Define table header strings.
        var strDateColumn = 'Date';
        var strTotalNumberOfUnits = 'UnitPrice';
        var strAverageUnitPriceColumn = 'AvgUnitPrice';
        var strTotalTaxColumn = 'Tax';
        var strTotalDiscountColumn = 'Discount';
        var strTotalProductSalesColumn = 'ProductSales';
        var strTotalServiceSalesColumn = 'ServiceSales';
       
        // Object used to store chart settings.
        var chartSettings = {
            dataSourceDisplayed: strTotalProductSalesColumn,
            type: "Line",
            solidColor: "White",
            fontColor: "#41AEBD",
            lineColor: "#2E81AD",
            categoryAxisFontColor: '#000000',
            valueAxisFontColor: '#000000',
            title: 'Total sales',
            name: 'SalesChart',
            upperLeftCorner: 'J2',
            lowerRightCorner: 'O20'
        }
        
        $(document).ready(function () {
            
            $('#button-text').text("Get sales data");
            $('#button-desc').text("Gets the sales data.");

            // Register event handlers.
            $('#get-sales-button').click(
                {
                    strTempWorksheetName: strTempWorksheetName, strTempTableName: strTempTableName, chartSettings: chartSettings,
                    strSalesSheet: strSalesSheet, strTotalNumberOfUnits: strTotalNumberOfUnits,
                    strAverageUnitPriceColumn: strAverageUnitPriceColumn, strTotalTaxColumn: strTotalTaxColumn,
                    strTotalDiscountColumn: strTotalDiscountColumn, strTotalProductSalesColumn: strTotalProductSalesColumn,
                    strTotalServiceSalesColumn: strTotalServiceSalesColumn, strDateColumn: strDateColumn,
                    isDataLoaded: isDataLoaded, tableName: tableName
                },
                SalesTrackerApp.getSalesData);
            $('#start-date').change(SalesTrackerApp.setSelectedDate);
            $('.column-selector').change(
                {
                    tableName: tableName
                },
                SalesTrackerApp.toggleColumnVisibility);
            $('[name=colorChoice]').click(
                {
                    tableName: tableName, strSalesSheet: strSalesSheet
                },
                SalesTrackerApp.setTableColor);
            $('[name=chartColorChoice]').click(
               {
                   chartSettings: chartSettings
               },
               SalesTrackerApp.setChartColorTheme);
            $('[name=chartType]').click(
                {
                    chartSettings: chartSettings, strTempTableName: strTempTableName,
                    strSalesSheet: strSalesSheet, strDateColumn: strDateColumn
                },
                SalesTrackerApp.setChartType);
            $('[name=chartDataSource]').click(
                {
                    strTempWorksheetName: strTempWorksheetName, strTempTableName: strTempTableName, chartSettings: chartSettings,
                    strSalesSheet: strSalesSheet, strTotalNumberOfUnits: strTotalNumberOfUnits,
                    strAverageUnitPriceColumn: strAverageUnitPriceColumn, strTotalTaxColumn: strTotalTaxColumn,
                    strTotalDiscountColumn: strTotalDiscountColumn, strTotalProductSalesColumn: strTotalProductSalesColumn,
                    strTotalServiceSalesColumn: strTotalServiceSalesColumn, strDateColumn: strDateColumn
                },
                SalesTrackerApp.setChartDataSource);
            $('#DateTab').click(SalesTrackerApp.newPage);
            $('#TableTab').click(SalesTrackerApp.newPage);
            $('#ChartTab').click(SalesTrackerApp.newPage);

            var DatePickerElements = document.querySelectorAll(".ms-DatePicker");
            for (var i = 0; i < DatePickerElements.length; i++) {
                new fabric['DatePicker'](DatePickerElements[i]);
            }

            document.getElementById("ChartTab").hidden = true;
            document.getElementById("TableTab").hidden = true;
        });      

    } 

    window.SalesTrackerApp = SalesTrackerApp;
})();


