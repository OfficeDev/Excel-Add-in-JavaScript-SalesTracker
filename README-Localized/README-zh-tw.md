# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>處理表格和圖表格式設定的 Excel Web 增益集

了解如何在 Excel Web 增益集中以程式控制方式格式化表格和圖表，以及如何將資料匯入試算表。在 VSTO 增益集的[表格和圖表](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3)中，對這些工作的完成方法進行比較。此 Excel Web 增益集也會顯示如何使用 [Office 增益集 UX 設計模式程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 的設計範例。 

## <a name="table-of-contents"></a>目錄
* [變更歷程記錄](#change-history)
* [必要條件](#prerequisites)
* [此增益集使用的設計範本](#design-templates-used-in-this-add-in)
* [取得 PickADate 程式庫](get-the-pickadate-library)
* [執行專案](#run-the-project)
* [比較 Excel Web 增益集程式碼和 VSTO 增益集範例](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## <a name="change-history"></a>變更歷程記錄

2016 年 11 月 2 日：

* 初始版本。

## <a name="prerequisites"></a>必要條件

* Excel 2016 for Windows (組建 16.0.6727.1000 或更新版本)、Excel Online 或 Excel for Mac (組建 15.26 或更新版本)。
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>此增益集使用的設計範本

- 登陸頁面
- 商標列
- 索引標籤列
- 設定

如需有關設計模式的詳細資訊，請參閱 [Office 增益集的 UX 設計模式範本](https://dev.office.com/docs/add-ins/design/ux-design-patterns)。欲取得實作範例，請參閱 [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。

## <a name="get-the-pickadate-library"></a>取得 PickADate 程式庫

Office Fabric 日期選擇器控制項對 PickADate 程式庫具有相依性。下載這個範例*之後*，請執行下列步驟。

1. 從 [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3) 下載 3.5.3 版的程式庫。 
2. 解壓縮套件，並瀏覽到 `\pickadate.js-3.5.3\lib` 資料夾。 
3. 複製在該資料夾中的所有檔案和資料夾 (*除了* `compressed` 和 `themes-source` 資料夾) 到專案資料夾︰`SalesTrackerWeb\Scripts\PickADate`。

## <a name="run-the-project"></a>執行專案

1. 開啟 Visual Studio 解決方案檔案。 
2. 按 [F5]。 
3. 當 Excel 開啟時，在 [首頁] 功能區右端上，按一下 [追蹤銷售] 按鈕。增益集會在工作窗格中開啟。

#### <a name="import-data"></a>匯入資料

4. 在 [首頁] 頁面上的 [產品名稱] 方塊中，輸入下列其中一種產品名稱 (區分大小寫)︰[鍵盤]、[滑鼠]、[監視器]、[膝上型電腦]、
5. 超過 2016 年 9 月 16 日之後沒有銷售範例資料，所以請使用日期選擇器控制項選取一個在此日之前的日期。
6. 選取 [取得銷售資料] 按鈕。幾秒鐘之後，活頁簿將焦點切換至新的**銷售**工作表。 

#### <a name="change-table-settings"></a>變更表格設定

1. 在索引標籤列上選取 [表格]。 
2. 如需隱藏對應欄位，請取消選取選項按鈕。
3. 選取表格的色彩。

#### <a name="change-chart-settings"></a>變更圖表的設定

1. 在索引標籤列上選取 [圖表]。 
2. 選取圖表的資料來源。
3. 選取圖表類型。
4. 選取圖表色彩主題。

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>比較 Excel Web 增益集程式碼和 VSTO 增益集範例

使用 Office 和  Word JavaScript API 的程式碼是在 Home.js 和 Helpers.js。所有的樣式是以 HTML5 與樣式表檔案完成︰settings.css、tab.bar.css 和數個 Office Fabric css 檔案。

比較這個程式碼和[表格和圖表](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3)中的程式碼。請注意下列事項：


- Excel Web 增益集支援跨多個平台，包括視窗、Mac 和 Office Online。只有 Windows 支援 VSTO 增益集。
- 變更表格樣式的方法與在 VSTO 和 Web 增益集類似。在這兩種情況下，您的程式碼會指派樣式名稱，例如指派 TableStyleMedium3 給屬性。在 VSTO 增益集中，此為表格物件。(請參閱 TableAndChartPane.cs 檔案中的各種 `*_CheckedChanged` 方法。)在此 Web 增益集，屬性位於 JavaScript 物件內，而物件則會傳遞至 `setTableOptionsAsync` 函數。(請參閱 helpers.js 檔案中的 `setTableColor` 函數。)
- 切換表格欄位可見性的方法與在 VSTO 和 Web 增益集非常類似。對 VSTO 範例中的 `ListObjectHeaders_Click` 方法邏輯和在 Web 增益集中的 `toggleColumnVisibility` 函數進行比較。
- 若要變更 VSTO 增益集中的圖表樣式，您必須指派整數給圖表物件的 `ChartStyle` 屬性。整數是指樣式設定的集合。參閱 VSTO 增益集中的 TableAndChartPane.cs 檔案。在 Excel Web 增益集中您會以新的圖表取代原有的圖表，該新圖表有您想要的樣式。您可以在 JavaScript 物件中的圖表記錄目前的樣式設定，方式如同在 home.js 檔案中的範例所示。若要變更單一樣式設定，您的程式碼會在設定物件中變更設定，然後設定會傳遞至 helpers.js 中的 `changeChart` 函數。
- 變更圖表*類型*的方法 (例如線條、區域、ClusteredColumn) 與在 VSTO 增益集和此 Web 增益集非常類似。在這兩種情況下，`switch-case` 結構是用來將值指派給類型屬性。比較 VSTO 範例中的 `ChartStyleComboBox_SelectedIndexChanged` 方法和此 Web 範例中的 `setChartType` 方法。 
- 如同本範例，在 Excel Web 增益集中，當您想要在不同時間 (或任何水平方向軸) 追蹤*單一*值，圖表必須建置出只有兩個可見欄位的表格；第一個欄位提供水平軸 (本例為日期)，而第二個欄位則提供圖表中所顯示的值。基於這個理由，增益集建立*隱藏*工作表，包含了銷售資料表格副本。此表格只能看到兩個欄位；**日期**欄位，以及包含圖表選定資料來源的欄位。雖然圖表會出現在表格旁邊的**銷售**工作表，它會從隱藏的 (暫存) 工作表上取得表格資料。
- 若要變更在 Excel Web 增益集中的圖表資料來源，您的程式碼會切換隱藏表格上的欄位可見性。請參閱 helpers.js 檔案中的 `setChartDataSource` 方法。在 VSTO 增益集中，您的程式碼會藉由呼叫圖表物件的 `SetSourceData` 函數，指定哪一個表格中的欄位是做為資料來源。請參閱 `chartDataSourceComboBox_SelectedIndexChanged` 方法。
- 在 Office Web 增益集中，您可以利用 HTML5、JavaScript 和 CSS 讓 UI 更加豐富，就像程式碼範例中的 UI 一樣。 
- 因為 Office Web 增益集進行非同步方法呼叫，UI 就永遠不會封鎖。
- Office Web 增益集進行 AJAX 呼叫，以擷取線上服務提供者的資料。此範例只會從本機的 JSON 檔案提取 JSON 資料。請參閱 helpers.js 中的 `getSalesData` 方法。VSTO 增益集在 C# 中使用 WebClient 存取線上資源。請參閱 TableAndChartPane.cs 中的 `GetDataUpdatesFoOneDataSource` 方法。   


## <a name="questions-and-comments"></a>問題和建議

我們很樂於收到您對於此範例的意見反應。您可以在此儲存機制的 [問題]** 區段中，將您的意見反應傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) 提出有關 Microsoft Office 365 開發的一般問題。如果您的問題是關於 Office JavaScript API，請確定您的問題標記有 [office js] 與 [API]。

## <a name="additional-resources"></a>其他資源

* [Office 增益集文件](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Office 開發人員中心](http://dev.office.com/)
* 在 [Github 上的 OfficeDev](https://github.com/officedev) 中有更多 Office 增益集範例

## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft Corporation 著作權所有，並保留一切權利。

