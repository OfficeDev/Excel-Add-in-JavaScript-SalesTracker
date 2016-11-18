# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>用于操控表和图表格式化的 Excel Web 外接程序

了解如何在 Excel Web 外接程序中以程序化方式来设置表和图表的格式，以及如何将数据导入电子表格。对比一下，是如何在[表和图表](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3) VSTO 外接程序中完成这些任务的。此 Excel Web 外接程序还展示了如何使用 [Office 外接程序的用户体验设计模式代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)中的设计示例。 

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [此外接程序中使用的设计模板](#design-templates-used-in-this-add-in)
* [获取 PickADate 库](get-the-pickadate-library)
* [运行项目](#run-the-project)
* [比较此 Excel Web 外接程序代码与 VSTO 外接程序示例](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>修订记录

2016 年 11 月 2 日：

* 首版。

## <a name="prerequisites"></a>先决条件

* Excel 2016 for Windows（内部版本 16.0.6727.1000 或更高版本）、Excel Online 或 Excel for Mac（内部版本 15.26 或更高版本）。
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>此外接程序中使用的设计模板

- 着陆页
- 品牌栏
- 选项卡栏
- 设置

有关设计模式的详细信息，请参阅 [Office 外接程序的用户体验设计模式模板](https://dev.office.com/docs/add-ins/design/ux-design-patterns)。有关实现示例，请参阅 [Office 外接程序的用户体验设计模式代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)。

## <a name="get-the-pickadate-library"></a>获取 PickADate 库

Office Fabric 日期选取器控件与 PickADate 库之间有依存关系。下载此示例*后*，按以下步骤操作。

1. 从 [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3) 下载第 3.5.3 版库。 
2. 将包解压缩，然后转到 `\pickadate.js-3.5.3\lib` 文件夹。 
3. 将此文件夹中的所有文件和文件夹都复制到项目文件夹 `SalesTrackerWeb\Scripts\PickADate` 中，`compressed` 和 `themes-source` 文件夹*除外*。

## <a name="run-the-project"></a>运行项目

1. 打开 Visual Studio 解决方案文件。 
2. 按 **F5**。 
3. 当 Excel 打开后，单击“**主页**”功能区右端的“**跟踪销售**”。此时，系统会在任务窗格中打开外接程序。

#### <a name="import-data"></a>导入数据

4. 在“**主页**”的“**产品名称**”框中，输入以下产品名称之一（区分大小写）：“**键盘**”、“**鼠标**”、“**显示器**”、“**笔记本电脑**”
5. 使用日期选取器控件选择 2016 年 9 月 16 日之前的一个日期，因为示例数据中没有这一天之后的销售数据。
6. 选择“**获取销售数据**”按钮。几秒后，工作簿便会将焦点切换到新的“**销售**”工作表。 

#### <a name="change-table-settings"></a>更改表设置

1. 选择选项卡栏上的“**表**”。 
2. 根据需要取消选中各个单选按钮，从而隐藏相应的列。
3. 选择表的颜色。

#### <a name="change-chart-settings"></a>更改图表设置

1. 选择选项卡栏上的“**图表**”。 
2. 选择图表的数据源。
3. 选择图表类型。
4. 选择图表颜色主题。

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>比较此 Excel Web 外接程序代码与 VSTO 外接程序示例

使用 Office 和 Word JavaScript API 的代码位于 Home.js 和 Helpers.js 中。所有的样式设置均通过 HTML5 和样式表文件（settings.css、tab.bar.css 和多个 Office Fabric css 文件）完成。

比较此代码与[表和图表](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3)中的代码。请注意以下事项：


- Excel Web 外接程序受多个平台支持，包括 Windows、Mac 和 Office Online。VSTO 外接程序只受 Windows 支持。
- 在 VSTO 和此 Web 外接程序中更改表样式的操作相似。在这两个外接程序中，代码均会向属性分配样式名称，如 TableStyleMedium3。在 VSTO 外接程序中，此为表对象。（请参阅 TableAndChartPane.cs 文件中的各种 `*_CheckedChanged` 方法。）在此 Web 外接程序中，这一属性位于传递给 `setTableOptionsAsync` 函数的 JavaScript 对象中。（请参阅 helpers.js 文件中的 `setTableColor` 函数。）
- 在 VSTO 和此 Web 外接程序中切换表中列的可见性的操作非常相似。比较 VSTO 示例中 `ListObjectHeaders_Click` 方法的逻辑与此 Web 外接程序中的 `toggleColumnVisibility` 函数。
- 若要在 VSTO 外接程序中更改图表的样式，请向图表对象的 `ChartStyle` 属性分配一个整数。此整数表示的是样式设置集合。请参阅 VSTO 外接程序中的 TableAndChartPane.cs 文件。在 Excel Web 外接程序中，将图表替换为采用所需样式的新图表。可以在 JavaScript 对象中记录图表的当前样式设置，如此示例中的 home.js 文件所示。若要更改单个样式设置，代码会在之后传递给 helpers.js 中 `changeChart` 函数的设置对象中进行更改。
- 在 VSTO 外接程序和此 Web 外接程序中更改图表*类型*（如折线图、分区图和簇状柱形图）的操作非常相似。在这两个外接程序中，`switch-case` 结构均用于向类型属性分配值。比较 VSTO 示例中的 `ChartStyleComboBox_SelectedIndexChanged` 方法与此 Web 示例中的 `setChartType` 方法。 
- 在类似于这样的 Excel Web 外接程序中，若要跟踪*单个*值在一段时间内的变化情况（或任意水平轴），必须使用仅包含两个可见列的表生成图表：一列用于生成水平轴（在此示例中为“日期”），另一列提供图表中显示的值。因此，此外接程序使用销售数据表的副本创建*隐藏*工作表。此表仅包含两个可见列：“**日期**”列和包含图表的选定数据源的列。虽然图表显示在表旁边的“**销售**”工作表中，但却从隐藏（临时）工作表的表中获取数据。
- 若要在 Excel Web 外接程序中更改图表的数据源，代码会切换隐藏表中列的可见性。请参阅 helpers.js 文件中的 `setChartDataSource` 方法。在 VSTO 外接程序中，代码会调用图表对象的 `SetSourceData` 函数，指定将表中的哪一列用作数据源。请参阅 `chartDataSourceComboBox_SelectedIndexChanged` 方法。
- 在 Office Web 外接程序，可以利用 HTML5、JavaScript 和 CSS 来丰富 UI，如此示例代码中所示。 
- 由于 Office Web 外接程序执行的是异步方法调用，因此 UI 一律不会受到影响。
- Office Web 外接程序进行 AJAX 调用，从联机服务提供商检索数据。此示例仅从本地 JSON 文件提取 JSON 数据。请参阅 Helpers.js 文件中的 `getSalesData` 方法。VSTO 外接程序通过在 C# 中使用 WebClient 来访问联机资源。请参阅 TableAndChartPane.cs 中的 `GetDataUpdatesFoOneDataSource` 方法。   


## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你对此示例的反馈。你可以在此存储库中的“*问题*”部分向我们发送反馈。

与 Microsoft Office 365 开发相关的一般问题应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。如果你的问题是关于 Office JavaScript API，请务必为问题添加 [office-js] 和 [API].标记。

## <a name="additional-resources"></a>其他资源

* [Office 外接程序文档](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Office 开发人员中心](http://dev.office.com/)
* 有关更多 Office 外接程序示例，请访问 [Github 上的 OfficeDev](https://github.com/officedev)。

## <a name="copyright"></a>版权
版权所有 © 2016 Microsoft Corporation。保留所有权利。

