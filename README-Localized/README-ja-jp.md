# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>テーブルとグラフの書式設定を操作するための Excel Web アドイン

Excel Web アドインでプログラムを使用してテーブルとグラフの書式を設定する方法と、スプレッドシートにデータをインポートする方法について説明します。これらのタスクが[テーブルとグラフ](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3)の VSTO アドインで実行される方法と比較します。この Excel Web アドインは、[Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) からの設計サンプルを使用する方法も表示します。 

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [このアドインで使用されるデザイン テンプレート](#design-templates-used-in-this-add-in)
* [PickADate ライブラリの取得](get-the-pickadate-library)
* [プロジェクトを実行する](#run-the-project)
* [この Excel Web アドインのコードを VSTO アドイン サンプルと比較する](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

2016 年 11 月 2 日:

* 初期バージョン。

## <a name="prerequisites"></a>前提条件

* Excel 2016 for Windows (ビルド 16.0.6727.1000 以降)、Excel Online、または Excel for Mac (ビルド 15.26 以降)。
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>このアドインで使用されるデザイン テンプレート

- ランディング ページ
- ブランド バー
- タブ バー
- 設定

デザイン パターンの詳細については、「[Office アドインの UX デザイン パターンのテンプレート](https://dev.office.com/docs/add-ins/design/ux-design-patterns)」を参照してください。サンプルの実装については、「[Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)」を参照してください。

## <a name="get-the-pickadate-library"></a>PickADate ライブラリの取得

Office Fabric の日付の選択コントロールは PickADate ライブラリに依存関係があります。このサンプルをダウンロードした*後*、次の手順を実行します。

1. バージョン 3.5.3 のライブラリを [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3) からダウンロードします。 
2. パッケージを解凍し、`\pickadate.js-3.5.3\lib` フォルダーに移動します。 
3. そのフォルダーにある、`compressed` フォルダーと `themes-source` フォルダー*以外*のすべてのファイルとフォルダーをプロジェクト フォルダー `SalesTrackerWeb\Scripts\PickADate` にコピーします。

## <a name="run-the-project"></a>プロジェクトを実行する

1. Visual Studio ソリューション ファイルを開きます。 
2. **F5** キーを押します。 
3. Excel が開いたら、**[ホーム]** リボンの右端にある **[売上の追跡]** ボタンをクリックします。アドインが作業ウィンドウで開きます。

#### <a name="import-data"></a>データをインポートする

4. **[ホーム]** ページで、次の製品名 (大文字と小文字を区別) のいずれかを **[製品名]** ボックスに入力します。**Keyboard**、**Mouse**、**Monitor**、**Laptop**、
5. 2016 年 9 月 16 日以降、サンプル データでの売上はないため、それより前の日付を日付の選択コントロールを使用して選択します。
6. **[売上データの取得]** ボタンを選択します。数秒後、ブックのスイッチのフォーカスが新しい **[売上]** ワークシートに移動します。 

#### <a name="change-table-settings"></a>テーブルの設定を変更する

1. タブ バーで **[テーブル]** を選択します。 
2. 必要に応じて、ラジオ ボタンの選択を解除し、対応する列を非表示にします。
3. テーブルの色を選択します。

#### <a name="change-chart-settings"></a>グラフの設定を変更する

1. タブ バーで **[グラフ]** を選択します。 
2. グラフのデータ ソースを選択します。
3. グラフの種類を選択します。
4. グラフの色のテーマを選択します。

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>この Excel Web アドインのコードを VSTO アドイン サンプルと比較する

Office および Word JavaScript API を使用するコードは、Home.js と Helpers.js に存在します。すべてのスタイル指定は、HTML5 とスタイルシート ファイル (settings.css ファイル、tab.bar.css ファイル、およびいくつかの Office Fabric css ファイル) によって行われます。

このコードを[テーブルとグラフ](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3)内のコードと比較します。以下の点に注意してください。


- Excel Web アドインは、Windows、Mac、および Office Online を含むいくつかのプラットフォームでサポートされています。VSTO アドインは、Windows でのみサポートされます。
- テーブルのスタイルの変更は、VSTO アドインと Web アドインとで似ています。どちらの場合も、コードはプロパティに TableStyleMedium3 などのスタイル名を割り当てます。VSTO アドインでは、これはテーブル オブジェクトになります。(TableAndChartPane.cs ファイル内のさまざまな `*_CheckedChanged` メソッドを参照してください。)この Web アドインでは、プロパティは `setTableOptionsAsync` 関数に渡される JavaScript オブジェクトにあります。(helpers.js ファイル内の `setTableColor` 関数を参照してください。)
- テーブル内の列の可視性の切り替えは、VSTO アドインと Web アドインとで非常に似ています。VSTO サンプルでの `ListObjectHeaders_Click` メソッドのロジックを Web アドインの `toggleColumnVisibility` 関数と比較します。
- VSTO アドインでグラフのスタイルを変更するには、グラフ オブジェクトの `ChartStyle` プロパティに整数を割り当てます。整数は、スタイルの設定のコレクションを表します。VSTO アドインの TableAndChartPane.cs ファイルを参照してください。Excel Web アドインで、グラフを目的のスタイルの新しいグラフと置き換えます。このサンプルが home.js ファイルで行うのと同様、グラフの現在のスタイル設定を JavaScript オブジェクトに記録できます。1 つのスタイル設定を変更する場合は、コードによって設定オブジェクトで変更され、次に helpers.js の `changeChart` 関数に渡されます。
- グラフの*種類* (折れ線、面、集合縦棒など) の変更は、VSTO アドインとこの Web アドインとで非常に似ています。どちらの場合も、種類のプロパティに値を割り当てるのに、`switch-case` 構造体が使用されます。VSTO サンプルの `ChartStyleComboBox_SelectedIndexChanged` メソッドをこの Web サンプルの `setChartType` メソッドと比較します。 
- Excel Web アドインでは、*1 つ*の値を一定期間 (または任意の横軸) 追跡する場合、グラフは表示されている 2 つの列、つまり、横軸 (この場合は日付) を指定する列と、グラフに表示されている値を指定する列のみで構成されるテーブルを基に作成される必要があります。このため、アドインは売上データ テーブルを含む*非表示*のワークシートを作成します。このテーブルに含まれる表示可能な列は、**日付**列と、グラフに対して選択されたデータ ソースを含む列の 2 つだけです。グラフは**売上**ワークシートのテーブルの横に表示されますが、そのデータは非表示 (一時) ワークシートのテーブルから取得されます。
- Excel Web アドインでグラフのデータ ソースを変更する場合、非表示テーブルの列の可視性がコードによって切り替えられます。helpers.js ファイル内の `setChartDataSource` メソッドを参照してください。VSTO アドインでは、グラフ オブジェクトの `SetSourceData` 関数を呼び出すことによって、テーブル内のどの列をデータ ソースとして使用するかをコードが指定します。`chartDataSourceComboBox_SelectedIndexChanged` メソッドを参照してください。
- Office Web アドインでは、HTML5、JavaScript および CSS を活用して、このコード サンプルの UI のような優れた UI を作成できます。 
- Office Web アドインでは非同期メソッドの呼び出しを行うため、UI がブロックすることはありません。
- Office Web アドインでは AJAX 呼び出しを行って、オンライン サービス プロバイダーからデータを取得します。このサンプルでは、単にローカル JSON ファイルから JSON データをフェッチします。Helpers.js の `getSalesData` メソッドを参照してください。VSTO アドインでは、C# の WebClient を使用してオンライン リソースにアクセスします。TableAndChartPane.cs の `GetDataUpdatesFoOneDataSource` メソッドを参照してください。   


## <a name="questions-and-comments"></a>質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。Office JavaScript API に関する質問の場合は、必ず質問に [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>追加リソース

* [Office アドインのドキュメント](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Office デベロッパー センター](http://dev.office.com/)
* [Github の OfficeDev](https://github.com/officedev) にあるその他の Office アドイン サンプル

## <a name="copyright"></a>著作権
Copyright (c) 2016 Microsoft Corporation.All rights reserved.

