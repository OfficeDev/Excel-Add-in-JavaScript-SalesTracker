# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>Suplemento Web do Excel para manipular formatação de tabelas e gráficos

Saiba como formatar programaticamente tabelas e gráficos e como importar dados para uma planilha em suplementos Web do Excel. Compare com a maneira como estas tarefas são realizadas no Suplemento VSTO [Tabelas e Gráficos](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Este suplemento Web do Excel também mostra como usar os exemplos de design do [Código de Padrões de Design da Experiência do Usuário de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 

## <a name="table-of-contents"></a>Sumário
* [Histórico de Alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Modelos de design usados neste suplemento](#design-templates-used-in-this-add-in)
* [Obter a biblioteca PickADate](get-the-pickadate-library)
* [Executar o projeto](#run-the-project)
* [Compare este código de suplemento Web do Excel ao exemplo de suplemento VSTO](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="change-history"></a>Histórico de Alterações

2 de novembro de 2016:

* Versão inicial.

## <a name="prerequisites"></a>Pré-requisitos

* Excel 2016 para Windows (compilação 16.0.6727.1000 ou posterior), Excel Online ou Excel para Mac (compilação 15.26 ou posterior).
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>Modelos de design usados neste suplemento

- Página de aterrissagem
- Barra da marca
- Barra de guias
- Configurações

Para obter mais informações sobre os padrões de design, confira [Modelos de padrão de design da experiência do usuário para suplementos do Office](https://dev.office.com/docs/add-ins/design/ux-design-patterns). Para implementações de exemplo, confira [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="get-the-pickadate-library"></a>Obter a biblioteca PickADate

O controle de seletor de data do Office Fabric tem uma dependência na biblioteca PickADate. Siga estas etapas *após* baixar este exemplo.

1. Baixe a versão 3.5.3 da biblioteca de [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3). 
2. Descompacte o pacote e navegue até a pasta `\pickadate.js-3.5.3\lib`. 
3. Copie todos os arquivos e pastas para essa pasta, *exceto* as pastas `compressed` e `themes-source`, para a pasta do projeto: `SalesTrackerWeb\Scripts\PickADate`.

## <a name="run-the-project"></a>Executar o projeto

1. Abra o arquivo de solução do Visual Studio. 
2. Pressione **F5**. 
3. Quando o Excel for aberto, clique no botão **Acompanhar vendas** na extremidade direita da faixa de opções **Início**. O suplemento é aberto em um painel de tarefas.

#### <a name="import-data"></a>Importar dados

4. Na página **Início**, digite um dos seguintes nomes de produtos (que diferenciam maiúsculas de minúsculas) na caixa **Nome do produto**: **Teclado**, **Mouse**, **Monitor**, **Laptop**,
5. Use o controle de seletor de data para selecionar uma data que não seja posterior a 16 de setembro de 2016, pois não há vendas após essa data nos dados de exemplo.
6. Selecione o botão **Obter dados de vendas**. Após alguns segundos, a pasta de trabalho mudará o foco para uma nova planilha **Vendas**. 

#### <a name="change-table-settings"></a>Alterar as configurações de tabela

1. Selecione **Tabela** na barra de guias. 
2. Desmarque os botões de opção conforme necessário, para ocultar as colunas correspondentes.
3. Selecione uma cor para a tabela.

#### <a name="change-chart-settings"></a>Alterar as configurações de gráfico

1. Selecione **Gráfico** na barra de guias. 
2. Selecione uma fonte de dados para o gráfico.
3. Selecione um tipo de gráfico.
4. Selecione um tema de cores de gráfico.

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>Compare este código de suplemento Web do Excel ao exemplo de suplemento VSTO

O código que usa as APIs JavaScript do Office e do Word está em Home.js e Helpers.js. Todos os estilos são feitos com HTML5 e os arquivos de folha de estilos: settings.css, tab.bar.css e vários arquivos css do Office Fabric.

Compare este código ao código em [Tabelas e Gráficos](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Observe o seguinte:


- Os suplementos Web do Excel têm suporte em várias plataformas, incluindo Windows, Mac e Office Online. Só há suporte para suplementos VSTO no Windows.
- A alteração do estilo de uma tabela é semelhante nos suplementos VSTO e Web. Em ambos os casos, o código atribui um nome de estilo, como TableStyleMedium3, a uma propriedade. No suplemento VSTO, esse é um objeto de tabela. (Confira os vários métodos `*_CheckedChanged` no arquivo TableAndChartPane.cs.) Neste suplemento Web, a propriedade está em um objeto JavaScript que é passado para a função `setTableOptionsAsync`. (Confira a função `setTableColor` no arquivo helpers.js.)
- Alternar a visibilidade das colunas em uma tabela é muito semelhante em suplementos VSTO e Web. Compare a lógica do método `ListObjectHeaders_Click` no exemplo VSTO à função `toggleColumnVisibility` no suplemento Web.
- Para alterar o estilo de um gráfico em um suplemento VSTO, você atribui um número inteiro à propriedade `ChartStyle` do objeto de gráfico. O número inteiro se refere a um conjunto de configurações de estilo. Confira o arquivo TableAndChartPane.cs no suplemento VSTO. Em um suplemento Web do Excel, você substitui o gráfico por um novo com o estilo desejado. Você pode gravar as configurações de estilo atuais do gráfico em um objeto JavaScript, como este exemplo faz no arquivo home.js. Para alterar uma configuração de estilo única, o código é alterado no objeto de configurações que é passado para a função `changeChart` em helpers.js.
- A alteração de um *tipo* de gráfico, como Linha, Área ou ClusteredColumn, é muito semelhante no suplemento VSTO e neste suplemento Web. Em ambos os casos, uma estrutura `switch-case` é usada para atribuir um valor a uma propriedade de tipo. Compare o método `ChartStyleComboBox_SelectedIndexChanged` no exemplo VSTO ao método `setChartType` neste exemplo da Web. 
- Em um suplemento Web do Excel, como este, quando você deseja acompanhar um *único* valor ao longo do tempo (ou em qualquer eixo horizontal), o gráfico deve ser criado com base em uma tabela com apenas duas colunas visíveis: uma que fornece o eixo horizontal (neste caso, datas) e uma segunda, que fornece o valor que está sendo exibido no gráfico. Por esse motivo, o suplemento cria uma planilha *ocultos* com uma cópia da tabela de dados de vendas. Esta tabela tem apenas duas colunas visíveis: a coluna **Data** e a coluna com a fonte de dados escolhida para o gráfico. Embora o gráfico seja mostrado na planilha **Vendas** ao lado da tabela, ele obtém dados da tabela na planilha oculta (temp).
- Para alterar a fonte de dados de um gráfico no suplemento Web do Excel, o código alterna a visibilidade das colunas na tabela oculta. Confira o método `setChartDataSource` no arquivo helpers.js. Em um suplemento VSTO, seu código especifica qual coluna da tabela deve ser usada como fonte de dados chamando a função `SetSourceData` do objeto de gráfico. Confira o método `chartDataSourceComboBox_SelectedIndexChanged`.
- Em suplementos Web do Office, você pode tirar proveito de HTML5, JavaScript e CSS para criar interfaces do usuário avançadas, como a interface do usuário neste exemplo de código. 
- Como os suplementos Web do Office fazem chamadas de método assíncronas, a interface do usuário nunca é bloqueada.
- Os suplementos Web do Office fazem chamadas AJAX para recuperar dados de provedores de serviços online. Este exemplo simplesmente busca dados JSON de um arquivo JSON local. Confira o método `getSalesData` em Helpers.js. Os suplementos VSTO usam um WebClient em C# para acessar recursos online. Confira o método `GetDataUpdatesFoOneDataSource` em TableAndChartPane.cs.   


## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode nos enviar comentários na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Microsoft Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Se sua pergunta estiver relacionada às APIs JavaScript para Office, não deixe de marcá-la com as tags [office-js] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação dos suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* Confira outros exemplos de Suplemento do Office em [OfficeDev no Github](https://github.com/officedev)

## <a name="copyright"></a>Direitos autorais
Copyright (C) 2016 Microsoft Corporation. Todos os direitos reservados.

