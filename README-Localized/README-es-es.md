# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>Complemento web de Excel para manipular el formato de tablas y gráficos

Obtenga información sobre cómo dar formato con programación a tablas y gráficos y cómo importar datos en una hoja de cálculo en los complementos web Excel. Compárelo con la forma en que estas tareas se realizan en el complemento VSTO [Tablas y gráficos](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). En este complemento web de Excel también se muestra cómo usar los ejemplos de diseño de [Código de modelos de diseño de la experiencia del usuario para complementos de Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 

## <a name="table-of-contents"></a>Tabla de contenido
* [Historial de cambios](#change-history)
* [Requisitos previos](#prerequisites)
* [Plantillas de diseño usadas en este complemento](#design-templates-used-in-this-add-in)
* [Obtener la biblioteca de PickADate](get-the-pickadate-library)
* [Ejecutar el proyecto](#run-the-project)
* [Comparar este código de complemento web de Excel con el ejemplo del complemento VSTO](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## <a name="change-history"></a>Historial de cambios

2 de noviembre de 2016:

* Versión inicial.

## <a name="prerequisites"></a>Requisitos previos

* Excel 2016 para Windows (compilación 16.0.6727.1000 o posteriores), Excel Online o Excel para Mac (compilación 15.26 o posteriores).
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>Plantillas de diseño usadas en este complemento

- Página de aterrizaje
- Barra de marca
- Barra de pestañas
- Configuración

Para obtener más información sobre los modelos de diseño, vea [Plantillas de modelos de diseño de la experiencia del usuario para complementos de Office](https://dev.office.com/docs/add-ins/design/ux-design-patterns). También puede ver las implementaciones de ejemplos en [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="get-the-pickadate-library"></a>Obtener la biblioteca de PickADate

El control de selector de fecha de Office Fabric tiene una dependencia en la biblioteca PickADate. Siga estos pasos *después* de descargar el ejemplo.

1. Descargue la versión 3.5.3 de la biblioteca desde [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3). 
2. Descomprima el paquete y abra la carpeta `\pickadate.js-3.5.3\lib`. 
3. Copie todos los archivos y carpetas en esa carpeta (*excepto* las carpetas `compressed` y `themes-source`) en la carpeta de proyecto: `SalesTrackerWeb\Scripts\PickADate`.

## <a name="run-the-project"></a>Ejecutar el proyecto

1. Abra el archivo de solución de Visual Studio. 
2. Presione **F5**. 
3. Cuando se abra Excel, haga clic en el botón **Realizar seguimiento de ventas** en el extremo derecho de la pestaña **Inicio** de la cinta de opciones. Se abrirá el complemento en un panel de tareas.

#### <a name="import-data"></a>Importación de datos

4. En la **página principal**, escriba uno de los siguientes nombres de producto (se distinguen mayúsculas de minúsculas) en el cuadro **Nombre de producto**: **Teclado**, **Mouse**, **Monitor**, **Portátil**,
5. Use el control de selector de fecha para elegir una fecha que no sea posterior al 16 de septiembre de 2016, ya que en los datos de ejemplo no hay ventas después de esta fecha.
6. Seleccione el botón **Obtener datos de ventas**. Después de unos segundos, el libro cambiará el foco a una nueva hoja de cálculo llamada **Ventas**. 

#### <a name="change-table-settings"></a>Cambiar la configuración de la tabla

1. Seleccione **Tabla** en la barra de pestañas. 
2. Anule la selección de los botones de radio según sea necesario para ocultar las columnas correspondientes.
3. Seleccione un color por la tabla.

#### <a name="change-chart-settings"></a>Cambiar la configuración de los gráficos

1. Seleccione **Gráfico** en la barra de pestañas. 
2. Seleccione un origen de datos para el gráfico.
3. Seleccione un tipo de gráfico.
4. Seleccione un tema de color para el gráfico.

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>Comparar este código de complemento web de Excel con el ejemplo del complemento VSTO

El código que usan las API de JavaScript de Office y Word se encuentra en Home.js y Helpers.js. Toda la configuración de estilos se realiza con HTML5 y los archivos de hojas de estilo: settings.css, tab.bar.css y varios archivos CSS de Office Fabric.

Compare este código con el código en [Tablas y gráficos](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Tenga en cuenta lo siguiente:


- Los complementos web de Excel son compatibles con varias plataformas, como Windows, Mac y Office Online. Los complementos VSTO solo son compatibles con Windows.
- El procedimiento para cambiar el estilo de una tabla es similar en VSTO y en los complementos web. En los dos casos, el código asigna un nombre de estilo (como TableStyleMedium3) a una propiedad. En el complemento VSTO, esto es un objeto de tabla. (Vea los diferentes métodos de `*_CheckedChanged` en el archivo TableAndChartPane.cs). En este complemento web, la propiedad se encuentra en un objeto de JavaScript que se pasó a la función `setTableOptionsAsync`. (Vea la función `setTableColor` en el archivo helpers.js).
- El procedimiento para cambiar la visibilidad de las columnas en una tabla es muy similar en los complementos VSTO y en los complementos web. Compare la lógica del método `ListObjectHeaders_Click` en el ejemplo de VSTO con la función `toggleColumnVisibility` en el complemento web.
- Para cambiar el estilo de un gráfico en un complemento VSTO, se asigna un entero a la propiedad `ChartStyle` del objeto del gráfico. El entero hace referencia a una colección de opciones de configuración de estilo. Vea el archivo TableAndChartPane.cs en el complemento VSTO. En un complemento web de Excel, el gráfico se reemplaza por uno nuevo que tenga el estilo que prefiera. Puede guardar la configuración del estilo actual para el gráfico en un objeto de JavaScript, como se realiza en este ejemplo en el archivo home.js. Para cambiar una sola configuración de estilo, el código realiza el cambio en el objeto de configuración que, después, se pasa a la función `changeChart` en helpers.js.
- Cambiar un *tipo* de gráfico (como Line, Area o ClusteredColumn) es muy similar en el complemento VSTO y en el complemento web. En los dos casos, se usa una estructura de `switch-case` para asignar un valor a una propiedad de tipo. Compare el método `ChartStyleComboBox_SelectedIndexChanged` del ejemplo de VSTO con el método `setChartType` en este ejemplo web. 
- En un complemento web de Excel (como este), para realizar un seguimiento de un valor *único* de un período de tiempo (o cualquier eje horizontal) es necesario generar el gráfico a partir de una tabla con solo dos columnas visibles: una columna que proporcione el eje horizontal (en este caso, fechas) y una segunda columna que proporcione el valor que se muestra en el gráfico. Por este motivo, el complemento crea una hoja de cálculo *oculta* con una copia de la tabla de datos de ventas. Esta tabla solo tiene dos columnas visibles: la columna **Fecha** y la columna con el origen de datos elegido por el gráfico. Aunque el gráfico aparece en la hoja de cálculo **Ventas** junto a la tabla, obtiene los datos de la tabla en la hoja de cálculo oculta (temporal).
- Para cambiar el origen de datos de un gráfico en el complemento web de Excel, el código cambia la visibilidad de las columnas en la tabla oculta. Vea el método `setChartDataSource` en el archivo helpers.js. En un complemento VSTO, el código especifica la columna de la tabla que se usará como el origen de datos con una llamada a la función `SetSourceData` del objeto de gráfico. Vea el método `chartDataSourceComboBox_SelectedIndexChanged`.
- En los complementos web de Office se puede usar HTML5, JavaScript y CSS para crear interfaces de usuario avanzadas, como la interfaz de usuario de este ejemplo de código. 
- Como los complementos web de Office usan llamadas de método asincrónico, la interfaz de usuario nunca se bloquea.
- Los complementos web de Office realizan llamadas AJAX para recuperar datos de proveedores de servicios en línea. En este ejemplo solo se capturan datos JSON de un archivo JSON local. Vea el método `getSalesData` en Helpers.js. Los complementos VSTO usan un cliente web en C# para obtener acceso a recursos en línea. Vea el método `GetDataUpdatesFoOneDataSource` en TableAndChartPane.cs.   


## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre el desarrollo de Microsoft Office 365 deben publicarse en [Desbordamiento de pila](http://stackoverflow.com/questions/tagged/office-js+API). Si su pregunta trata sobre las API de JavaScript para Office, asegúrese de que su pregunta se etiqueta con [office-js] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* [Documentación de complementos de Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Centro de desarrollo de Office](http://dev.office.com/)
* Más ejemplos de complementos de Office en [OfficeDev en GitHub](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.

