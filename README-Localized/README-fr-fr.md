# <a name="excel-web-addin-for-manipulating-table-and-chart-formatting"></a>Complément web Excel pour manipuler la mise en forme de tableau et de graphique

Découvrez comment mettre en forme des tableaux et des graphiques par programme, et comment importer des données dans une feuille de calcul, dans des compléments web Excel. Comparez avec la façon dont ces tâches sont réalisées dans le complément VSTO relatif aux [tableaux et aux graphiques](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Ce complément web Excel montre également comment utiliser les exemples de conception dans le [code des modèles de conception de l’expérience utilisateur de complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code). 

## <a name="table-of-contents"></a>Sommaire
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Modèles de conception utilisés dans ce complément](#design-templates-used-in-this-add-in)
* [Recherche de la bibliothèque PickADate](get-the-pickadate-library)
* [Exécution du projet](#run-the-project)
* [Comparaison du code de complément web Excel avec l’exemple de complément VSTO](#compare-this-web-add-in-code-with-the-VSTO-add-in-sample)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="change-history"></a>Historique des modifications

2 novembre 2016 :

* Version d’origine.

## <a name="prerequisites"></a>Conditions préalables

* Excel 2016 pour Windows (version 16.0.6727.1000 ou ultérieure), Excel Online ou Excel pour Mac (version 15.26 ou ultérieure).
* Visual Studio 2015 

## <a name="design-templates-used-in-this-addin"></a>Modèles de conception utilisés dans le complément

- Page d’accueil
- Barre de marque
- Barre d’onglets
- Paramètres

Pour plus d’informations sur les modèles de conception, voir l’article sur les [modèles de conception de l’expérience utilisateur pour les compléments Office](https://dev.office.com/docs/add-ins/design/ux-design-patterns). Pour obtenir des exemples d’implémentation, voir [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="get-the-pickadate-library"></a>Recherche de la bibliothèque PickADate

Le contrôle de sélecteur de dates de la structure Office a une dépendance sur la bibliothèque PickADate. Réalisez les étapes suivantes *après* que vous avez téléchargé cet exemple.

1. Téléchargez la version 3.5.3 de la bibliothèque à partir de [pickadate.js](https://github.com/amsul/pickadate.js/releases/tag/3.5.3). 
2. Décompressez le package et accédez au dossier `\pickadate.js-3.5.3\lib`. 
3. Copiez tous les fichiers et dossiers de ce dossier, *sauf* les dossiers `compressed` et `themes-source`, dans le dossier du projet : `SalesTrackerWeb\Scripts\PickADate`.

## <a name="run-the-project"></a>Exécuter le projet

1. Ouvrez le fichier de solution Visual Studio. 
2. Appuyez sur la touche **F5**. 
3. Quand Excel s’ouvre, cliquez sur le bouton **Suivre les ventes** situé sur l’extrémité droite du ruban **Accueil**. Le complément s’ouvre dans un volet Office.

#### <a name="import-data"></a>Importer des données

4. Sur la page **Accueil**, entrez l’un des noms de produit suivants (respectez la casse) dans la case **Nom de produit** : **clavier**, **souris**, **moniteur**, **portable**.
5. Utilisez le contrôle de sélecteur de dates pour sélectionner une date qui n’est pas ultérieure au 16 septembre 2016, car aucune vente n’existe après cette date dans l’exemple de données.
6. Sélectionnez le bouton **Obtenir les données de vente**. Au bout de quelques secondes, le classeur bascule vers une nouvelle feuille de calcul **Ventes**. 

#### <a name="change-table-settings"></a>Modifier les paramètres de tableau

1. Sélectionnez **Tableau** dans la barre d’onglets. 
2. Désactivez les cases d’option selon vos besoins afin de masquer les colonnes correspondantes.
3. Sélectionnez une couleur pour le tableau.

#### <a name="change-chart-settings"></a>Modifier les paramètres de graphique

1. Sélectionnez **Graphique** dans la barre d’onglets. 
2. Sélectionnez une source de données pour le graphique.
3. Sélectionnez un type de graphique.
4. Sélectionnez un thème de couleur de graphique.

## <a name="compare-this-excel-web-addin-code-with-the-vsto-addin-sample"></a>Comparaison du code de complément web Excel avec l’exemple de complément VSTO

Les codes utilisés par les API Office et Word JavaScript sont Home.js et Helpers.js. Tous les styles sont effectués avec HTML5 et les fichiers de feuille de style : settings.css, tab.bar.css et plusieurs fichiers CSS de structure Office.

Comparez ce code avec le code dans [Tableaux et graphiques](https://code.msdn.microsoft.com/VSTO-Generate-tables-and-f19859b3). Remarques :


- Les compléments web Excel sont pris en charge sur différentes plateformes, y compris Windows, Mac et Office Online. Les compléments VSTO sont uniquement pris en charge dans Windows.
- La modification du style d’un tableau est similaire dans les compléments VSTO et web. Dans les deux cas, votre code affecte un nom de style, par exemple TableStyleMedium3, à une propriété. Dans le complément VSTO, il s’agit d’un objet de tableau. (Voir les différentes méthodes `*_CheckedChanged` dans le fichier TableAndChartPane.cs.) Dans ce complément web, la propriété est un objet JavaScript qui est transmis à la fonction `setTableOptionsAsync`. (Voir la fonction `setTableColor` dans le fichier helpers.js.)
- L’activation et la désactivation de la visibilité des colonnes dans un tableau est très similaire dans les compléments web et VSTO. Comparez la logique de la méthode `ListObjectHeaders_Click` dans l’exemple VSTO à la fonction `toggleColumnVisibility` dans le complément web.
- Pour changer le style d’un graphique dans un complément VSTO, vous affectez un entier à la propriété `ChartStyle` de l’objet de graphique. L’entier fait référence à une collection de paramètres de style. Consultez le fichier TableAndChartPane.cs dans le complément VSTO. Dans un complément web Excel vous remplacez le graphique par un nouveau ayant le style souhaité. Vous pouvez enregistrer les paramètres de style actuels du graphique dans un objet JavaScript, à l’instar de cet exemple dans le fichier home.js. Pour modifier un paramètre de style unique, votre code le modifie dans l’objet de paramètres qui est ensuite transmis à la fonction `changeChart` dans helpers.js.
- La modification d’un *type* de graphique, tel que Line, Area et ClusteredColumn, est très similaire dans le complément VSTO et le complément web. Dans les deux cas, une structure `switch-case` est utilisée pour affecter une valeur à une propriété de type. Comparez la méthode `ChartStyleComboBox_SelectedIndexChanged` dans l’exemple VSTO à la méthode `setChartType` dans cet exemple web. 
- Dans un complément web Excel, comme celui-ci, lorsque vous voulez suivre une valeur *unique* dans le temps (ou sur n’importe quel axe horizontal), le graphique doit être créé à partir d’un tableau avec seulement deux colonnes visibles : une qui fournit l’axe horizontal (dans ce cas, les dates) et une seconde qui fournit la valeur affichée dans le graphique. Pour cette raison, le complément crée une feuille de calcul *masquée* avec une copie de la table de données des ventes. Ce tableau comporte seulement deux colonnes visibles : la colonne **Date** et la colonne avec la source de données sélectionnée pour le graphique. Bien que le graphique apparaisse sur la feuille de calcul **Sales** en regard du tableau, ses données proviennent du tableau figurant sur la feuille de calcul (temporaire) masquée.
- Pour changer la source de données d’un graphique dans le complément web Excel, votre code active ou désactive la visibilité des colonnes sur le tableau masqué. Consultez la méthode `setChartDataSource` dans le fichier helpers.js. Dans un complément VSTO, votre code spécifie la colonne dans le tableau à utiliser comme source de données en appelant la fonction `SetSourceData` de l’objet de graphique. Voir la méthode `chartDataSourceComboBox_SelectedIndexChanged`.
- Dans les compléments web Office, vous pouvez utiliser HTML5, JavaScript et CSS pour créer des interfaces utilisateur enrichies, comme l’interface utilisateur dans cet exemple de code. 
- Étant donné que les compléments web Office émettent des appels de méthode asynchrone, l’interface utilisateur n’est jamais bloquée.
- Les compléments web Office émettent des appels AJAX pour extraire des données des fournisseurs de services en ligne. Cet exemple récupère simplement les données JSON d’un fichier JSON local. Voir la méthode `getSalesData` dans Helpers.js. Les compléments VSTO utilisent un client web dans C# pour accéder aux ressources en ligne. Voir la méthode `GetDataUpdatesFoOneDataSource` dans TableAndChartPane.cs.   


## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle est marquée avec les balises [office js] et [API].

## <a name="additional-resources"></a>Ressources supplémentaires

* [Documentation de complément Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.

