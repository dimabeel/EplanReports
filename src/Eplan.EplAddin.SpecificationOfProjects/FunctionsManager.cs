using System;
using System.Collections.Generic;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System.Linq;
using DataBaseLibrary;

namespace Eplan.EplAddin.SpecificationOfProjects
{
    public class FunctionsManager
    {
        // Функция получения проекта
        public Project GetProject()
        {
            try
            {
                SelectionSet selectionSetProject = new SelectionSet();
                Project project = selectionSetProject.GetCurrentProject(true);
                if (project == null) // Проверяем, получили или нет проект
                {
                    BaseException projectIsNull = new BaseException("Проект пуст (функция GetProject)", MessageLevel.Error);
                    throw projectIsNull;
                }
                return project;
            }
            catch
            {
                BaseException projectErr = new BaseException("Ошибка в функции GetProject", MessageLevel.Error);
                throw projectErr;
            }
        }

        // Функция получения ссылок в проекте
        public ArticleReference[] GetArticleReferences(Project project)
        {
            try
            {
                DMObjectsFinder objectFinder = new DMObjectsFinder(project);
                ArticleReference[] articleReferences = objectFinder.GetArticleReferences(null);
                if (articleReferences.Length < 1) // Если ссылок нет
                {
                    BaseException articleReferencesIsNull = new BaseException("Ссылки ArticleReferences пусты! (функция GetArticleReferences)", MessageLevel.Error);
                    throw articleReferencesIsNull;
                }
                return articleReferences;
            }
            catch
            {
                BaseException articleReferencesErr = new BaseException("Ошибка в функции GetArticleReferences", MessageLevel.Error);
                throw articleReferencesErr;
            }
        }

        // Функция фильтрации ссылок по необходимым параметрам
        public List<ArticleReference> FilterArticleReferences(ArticleReference[] articleReferences)
        {
            try
            {
                List<ArticleReference> filteredArticleReferences = new List<ArticleReference>();
                foreach (ArticleReference reference in articleReferences)
                {
                    // Example: +CAB5-F8 || +CAB1-1-F8
                    string[] splitedReferenceIdentifyingName = reference.IdentifyingName.Split('-');
                    if (splitedReferenceIdentifyingName.Length > 1)
                    {
                        string articleDeviceType = splitedReferenceIdentifyingName[1];
                        string articleType = splitedReferenceIdentifyingName[0];
                        // Проверяем корректность, буква 'W' - кабели, они не нужны
                        // Ловим символ между CAB1 и F8 (пример), что бы понять, что пришло
                        int fictiv;
                        bool isDigit = int.TryParse(articleDeviceType.ToString(), out fictiv);
                        if ((isDigit == false) && (fictiv < 100))
                        {
                            // Цифры нет, все штатно
                            if ((articleDeviceType[0] != 'W') && (articleType.Length > 1))
                            {
                                filteredArticleReferences.Add(reference);
                            }
                        }
                        else
                        {
                            // Если цифра есть, то сдвигаем проверку на 1 индекс вправо
                            // Так как теперь компонент лежит там
                            articleDeviceType = splitedReferenceIdentifyingName[2];
                            if ((articleDeviceType[0] != 'W') && (articleType.Length > 1))
                            {
                                filteredArticleReferences.Add(reference);
                            }
                        }
                    }
                }
                if (filteredArticleReferences.Count < 1) // Если список пуст
                {
                    BaseException filteredArticleReferencesListIsNull = new BaseException("Отфильтрованный список пуст! (функция FilterArticleReferences)", MessageLevel.Error);
                    throw filteredArticleReferencesListIsNull;
                }
                return filteredArticleReferences;
            }
            catch
            {
                BaseException filterArticleReferencesErr = new BaseException("Ошибка в функции FilterArticleReferences", MessageLevel.Error);
                throw filterArticleReferencesErr;
            }
        }

        // Функция получения имени изделия (одного)
        public string GetProjectProductName(ArticleReference articleReference) // Получение имени изделия из IdefName для 1 ссылки
        {
            string productName = string.Empty;
            try
            {
                // +CAB1-1-FQT1 - нештатно; +CAB1-FQT1 - штатно
                // [0] +CAB1; [1] 1; [2] FQT1;
                string[] mainStringSplit = articleReference.IdentifyingName.Split('-');
                string[] noReadyName = mainStringSplit[0].Split('+');
                int fictiv;
                bool isDigit = int.TryParse(mainStringSplit[1].ToString(), out fictiv);
                if ((isDigit == false) && (fictiv < 100))
                {
                    productName = noReadyName[1];
                }
                else
                {
                    productName = noReadyName[1] + "-" + mainStringSplit[1].ToString();
                }
            }
            catch
            {
                BaseException getProjProdNameExcept = new BaseException("Ошибка в функции GetProjectProductName", MessageLevel.Error);
                throw getProjProdNameExcept;
            }
            if (productName == string.Empty)
            {
                BaseException productNameException = new BaseException("productName пусто! (функция GetProjectProductName)", MessageLevel.Error);
                throw productNameException;
            }
            return productName;
        }

        // Функция получения массива имен изделий, которые содержатся в проекте (шкаф,коагулятор и др)
        // Получу все изделия в т.ч с повторением
        public string[] GetProjectProductNames(List<ArticleReference> articleReferences)
        {
            string[] filteredProjectProductNames = new string[0];
            try
            {
                string[] projectProductNames = new string[articleReferences.Count];
                for (int i = 0; i < articleReferences.Count; i++)
                {
                    projectProductNames[i] = GetProjectProductName(articleReferences[i]);
                }
                filteredProjectProductNames = projectProductNames.Distinct().ToArray();
            }
            catch
            {
                BaseException getProjectProductNamesException = new BaseException("Ошибка в функции GetProjectProductNames", MessageLevel.Error);
                throw getProjectProductNamesException;
            }
            return filteredProjectProductNames;
        }

        // Функция получения списка списков всех компонентов по каждому изделию
        public List<List<string>> GetProductArticles(string[] projectProductNames, List<ArticleReference> articleReferences)
        {
            try
            {
                // Первый список - соответствует string[] projectProductNames
                // Второй - содержит элементы которые есть в projectProductNames
                List<List<string>> productArticles = new List<List<string>>();
                for (int i = 0; i < projectProductNames.Length; i++)
                {
                    List<string> secondList = new List<string>();
                    for (int j = 0; j < articleReferences.Count; j++)
                    {
                        if (projectProductNames[i] == GetProjectProductName(articleReferences[j]))
                        {
                            try
                            {
                                secondList.Add(articleReferences[j].Article.PartNr);
                            }
                            catch
                            {
                                BaseException baseException = new BaseException("Ошибка во внутреннем списке (функция GetProductArticles)", MessageLevel.Error);
                                throw baseException;
                            }
                        }
                    }
                    secondList.Sort();
                    productArticles.Add(secondList);
                }
                return productArticles;
            }
            catch
            {
                BaseException productArticlesErr = new BaseException("Ошибка в функции GetProductArticles", MessageLevel.Error);
                throw productArticlesErr;
            }
        }

        // Функция получения количества каждого из компонентов в изделии
        public List<List<ComponentInfo>> GetProductArticleCount(string[] projectProductNames, List<List<string>> productArticles)
        {
            List<List<ComponentInfo>> productArticleInfo = new List<List<ComponentInfo>>();
            for (int i = 0; i < projectProductNames.Length; i++)
            {
                string previouseElem = string.Empty;
                string currentElem = string.Empty;
                try
                {
                    List<ComponentInfo> innerList = new List<ComponentInfo>();
                    for (int j = 0; j < productArticles[i].Count; j++)
                    {
                        currentElem = productArticles[i][j];
                        if (currentElem != previouseElem)
                        {
                            ComponentInfo componentInfo = new ComponentInfo
                            {
                                Count = 0,
                                PartNumber = productArticles[i][j]
                            };
                            for (int k = 0; k < productArticles[i].Count; k++)
                            {
                                if (productArticles[i][k] == componentInfo.PartNumber)
                                {
                                    componentInfo.Count++;
                                }
                            }
                            previouseElem = currentElem;
                            innerList.Add(componentInfo);
                        }
                    }
                    productArticleInfo.Add(innerList);
                }
                catch
                {
                    BaseException ProdArticleCountErr = new BaseException("Ошибка в функции GetProductArticleCount", MessageLevel.Error);
                    throw ProdArticleCountErr;
                }
            }
            return productArticleInfo;
        }

        // Функция получения cписка компонентов в проекте
        public Article[] GetProjectProductArticleList(Project project, List<ArticleReference> articleReferences)
        {
            try
            {
                // Инициализация
                string[] articleStrNames = new string[0];
                string[] countOfReferences = new string[articleReferences.Count];
                // Записываю все PartNr компонентов в массив
                for (int i = 0; i < articleReferences.Count; i++)
                {
                    countOfReferences[i] = articleReferences[i].PartNr;
                }
                // Убираю повторения и сортирую
                articleStrNames = countOfReferences.Distinct().ToArray();
                Array.Sort(articleStrNames);
                // Получаю данные по каждому из компонентов в массиве
                Article[] ProductArticleList = new Article[articleStrNames.Length];
                Article[] AllProductArticlesInproject = project.Articles;
                for (int i = 0; i < articleStrNames.Length; i++)
                {
                    for (int j = 0; j < AllProductArticlesInproject.Length; j++)
                    {
                        if (articleStrNames[i] == AllProductArticlesInproject[j].PartNr)
                        {
                            ProductArticleList[i] = AllProductArticlesInproject[j];
                        }
                    }
                }
                return ProductArticleList;
            }
            catch
            {
                BaseException ProjArticleListErr = new BaseException("Ошибка в функции GetProjectArticleList", MessageLevel.Error);
                throw ProjArticleListErr;
            }

        }

        // Функция получения свойств компонентов для справочника
        public ComponentCatalogInfo[] GetArticleProperties(Article[] articles)
        {
            try
            {
                ComponentCatalogInfo[] componentCatalogInfos = new ComponentCatalogInfo[articles.Length];
                for (int i = 0; i < componentCatalogInfos.Length; i++)
                {
                    MultiLangString langString; // Для получения строк по языку
                    componentCatalogInfos[i] = new ComponentCatalogInfo();
                    componentCatalogInfos[i].PartNumber = articles[i].Properties.ARTICLE_PARTNR.ToString();

                    if (articles[i].Properties.ARTICLE_TYPENR.IsEmpty == false)
                    {
                        componentCatalogInfos[i].TypeNumber = articles[i].Properties.ARTICLE_TYPENR;
                    }
                    else
                    {
                        componentCatalogInfos[i].TypeNumber = null;
                    }

                    if (articles[i].Properties.ARTICLE_ORDERNR.IsEmpty == false)
                    {
                        componentCatalogInfos[i].OrderNumber = articles[i].Properties.ARTICLE_ORDERNR;

                    }
                    else
                    {
                        componentCatalogInfos[i].OrderNumber = null;
                    }

                    if (articles[i].Properties.ARTICLE_MANUFACTURER.IsEmpty == false)
                    {
                        componentCatalogInfos[i].ManufacturerSmallName = articles[i].Properties.ARTICLE_MANUFACTURER;
                    }
                    else
                    {
                        componentCatalogInfos[i].ManufacturerSmallName = null;
                    }

                    if (articles[i].Properties.ARTICLE_MANUFACTURER_NAME.IsEmpty == false)
                    {
                        componentCatalogInfos[i].ManufacturerFullName = articles[i].Properties.ARTICLE_MANUFACTURER_NAME;
                    }
                    else
                    {
                        componentCatalogInfos[i].ManufacturerFullName = null;
                    }

                    if (articles[i].Properties.ARTICLE_SUPPLIER.IsEmpty == false)
                    {
                        componentCatalogInfos[i].SupplierSmallName = articles[i].Properties.ARTICLE_SUPPLIER;
                    }
                    else
                    {
                        componentCatalogInfos[i].SupplierSmallName = null;
                    }

                    if (articles[i].Properties.ARTICLE_SUPPLIER_NAME.IsEmpty == false)
                    {
                        componentCatalogInfos[i].SupplierFullName = articles[i].Properties.ARTICLE_SUPPLIER_NAME;
                    }
                    else
                    {
                        componentCatalogInfos[i].SupplierFullName = null;
                    }

                    if (articles[i].Properties.ARTICLE_DESCR1.IsEmpty == false)
                    {
                        langString = articles[i].Properties.ARTICLE_DESCR1;
                        componentCatalogInfos[i].Description1 = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[i].Description1 = null;
                    }

                    if (articles[i].Properties.ARTICLE_DESCR2.IsEmpty == false)
                    {
                        langString = articles[i].Properties.ARTICLE_DESCR2;
                        componentCatalogInfos[i].Description2 = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[i].Description2 = null;
                    }

                    if (articles[i].Properties.ARTICLE_CHARACTERISTICS.IsEmpty == false)
                    {
                        componentCatalogInfos[i].TechnicalCharacteristics = articles[i].Properties.ARTICLE_CHARACTERISTICS;
                    }
                    else
                    {
                        componentCatalogInfos[i].TechnicalCharacteristics = null;
                    }

                    if (articles[i].Properties.ARTICLE_NOTE.IsEmpty == false)
                    {
                        langString = articles[i].Properties.ARTICLE_NOTE;
                        componentCatalogInfos[i].Note = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[i].Note = null;
                    }

                    if (articles[i].Properties.ARTICLE_CROSSSECTIONFROM.IsEmpty == false)
                    {
                        componentCatalogInfos[i].TerminalCrossSectionFrom = articles[i].Properties.ARTICLE_CROSSSECTIONFROM;
                    }
                    else
                    {
                        componentCatalogInfos[i].TerminalCrossSectionFrom = null;
                    }

                    if (articles[i].Properties.ARTICLE_CROSSSECTIONTILL.IsEmpty == false)
                    {
                        componentCatalogInfos[i].TerminalCrossSectionTo = articles[i].Properties.ARTICLE_CROSSSECTIONTILL;
                    }
                    else
                    {
                        componentCatalogInfos[i].TerminalCrossSectionTo = null;
                    }

                    if (articles[i].Properties.ARTICLE_ELECTRICALCURRENT.IsEmpty == false)
                    {
                        componentCatalogInfos[i].ElectricalCurrent = articles[i].Properties.ARTICLE_ELECTRICALCURRENT;
                    }
                    else
                    {
                        componentCatalogInfos[i].ElectricalCurrent = null;
                    }

                    if (articles[i].Properties.ARTICLE_ELECTRICALPOWER.IsEmpty == false)
                    {
                        componentCatalogInfos[i].ElectricalSwitchingCapacity = articles[i].Properties.ARTICLE_ELECTRICALPOWER;
                    }
                    else
                    {
                        componentCatalogInfos[i].ElectricalSwitchingCapacity = null;
                    }

                    if (articles[i].Properties.ARTICLE_VOLTAGE.IsEmpty == false)
                    {
                        componentCatalogInfos[i].Voltage = articles[i].Properties.ARTICLE_VOLTAGE;
                    }
                    else
                    {
                        componentCatalogInfos[i].Voltage = null;
                    }

                    if (articles[i].Properties.ARTICLE_VOLTAGETYPE.IsEmpty == false)
                    {
                        componentCatalogInfos[i].VoltageType = articles[i].Properties.ARTICLE_VOLTAGETYPE;
                    }
                    else
                    {
                        componentCatalogInfos[i].VoltageType = null;
                    }

                    if (articles[i].Properties.ARTICLE_HEIGHT.IsEmpty == false)
                    {
                        componentCatalogInfos[i].Height = articles[i].Properties.ARTICLE_HEIGHT;
                    }
                    else
                    {
                        componentCatalogInfos[i].Height = 0;
                    }

                    if (articles[i].Properties.ARTICLE_WIDTH.IsEmpty == false)
                    {
                        componentCatalogInfos[i].Width = articles[i].Properties.ARTICLE_WIDTH.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].Width = 0;
                    }

                    if (articles[i].Properties.ARTICLE_DEPTH.IsEmpty == false)
                    {
                        componentCatalogInfos[i].Depth = articles[i].Properties.ARTICLE_DEPTH.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].Depth = 0;
                    }

                    if (articles[i].Properties.ARTICLE_WEIGHT.IsEmpty == false)
                    {
                        componentCatalogInfos[i].Weight = articles[i].Properties.ARTICLE_WEIGHT.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].Weight = 0;
                    }

                    if (articles[i].Properties.ARTICLE_MOUNTINGSITE.IsEmpty == false)
                    {
                        componentCatalogInfos[i].MountingSiteID = articles[i].Properties.ARTICLE_MOUNTINGSITE;
                    }
                    else
                    {
                        componentCatalogInfos[i].MountingSiteID = 0;
                    }

                    if (articles[i].Properties.ARTICLE_MOUNTINGSPACE.IsEmpty == false)
                    {
                        componentCatalogInfos[i].MountingSpace = articles[i].Properties.ARTICLE_MOUNTINGSPACE.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].MountingSpace = 0;
                    }

                    componentCatalogInfos[i].PartGroup = GetPartGroup(articles[i]); // 0-00-000

                    if (articles[i].Properties.ARTICLE_QUANTITYUNIT.IsEmpty == false)
                    {
                        langString = articles[i].Properties.ARTICLE_QUANTITYUNIT;
                        string filterString = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU).Trim();
                        componentCatalogInfos[i].QuantityUnit = filterString;
                    }
                    else
                    {
                        componentCatalogInfos[i].QuantityUnit = null;
                    }

                    if (articles[i].Properties.ARTICLE_PACKAGINGQUANTITY.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PackagingQuantity = articles[i].Properties.ARTICLE_PACKAGINGQUANTITY.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].PackagingQuantity = 0;
                    }

                    if (articles[i].Properties.PART_LASTCHANGE.IsEmpty == false)
                    {
                        componentCatalogInfos[i].LastChange = articles[i].Properties.PART_LASTCHANGE;
                    }
                    else
                    {
                        componentCatalogInfos[i].LastChange = null;
                    }

                    if (articles[i].Properties.ARTICLE_SALESPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[i].SalesPrice1 = articles[i].Properties.ARTICLE_SALESPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].SalesPrice1 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_SALESPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[i].SalesPrice2 = articles[i].Properties.ARTICLE_SALESPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].SalesPrice2 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_PURCHASEPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PurchasePrice1 = articles[i].Properties.ARTICLE_PURCHASEPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].PurchasePrice1 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_PURCHASEPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PurchasePrice2 = articles[i].Properties.ARTICLE_PURCHASEPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].PurchasePrice2 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_PACKAGINGPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PackagingPrice1 = articles[i].Properties.ARTICLE_PACKAGINGPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].PackagingPrice1 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_PACKAGINGPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PackagingPrice2 = articles[i].Properties.ARTICLE_PACKAGINGPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[i].PackagingPrice2 = 0;
                    }

                    if (articles[i].Properties.ARTICLE_PRICEUNIT.IsEmpty == false)
                    {
                        componentCatalogInfos[i].PriceUnit = articles[i].Properties.ARTICLE_PRICEUNIT;
                    }
                    else
                    {
                        componentCatalogInfos[i].PriceUnit = 0;
                    }
                }
                return componentCatalogInfos;
            }
            catch
            {
                BaseException ArtPropErr = new BaseException("Ошибка в функции GetArticleProperties", MessageLevel.Error);
                throw ArtPropErr;
            }
        }

        // Функция получения шаблона типа "0-00-000"
        public string GetPartGroup(Article article)
        {
            try
            {
                // Распарсили группы и получили "дерево"
                string partGroup = String.Format("{0:0}-{1:00}-{2:000}",
                article.Properties.ARTICLE_PRODUCTTOPGROUP.ToInt(),
                article.Properties.ARTICLE_PRODUCTGROUP.ToInt(),
                article.Properties.ARTICLE_PRODUCTSUBGROUP.ToInt());
                return partGroup;
            }
            catch
            {
                BaseException GetPartGroupErr = new BaseException("Ошибка в функции GetPartGroup", MessageLevel.Error);
                throw GetPartGroupErr;
            }
        }

        // Функция записи данных в справочник
        public void FillComponentCatalog(ComponentCatalogInfo[] componentCatalogInfos)
        {
            using (DBContext DBCon = new DBContext())
            {
                // Если БД нет - создать и инициализировать
                if (!DBCon.Database.Exists())
                {
                    DBCon.Database.Initialize(false);
                }
                foreach (ComponentCatalogInfo componentCatalogInfo in componentCatalogInfos)
                {
                    // Если в справочнике количество таких PartNumber = 0 (нет таких), то пишем в справочник
                    if (DBCon.ComponentCatalogs.Where(o => o.PartNumber == componentCatalogInfo.PartNumber).Count() < 1)
                    {
                        PriceList priceList = new PriceList(); // Для прайса
                        ComponentCatalog componentCatalog = new ComponentCatalog(); // Для свойств
                        QuantityUnit quantityUnit = new QuantityUnit(); // Для ед. измерения
                        MountingSite mountingSite; // Для монтажной поверхности

                        // Обработка QuantityUnit для получения ID
                        try
                        {
                            if ((componentCatalogInfo.QuantityUnit != null) && (componentCatalogInfo.QuantityUnit != ""))
                            {
                                quantityUnit = DBCon.QuantityUnits.Where(
                                    o => o.Name == componentCatalogInfo.QuantityUnit)
                                    .FirstOrDefault();
                                if (quantityUnit != null)
                                {
                                    componentCatalog.QuantityUnitID = quantityUnit.QuantityUnitID;
                                }
                                else
                                {
                                    quantityUnit = new QuantityUnit();
                                    quantityUnit.Name = componentCatalogInfo.QuantityUnit;
                                    DBCon.QuantityUnits.Add(quantityUnit);
                                    DBCon.SaveChanges();
                                    quantityUnit = DBCon.QuantityUnits.Where(
                                        o => o.Name == componentCatalogInfo.QuantityUnit).FirstOrDefault();
                                    componentCatalog.QuantityUnitID = quantityUnit.QuantityUnitID;
                                    quantityUnit = null;
                                }
                            }
                            else
                            {
                                componentCatalog.QuantityUnitID = null;
                            }

                        }
                        catch
                        {
                            BaseException QuantityUnitCheckErr = new BaseException("Ошибка в преобразовании и поиске QuantityUnitID", MessageLevel.Error);
                            throw QuantityUnitCheckErr;
                        }

                        // Поиск монтажной поверхности по Internal value
                        try
                        {
                            mountingSite = DBCon.MountingSites.Where(
                                o => o.InternalValue == componentCatalogInfo.MountingSiteID)
                                .First();
                            componentCatalog.MountingSiteID = mountingSite.MountingSiteID;
                            mountingSite = null;
                        }
                        catch
                        {
                            BaseException mountingSiteErr = new BaseException("Ошибка в поиске и преобразовании MountingSiteID", MessageLevel.Error);
                            throw mountingSiteErr;
                        }

                        // Заполнение свойств справочника для вставки в контекст
                        try
                        {
                            componentCatalog.Depth = componentCatalogInfo.Depth;
                            componentCatalog.Description1 = componentCatalogInfo.Description1;
                            componentCatalog.Description2 = componentCatalogInfo.Description2;
                            componentCatalog.ElectricalCurrent = componentCatalogInfo.ElectricalCurrent;
                            componentCatalog.ElectricalSwitchingCapacity = componentCatalogInfo.ElectricalSwitchingCapacity;
                            componentCatalog.Height = componentCatalogInfo.Height;
                            componentCatalog.LastChange = componentCatalogInfo.LastChange;
                            componentCatalog.ManufacturerFullName = componentCatalogInfo.ManufacturerFullName;
                            componentCatalog.ManufacturerSmallName = componentCatalogInfo.ManufacturerSmallName;
                            componentCatalog.MountingSpace = componentCatalogInfo.MountingSpace;
                            componentCatalog.Note = componentCatalogInfo.Note;
                            componentCatalog.OrderNumber = componentCatalogInfo.OrderNumber;
                            componentCatalog.PackagingQuantity = componentCatalogInfo.PackagingQuantity;
                            componentCatalog.PartGroup = componentCatalogInfo.PartGroup;
                            componentCatalog.PartNumber = componentCatalogInfo.PartNumber;
                            componentCatalog.SupplierFullName = componentCatalogInfo.SupplierFullName;
                            componentCatalog.SupplierSmallName = componentCatalogInfo.SupplierSmallName;
                            componentCatalog.TechnicalCharacteristics = componentCatalogInfo.TechnicalCharacteristics;
                            componentCatalog.TerminalCrossSectionFrom = componentCatalogInfo.TerminalCrossSectionFrom;
                            componentCatalog.TerminalCrossSectionTo = componentCatalogInfo.TerminalCrossSectionTo;
                            componentCatalog.TypeNumber = componentCatalogInfo.TypeNumber;
                            componentCatalog.Voltage = componentCatalogInfo.Voltage;
                            componentCatalog.VoltageType = componentCatalogInfo.VoltageType;
                            componentCatalog.Weight = componentCatalogInfo.Weight;
                            componentCatalog.Width = componentCatalogInfo.Width;
                        }
                        catch
                        {
                            BaseException comCatErr = new BaseException("Ошибка в присваивании данных ComponentCatalog", MessageLevel.Error);
                            throw comCatErr;
                        }

                        // Проверяем как записывается в справочник компонентов
                        try
                        {
                            DBCon.ComponentCatalogs.Add(componentCatalog);
                            DBCon.SaveChanges();
                        }
                        catch
                        {
                            BaseException comCatAddErr = new BaseException("Ошибка в добавлении и сохранении ComponentCatalog в БД", MessageLevel.Error);
                            throw comCatAddErr;
                        }

                        // Записываем прайс-лист компонента
                        try
                        {
                            priceList.PartNumber = componentCatalogInfo.PartNumber;
                            priceList.LastChange = GetDateTimeFromString(componentCatalogInfo.LastChange);
                            priceList.PackagingPrice1 = componentCatalogInfo.PackagingPrice1;
                            priceList.PackagingPrice2 = componentCatalogInfo.PackagingPrice2;
                            priceList.PriceUnit = componentCatalogInfo.PriceUnit;
                            priceList.PurchasePrice1 = componentCatalogInfo.PurchasePrice1;
                            priceList.PurchasePrice2 = componentCatalogInfo.PurchasePrice2;
                            priceList.SalesPrice1 = componentCatalogInfo.SalesPrice1;
                            priceList.SalesPrice2 = componentCatalogInfo.SalesPrice2;
                            DBCon.PriceLists.Add(priceList);
                            DBCon.SaveChanges();
                        }
                        catch
                        {
                            BaseException PriceListErr = new BaseException("Ошибка в записи прайс-листа после добавления нового компонента в справочник", MessageLevel.Error);
                            throw PriceListErr;
                        }
                    }
                    else
                    {
                        // Раз объект в справочнике есть, проверим его прайсы
                        try
                        {
                            // Получаю последний прайс объекта
                            List<PriceList> prisesForCheck = DBCon.PriceLists.Where(
                                o => o.PartNumber == componentCatalogInfo.PartNumber).OrderBy(
                                o1 => o1.LastChange).ToList();
                            PriceList lastPriceList = prisesForCheck.Last();
                            bool priseIsChanged = false; // Флаг на изменение прайса
                            // Проверяю, отличается хотя бы один прайс от того, что есть
                            if (lastPriceList.PackagingPrice1 != componentCatalogInfo.PackagingPrice1) priseIsChanged = true;
                            if (lastPriceList.PackagingPrice2 != componentCatalogInfo.PackagingPrice2) priseIsChanged = true;
                            if (lastPriceList.PriceUnit != componentCatalogInfo.PriceUnit) priseIsChanged = true;
                            if (lastPriceList.PurchasePrice1 != componentCatalogInfo.PurchasePrice1) priseIsChanged = true;
                            if (lastPriceList.PurchasePrice2 != componentCatalogInfo.PurchasePrice2) priseIsChanged = true;
                            if (lastPriceList.SalesPrice1 != componentCatalogInfo.SalesPrice1) priseIsChanged = true;
                            if (lastPriceList.SalesPrice2 != componentCatalogInfo.SalesPrice2) priseIsChanged = true;

                            // Если все таки у нас есть изменения, добавляем их в справочник прайсов
                            if (priseIsChanged == true)
                            {
                                PriceList updPriseList = new PriceList();
                                updPriseList.PartNumber = componentCatalogInfo.PartNumber;
                                updPriseList.LastChange = DateTime.Now;
                                updPriseList.PackagingPrice1 = componentCatalogInfo.PackagingPrice1;
                                updPriseList.PackagingPrice2 = componentCatalogInfo.PackagingPrice2;
                                updPriseList.PriceUnit = componentCatalogInfo.PriceUnit;
                                updPriseList.PurchasePrice1 = componentCatalogInfo.PurchasePrice1;
                                updPriseList.PurchasePrice2 = componentCatalogInfo.PurchasePrice2;
                                updPriseList.SalesPrice1 = componentCatalogInfo.SalesPrice1;
                                updPriseList.SalesPrice2 = componentCatalogInfo.SalesPrice2;
                                DBCon.PriceLists.Add(updPriseList);
                                DBCon.SaveChanges();
                            }
                        }
                        catch
                        {
                            BaseException checkPriseListErr = new BaseException("Ошибка при проверке прайс-листов", MessageLevel.Error);
                            throw checkPriseListErr;
                        }
                    }
                }
            }
        }

        // Функция получения времени из ARTICLE_LASTCHANGE
        public DateTime GetDateTimeFromString(string strForParse)
        {
            try
            {
                // "TSAM / 2010.10.03 10:09:44", поэтому нужна вторая часть строки
                string[] splittedStr = strForParse.Split('/');
                string strForConvert = splittedStr[1].Trim();
                DateTime gettedDateTime = Convert.ToDateTime(strForConvert);
                return gettedDateTime;
            }
            catch
            {
                BaseException dateTimeFromStrErr = new BaseException("Ошибка в функции GetDateTimeFromString", MessageLevel.Error);
                throw dateTimeFromStrErr;
            }
        }

        // Функция записи данных спецификации в БД
        public void FillSpecification(string[] projectProductNames, List<List<ComponentInfo>> specificationInfo, string projName, List<LocationInfo> locationInfos)
        {
            using (DBContext DBCon = new DBContext())
            {
                Proj project = new Proj(); // Модель объекта
                project.Name = projName;
                // Проверяю, есть ли у нас такой проект уже
                if (DBCon.Projs.Where(o => o.Name == project.Name).Count() == 0)
                {
                    try
                    {
                        // Если нету, добавляю проект
                        DBCon.Projs.Add(project);
                        DBCon.SaveChanges();
                    }
                    catch
                    {
                        BaseException projAddErr = new BaseException("Ошибка при добавлении проекта в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw projAddErr;
                    }

                    try
                    {
                        // Теперь добавляю изделия из проекта
                        int projID = DBCon.Projs.Where(
                            o => o.Name == project.Name)
                            .Select(o1 => o1.ProjID).FirstOrDefault();
                        foreach (string articleName in projectProductNames)
                        {
                            // Для поиска описания изделия
                            LocationDescription locationInfo = new LocationDescription();
                            PArticle projArticle = new PArticle(); // Моделька объекта
                            // Ищу, есть ли такое имя в таблице имен
                            locationInfo = DBCon.LocationDescriptions.Where(
                                o => o.Name == articleName).FirstOrDefault();
                            if (locationInfo == null)
                            {
                                locationInfo = new LocationDescription();
                                locationInfo.Name = articleName;
                                try
                                {   // Если не найдено описание такого элемента, то ошибка
                                    locationInfo.Description = locationInfos.Find(o => o.Name == articleName).Description;
                                }
                                catch
                                {
                                    // Description = Name, ставим описание равным имени
                                    locationInfo.Description = locationInfo.Name;
                                }
                                DBCon.LocationDescriptions.Add(locationInfo);
                                DBCon.SaveChanges();
                                locationInfo = DBCon.LocationDescriptions.Where(
                                o => o.Name == articleName).FirstOrDefault();
                                projArticle.LocationDesriptionID = locationInfo.LocationDescriptionID;
                                locationInfo = null;
                            }
                            else
                            {
                                // Если обозначение есть, проверяю равняется ли его имя и описание
                                string locationName = locationInfo.Name;
                                string locationDescription = locationInfo.Description;
                                if (locationName.Equals(locationDescription) == true)
                                {
                                    try
                                    {
                                        // Получаю описание текущего изделия (структурное обозначение)
                                        string currentLocationDescription = locationInfos.Find(o => o.Name == articleName).Description;
                                        // Если сработало без ошибок, обновляю его
                                        locationInfo.Description = currentLocationDescription;
                                        DBCon.SaveChanges();
                                        // И присваиваю ID projArticle
                                        projArticle.LocationDesriptionID = locationInfo.LocationDescriptionID;
                                    }
                                    catch
                                    {
                                        // Если ошибка, то значит изделие все равно без описания структурного
                                        // Значит, оставляем сопоставление Name = Description
                                        projArticle.LocationDesriptionID = locationInfo.LocationDescriptionID;
                                    }
                                }
                                else
                                {
                                    // Если же описание и имя обозначений разные (введено корректное описание)
                                    projArticle.LocationDesriptionID = locationInfo.LocationDescriptionID;
                                }
                            }
                            projArticle.ProjectID = projID;
                            DBCon.PArticles.Add(projArticle);
                        }
                        DBCon.SaveChanges();
                    }
                    catch
                    {
                        BaseException articleAddErr = new BaseException("Ошибка при добавлении изделий в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw articleAddErr;
                    }

                    try
                    {
                        // Добавим компоненты к изделиям, то есть, завершим спецификацию
                        for (int i = 0; i < projectProductNames.Length; i++)
                        {
                            // Айди проекта ищу
                            int projID = DBCon.Projs.Where(
                            o => o.Name == project.Name)
                            .Select(o1 => o1.ProjID).FirstOrDefault();
                            string searchArticle = projectProductNames[i];
                            // Ищу айди изделия, в которое буду добавлять
                            // Но надо проверить, есть ли это изделие в структурных обозначениях               
                            if (DBCon.LocationDescriptions.Where(
                                o => o.Name == searchArticle).FirstOrDefault() != null)
                            {
                                int LocationID = DBCon.LocationDescriptions.Where(
                                    o => o.Name == searchArticle)
                                    .Select(o1 => o1.LocationDescriptionID).FirstOrDefault();
                                int articleID = DBCon.PArticles.Where(
                                    o => o.ProjectID == projID).Where(
                                    o1 => o1.LocationDesriptionID == LocationID).Select(
                                    o2 => o2.PArticleID).FirstOrDefault();
                                for (int j = 0; j < specificationInfo[i].Count; j++)
                                {
                                    // Так как индекс projectProductNames соответствует
                                    // индексу списка списков ComponentInfo
                                    Component component = new Component();
                                    component.PArticleID = articleID;
                                    // Перебираю внутренний список
                                    component.PartNumber = specificationInfo[i][j].PartNumber;
                                    component.Count = specificationInfo[i][j].Count;
                                    DBCon.Components.Add(component);
                                    DBCon.SaveChanges();
                                }
                            }
                            else
                            {
                                // Так же, как и в изделиях. Нету в списке структурных обозначений - пропускаем
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        BaseException componentAddErr = new BaseException("Ошибка при добавлении компонентов в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw componentAddErr;
                    }
                }
                else
                {
                    // Если проект есть, удалю его спецификацию
                    // И запишу новую (обновленную).
                    Proj proj = DBCon.Projs.Where(
                        o => o.Name == project.Name).FirstOrDefault();
                    DBCon.Projs.Remove(proj);
                    DBCon.SaveChanges();
                    FillSpecification(projectProductNames, specificationInfo, projName, locationInfos);
                }
            }
        }

        // Функция получения списка структурных идентификаторов изделий
        public List<LocationInfo> GetLocationDescriptions(Project project, string[] projectProductNames)
        {
            try
            {
                Location[] locations = project.GetLocationObjects();
                List<LocationInfo> locationInfos = new List<LocationInfo>();
                MultiLangString descriptionString;
                foreach (Location location in locations)
                {
                    // Если описание не пусто
                    if (location.Properties.LOCATION_DESCRIPTION != "")
                    {
                        LocationInfo locationInfo = new LocationInfo();
                        for (int i = 0; i < projectProductNames.Length; i++)
                        {
                            // Если нашли совпадение в списке изделий, записываем в список
                            if (projectProductNames[i].Equals(location.Properties.LOCATION_FULLNAME))
                            {
                                descriptionString = location.Properties.LOCATION_DESCRIPTION;
                                locationInfo.Name = location.Properties.LOCATION_FULLNAME;
                                locationInfo.Description = descriptionString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                                locationInfos.Add(locationInfo);
                            }
                        }
                    }
                }
                return locationInfos;
            }
            catch
            {
                BaseException getLocDescrErr = new BaseException("Ошибка в функции GetLocationDescriptions", MessageLevel.Error);
                throw getLocDescrErr;
            }
        }
    }
}