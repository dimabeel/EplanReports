using System;
using System.Collections.Generic;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System.Linq;
using DataBaseLibrary;
using Eplan.EplAddin.SpecificationListOfObjects;

namespace Eplan.EplAddin.SpecificationOfProjects
{
    public class Functions
    {
        // Путь к надстройке
        string[] originalAssemblypath = AddInModule.OriginalAssemblyPath.Split('\\');
        
        // Название файла конфигурации
        string configFileName = @"DataBaseConnectionConfig.ini";
        
        // Строка подключения
        string connectionString = @"";

        // Получение полного пути к файлу настроек
        public string GetConfigFilePath()
        {
            var sourceEnd = originalAssemblypath.Length - 1;
            var path = @"";
            for (int source = 0; source < sourceEnd; source++)
            {
                path += originalAssemblypath[source].ToString() + "\\";
            }
            path += configFileName;
            return path;
        }
        
        // Функция проверки ini файла надстройки
        public void ChekAddInIniFile()
        {
            var configFilePath = GetConfigFilePath();
            // Обращаемся к INI файлу
            var iniFile = new ConfigIniFile(configFilePath);

            if (iniFile.KeyExists("Data_Source", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Data_Source");
                connectionString += "Data Source=" + iniProperty.ToString() + ";";
            }
            else
            {
                var iniErr = new BaseException("Ошибка в конфигурационном файле надстройки, нет Data Souce",MessageLevel.Error);
                throw iniErr;
            }

            if (iniFile.KeyExists("Initial_Catalog", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Initial_Catalog");
                connectionString += "Initial Catalog=" + iniProperty.ToString() + ";";
            }
            else
            {
                var iniErr = new BaseException("Ошибка в конфигурационном файле надстройки, нет InitialCatalog", MessageLevel.Error);
                throw iniErr;
            }

            if (iniFile.KeyExists("Persist_Security_Info", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Persist_Security_Info");
                connectionString += "Persist Security Info=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Integrated_Security", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Integrated_Security");
                connectionString += "Integrated Security=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("User_ID", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "User_ID");
                connectionString += "User ID=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Password", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Password");
                connectionString += "Password=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Trusted_Connection", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Trusted_Connection");
                connectionString += "Trusted_Connection=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Encrypt", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Encrypt");
                connectionString += "Encrypt=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Trust_Server_Certificate", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Trust_Server_Certificate");
                connectionString += "TrustServerCertificate=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Application_Name", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Application_Name");
                connectionString += "Application Name=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Application_Intent", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Application_Intent");
                connectionString += "ApplicationIntent=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Asynchronous_Processing", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Asynchronous_Processing");
                connectionString += "Asynchronous Processing=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Column_Encryption_Setting", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Column_Encryption_Setting");
                connectionString += "Column Encryption Setting=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Connect_Timeout", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Connect_Timeout");
                connectionString += "Connect Timeout=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Connection_Lifetime", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Connection_Lifetime");
                connectionString += "Connection Lifetime=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Connect_Retry_Count", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Connect_Retry_Count");
                connectionString += "ConnectRetryCount=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Connect_Retry_Interval", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Connect_Retry_Interval");
                connectionString += "ConnectRetryInterval=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Context_Connection", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Context_Connection");
                connectionString += "Context Connection=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Current_Language", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Current_Language");
                connectionString += "Current Language=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Enlist", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Enlist");
                connectionString += "Enlist=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Failover_Partner", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Failover_Partner");
                connectionString += "Failover Partner=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Max_Pool_Size", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Max_Pool_Size");
                connectionString += "Max Pool Size=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Min_Pool_Size", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Min_Pool_Size");
                connectionString += "Min Pool Size=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Multiple_Active_Result_Sets", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Multiple_Active_Result_Sets");
                connectionString += "MultipleActiveResultSets=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Multi_Subnet_Failover", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Multi_Subnet_Failover");
                connectionString += "MultiSubnetFailover=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Network_Library", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Network_Library");
                connectionString += "Network Library=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Packet_Size", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Packet_Size");
                connectionString += "Packet Size=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Pool_Blocking_Period", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Pool_Blocking_Period");
                connectionString += "PoolBlockingPeriod=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Pooling", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Pooling");
                connectionString += "Pooling=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Replication", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Replication");
                connectionString += "Replication=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Transaction_Binding", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Transaction_Binding");
                connectionString += "Transaction Binding=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Transparent_Network_IP_Resolution", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Transparent_Network_IP_Resolution");
                connectionString += "TransparentNetworkIPResolution=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Type_System_Version", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Type_System_Version");
                connectionString += "Type System Version=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("User_Instance", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "User_Instance");
                connectionString += "User Instance=" + iniProperty.ToString() + ";";
            }

            if (iniFile.KeyExists("Workstation_ID", "Config"))
            {
                var iniProperty = iniFile.ReadINI("Config", "Workstation_ID");
                connectionString += "Workstation ID=" + iniProperty.ToString() + ";";
            }
        }

        // Функция получения проекта
        public Project GetProject()
        {
            try
            {
                var selectionSetProject = new SelectionSet();
                var project = selectionSetProject.GetCurrentProject(true);
                if (project == null) // Проверяем, получили или нет проект
                {
                    var projectIsNull = new BaseException("Проект пуст (функция GetProject)", MessageLevel.Error);
                    throw projectIsNull;
                }
                return project;
            }
            catch
            {
                var projectErr = new BaseException("Ошибка в функции GetProject", MessageLevel.Error);
                throw projectErr;
            }
        }

        // Функция получения ссылок на изделия в проекте
        public ArticleReference[] GetArticleReferences(Project project)
        {
            const int MinimalLength = 1;
            try
            {
                var objectFinder = new DMObjectsFinder(project);
                var articleReferences = objectFinder.GetArticleReferences(null);
                if (articleReferences.Length < MinimalLength) // Если ссылок нет
                {
                    var articleReferencesIsNull = new BaseException("Ссылки ArticleReferences пусты! (функция GetArticleReferences)", MessageLevel.Error);
                    throw articleReferencesIsNull;
                }
                return articleReferences;
            }
            catch
            {
                var articleReferencesErr = new BaseException("Ошибка в функции GetArticleReferences", MessageLevel.Error);
                throw articleReferencesErr;
            }
        }

        // Функция фильтрации ссылок на изделия по необходимым параметрам
        public List<ArticleReference> FilterArticleReferences(ArticleReference[] articleReferences)
        {
            const int MinimalLength = 1;
            const int ArticleNameStringNumber = 0;
            const int ArticleDeviceTypeStringNumber = 1;
            const int ArticleTypeStringNumber = 0;
            try
            {
                var filteredArticleReferences = new List<ArticleReference>();
                foreach (ArticleReference reference in articleReferences)
                {
                    // Example: +CAB5-F8 || +CAB1-1-F8
                    var splitedReferenceIdentifyingName = reference.IdentifyingName.Split('-');
                    if (splitedReferenceIdentifyingName.Length > MinimalLength)
                    {
                        var articleDeviceType = splitedReferenceIdentifyingName[ArticleDeviceTypeStringNumber];
                        var articleType = splitedReferenceIdentifyingName[ArticleTypeStringNumber];
                        // Проверяем корректность, буква 'W' - кабели, они не нужны
                        // Ловим символ между CAB1 и F8 (пример), что бы понять, что пришло
                        var isDigit = int.TryParse(articleDeviceType.ToString(), out int fictiv);
                        if ((isDigit == false) && (fictiv < 100))
                        {
                            // Цифры нет, все штатно
                            if ((articleDeviceType[ArticleNameStringNumber] != 'W') && (articleType.Length > MinimalLength))
                            {
                                filteredArticleReferences.Add(reference);
                            }
                        }
                        else
                        {
                            // Если цифра есть, то сдвигаем проверку на 1 индекс вправо
                            // Так как теперь компонент лежит там
                            articleDeviceType = splitedReferenceIdentifyingName[ArticleDeviceTypeStringNumber + 1];
                            if ((articleDeviceType[ArticleNameStringNumber] != 'W') && (articleType.Length > MinimalLength))
                            {
                                filteredArticleReferences.Add(reference);
                            }
                        }
                    }
                }
                if (filteredArticleReferences.Count < MinimalLength) // Если список пуст
                {
                    var filteredArticleReferencesListIsNull = new BaseException("Отфильтрованный список пуст! (функция FilterArticleReferences)", MessageLevel.Error);
                    throw filteredArticleReferencesListIsNull;
                }
                return filteredArticleReferences;
            }
            catch
            {
                var filterArticleReferencesErr = new BaseException("Ошибка в функции FilterArticleReferences", MessageLevel.Error);
                throw filterArticleReferencesErr;
            }
        }

        // Функция получения имени изделия (одного)
        public string GetProjectArticleName(ArticleReference articleReference) // Получение имени изделия из IdentifyingName для 1 ссылки
        {
            const int ArticleNameStringNumber = 1;
            const int ExtraArticleNameStringNumber = 1;
            var articleName = string.Empty;
            try
            {
                // +CAB1-1-FQT1 - нештатно; +CAB1-FQT1 - штатно
                // [0] +CAB1; [1] 1; [2] FQT1;
                var mainStringSplit = articleReference.IdentifyingName.Split('-');
                var noReadyName = mainStringSplit[0].Split('+');
                var isDigit = int.TryParse(mainStringSplit[ExtraArticleNameStringNumber].ToString(), out int fictiv);
                if ((isDigit == false) && (fictiv < 100))
                {
                    articleName = noReadyName[ArticleNameStringNumber];
                }
                else
                {
                    articleName = noReadyName[ArticleNameStringNumber] + "-" + mainStringSplit[ExtraArticleNameStringNumber].ToString();
                }
            }
            catch
            {
                var getProjProdNameExcept = new BaseException("Ошибка в функции GetProjectArticleName", MessageLevel.Error);
                throw getProjProdNameExcept;
            }
            if (articleName == string.Empty)
            {
                var productNameException = new BaseException("articleName пусто! (функция GetProjectArticleName)", MessageLevel.Error);
                throw productNameException;
            }
            return articleName;
        }

        // Функция получения массива имен изделий, которые содержатся в проекте (шкаф,коагулятор и др)
        // Получу все изделия в т.ч с повторением
        public string[] GetProjectArticlesNames(List<ArticleReference> articleReferences)
        {
            var filteredProjectArticlesNames = new string[0];
            try
            {
                var projectArticlesNames = new string[articleReferences.Count];
                for (int item = 0; item < articleReferences.Count; item++)
                {
                    projectArticlesNames[item] = GetProjectArticleName(articleReferences[item]);
                }
                filteredProjectArticlesNames = projectArticlesNames.Distinct().ToArray();
            }
            catch
            {
                var getProjectProductNamesException = new BaseException("Ошибка в функции GetProjectArticlesNames", MessageLevel.Error);
                throw getProjectProductNamesException;
            }
            return filteredProjectArticlesNames;
        }

        // Функция получения списка списков всех компонентов по каждому изделию
        public List<List<string>> GetArticlesComponents(string[] projectArticlesNames, List<ArticleReference> articleReferences)
        {
            try
            {
                // Первый список - соответствует string[] projectProductNames
                // articleReferences - содержит элементы которые есть в projectProductNames
                var articlesComponents = new List<List<string>>();
                for (int firstItem = 0; firstItem < projectArticlesNames.Length; firstItem++)
                {
                    var components = new List<string>();
                    for (int secondItem = 0; secondItem < articleReferences.Count; secondItem++)
                    {
                        if (projectArticlesNames[firstItem] == GetProjectArticleName(articleReferences[secondItem]))
                        {
                            try
                            {
                                components.Add(articleReferences[secondItem].Article.PartNr);
                            }
                            catch
                            {
                                var baseException = new BaseException("Ошибка во внутреннем списке (функция GetArticlesComponents)", MessageLevel.Error);
                                throw baseException;
                            }
                        }
                    }
                    components.Sort();
                    articlesComponents.Add(components);
                }
                return articlesComponents;
            }
            catch
            {
                var productArticlesErr = new BaseException("Ошибка в функции GetArticlesComponents", MessageLevel.Error);
                throw productArticlesErr;
            }
        }

        // Функция получения количества каждого из компонентов в изделии
        public List<List<ComponentShortDescription>> GetArticleComponentsCount(string[] projectArticlesNames, List<List<string>> articleComponents)
        {
            var articleComponentsShortDesctiption = new List<List<ComponentShortDescription>>();
            for (int item = 0; item < projectArticlesNames.Length; item++)
            {
                var previouseElem = string.Empty;
                var currentElem = string.Empty;
                try
                {
                    var shortDescriptions = new List<ComponentShortDescription>();
                    for (int innerItem = 0; innerItem < articleComponents[item].Count; innerItem++)
                    {
                        currentElem = articleComponents[item][innerItem];
                        if (currentElem != previouseElem)
                        {
                            var componentShortDescription = new ComponentShortDescription
                            {
                                Count = 0,
                                PartNumber = articleComponents[item][innerItem]
                            };
                            for (int componentNumber = 0; componentNumber < articleComponents[item].Count; componentNumber++)
                            {
                                if (articleComponents[item][componentNumber] == componentShortDescription.PartNumber)
                                {
                                    componentShortDescription.Count++;
                                }
                            }
                            previouseElem = currentElem;
                            shortDescriptions.Add(componentShortDescription);
                        }
                    }
                    articleComponentsShortDesctiption.Add(shortDescriptions);
                }
                catch
                {
                    var prodArticleCountErr = new BaseException("Ошибка в функции GetArticleComponentsCount", MessageLevel.Error);
                    throw prodArticleCountErr;
                }
            }
            return articleComponentsShortDesctiption;
        }

        // Функция получения cписка компонентов в проекте
        public Article[] GetProjectComponents(Project project, List<ArticleReference> articleReferences)
        {
            try
            {
                // Инициализация
                var componentsNames = new string[0];
                var componentsReferences = new string[articleReferences.Count];
                // Записываю все PartNr компонентов в массив
                for (int item = 0; item < articleReferences.Count; item++)
                {
                    componentsReferences[item] = articleReferences[item].PartNr;
                }
                // Убираю повторения и сортирую
                componentsNames = componentsReferences.Distinct().ToArray();
                Array.Sort(componentsNames);
                // Получаю данные по каждому из компонентов в массиве
                var projectComponents = new Article[componentsNames.Length];
                var allComponentsInProject = project.Articles;
                for (int firstCounter = 0; firstCounter < componentsNames.Length; firstCounter++)
                {
                    for (int secondCounter = 0; secondCounter < allComponentsInProject.Length; secondCounter++)
                    {
                        if (componentsNames[firstCounter] == allComponentsInProject[secondCounter].PartNr)
                        {
                            projectComponents[firstCounter] = allComponentsInProject[secondCounter];
                        }
                    }
                }
                return projectComponents;
            }
            catch
            {
                var projArticleListErr = new BaseException("Ошибка в функции GetProjectComponents", MessageLevel.Error);
                throw projArticleListErr;
            }
        }

        // Функция получения свойств компонентов для справочника
        public ComponentsFullDescription[] GetComponentsProperties(Article[] articles)
        {
            const string NullValue = null;
            const int ZeroValue = 0;
            try
            {
                var componentCatalogInfos = new ComponentsFullDescription[articles.Length];
                for (int item = 0; item < componentCatalogInfos.Length; item++)
                {
                    var langString = new MultiLangString(); // Для получения строк по языку
                    componentCatalogInfos[item] = new ComponentsFullDescription();
                    componentCatalogInfos[item].PartNumber = articles[item].Properties.ARTICLE_PARTNR.ToString();

                    if (articles[item].Properties.ARTICLE_TYPENR.IsEmpty == false)
                    {
                        componentCatalogInfos[item].TypeNumber = articles[item].Properties.ARTICLE_TYPENR;
                    }
                    else
                    {
                        componentCatalogInfos[item].TypeNumber = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_ORDERNR.IsEmpty == false)
                    {
                        componentCatalogInfos[item].OrderNumber = articles[item].Properties.ARTICLE_ORDERNR;

                    }
                    else
                    {
                        componentCatalogInfos[item].OrderNumber = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_MANUFACTURER.IsEmpty == false)
                    {
                        componentCatalogInfos[item].ManufacturerSmallName = articles[item].Properties.ARTICLE_MANUFACTURER;
                    }
                    else
                    {
                        componentCatalogInfos[item].ManufacturerSmallName = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_MANUFACTURER_NAME.IsEmpty == false)
                    {
                        componentCatalogInfos[item].ManufacturerFullName = articles[item].Properties.ARTICLE_MANUFACTURER_NAME;
                    }
                    else
                    {
                        componentCatalogInfos[item].ManufacturerFullName = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_SUPPLIER.IsEmpty == false)
                    {
                        componentCatalogInfos[item].SupplierSmallName = articles[item].Properties.ARTICLE_SUPPLIER;
                    }
                    else
                    {
                        componentCatalogInfos[item].SupplierSmallName = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_SUPPLIER_NAME.IsEmpty == false)
                    {
                        componentCatalogInfos[item].SupplierFullName = articles[item].Properties.ARTICLE_SUPPLIER_NAME;
                    }
                    else
                    {
                        componentCatalogInfos[item].SupplierFullName = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_DESCR1.IsEmpty == false)
                    {
                        langString = articles[item].Properties.ARTICLE_DESCR1;
                        componentCatalogInfos[item].Description1 = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[item].Description1 = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_DESCR2.IsEmpty == false)
                    {
                        langString = articles[item].Properties.ARTICLE_DESCR2;
                        componentCatalogInfos[item].Description2 = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[item].Description2 = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_CHARACTERISTICS.IsEmpty == false)
                    {
                        componentCatalogInfos[item].TechnicalCharacteristics = articles[item].Properties.ARTICLE_CHARACTERISTICS;
                    }
                    else
                    {
                        componentCatalogInfos[item].TechnicalCharacteristics = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_NOTE.IsEmpty == false)
                    {
                        langString = articles[item].Properties.ARTICLE_NOTE;
                        componentCatalogInfos[item].Note = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                    }
                    else
                    {
                        componentCatalogInfos[item].Note = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_CROSSSECTIONFROM.IsEmpty == false)
                    {
                        componentCatalogInfos[item].TerminalCrossSectionFrom = articles[item].Properties.ARTICLE_CROSSSECTIONFROM;
                    }
                    else
                    {
                        componentCatalogInfos[item].TerminalCrossSectionFrom = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_CROSSSECTIONTILL.IsEmpty == false)
                    {
                        componentCatalogInfos[item].TerminalCrossSectionTo = articles[item].Properties.ARTICLE_CROSSSECTIONTILL;
                    }
                    else
                    {
                        componentCatalogInfos[item].TerminalCrossSectionTo = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_ELECTRICALCURRENT.IsEmpty == false)
                    {
                        componentCatalogInfos[item].ElectricalCurrent = articles[item].Properties.ARTICLE_ELECTRICALCURRENT;
                    }
                    else
                    {
                        componentCatalogInfos[item].ElectricalCurrent = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_ELECTRICALPOWER.IsEmpty == false)
                    {
                        componentCatalogInfos[item].ElectricalSwitchingCapacity = articles[item].Properties.ARTICLE_ELECTRICALPOWER;
                    }
                    else
                    {
                        componentCatalogInfos[item].ElectricalSwitchingCapacity = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_VOLTAGE.IsEmpty == false)
                    {
                        componentCatalogInfos[item].Voltage = articles[item].Properties.ARTICLE_VOLTAGE;
                    }
                    else
                    {
                        componentCatalogInfos[item].Voltage = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_VOLTAGETYPE.IsEmpty == false)
                    {
                        componentCatalogInfos[item].VoltageType = articles[item].Properties.ARTICLE_VOLTAGETYPE;
                    }
                    else
                    {
                        componentCatalogInfos[item].VoltageType = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_HEIGHT.IsEmpty == false)
                    {
                        componentCatalogInfos[item].Height = articles[item].Properties.ARTICLE_HEIGHT;
                    }
                    else
                    {
                        componentCatalogInfos[item].Height = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_WIDTH.IsEmpty == false)
                    {
                        componentCatalogInfos[item].Width = articles[item].Properties.ARTICLE_WIDTH.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].Width = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_DEPTH.IsEmpty == false)
                    {
                        componentCatalogInfos[item].Depth = articles[item].Properties.ARTICLE_DEPTH.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].Depth = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_WEIGHT.IsEmpty == false)
                    {
                        componentCatalogInfos[item].Weight = articles[item].Properties.ARTICLE_WEIGHT.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].Weight = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_MOUNTINGSITE.IsEmpty == false)
                    {
                        componentCatalogInfos[item].MountingSiteID = articles[item].Properties.ARTICLE_MOUNTINGSITE;
                    }
                    else
                    {
                        componentCatalogInfos[item].MountingSiteID = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_MOUNTINGSPACE.IsEmpty == false)
                    {
                        componentCatalogInfos[item].MountingSpace = articles[item].Properties.ARTICLE_MOUNTINGSPACE.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].MountingSpace = ZeroValue;
                    }

                    componentCatalogInfos[item].PartGroup = GetPartGroup(articles[item]); // 0-00-000

                    if (articles[item].Properties.ARTICLE_QUANTITYUNIT.IsEmpty == false)
                    {
                        langString = articles[item].Properties.ARTICLE_QUANTITYUNIT;
                        var filterString = langString.GetStringToDisplay(ISOCode.Language.L_ru_RU).Trim();
                        componentCatalogInfos[item].QuantityUnit = filterString;
                    }
                    else
                    {
                        componentCatalogInfos[item].QuantityUnit = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_PACKAGINGQUANTITY.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PackagingQuantity = articles[item].Properties.ARTICLE_PACKAGINGQUANTITY.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].PackagingQuantity = ZeroValue;
                    }

                    if (articles[item].Properties.PART_LASTCHANGE.IsEmpty == false)
                    {
                        componentCatalogInfos[item].LastChange = articles[item].Properties.PART_LASTCHANGE;
                    }
                    else
                    {
                        componentCatalogInfos[item].LastChange = NullValue;
                    }

                    if (articles[item].Properties.ARTICLE_SALESPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[item].SalesPrice1 = articles[item].Properties.ARTICLE_SALESPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].SalesPrice1 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_SALESPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[item].SalesPrice2 = articles[item].Properties.ARTICLE_SALESPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].SalesPrice2 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_PURCHASEPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PurchasePrice1 = articles[item].Properties.ARTICLE_PURCHASEPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].PurchasePrice1 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_PURCHASEPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PurchasePrice2 = articles[item].Properties.ARTICLE_PURCHASEPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].PurchasePrice2 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_PACKAGINGPRICE_1.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PackagingPrice1 = articles[item].Properties.ARTICLE_PACKAGINGPRICE_1.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].PackagingPrice1 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_PACKAGINGPRICE_2.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PackagingPrice2 = articles[item].Properties.ARTICLE_PACKAGINGPRICE_2.ToDouble();
                    }
                    else
                    {
                        componentCatalogInfos[item].PackagingPrice2 = ZeroValue;
                    }

                    if (articles[item].Properties.ARTICLE_PRICEUNIT.IsEmpty == false)
                    {
                        componentCatalogInfos[item].PriceUnit = articles[item].Properties.ARTICLE_PRICEUNIT;
                    }
                    else
                    {
                        componentCatalogInfos[item].PriceUnit = ZeroValue;
                    }
                }
                return componentCatalogInfos;
            }
            catch
            {
                var artPropErr = new BaseException("Ошибка в функции GetComponentsProperties", MessageLevel.Error);
                throw artPropErr;
            }
        }

        // Функция получения шаблона типа "0-00-000"
        public string GetPartGroup(Article article)
        {
            try
            {
                // Распарсили группы и получили "дерево"
                var partGroup = String.Format("{0:0}-{1:00}-{2:000}",
                article.Properties.ARTICLE_PRODUCTTOPGROUP.ToInt(),
                article.Properties.ARTICLE_PRODUCTGROUP.ToInt(),
                article.Properties.ARTICLE_PRODUCTSUBGROUP.ToInt());
                return partGroup;
            }
            catch
            {
                var getPartGroupErr = new BaseException("Ошибка в функции GetPartGroup", MessageLevel.Error);
                throw getPartGroupErr;
            }
        }

        // Функция записи данных в справочник
        public void FillComponentCatalog(ComponentsFullDescription[] componentCatalogInfos)
        {
            const int MinimalLength = 1;
            using (DataBaseContext dataBaseConnection = new DataBaseContext(connectionString))
            {
                foreach (ComponentsFullDescription componentCatalogInfo in componentCatalogInfos)
                {
                    // Если в справочнике количество таких PartNumber = 0 (нет таких), то пишем в справочник
                    if (dataBaseConnection.ComponentCatalogs.Where(o => o.PartNumber == componentCatalogInfo.PartNumber).Count() < MinimalLength)
                    {
                        var priceList = new PriceList(); // Для прайса
                        var componentCatalog = new ComponentCatalog(); // Для свойств
                        var quantityUnit = new QuantityUnit(); // Для ед. измерения
                        var mountingSite = new MountingSite(); // Для монтажной поверхности

                        // Обработка QuantityUnit для получения ID
                        try
                        {
                            if ((componentCatalogInfo.QuantityUnit != null) && (componentCatalogInfo.QuantityUnit != string.Empty))
                            {
                                quantityUnit = dataBaseConnection.QuantityUnits.Where(
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
                                    dataBaseConnection.QuantityUnits.Add(quantityUnit);
                                    dataBaseConnection.SaveChanges();
                                    quantityUnit = dataBaseConnection.QuantityUnits.Where(
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
                            var QuantityUnitCheckErr = new BaseException("Ошибка в преобразовании и поиске QuantityUnitID", MessageLevel.Error);
                            throw QuantityUnitCheckErr;
                        }

                        // Поиск монтажной поверхности по Internal value
                        try
                        {
                            mountingSite = dataBaseConnection.MountingSites.Where(
                                o => o.InternalValue == componentCatalogInfo.MountingSiteID)
                                .First();
                            componentCatalog.MountingSiteID = mountingSite.MountingSiteID;
                            mountingSite = null;
                        }
                        catch
                        {
                            var mountingSiteErr = new BaseException("Ошибка в поиске и преобразовании MountingSiteID", MessageLevel.Error);
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
                            var comCatErr = new BaseException("Ошибка в присваивании данных ComponentCatalog", MessageLevel.Error);
                            throw comCatErr;
                        }

                        // Проверяем как записывается в справочник компонентов
                        try
                        {
                            dataBaseConnection.ComponentCatalogs.Add(componentCatalog);
                            dataBaseConnection.SaveChanges();
                        }
                        catch
                        {
                            var comCatAddErr = new BaseException("Ошибка в добавлении и сохранении ComponentCatalog в БД", MessageLevel.Error);
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
                            dataBaseConnection.PriceLists.Add(priceList);
                            dataBaseConnection.SaveChanges();
                        }
                        catch
                        {
                            var PriceListErr = new BaseException("Ошибка в записи прайс-листа после добавления нового компонента в справочник", MessageLevel.Error);
                            throw PriceListErr;
                        }
                    }
                    else
                    {
                        // Раз объект в справочнике есть, проверим его прайсы
                        try
                        {
                            // Получаю последний прайс объекта
                            var prisesForCheck = dataBaseConnection.PriceLists.Where(
                                o => o.PartNumber == componentCatalogInfo.PartNumber).OrderBy(
                                o1 => o1.LastChange).ToList();
                            var lastPriceList = prisesForCheck.Last();
                            var priseIsChanged = false; // Флаг на изменение прайса
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
                                var updPriseList = new PriceList();
                                updPriseList.PartNumber = componentCatalogInfo.PartNumber;
                                updPriseList.LastChange = DateTime.Now;
                                updPriseList.PackagingPrice1 = componentCatalogInfo.PackagingPrice1;
                                updPriseList.PackagingPrice2 = componentCatalogInfo.PackagingPrice2;
                                updPriseList.PriceUnit = componentCatalogInfo.PriceUnit;
                                updPriseList.PurchasePrice1 = componentCatalogInfo.PurchasePrice1;
                                updPriseList.PurchasePrice2 = componentCatalogInfo.PurchasePrice2;
                                updPriseList.SalesPrice1 = componentCatalogInfo.SalesPrice1;
                                updPriseList.SalesPrice2 = componentCatalogInfo.SalesPrice2;
                                dataBaseConnection.PriceLists.Add(updPriseList);
                                dataBaseConnection.SaveChanges();
                            }
                        }
                        catch
                        {
                            var checkPriseListErr = new BaseException("Ошибка при проверке прайс-листов", MessageLevel.Error);
                            throw checkPriseListErr;
                        }
                    }
                }
            }
        }

        // Функция получения времени из ARTICLE_LASTCHANGE
        public DateTime GetDateTimeFromString(string strForParse)
        {
            const int DateTimeStringNumber = 1;
            try
            {
                // "TSAM / 2010.10.03 10:09:44", поэтому нужна вторая часть строки
                var splittedStr = strForParse.Split('/');
                var strForConvert = splittedStr[DateTimeStringNumber].Trim();
                var gettedDateTime = Convert.ToDateTime(strForConvert);
                return gettedDateTime;
            }
            catch
            {
                var dateTimeFromStrErr = new BaseException("Ошибка в функции GetDateTimeFromString", MessageLevel.Error);
                throw dateTimeFromStrErr;
            }
        }

        // Функция записи данных спецификации в БД
        public void FillSpecification(string[] projectArticlesNames, List<List<ComponentShortDescription>> сomponentShortDescriptions, List<StructuralDescription> structuralDescriptions)
        {
            using (DataBaseContext dataBaseConnection = new DataBaseContext(connectionString))
            {
                var selectionSetProject = new SelectionSet();
                var currentProject = selectionSetProject.GetCurrentProject(true);

                var project = new Proj(); // Модель объекта
                project.Name = currentProject.ProjectName;
                project.Executor = currentProject.Properties.PROJ_CREATOR;
                project.DateTime = currentProject.Properties.PROJ_PROJECTBEGIN;

                // Проверяю, есть ли у нас такой проект уже
                if (dataBaseConnection.Projs.Where(o => o.Name == project.Name).Count() == 0)
                {
                    try
                    {
                        // Если нету, добавляю проект
                        dataBaseConnection.Projs.Add(project);
                        dataBaseConnection.SaveChanges();
                    }
                    catch
                    {
                        var projAddErr = new BaseException("Ошибка при добавлении проекта в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw projAddErr;
                    }

                    try
                    {
                        // Теперь добавляю изделия из проекта
                        var projID = dataBaseConnection.Projs.Where(
                            o => o.Name == project.Name)
                            .Select(o1 => o1.ProjID).FirstOrDefault();
                        foreach (string articleName in projectArticlesNames)
                        {
                            // Для поиска описания изделия
                            var location = new LocationDescription();
                            var projArticle = new PArticle(); // Моделька объекта
                            // Ищу, есть ли такое имя в таблице имен
                            location = dataBaseConnection.LocationDescriptions.Where(
                                o => o.Name == articleName).FirstOrDefault();
                            if (location == null)
                            {
                                location = new LocationDescription();
                                location.Name = articleName;
                                try
                                {   // Если не найдено описание такого элемента, то ошибка
                                    location.Description = structuralDescriptions.Find(o => o.Name == articleName).Description;
                                }
                                catch
                                {
                                    // Description = Name, ставим описание равным имени
                                    location.Description = location.Name;
                                }
                                dataBaseConnection.LocationDescriptions.Add(location);
                                dataBaseConnection.SaveChanges();
                                location = dataBaseConnection.LocationDescriptions.Where(
                                o => o.Name == articleName).FirstOrDefault();
                                projArticle.LocationDesriptionID = location.LocationDescriptionID;
                                location = null;
                            }
                            else
                            {
                                // Если обозначение есть, проверяю равняется ли его имя и описание
                                var locationName = location.Name;
                                var locationDescription = location.Description;
                                if (locationName.Equals(locationDescription) == true)
                                {
                                    try
                                    {
                                        // Получаю описание текущего изделия (структурное обозначение)
                                        var currentLocationDescription = structuralDescriptions.Find(o => o.Name == articleName).Description;
                                        // Если сработало без ошибок, обновляю его
                                        location.Description = currentLocationDescription;
                                        dataBaseConnection.SaveChanges();
                                        // И присваиваю ID projArticle
                                        projArticle.LocationDesriptionID = location.LocationDescriptionID;
                                    }
                                    catch
                                    {
                                        // Если ошибка, то значит изделие все равно без описания структурного
                                        // Значит, оставляем сопоставление Name = Description
                                        projArticle.LocationDesriptionID = location.LocationDescriptionID;
                                    }
                                }
                                else
                                {
                                    // Если же описание и имя обозначений разные (введено корректное описание)
                                    projArticle.LocationDesriptionID = location.LocationDescriptionID;
                                }
                            }
                            projArticle.ProjectID = projID;
                            dataBaseConnection.PArticles.Add(projArticle);
                        }
                        dataBaseConnection.SaveChanges();
                    }
                    catch
                    {
                        var articleAddErr = new BaseException("Ошибка при добавлении изделий в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw articleAddErr;
                    }

                    try
                    {
                        // Добавим компоненты к изделиям, то есть, завершим спецификацию
                        for (int item = 0; item < projectArticlesNames.Length; item++)
                        {
                            // Айди проекта ищу
                            var projID = dataBaseConnection.Projs.Where(
                            o => o.Name == project.Name)
                            .Select(o1 => o1.ProjID).FirstOrDefault();
                            var searchArticle = projectArticlesNames[item];
                            // Ищу айди изделия, в которое буду добавлять
                            // Но надо проверить, есть ли это изделие в структурных обозначениях               
                            if (dataBaseConnection.LocationDescriptions.Where(
                                o => o.Name == searchArticle).FirstOrDefault() != null)
                            {
                                var LocationID = dataBaseConnection.LocationDescriptions.Where(
                                    o => o.Name == searchArticle)
                                    .Select(o1 => o1.LocationDescriptionID).FirstOrDefault();
                                var articleID = dataBaseConnection.PArticles.Where(
                                    o => o.ProjectID == projID).Where(
                                    o1 => o1.LocationDesriptionID == LocationID).Select(
                                    o2 => o2.PArticleID).FirstOrDefault();
                                for (int itemCounter = 0; itemCounter < сomponentShortDescriptions[item].Count; itemCounter++)
                                {
                                    // Так как индекс projectProductNames соответствует
                                    // индексу списка списков ComponentInfo
                                    var component = new Component();
                                    component.PArticleID = articleID;
                                    // Перебираю внутренний список
                                    component.PartNumber = сomponentShortDescriptions[item][itemCounter].PartNumber;
                                    component.Count = сomponentShortDescriptions[item][itemCounter].Count;
                                    dataBaseConnection.Components.Add(component);
                                    dataBaseConnection.SaveChanges();
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
                        var componentAddErr = new BaseException("Ошибка при добавлении компонентов в БД (Функция FillSpecification)", MessageLevel.Error);
                        throw componentAddErr;
                    }
                }
                else
                {
                    // Если проект есть, удалю его спецификацию
                    // И запишу новую (обновленную).
                    var proj = dataBaseConnection.Projs.Where(
                        o => o.Name == project.Name).FirstOrDefault();
                    dataBaseConnection.Projs.Remove(proj);
                    dataBaseConnection.SaveChanges();
                    FillSpecification(projectArticlesNames, сomponentShortDescriptions, structuralDescriptions);
                }
            }
        }

        // Функция получения списка структурных идентификаторов изделий
        public List<StructuralDescription> GetLocationDescriptions(Project project, string[] projectProductNames)
        {
            try
            {
                var locations = project.GetLocationObjects();
                var structuralDescriptions = new List<StructuralDescription>();
                var descriptionString = new MultiLangString();
                foreach (Location location in locations)
                {
                    // Если описание не пусто
                    if (location.Properties.LOCATION_DESCRIPTION != "")
                    {
                        var structuralDescription = new StructuralDescription();
                        for (int item = 0; item < projectProductNames.Length; item++)
                        {
                            // Если нашли совпадение в списке изделий, записываем в список
                            if (projectProductNames[item].Equals(location.Properties.LOCATION_FULLNAME))
                            {
                                descriptionString = location.Properties.LOCATION_DESCRIPTION;
                                structuralDescription.Name = location.Properties.LOCATION_FULLNAME;
                                structuralDescription.Description = descriptionString.GetStringToDisplay(ISOCode.Language.L_ru_RU);
                                structuralDescriptions.Add(structuralDescription);
                            }
                        }
                    }
                }
                return structuralDescriptions;
            }
            catch
            {
                var getLocDescrErr = new BaseException("Ошибка в функции GetLocationDescriptions", MessageLevel.Error);
                throw getLocDescrErr;
            }
        }
    }
}