using System.Collections.Generic;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplAddin.SpecificationOfProjects;

namespace Eplan.AplAddin.SpecificationListOfObjects
{
    class TestAction : IEplAction
    {
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            // Имя действия для связи с меню
            Name = "GetSpecification";
            Ordinal = 20;
            return true;
        }

        public bool Execute(ActionCallingContext actionCallingContext)
        {
            // Класс со всеми функциями данной dll
            FunctionsManager FM = new FunctionsManager();
            
            // Получить проект
            Project currentProject = FM.GetProject();

            // Получить все ссылки на компоненты в текущем проекте
            ArticleReference[] currentArticleReferences = FM.GetArticleReferences(currentProject);

            // Отфильтровать ссылки, оставив только те, которые привязаны к реальным объектам
            List<ArticleReference> filteredArticleReferences = FM.FilterArticleReferences(currentArticleReferences);

            // Получить список всех изделий, которые есть в проекте
            string[] projectProductNames = FM.GetProjectProductNames(filteredArticleReferences);

            // Получить список списков всех компонентов для каждого изделия.
            List<List<string>> productArticles = FM.GetProductArticles(projectProductNames, filteredArticleReferences);

            // Получить список списков компонентов по изделиям
            List<List<ComponentInfo>> specificationInfo = FM.GetProductArticleCount(projectProductNames, productArticles);

            // Получить компоненты, которые есть в проекте
            Article[] articleList = FM.GetProjectProductArticleList(currentProject, filteredArticleReferences);

            // Получить данные для записи в справочник компонентов
            ComponentCatalogInfo[] componentCatalogInfos = FM.GetArticleProperties(articleList);

            // Получить список структурных обозначений изделий
            List<LocationInfo> locationInfos = FM.GetLocationDescriptions(currentProject, projectProductNames);
            
            // Заполнить справочник компонентов
            FM.FillComponentCatalog(componentCatalogInfos);

            // Заполнить спецификацию в БД
            FM.FillSpecification(projectProductNames, specificationInfo, currentProject.ProjectName, locationInfos);

            // Оповестить об успешности
            new Decider().Decide(
                EnumDecisionType.eOkDecision,
                "Спецификация выгружена успешно!\n" + "Проект: " + currentProject.ProjectName,
                "Оповещение",
                EnumDecisionReturn.eOK,
                EnumDecisionReturn.eOK,
                "",
                false,
                EnumDecisionIcon.eINFORMATION);
            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties) { }

    }
}

    // Метод OnRegister - регистрирует наш Action под указанным именем.
    // Метод Execute - выполняются при вызове Actiona из платформы EPLAN. 
    // Метод GetActionProperties - возвращает описание нашего Actiona