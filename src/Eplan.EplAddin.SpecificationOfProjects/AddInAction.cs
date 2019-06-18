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
            var functions = new Functions();

            // Проверяем ini настройки
            functions.chekAddInIniFile();

            // Получить проект
            var currentProject = functions.GetProject();

            // Получить все ссылки на компоненты в текущем проекте
            var currentArticleReferences = functions.GetArticleReferences(currentProject);

            // Отфильтровать ссылки, оставив только те, которые привязаны к реальным объектам
            var filteredArticleReferences = functions.FilterArticleReferences(currentArticleReferences);

            // Получить список всех изделий, которые есть в проекте
            var projectArticlesNames = functions.GetProjectArticlesNames(filteredArticleReferences);

            // Получить список списков всех компонентов для каждого изделия.
            var articlesComponents = functions.GetArticlesComponents(projectArticlesNames, filteredArticleReferences);

            // Получить список списков компонентов по изделиям
            var сomponentShortDescriptions = functions.GetArticleComponentsCount(projectArticlesNames, articlesComponents);

            // Получить компоненты, которые есть в проекте
            var articleList = functions.GetProjectComponents(currentProject, filteredArticleReferences);

            // Получить данные для записи в справочник компонентов
            var ComponentsFullDescriptions = functions.GetComponentsProperties(articleList);

            // Получить список структурных обозначений изделий
            var structuralDescriptions = functions.GetLocationDescriptions(currentProject, projectArticlesNames);
            
            // Заполнить справочник компонентов
            functions.FillComponentCatalog(ComponentsFullDescriptions);

            // Заполнить спецификацию в БД
            functions.FillSpecification(projectArticlesNames, сomponentShortDescriptions, structuralDescriptions);

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
