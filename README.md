# Проект экспорта спецификации проекта EPLAN
Ветка содержит решение, состоящее на данный момент из двух проектов собственной разработки, а именно:
1. Надстройка для EPLAN, выгружающая данные в базу данных.
2. Библиотека классов Entity Framework для работы с базой данных.
3. Клиентское приложения для формирования документов.

## Что сделано
### Со стороны EPLAN:
1. Добавление пункта меню при регистрации Add-in в EPLAN.
2. Получение проекта и проверка его получения вынесена в отдельную функцию (*GerProject*).
3. Получение списка ссылок (*ArticleReferences*) вынесено в отдельную функцию (*GetArticleReferences*).
4. Фильтрация списка ссылок, подразумевает получение ссылок, которые непосредственно связаны с реальным объектом, вынесено в отдельную функцию  (*FilterArticleReferences*).
4. Получение списка всех изделий, которые есть в проекте, вынесено в отдельную функцию  (*GetProjectProductNames*).
5. Получение списка всех компонентов каждого изделия, вынесено в отдельную функцию  (*GetProductArticles*).
6. Получение списка компонентов по изделиям в виде "компонент - количество", вынесено в отдельную функцию  (*GetProductArticleCount*).
7. Получение компонентов, которые есть в проекте, вынесено в отдельную функцию  (*GetProjectProductArticleList*).
8. Получение данных для записи в справочник компонентов, вынесено в отдельную функцию (*GetArticleProperties*);
9. Заполнение справочника компонентов, вынесено в отдельную функцию (*FillComponentCatalog*).
10. Получение информации о структурной идентификации изделий, вынесено в отдельную функцию (*GetLocationDescriptions*)
11. Заполнение таблиц для спецификации, вынесено в отдельную функцию (*FillSpecification*).

### Со стороны клиента:
1. Создано клиентское оконное приложения Windows Forms.
2. Реализована возможность просмотра информации по таким параметрам:
2.1 Проект (*просмотр изделий и общего количества компонентов в них*).
2.2 Изделие (*просмотр компонентов в изделии и их количества*).
2.3 Компонент (*просмотр полной информации по компоненту*).
3. Сбор данных для вывода по компоненту, вынесен в отдельную функцию (*GetPropertiesInfos*).
4. Формирование спецификации и заявки на склад по проекту.
5. Хранения, добавления, удаления  документации по проекту и просмотр её через интернет.
6. Открытие хранимых документов, инициируя запуск программы из клиента.

### Последние изменения (upd. 24.05.2019)
1. Настройка доступа к каталогу с документами через ini-файл (*для клиента*).
2. Настройка доступа к базе данных (*строки подключения*) SQL Server через ini-файл (*для надстройки*).
3. Доступ к базе данных SQL Server для клиента организован последством изменения строки подключения в файле config формата.
4. Исправлены мелкие баги и недочеты, найденные в ходе локального тестирования.

## Примечание
1. Комментарии в программном коде служат облегчить понимание кода.
2. Возможны некорректные для понимания имен.