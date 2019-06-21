using DataBaseLibrary;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Component = DataBaseLibrary.Component;

namespace SpecificationOfProject.Client
{
    class Functions
    {
        // Функция получения свойств выбранного компонента
        public List<ComponentProperties> GetSelectedComponentProperties(string componentName)
        {
            const int MinStringManufacturerLength = 4;
            const int MinStringMountingSpaceLength = 6;
            const int MinStringProportionsLength = 5;
            const string EmptyStringWithZero = "0";
            try
            {
                var componentPropertiesInfos = new List<ComponentProperties>();
                using (DataBaseContext dataBaseConnection = new DataBaseContext())
                {
                    var componentCatalog = dataBaseConnection.ComponentCatalogs.Where(
                        o => o.PartNumber == componentName).FirstOrDefault();
                    var componentProperty = new ComponentProperties();

                    // Номер изделия
                    componentProperty.Property = "Номер изделия";
                    componentProperty.Value = componentCatalog.PartNumber;
                    componentPropertiesInfos.Add(componentProperty);

                    // Тип изделия
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Номер типа";
                    componentProperty.Value = componentCatalog.TypeNumber;
                    componentPropertiesInfos.Add(componentProperty);

                    // Номер для заказа
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Номер для заказа";
                    // Если его нет, то ставим его пустым
                    if (componentCatalog.OrderNumber != string.Empty)
                    {
                        componentProperty.Value = componentCatalog.OrderNumber;
                    }
                    else
                    {
                        componentProperty.Value = string.Empty;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Производитель
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Производитель";
                    componentProperty.Value = componentCatalog.ManufacturerFullName
                         + " (" + componentCatalog.ManufacturerSmallName + ")";
                    // Проверяем, кроме скобок есть что нибудь или нет
                    if (componentProperty.Value.Length < MinStringManufacturerLength)
                    {
                        componentProperty.Value = string.Empty;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Поставщик
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Поставщик";
                    componentProperty.Value = componentCatalog.SupplierFullName
                        + " (" + componentCatalog.SupplierSmallName + ")";
                    if (componentProperty.Value.Length < MinStringManufacturerLength)
                    {
                        componentProperty.Value = string.Empty;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Обозначение 1
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Обозначение 1";
                    componentProperty.Value = componentCatalog.Description1;
                    componentPropertiesInfos.Add(componentProperty);

                    // Обозначение 2
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Обозначение 2";
                    componentProperty.Value = componentCatalog.Description2;
                    componentPropertiesInfos.Add(componentProperty);

                    // Описание
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Описание";
                    componentProperty.Value = componentCatalog.Note;
                    componentPropertiesInfos.Add(componentProperty);

                    // Клеммы: сечение от
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Клеммы: сечение от";
                    componentProperty.Value = componentCatalog.TerminalCrossSectionFrom;
                    componentPropertiesInfos.Add(componentProperty);

                    // Клеммы: сечение до
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Клеммы: сечение до";
                    componentProperty.Value = componentCatalog.TerminalCrossSectionTo;
                    componentPropertiesInfos.Add(componentProperty);

                    // Ток
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Ток";
                    componentProperty.Value = componentCatalog.ElectricalCurrent;
                    componentPropertiesInfos.Add(componentProperty);

                    // Коммутационная способность
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Коммутационная способность";
                    componentProperty.Value = componentCatalog.ElectricalSwitchingCapacity;
                    componentPropertiesInfos.Add(componentProperty);

                    // Напряжение
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Напряжение";
                    componentProperty.Value = componentCatalog.Voltage;
                    componentPropertiesInfos.Add(componentProperty);

                    // Тип напряжения
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Тип напряжения";
                    componentProperty.Value = componentCatalog.VoltageType;
                    componentPropertiesInfos.Add(componentProperty);

                    // Высота
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Высота";
                    componentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Height, 3) + " мм");
                    if (componentProperty.Value.Length < MinStringProportionsLength)
                    {
                        componentProperty.Value = EmptyStringWithZero;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Ширина
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Ширина";
                    componentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Width, 3) + " мм");
                    if (componentProperty.Value.Length < MinStringProportionsLength)
                    {
                        componentProperty.Value = EmptyStringWithZero;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Глубина
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Глубина";
                    componentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Depth, 3) + " мм");
                    if (componentProperty.Value.Length < MinStringProportionsLength)
                    {
                        componentProperty.Value = EmptyStringWithZero;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Вес
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Вес";
                    componentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Weight, 3) + " кг");
                    if (componentProperty.Value.Length < MinStringProportionsLength)
                    {
                        componentProperty.Value = EmptyStringWithZero;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Монтажная поверхность
                    componentProperty = new ComponentProperties();
                    string mountingSite = dataBaseConnection.MountingSites.Where(
                        o => o.InternalValue == componentCatalog.MountingSiteID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentProperty.Property = "Монтажная поверхность";
                    componentProperty.Value = mountingSite;
                    componentPropertiesInfos.Add(componentProperty);

                    // Занимаемое пространство
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Занимаемое пространство";
                    componentProperty.Value = Convert.ToString(Math.Round(componentCatalog.MountingSpace, 3) + " мм2");
                    if (componentProperty.Value.Length < MinStringMountingSpaceLength)
                    {
                        componentProperty.Value = EmptyStringWithZero;
                    }
                    componentPropertiesInfos.Add(componentProperty);

                    // Область - раздел - сфера
                    componentProperty = new ComponentProperties();
                    // Спличу строку
                    var PartGroup = componentCatalog.PartGroup.Split('-');
                    var areaID = Convert.ToInt32(PartGroup[0]);
                    var sectionID = Convert.ToInt32(PartGroup[1]);
                    var sphereID = Convert.ToInt32(PartGroup[2]);
                    var area = dataBaseConnection.Areas.Where(o => o.InternalValue == areaID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    var section = dataBaseConnection.Sections.Where(o => o.InternalValue == sectionID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    var sphere = dataBaseConnection.Areas.Where(o => o.InternalValue == sphereID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentProperty.Property = "Область";
                    componentProperty.Value = area;
                    componentPropertiesInfos.Add(componentProperty);
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Раздел";
                    componentProperty.Value = section;
                    componentPropertiesInfos.Add(componentProperty);
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Сфера";
                    componentProperty.Value = sphere;
                    componentPropertiesInfos.Add(componentProperty);

                    // Единица измерения
                    componentProperty = new ComponentProperties();
                    var quantityUnit = dataBaseConnection.QuantityUnits.Where(
                        o => o.QuantityUnitID == componentCatalog.QuantityUnitID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentProperty.Property = "Единица измерения";
                    componentProperty.Value = quantityUnit;
                    componentPropertiesInfos.Add(componentProperty);

                    // Количество упаковки
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Количество в упаковке";
                    componentProperty.Value = Convert.ToString(componentCatalog.PackagingQuantity);
                    componentPropertiesInfos.Add(componentProperty);

                    // Последнее изменение
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Последнее изменение";
                    componentProperty.Value = componentCatalog.LastChange;
                    componentPropertiesInfos.Add(componentProperty);

                    // Цена продажи, Валюта 1
                    // Выводится ПОСЛЕДНИЙ актуальный прайс
                    // Нашел количество прайсов в таблице
                    // Выбрал последний (сортировка по возрастанию даты)
                    var prices = dataBaseConnection.PriceLists.Where(
                        o => o.PartNumber == componentCatalog.PartNumber).OrderBy(
                        o1 => o1.LastChange).ToList();
                    var salesPrice1 = prices.Select(
                        o => o.SalesPrice1).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Цена продажи, Валюта 1";
                    componentProperty.Value = Convert.ToString(Math.Round(salesPrice1, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Цена продажи, Валюта 2
                    var salesPrice2 = prices.Select(
                        o => o.SalesPrice2).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Цена продажи, Валюта 2";
                    componentProperty.Value = Convert.ToString(Math.Round(salesPrice2, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Закупочная цена, Валюта 1
                    var purchasePrice1 = prices.Select(
                        o => o.PurchasePrice1).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Закупочная цена, Валюта 1";
                    componentProperty.Value = Convert.ToString(Math.Round(purchasePrice1, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Закупочная цена, Валюта 2
                    var purchasePrice2 = prices.Select(
                        o => o.PurchasePrice2).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Закупочная цена, Валюта 2";
                    componentProperty.Value = Convert.ToString(Math.Round(purchasePrice2, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Цена упаковки, Валюта 1
                    var packagingPrice1 = prices.Select(
                        o => o.PackagingPrice1).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Цена упаковки, Валюта 1";
                    componentProperty.Value = Convert.ToString(Math.Round(packagingPrice1, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Цена упаковки, Валюта 2
                    var packagingPrice2 = prices.Select(
                        o => o.PackagingPrice2).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Цена упаковки, Валюта 2";
                    componentProperty.Value = Convert.ToString(Math.Round(packagingPrice2, 3));
                    componentPropertiesInfos.Add(componentProperty);

                    // Единица цены
                    var priceUnit = prices.Select(
                        o => o.PriceUnit).LastOrDefault();
                    componentProperty = new ComponentProperties();
                    componentProperty.Property = "Единица цены";
                    componentProperty.Value = Convert.ToString(priceUnit);
                    componentPropertiesInfos.Add(componentProperty);
                }
                return componentPropertiesInfos;
            }
            catch
            {
                var getPropertiesInfosErr = new Exception("Ошибка в функции GetPropertiesInfos");
                throw getPropertiesInfosErr;
            }  
        }

        // Функция формирования спецификации по проекту
        // И загрузки в перечень документов
        public void CreateProjectSpecificationDocument(string projID, string savePath, string projName)
        {
            try
            {
                //Ширина колонок от А до J
                const double WidthCellA = 4;
                const double WidthCellB = 12.8;
                const double WidthCellC = 23;
                const double WidthCellD = 15.57;
                const double WidthCellE = 15.7;
                const double WidthCellF = 11.5;
                const double WidthCellG = 5.29;
                const double WidthCellH = 5.57;
                const double WidthCellI = 24;
                const double WidthCellJ = 17;

                // Использую библиотеку EPPlus
                using (var spec = new ExcelPackage())
                {
                    // Добавляю лист в spec
                    var workSheet = spec.Workbook.Worksheets.Add("Спецификация");

                    // Глобальные настройки документа                
                    // Настраиваю ширину колонок и переносы
                    workSheet.Column(1).Width = WidthCellA;
                    workSheet.Column(2).Width = WidthCellB;
                    workSheet.Column(3).Width = WidthCellC;
                    workSheet.Column(4).Width = WidthCellD;
                    workSheet.Column(5).Width = WidthCellE;
                    workSheet.Column(6).Width = WidthCellF;
                    workSheet.Column(7).Width = WidthCellG;
                    workSheet.Column(8).Width = WidthCellH;
                    workSheet.Column(9).Width = WidthCellI;
                    workSheet.Column(10).Width = WidthCellJ;

                    // Дефолтный шрифт, размер шрифта
                    workSheet.Cells.Style.Font.Name = "Times New Roman";
                    workSheet.Cells.Style.Font.Size = 12;

                    // Ориентация и поля документа (для печати)
                    workSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    workSheet.PrinterSettings.BottomMargin = 0.7M;
                    workSheet.PrinterSettings.FooterMargin = 0.2M;
                    workSheet.PrinterSettings.HeaderMargin = 0.2M;
                    workSheet.PrinterSettings.LeftMargin = 0.2M;
                    workSheet.PrinterSettings.RightMargin = 0.2M;
                    workSheet.PrinterSettings.TopMargin = 0.7M;

                    // Заполняю шаблон
                    workSheet.Cells["B1:C1"].Merge = true;
                    workSheet.Cells["B1"].Value = "<укажите службу>";
                    workSheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A4:J4"].Merge = true;
                    workSheet.Cells["A4"].Value = "Спецификация";
                    workSheet.Cells["A4"].Style.Font.Bold = true;
                    workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A5:J5"].Merge = true;
                    workSheet.Cells["A5"].Value = "по проекту " + projName;
                    workSheet.Cells["A5"].Style.Font.Bold = true;
                    workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A6:J6"].Merge = true;
                    workSheet.Cells["A6"].Value = "№<укажите номер> от " + DateTime.Now.ToShortDateString();
                    workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A8:J5000"].Style.WrapText = true;
                    workSheet.Cells["A8:J8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A8:J8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Row(8).Height = 45;
                    workSheet.Cells["A8"].Value = "№ п\\п";
                    workSheet.Cells["B8"].Value = "Код ТНВЭД";
                    workSheet.Cells["C8"].Value = "Полное наименование ТМЦ";
                    workSheet.Cells["D8"].Value = "№ по каталогу производителя";
                    workSheet.Cells["E8"].Value = "Производитель";
                    workSheet.Cells["F8"].Value = "Поставщик";
                    workSheet.Cells["G8"].Value = "Ед. изм.";
                    workSheet.Cells["H8"].Value = "Кол-во";
                    workSheet.Cells["I8"].Value = "Доп. характеристик.";
                    workSheet.Cells["J8"].Value = "Текущий остаток у указанного получателя";

                    const int StartedRow = 8; // Стартовая строка для записи
                    int endRow = 9; // Конечная строка (на текущий момент)
                    int counterOfComponents = 0; // Количество компонентов в общем по спецификации
                    int currentRow = StartedRow + 1; // Текущая строка для записи

                    // Устанавливаю границу для шапки таблицы (десять столбцов)
                    for (int column = 0 + 1; column <= 10; column++)
                    {
                        workSheet.Cells[StartedRow, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Получаю данные компонентов для записи
                    var componentsPropertiesForSpecification = GetComponentPropertiesForSpecification(projID);

                    // Получаю данные по изделиям
                    var articlePropertiesForSpecification = GetArticlePropertiesForSpecification(projID);
                    // Под одинаковыми индексами у каждого списка находится информация по одному изделию
                    
                    // Записываю данные в документ
                    // Перебираю изделия и компоненты внутри друг друга т.к они сопоставимы
                    foreach (ArticlePropertiesForSpecification articleProperty in articlePropertiesForSpecification)
                    {
                        workSheet.Cells[currentRow, 1, currentRow, 10].Merge = true;
                        workSheet.Cells[currentRow, 1].Value =
                            "Изделие: " + "\"" + articleProperty.ArticleDescription + "\" "+ "(" + articleProperty.ArticleName + ")";
                        workSheet.Cells[currentRow, 1, currentRow, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        currentRow += 1; // Записал строку, увеличил счетчик на единицу
                        foreach (ComponentPropertiesForSpecification componentPropery in componentsPropertiesForSpecification)
                        {
                            // Записываем только, если articleID изделия = articleID компонента
                            if (articleProperty.ArticleID == componentPropery.ArticleID)
                            {
                                // Настройка строки для записи и столбцов
                                for (int column = 0 + 1; column <= 10; column++)
                                {
                                    workSheet.Cells[currentRow, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[currentRow, column].Style.Font.Name = "Times New Roman";
                                    workSheet.Cells[currentRow, column].Style.Font.Size = 10;
                                    workSheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    workSheet.Cells[currentRow, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                const int ItemNumberColumn = 1;
                                const int Description1Column = 3;
                                const int OrderNumberColumn = 4;
                                const int ManufacturerColumn = 5;
                                const int QuantityUnitColumn = 7;
                                const int ItemCountColumn = 8;
                                const int Description2Column = 9;

                                // Запись по столбцам (от 1 до 10, A > J). Не все столбцы
                                workSheet.Cells[currentRow, ItemNumberColumn].Value = counterOfComponents + 1;
                                workSheet.Cells[currentRow, Description1Column].Value = componentPropery.Description1 +
                                    " " + componentPropery.PartNumber;
                                workSheet.Cells[currentRow, OrderNumberColumn].Value = componentPropery.OrderNumber;
                                workSheet.Cells[currentRow, ManufacturerColumn].Value = componentPropery.ManufacturerFullName;
                                workSheet.Cells[currentRow, QuantityUnitColumn].Value = componentPropery.QuantityUnit;
                                workSheet.Cells[currentRow, ItemCountColumn].Value = componentPropery.Count;
                                workSheet.Cells[currentRow, Description2Column].Value = componentPropery.Description2;

                                currentRow += 1;
                                counterOfComponents++;
                            }      
                        }
                        endRow = currentRow;
                    }

                    const int FIOColumn = 9;
                    const int ServiceNameColumn = 2;

                    // Указываем начальника службы и ФИО
                    workSheet.Cells[endRow + 1, ServiceNameColumn, endRow + 1, ServiceNameColumn + 1].Merge = true;
                    workSheet.Cells[endRow + 1, ServiceNameColumn].Value = "<укажите начальника службы>";
                    workSheet.Cells[endRow + 1, FIOColumn].Value = "И.О. Фамилия";
                    // Настраиваем отображение
                    workSheet.Cells[endRow + 1, ServiceNameColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[endRow + 1, FIOColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // Сохраняем документ
                    spec.SaveAs(new FileInfo(savePath + "\\Спецификация.xlsx"));
                }
            }
            catch
            {
                var createProjectSpecificationErr = new Exception("Ошибка в функции CreateProjectSpecification");
                throw createProjectSpecificationErr;
            }
        }

        // Функция формирования заявки на склад по проекту
        // И загрузки в перечень документов
        public void CreateProjectWarehouseRequestDocument(string projID, string savePath)
        {
            try
            {
                //Ширина колонок от А до J
                const double WidthCellA = 4;
                const double WidthCellB = 12.8;
                const double WidthCellC = 23;
                const double WidthCellD = 15.57;
                const double WidthCellE = 15.7;
                const double WidthCellF = 11.5;
                const double WidthCellG = 5.29;
                const double WidthCellH = 5.57;
                const double WidthCellI = 24;
                const double WidthCellJ = 17;

                // Использую библиотеку EPPlus
                using (var spec = new ExcelPackage())
                {
                    // Добавляю лист в spec
                    var workSheet = spec.Workbook.Worksheets.Add("Заявка");

                    // Глобальные настройки документа                
                    // Настраиваю ширину колонок и переносы
                    workSheet.Column(1).Width = WidthCellA;
                    workSheet.Column(2).Width = WidthCellB;
                    workSheet.Column(3).Width = WidthCellC;
                    workSheet.Column(4).Width = WidthCellD;
                    workSheet.Column(5).Width = WidthCellE;
                    workSheet.Column(6).Width = WidthCellF;
                    workSheet.Column(7).Width = WidthCellG;
                    workSheet.Column(8).Width = WidthCellH;
                    workSheet.Column(9).Width = WidthCellI;
                    workSheet.Column(10).Width = WidthCellJ;

                    // Дефолтный шрифт, размер шрифта
                    workSheet.Cells.Style.Font.Name = "Times New Roman";
                    workSheet.Cells.Style.Font.Size = 12;

                    // Ориентация и поля документа (для печати)
                    workSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    workSheet.PrinterSettings.BottomMargin = 0.7M;
                    workSheet.PrinterSettings.FooterMargin = 0.2M;
                    workSheet.PrinterSettings.HeaderMargin = 0.2M;
                    workSheet.PrinterSettings.LeftMargin = 0.2M;
                    workSheet.PrinterSettings.RightMargin = 0.2M;
                    workSheet.PrinterSettings.TopMargin = 0.7M;

                    // Заполняю шаблон
                    workSheet.Cells["B1:C1"].Merge = true;
                    workSheet.Cells["B1"].Value = "<укажите службу>";
                    workSheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A4:J4"].Merge = true;
                    workSheet.Cells["A4"].Value = "Заявка";
                    workSheet.Cells["A4"].Style.Font.Bold = true;
                    workSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A5:J5"].Merge = true;
                    workSheet.Cells["A5"].Value = "отделу снабжения на закупку товарно-материальных ценностей";
                    workSheet.Cells["A5"].Style.Font.Bold = true;
                    workSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["A6:J6"].Merge = true;
                    workSheet.Cells["A6"].Value = "№<укажите номер> от " + DateTime.Now.ToShortDateString();
                    workSheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    workSheet.Cells["B7"].Value = "Получатель:";
                    workSheet.Cells["C7:D7"].Merge = true;
                    workSheet.Cells["C7"].Value = "<укажите получателя>";

                    workSheet.Cells["B8"].Value = "Инициатор:";
                    workSheet.Cells["C8:D8"].Merge = true;
                    workSheet.Cells["C8"].Value = "<укажите инициатора>";

                    workSheet.Cells["B9"].Value = "Срок:";
                    workSheet.Cells["C9:D9"].Merge = true;
                    workSheet.Cells["C9"].Value = "<укажите срок>";

                    workSheet.Cells["A10:J5000"].Style.WrapText = true;
                    workSheet.Cells["A10:J10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells["A10:J10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Row(10).Height = 45;
                    workSheet.Cells["A10"].Value = "№ п\\п";
                    workSheet.Cells["B10"].Value = "Код ТНВЭД";
                    workSheet.Cells["C10"].Value = "Полное наименование ТМЦ";
                    workSheet.Cells["D10"].Value = "№ по каталогу производителя";
                    workSheet.Cells["E10"].Value = "Производитель";
                    workSheet.Cells["F10"].Value = "Поставщик";
                    workSheet.Cells["G10"].Value = "Ед. изм.";
                    workSheet.Cells["H10"].Value = "Кол-во";
                    workSheet.Cells["I10"].Value = "Доп. характеристик.";
                    workSheet.Cells["J10"].Value = "Текущий остаток у указанного получателя";

                    for (int column = 1; column <= 10; column++)
                    {
                        workSheet.Cells[10, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Получаю все компоненты
                    var nonFilteredComponents = GetNonFilteredComponents(projID);

                    // Фильтрую компоненты
                    var filteredComponents = GetFilteredComponents(nonFilteredComponents);

                    // Получаю полные данные для заявки
                    var dataForWarehouseRequests = GetPropertiesForWarehouseRequest(filteredComponents);

                    // Заполняю
                    const int StartedRow = 10;
                    int endRow = 11; // Конец строк с компонентами
                    for (int item = 0; item < dataForWarehouseRequests.Count; item++)
                    {
                        // Настройка строки для записи и столбцов
                        for (int column = 1; column <= 10; column++)
                        {
                            workSheet.Cells[StartedRow + item + 1, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[StartedRow + item + 1, column].Style.Font.Name = "Times New Roman";
                            workSheet.Cells[StartedRow + item + 1, column].Style.Font.Size = 10;
                            workSheet.Cells[StartedRow + item + 1, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[StartedRow + item + 1, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        const int ItemNumberColumn = 1;
                        const int Description1Column = 3;
                        const int OrderNumberColumn = 4;
                        const int ManufacturerColumn = 5;
                        const int QuantityUnitColumn = 7;
                        const int ItemCountColumn = 8;
                        const int Description2Column = 9;

                        // Запись по столбцам (от 1 до 10, A > J)
                        workSheet.Cells[StartedRow + item + 1, ItemNumberColumn].Value = item + 1;
                        workSheet.Cells[StartedRow + item + 1, Description1Column].Value = dataForWarehouseRequests[item].Description1 +
                            " " + dataForWarehouseRequests[item].PartNumber;
                        workSheet.Cells[StartedRow + item + 1, OrderNumberColumn].Value = dataForWarehouseRequests[item].OrderNumber;
                        workSheet.Cells[StartedRow + item + 1, ManufacturerColumn].Value = dataForWarehouseRequests[item].ManufacturerFullName;
                        workSheet.Cells[StartedRow + item + 1, QuantityUnitColumn].Value = dataForWarehouseRequests[item].QuantityUnit;
                        workSheet.Cells[StartedRow + item + 1, ItemCountColumn].Value = dataForWarehouseRequests[item].Count;
                        workSheet.Cells[StartedRow + item + 1, Description2Column].Value = dataForWarehouseRequests[item].Description2;
                        endRow = StartedRow + item + 1;
                    }

                    const int FIOColumn = 9;
                    const int ServiceNameColumn = 2;

                    // Указываем начальника службы и ФИО
                    workSheet.Cells[endRow + 2, ServiceNameColumn, endRow + 2, ServiceNameColumn + 1].Merge = true;
                    workSheet.Cells[endRow + 2, ServiceNameColumn].Value = "<укажите начальника службы>";
                    workSheet.Cells[endRow + 2, FIOColumn].Value = "И.О. Фамилия";
                    // Настраиваем отображение
                    workSheet.Cells[endRow + 2, ServiceNameColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[endRow + 2, FIOColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // Сохраняем документ
                    spec.SaveAs(new FileInfo(savePath + "\\Заявка на склад.xlsx"));
                }
            }
            catch
            {
                var createProjectWarehouseRequestErr = new Exception("Ошибка в функции CreateProjectWarehouseRequest");
                throw createProjectWarehouseRequestErr;
            }      
        }

        // Функция получения списка компонентов для заявки на склад
        public List<ComponentInitialProperties> GetNonFilteredComponents(string projID)
        {
            try
            {
                // Так как ID записан строкой в списке, то конвертируем
                var curProjectID = Convert.ToInt32(projID);
                var nonFilteredComponents = new List<ComponentInitialProperties>();
                using (DataBaseContext dataBaseConnection = new DataBaseContext())
                {
                    // Получил изделия по проекту
                    var pArticles = dataBaseConnection.PArticles.Where(
                        o => o.ProjectID == curProjectID).ToList();
                    // Перебираю проекты и заполняю компоненты
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Получил список всех компонентов в изделии
                        var components = dataBaseConnection.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToList();
                        // Записываю базовую информацию по компонентам
                        foreach (Component component in components)
                        {
                            var componentInitialProperty = new ComponentInitialProperties();
                            componentInitialProperty.PartNumber = component.PartNumber;
                            componentInitialProperty.Count = component.Count;
                            nonFilteredComponents.Add(componentInitialProperty);
                        }
                    }
                }
                return nonFilteredComponents;
            }
            catch
            {
                var getNonFilteredComponentsErr = new Exception("Ошибка в функции GetNonFilteredComponents");
                throw getNonFilteredComponentsErr;
            }   
        }
        
        // Функция фильтрации списка компонентов для заявки на склад
        public List<ComponentInitialProperties> GetFilteredComponents(List<ComponentInitialProperties> nonFilteredComponents)
        {
            try
            {
                var filteredComponents = new List<ComponentInitialProperties>();
                // Получил отфильтрованные компоненты, но не посчитал их количество
                var nonFilteredPartNumbers = new List<string>();
                foreach (ComponentInitialProperties componentInitialProperty in nonFilteredComponents)
                {
                    nonFilteredPartNumbers.Add(componentInitialProperty.PartNumber);
                }
                var partNumbersArray = nonFilteredPartNumbers.Distinct().ToArray();
                // Считаю количество компонентов и формирую список
                for (int counter = 0; counter < partNumbersArray.Length; counter++)
                {
                    var componentCount = 0; // Количество компонентов
                    var componentInitialProperty = new ComponentInitialProperties();
                    //Прохожу по всему списку неотфильтрованному и считаю количество компонентов
                    for (int component = 0; component < nonFilteredComponents.Count; component++)
                    {
                        // Если совпали названия, смотрим количество и добавляем
                        if (partNumbersArray[counter] == nonFilteredComponents[component].PartNumber)
                        {
                            componentCount += nonFilteredComponents[component].Count;
                        }
                    }
                    componentInitialProperty.PartNumber = partNumbersArray[counter];
                    componentInitialProperty.Count = componentCount;
                    filteredComponents.Add(componentInitialProperty);
                }
                return filteredComponents;
            }
            catch
            {
                var getFilteredComponentsErr = new Exception("Ошибка в функции GetFilteredComponents");
                throw getFilteredComponentsErr;
            }   
        }

        // Функция получения данных для заявки на склад
        public List<PropertyForWarehouseRequest> GetPropertiesForWarehouseRequest(List<ComponentInitialProperties> filteredComponents)
        {
            try
            {
                var propertiesForWarehouseRequest = new List<PropertyForWarehouseRequest>();
                for (int item = 0; item < filteredComponents.Count; item++)
                {
                    var property = new PropertyForWarehouseRequest();
                    var partNumber = filteredComponents[item].PartNumber;
                    property.Count = filteredComponents[item].Count;
                    property.PartNumber = filteredComponents[item].PartNumber;
                    using (DataBaseContext dataBaseConnection = new DataBaseContext())
                    {
                        property.Description1 = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description1).FirstOrDefault();
                        property.Description2 = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description2).FirstOrDefault();
                        property.OrderNumber = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.OrderNumber).FirstOrDefault();
                        property.ManufacturerFullName = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.ManufacturerFullName).FirstOrDefault();
                        int? quantityUnitID = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                        if (quantityUnitID == null)
                        {
                            property.QuantityUnit = string.Empty;
                        }
                        else
                        {
                            property.QuantityUnit = dataBaseConnection.QuantityUnits.Where(
                                o => o.QuantityUnitID == quantityUnitID).Select(
                                o1 => o1.Name).FirstOrDefault();
                        }
                    }
                    propertiesForWarehouseRequest.Add(property);
                }
                return propertiesForWarehouseRequest;
            }
            catch
            {
                var getPropertiesForWarehouseRequestErr = new Exception("Ошибка в функции GetPropertiesForWarehouseRequest");
                throw getPropertiesForWarehouseRequestErr;
            }
        }

        // Функция получения данных для спецификации по проекту
        public List<ComponentPropertiesForSpecification> GetComponentPropertiesForSpecification(string projID)
        {
            try
            {
                var componentPropertiesForSpecification = new List<ComponentPropertiesForSpecification>();
                var projectID = Convert.ToInt32(projID);
                using (DataBaseContext dataBaseConnection = new DataBaseContext())
                {
                    // Ищу изделия по проекту
                    var pArticles = dataBaseConnection.PArticles.Where(
                        o => o.ProjectID == projectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Ищу компоненты по изделиям в проекте
                        var components = dataBaseConnection.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToArray();
                        // Перебираю и записываю данные
                        foreach (Component component in components)
                        {
                            var propertyForSpecification = new ComponentPropertiesForSpecification();
                            var componentPartNumber = component.PartNumber;
                            propertyForSpecification.ArticleID = component.PArticleID;
                            propertyForSpecification.PartNumber = component.PartNumber;
                            propertyForSpecification.Count = component.Count;
                            propertyForSpecification.OrderNumber = dataBaseConnection.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.OrderNumber).FirstOrDefault();
                            propertyForSpecification.Description1 = dataBaseConnection.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description1).FirstOrDefault();
                            propertyForSpecification.Description2 = dataBaseConnection.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description2).FirstOrDefault();
                            propertyForSpecification.ManufacturerFullName = dataBaseConnection.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.ManufacturerFullName).FirstOrDefault();
                            int? quantityUnitID = dataBaseConnection.ComponentCatalogs.Where(
                            o => o.PartNumber == componentPartNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                            if (quantityUnitID == null)
                            {
                                propertyForSpecification.QuantityUnit = string.Empty;
                            }
                            else
                            {
                                propertyForSpecification.QuantityUnit = dataBaseConnection.QuantityUnits.Where(
                                    o => o.QuantityUnitID == quantityUnitID).Select(
                                    o1 => o1.Name).FirstOrDefault();
                            }
                            componentPropertiesForSpecification.Add(propertyForSpecification);
                        }
                    }
                }
                return componentPropertiesForSpecification;
            }
            catch
            {
                var getComponentPropertiesForSpecificationErr = new Exception("Ошибка в функции GetComponentPropertiesForSpecification");
                throw getComponentPropertiesForSpecificationErr;
            }
        }

        // Функция получения описания изделия для спецификации
        public List<ArticlePropertiesForSpecification> GetArticlePropertiesForSpecification(string projID)
        {
            try
            {
                var articlePropertiesForSpecification = new List<ArticlePropertiesForSpecification>();
                var currentProjectID = Convert.ToInt32(projID);
                using (DataBaseContext dataBaseConnection = new DataBaseContext())
                {
                    var pArticles = dataBaseConnection.PArticles.Where(
                        o => o.ProjectID == currentProjectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        var propertyForSpecification = new ArticlePropertiesForSpecification();
                        var pArticleDecription = dataBaseConnection.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Description).FirstOrDefault();
                        var pArticleName = dataBaseConnection.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Name).FirstOrDefault();
                        propertyForSpecification.ArticleName = pArticleName;
                        propertyForSpecification.ArticleDescription = pArticleDecription;
                        propertyForSpecification.ArticleID = pArticle.PArticleID;
                        articlePropertiesForSpecification.Add(propertyForSpecification);
                    }
                }
                return articlePropertiesForSpecification;
            }
            catch
            {
                var getArticlePropertiesForSpecificationErr = new Exception("Ошибка в функции GetArticlePropertiesForSpecification");
                throw getArticlePropertiesForSpecificationErr;
            }   
        }
    }
}
