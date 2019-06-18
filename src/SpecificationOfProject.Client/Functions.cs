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
            try
            {
                var componentPropertiesInfos = new List<ComponentProperties>();
                using (DataBaseContext DBCon = new DataBaseContext())
                {
                    var componentCatalog = DBCon.ComponentCatalogs.Where(
                        o => o.PartNumber == componentName).FirstOrDefault();
                    var ComponentProperty = new ComponentProperties();
                    const int minStringManufacturerLength = 4;
                    const int minStringMountingSpaceLength = 6;
                    const int minStringProportionsLength = 5;
                    // Номер изделия
                    ComponentProperty.Property = "Номер изделия";
                    ComponentProperty.Value = componentCatalog.PartNumber;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Тип изделия
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Номер типа";
                    ComponentProperty.Value = componentCatalog.TypeNumber;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Номер для заказа
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Номер для заказа";
                    // Если его нет, то ставим его пустым
                    if (componentCatalog.OrderNumber != "")
                    {
                        ComponentProperty.Value = componentCatalog.OrderNumber;
                    }
                    else
                    {
                        ComponentProperty.Value = "";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Производитель
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Производитель";
                    ComponentProperty.Value = componentCatalog.ManufacturerFullName
                         + " (" + componentCatalog.ManufacturerSmallName + ")";
                    // Проверяем, кроме скобок есть что нибудь или нет
                    if (ComponentProperty.Value.Length < minStringManufacturerLength)
                    {
                        ComponentProperty.Value = "";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Поставщик
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Поставщик";
                    ComponentProperty.Value = componentCatalog.SupplierFullName
                        + " (" + componentCatalog.SupplierSmallName + ")";
                    if (ComponentProperty.Value.Length < minStringManufacturerLength)
                    {
                        ComponentProperty.Value = "";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Обозначение 1
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Обозначение 1";
                    ComponentProperty.Value = componentCatalog.Description1;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Обозначение 2
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Обозначение 2";
                    ComponentProperty.Value = componentCatalog.Description2;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Описание
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Описание";
                    ComponentProperty.Value = componentCatalog.Note;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Клеммы: сечение от
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Клеммы: сечение от";
                    ComponentProperty.Value = componentCatalog.TerminalCrossSectionFrom;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Клеммы: сечение до
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Клеммы: сечение до";
                    ComponentProperty.Value = componentCatalog.TerminalCrossSectionTo;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Ток
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Ток";
                    ComponentProperty.Value = componentCatalog.ElectricalCurrent;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Коммутационная способность
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Коммутационная способность";
                    ComponentProperty.Value = componentCatalog.ElectricalSwitchingCapacity;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Напряжение
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Напряжение";
                    ComponentProperty.Value = componentCatalog.Voltage;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Тип напряжения
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Тип напряжения";
                    ComponentProperty.Value = componentCatalog.VoltageType;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Высота
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Высота";
                    ComponentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Height, 3) + " мм");
                    if (ComponentProperty.Value.Length < minStringProportionsLength)
                    {
                        ComponentProperty.Value = "0";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Ширина
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Ширина";
                    ComponentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Width, 3) + " мм");
                    if (ComponentProperty.Value.Length < minStringProportionsLength)
                    {
                        ComponentProperty.Value = "0";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Глубина
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Глубина";
                    ComponentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Depth, 3) + " мм");
                    if (ComponentProperty.Value.Length < minStringProportionsLength)
                    {
                        ComponentProperty.Value = "0";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Вес
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Вес";
                    ComponentProperty.Value = Convert.ToString(Math.Round(componentCatalog.Weight, 3) + " кг");
                    if (ComponentProperty.Value.Length < minStringProportionsLength)
                    {
                        ComponentProperty.Value = "0";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Монтажная поверхность
                    ComponentProperty = new ComponentProperties();
                    string mountingSite = DBCon.MountingSites.Where(
                        o => o.InternalValue == componentCatalog.MountingSiteID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    ComponentProperty.Property = "Монтажная поверхность";
                    ComponentProperty.Value = mountingSite;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Занимаемое пространство
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Занимаемое пространство";
                    ComponentProperty.Value = Convert.ToString(Math.Round(componentCatalog.MountingSpace, 3) + " мм2");
                    if (ComponentProperty.Value.Length < minStringMountingSpaceLength)
                    {
                        ComponentProperty.Value = "0";
                    }
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Область - раздел - сфера
                    ComponentProperty = new ComponentProperties();
                    // Спличу строку
                    var PartGroup = componentCatalog.PartGroup.Split('-');
                    var areaID = Convert.ToInt32(PartGroup[0]);
                    var sectionID = Convert.ToInt32(PartGroup[1]);
                    var sphereID = Convert.ToInt32(PartGroup[2]);
                    var area = DBCon.Areas.Where(o => o.InternalValue == areaID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    var section = DBCon.Sections.Where(o => o.InternalValue == sectionID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    var sphere = DBCon.Areas.Where(o => o.InternalValue == sphereID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    ComponentProperty.Property = "Область";
                    ComponentProperty.Value = area;
                    componentPropertiesInfos.Add(ComponentProperty);
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Раздел";
                    ComponentProperty.Value = section;
                    componentPropertiesInfos.Add(ComponentProperty);
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Сфера";
                    ComponentProperty.Value = sphere;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Единица измерения
                    ComponentProperty = new ComponentProperties();
                    var quantityUnit = DBCon.QuantityUnits.Where(
                        o => o.QuantityUnitID == componentCatalog.QuantityUnitID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    ComponentProperty.Property = "Единица измерения";
                    ComponentProperty.Value = quantityUnit;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Количество упаковки
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Количество в упаковке";
                    ComponentProperty.Value = Convert.ToString(componentCatalog.PackagingQuantity);
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Последнее изменение
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Последнее изменение";
                    ComponentProperty.Value = componentCatalog.LastChange;
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Цена продажи, Валюта 1
                    // Выводится ПОСЛЕДНИЙ актуальный прайс
                    // Нашел количество прайсов в таблице
                    // Выбрал последний (сортировка по возрастанию даты)
                    var prices = DBCon.PriceLists.Where(
                        o => o.PartNumber == componentCatalog.PartNumber).OrderBy(
                        o1 => o1.LastChange).ToList();
                    var salesPrice1 = prices.Select(
                        o => o.SalesPrice1).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Цена продажи, Валюта 1";
                    ComponentProperty.Value = Convert.ToString(Math.Round(salesPrice1, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Цена продажи, Валюта 2
                    var salesPrice2 = prices.Select(
                        o => o.SalesPrice2).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Цена продажи, Валюта 2";
                    ComponentProperty.Value = Convert.ToString(Math.Round(salesPrice2, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Закупочная цена, Валюта 1
                    var purchasePrice1 = prices.Select(
                        o => o.PurchasePrice1).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Закупочная цена, Валюта 1";
                    ComponentProperty.Value = Convert.ToString(Math.Round(purchasePrice1, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Закупочная цена, Валюта 2
                    var purchasePrice2 = prices.Select(
                        o => o.PurchasePrice2).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Закупочная цена, Валюта 2";
                    ComponentProperty.Value = Convert.ToString(Math.Round(purchasePrice2, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Цена упаковки, Валюта 1
                    var packagingPrice1 = prices.Select(
                        o => o.PackagingPrice1).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Цена упаковки, Валюта 1";
                    ComponentProperty.Value = Convert.ToString(Math.Round(packagingPrice1, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Цена упаковки, Валюта 2
                    var packagingPrice2 = prices.Select(
                        o => o.PackagingPrice2).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Цена упаковки, Валюта 2";
                    ComponentProperty.Value = Convert.ToString(Math.Round(packagingPrice2, 3));
                    componentPropertiesInfos.Add(ComponentProperty);

                    // Единица цены
                    var priceUnit = prices.Select(
                        o => o.PriceUnit).LastOrDefault();
                    ComponentProperty = new ComponentProperties();
                    ComponentProperty.Property = "Единица цены";
                    ComponentProperty.Value = Convert.ToString(priceUnit);
                    componentPropertiesInfos.Add(ComponentProperty);
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
                double widthCellA = 4;
                double widthCellB = 12.8;
                double widthCellC = 23;
                double widthCellD = 15.57;
                double widthCellE = 15.7;
                double widthCellF = 11.5;
                double widthCellG = 5.29;
                double widthCellH = 5.57;
                double widthCellI = 24;
                double widthCellJ = 17;

                // Использую библиотеку EPPlus
                using (var spec = new ExcelPackage())
                {
                    // Добавляю лист в spec
                    var workSheet = spec.Workbook.Worksheets.Add("Спецификация");

                    // Глобальные настройки документа                
                    // Настраиваю ширину колонок и переносы
                    workSheet.Column(1).Width = widthCellA;
                    workSheet.Column(2).Width = widthCellB;
                    workSheet.Column(3).Width = widthCellC;
                    workSheet.Column(4).Width = widthCellD;
                    workSheet.Column(5).Width = widthCellE;
                    workSheet.Column(6).Width = widthCellF;
                    workSheet.Column(7).Width = widthCellG;
                    workSheet.Column(8).Width = widthCellH;
                    workSheet.Column(9).Width = widthCellI;
                    workSheet.Column(10).Width = widthCellJ;

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

                    int startedRow = 8; // Стартовая строка для записи
                    int endRow = 9; // Конечная строка (на текущий момент)
                    int counterOfComponents = 0; // Количество компонентов в общем по спецификации
                    int currentRow = startedRow + 1; // Текущая строка для записи

                    // Устанавливаю границу для шапки таблицы (десять столбцов)
                    for (int column = 1; column <= 10; column++)
                    {
                        workSheet.Cells[startedRow, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
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
                            "Изделие: " + "\"" + articleProperty.articleDescription + "\" "+ "(" + articleProperty.articleName + ")";
                        workSheet.Cells[currentRow, 1, currentRow, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        currentRow += 1; // Записал строку, увеличил счетчик на единицу
                        foreach (ComponentPropertiesForSpecification componentPropery in componentsPropertiesForSpecification)
                        {
                            // Записываем только, если articleID изделия = articleID компонента
                            if (articleProperty.articleID == componentPropery.articleID)
                            {
                                // Настройка строки для записи и столбцов
                                for (int column = 1; column <= 10; column++)
                                {
                                    workSheet.Cells[currentRow, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[currentRow, column].Style.Font.Name = "Times New Roman";
                                    workSheet.Cells[currentRow, column].Style.Font.Size = 10;
                                    workSheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    workSheet.Cells[currentRow, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                // Запись по столбцам (от 1 до 10, A > J). Не все столбцы
                                workSheet.Cells[currentRow, 1].Value = counterOfComponents + 1;
                                workSheet.Cells[currentRow, 3].Value = componentPropery.Description1 +
                                    " " + componentPropery.PartNumber;
                                workSheet.Cells[currentRow, 4].Value = componentPropery.OrderNumber;
                                workSheet.Cells[currentRow, 5].Value = componentPropery.ManufacturerFullName;
                                workSheet.Cells[currentRow, 7].Value = componentPropery.QuantityUnit;
                                workSheet.Cells[currentRow, 8].Value = componentPropery.Count;
                                workSheet.Cells[currentRow, 9].Value = componentPropery.Description2;

                                currentRow += 1;
                                counterOfComponents++;
                            }      
                        }
                        endRow = currentRow;
                    }

                    // Указываем начальника службы и ФИО
                    workSheet.Cells[endRow + 1, 2, endRow + 1, 3].Merge = true;
                    workSheet.Cells[endRow + 1, 2].Value = "<укажите начальника службы>";
                    workSheet.Cells[endRow + 1, 9].Value = "И.О. Фамилия";
                    // Настраиваем отображение
                    workSheet.Cells[endRow + 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[endRow + 1, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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
                double widthCellA = 4;
                double widthCellB = 12.8;
                double widthCellC = 23;
                double widthCellD = 15.57;
                double widthCellE = 15.7;
                double widthCellF = 11.5;
                double widthCellG = 5.29;
                double widthCellH = 5.57;
                double widthCellI = 24;
                double widthCellJ = 17;

                // Использую библиотеку EPPlus
                using (var spec = new ExcelPackage())
                {
                    // Добавляю лист в spec
                    var workSheet = spec.Workbook.Worksheets.Add("Заявка");

                    // Глобальные настройки документа                
                    // Настраиваю ширину колонок и переносы
                    workSheet.Column(1).Width = widthCellA;
                    workSheet.Column(2).Width = widthCellB;
                    workSheet.Column(3).Width = widthCellC;
                    workSheet.Column(4).Width = widthCellD;
                    workSheet.Column(5).Width = widthCellE;
                    workSheet.Column(6).Width = widthCellF;
                    workSheet.Column(7).Width = widthCellG;
                    workSheet.Column(8).Width = widthCellH;
                    workSheet.Column(9).Width = widthCellI;
                    workSheet.Column(10).Width = widthCellJ;

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
                    const int startedRow = 10;
                    int endRow = 11; // Конец строк с компонентами
                    for (int item = 0; item < dataForWarehouseRequests.Count; item++)
                    {
                        // Настройка строки для записи и столбцов
                        for (int column = 1; column <= 10; column++)
                        {
                            workSheet.Cells[startedRow + item + 1, column].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[startedRow + item + 1, column].Style.Font.Name = "Times New Roman";
                            workSheet.Cells[startedRow + item + 1, column].Style.Font.Size = 10;
                            workSheet.Cells[startedRow + item + 1, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[startedRow + item + 1, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        // Запись по столбцам (от 1 до 10, A > J)
                        workSheet.Cells[startedRow + item + 1, 1].Value = item + 1;
                        workSheet.Cells[startedRow + item + 1, 3].Value = dataForWarehouseRequests[item].Description1 +
                            " " + dataForWarehouseRequests[item].PartNumber;
                        workSheet.Cells[startedRow + item + 1, 4].Value = dataForWarehouseRequests[item].OrderNumber;
                        workSheet.Cells[startedRow + item + 1, 5].Value = dataForWarehouseRequests[item].ManufacturerFullName;
                        workSheet.Cells[startedRow + item + 1, 7].Value = dataForWarehouseRequests[item].QuantityUnit;
                        workSheet.Cells[startedRow + item + 1, 8].Value = dataForWarehouseRequests[item].Count;
                        workSheet.Cells[startedRow + item + 1, 9].Value = dataForWarehouseRequests[item].Description2;
                        endRow = startedRow + item + 1;
                    }

                    // Указываем начальника службы и ФИО
                    workSheet.Cells[endRow + 2, 2, endRow + 2, 3].Merge = true;
                    workSheet.Cells[endRow + 2, 2].Value = "<укажите начальника службы>";
                    workSheet.Cells[endRow + 2, 9].Value = "И.О. Фамилия";
                    // Настраиваем отображение
                    workSheet.Cells[endRow + 2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[endRow + 2, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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
                using (DataBaseContext DBCon = new DataBaseContext())
                {
                    // Получил изделия по проекту
                    var pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == curProjectID).ToList();
                    // Перебираю проекты и заполняю компоненты
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Получил список всех компонентов в изделии
                        var components = DBCon.Components.Where(
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
                    using (DataBaseContext DBCon = new DataBaseContext())
                    {
                        property.Description1 = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description1).FirstOrDefault();
                        property.Description2 = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description2).FirstOrDefault();
                        property.OrderNumber = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.OrderNumber).FirstOrDefault();
                        property.ManufacturerFullName = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.ManufacturerFullName).FirstOrDefault();
                        int? quantityUnitID = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                        if (quantityUnitID == null)
                        {
                            property.QuantityUnit = "";
                        }
                        else
                        {
                            property.QuantityUnit = DBCon.QuantityUnits.Where(
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
                using (DataBaseContext DBCon = new DataBaseContext())
                {
                    // Ищу изделия по проекту
                    var pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == projectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Ищу компоненты по изделиям в проекте
                        var components = DBCon.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToArray();
                        // Перебираю и записываю данные
                        foreach (Component component in components)
                        {
                            var propertyForSpecification = new ComponentPropertiesForSpecification();
                            var componentPartNumber = component.PartNumber;
                            propertyForSpecification.articleID = component.PArticleID;
                            propertyForSpecification.PartNumber = component.PartNumber;
                            propertyForSpecification.Count = component.Count;
                            propertyForSpecification.OrderNumber = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.OrderNumber).FirstOrDefault();
                            propertyForSpecification.Description1 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description1).FirstOrDefault();
                            propertyForSpecification.Description2 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description2).FirstOrDefault();
                            propertyForSpecification.ManufacturerFullName = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.ManufacturerFullName).FirstOrDefault();
                            int? quantityUnitID = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == componentPartNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                            if (quantityUnitID == null)
                            {
                                propertyForSpecification.QuantityUnit = "";
                            }
                            else
                            {
                                propertyForSpecification.QuantityUnit = DBCon.QuantityUnits.Where(
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
                using (DataBaseContext DBCon = new DataBaseContext())
                {
                    var pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == currentProjectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        var propertyForSpecification = new ArticlePropertiesForSpecification();
                        var pArticleDecription = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Description).FirstOrDefault();
                        var pArticleName = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Name).FirstOrDefault();
                        propertyForSpecification.articleName = pArticleName;
                        propertyForSpecification.articleDescription = pArticleDecription;
                        propertyForSpecification.articleID = pArticle.PArticleID;
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
