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
    class FunctionManager
    {
        // Функция получения свойств выбранного компонента
        public List<ComponentPropertiesInfo> GetPropertiesInfos(string componentName)
        {
            try
            {
                List<ComponentPropertiesInfo> componentPropertiesInfos = new List<ComponentPropertiesInfo>();
                using (DBContext DBCon = new DBContext())
                {
                    ComponentCatalog componentCatalog = DBCon.ComponentCatalogs.Where(
                        o => o.PartNumber == componentName).FirstOrDefault();
                    ComponentPropertiesInfo componentPropertiesInfo = new ComponentPropertiesInfo();
                    const int minStringManufacturerLength = 4;
                    const int minStringMountingSpaceLength = 6;
                    const int minStringProportionsLength = 5;
                    // Номер изделия
                    componentPropertiesInfo.Property = "Номер изделия";
                    componentPropertiesInfo.Value = componentCatalog.PartNumber;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Тип изделия
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Номер типа";
                    componentPropertiesInfo.Value = componentCatalog.TypeNumber;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Номер для заказа
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Номер для заказа";
                    // Если его нет, то ставим его пустым
                    if (componentCatalog.OrderNumber != "")
                    {
                        componentPropertiesInfo.Value = componentCatalog.OrderNumber;
                    }
                    else
                    {
                        componentPropertiesInfo.Value = "";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Производитель
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Производитель";
                    componentPropertiesInfo.Value = componentCatalog.ManufacturerFullName
                         + " (" + componentCatalog.ManufacturerSmallName + ")";
                    // Проверяем, кроме скобок есть что нибудь или нет
                    if (componentPropertiesInfo.Value.Length < minStringManufacturerLength)
                    {
                        componentPropertiesInfo.Value = "";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Поставщик
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Поставщик";
                    componentPropertiesInfo.Value = componentCatalog.SupplierFullName
                        + " (" + componentCatalog.SupplierSmallName + ")";
                    if (componentPropertiesInfo.Value.Length < minStringManufacturerLength)
                    {
                        componentPropertiesInfo.Value = "";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Обозначение 1
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Обозначение 1";
                    componentPropertiesInfo.Value = componentCatalog.Description1;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Обозначение 2
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Обозначение 2";
                    componentPropertiesInfo.Value = componentCatalog.Description2;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Описание
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Описание";
                    componentPropertiesInfo.Value = componentCatalog.Note;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Клеммы: сечение от
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Клеммы: сечение от";
                    componentPropertiesInfo.Value = componentCatalog.TerminalCrossSectionFrom;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Клеммы: сечение до
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Клеммы: сечение до";
                    componentPropertiesInfo.Value = componentCatalog.TerminalCrossSectionTo;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Ток
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Ток";
                    componentPropertiesInfo.Value = componentCatalog.ElectricalCurrent;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Коммутационная способность
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Коммутационная способность";
                    componentPropertiesInfo.Value = componentCatalog.ElectricalSwitchingCapacity;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Напряжение
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Напряжение";
                    componentPropertiesInfo.Value = componentCatalog.Voltage;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Тип напряжения
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Тип напряжения";
                    componentPropertiesInfo.Value = componentCatalog.VoltageType;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Высота
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Высота";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(componentCatalog.Height, 3) + " мм");
                    if (componentPropertiesInfo.Value.Length < minStringProportionsLength)
                    {
                        componentPropertiesInfo.Value = "0";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Ширина
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Ширина";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(componentCatalog.Width, 3) + " мм");
                    if (componentPropertiesInfo.Value.Length < minStringProportionsLength)
                    {
                        componentPropertiesInfo.Value = "0";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Глубина
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Глубина";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(componentCatalog.Depth, 3) + " мм");
                    if (componentPropertiesInfo.Value.Length < minStringProportionsLength)
                    {
                        componentPropertiesInfo.Value = "0";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Вес
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Вес";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(componentCatalog.Weight, 3) + " кг");
                    if (componentPropertiesInfo.Value.Length < minStringProportionsLength)
                    {
                        componentPropertiesInfo.Value = "0";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Монтажная поверхность
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    string mountingSite = DBCon.MountingSites.Where(
                        o => o.InternalValue == componentCatalog.MountingSiteID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentPropertiesInfo.Property = "Монтажная поверхность";
                    componentPropertiesInfo.Value = mountingSite;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Занимаемое пространство
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Занимаемое пространство";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(componentCatalog.MountingSpace, 3) + " мм2");
                    if (componentPropertiesInfo.Value.Length < minStringMountingSpaceLength)
                    {
                        componentPropertiesInfo.Value = "0";
                    }
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Область - раздел - сфера
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    // Спличу строку
                    string[] PartGroup = componentCatalog.PartGroup.Split('-');
                    int areaID = Convert.ToInt32(PartGroup[0]);
                    int sectionID = Convert.ToInt32(PartGroup[1]);
                    int sphereID = Convert.ToInt32(PartGroup[2]);
                    string area = DBCon.Areas.Where(o => o.InternalValue == areaID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    string section = DBCon.Sections.Where(o => o.InternalValue == sectionID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    string sphere = DBCon.Areas.Where(o => o.InternalValue == sphereID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentPropertiesInfo.Property = "Область";
                    componentPropertiesInfo.Value = area;
                    componentPropertiesInfos.Add(componentPropertiesInfo);
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Раздел";
                    componentPropertiesInfo.Value = section;
                    componentPropertiesInfos.Add(componentPropertiesInfo);
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Сфера";
                    componentPropertiesInfo.Value = sphere;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Единица измерения
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    string quantityUnit = DBCon.QuantityUnits.Where(
                        o => o.QuantityUnitID == componentCatalog.QuantityUnitID).Select(
                        o1 => o1.Name).FirstOrDefault();
                    componentPropertiesInfo.Property = "Единица измерения";
                    componentPropertiesInfo.Value = quantityUnit;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Количество упаковки
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Количество в упаковке";
                    componentPropertiesInfo.Value = Convert.ToString(componentCatalog.PackagingQuantity);
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Последнее изменение
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Последнее изменение";
                    componentPropertiesInfo.Value = componentCatalog.LastChange;
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Цена продажи, Валюта 1
                    // Выводится ПОСЛЕДНИЙ актуальный прайс
                    // Нашел количество прайсов в таблице
                    // Выбрал последний (сортировка по возрастанию даты)
                    List<PriceList> prices = DBCon.PriceLists.Where(
                        o => o.PartNumber == componentCatalog.PartNumber).OrderBy(
                        o1 => o1.LastChange).ToList();
                    double salesPrice1 = prices.Select(
                        o => o.SalesPrice1).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Цена продажи, Валюта 1";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(salesPrice1, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Цена продажи, Валюта 2
                    double salesPrice2 = prices.Select(
                        o => o.SalesPrice2).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Цена продажи, Валюта 2";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(salesPrice2, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Закупочная цена, Валюта 1
                    double purchasePrice1 = prices.Select(
                        o => o.PurchasePrice1).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Закупочная цена, Валюта 1";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(purchasePrice1, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Закупочная цена, Валюта 2
                    double purchasePrice2 = prices.Select(
                        o => o.PurchasePrice2).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Закупочная цена, Валюта 2";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(purchasePrice2, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Цена упаковки, Валюта 1
                    double packagingPrice1 = prices.Select(
                        o => o.PackagingPrice1).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Цена упаковки, Валюта 1";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(packagingPrice1, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Цена упаковки, Валюта 2
                    double packagingPrice2 = prices.Select(
                        o => o.PackagingPrice2).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Цена упаковки, Валюта 2";
                    componentPropertiesInfo.Value = Convert.ToString(Math.Round(packagingPrice2, 3));
                    componentPropertiesInfos.Add(componentPropertiesInfo);

                    // Единица цены
                    int priceUnit = prices.Select(
                        o => o.PriceUnit).LastOrDefault();
                    componentPropertiesInfo = new ComponentPropertiesInfo();
                    componentPropertiesInfo.Property = "Единица цены";
                    componentPropertiesInfo.Value = Convert.ToString(priceUnit);
                    componentPropertiesInfos.Add(componentPropertiesInfo);
                }
                return componentPropertiesInfos;
            }
            catch
            {
                Exception getPropertiesInfosErr = new Exception("Ошибка в функции GetPropertiesInfos");
                throw getPropertiesInfosErr;
            }  
        }

        // Функция формирования спецификации по проекту
        // И загрузки в перечень документов
        public void CreateProjectSpecification(string projID, string savePath, string projName)
        {
            try
            {
                //Ширина колонок от А до J
                const double widthCellA = 4;
                const double widthCellB = 12.8;
                const double widthCellC = 23;
                const double widthCellD = 15.57;
                const double widthCellE = 15.7;
                const double widthCellF = 11.5;
                const double widthCellG = 5.29;
                const double widthCellH = 5.57;
                const double widthCellI = 24;
                const double widthCellJ = 17;

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
                    const int firstColumn = 1; // Первая колонка
                    const int lastColumn = 10; // Последняя колонка

                    // Устанавливаю границу для шапки таблицы (десять столбцов)
                    for (int i = 1; i <= 10; i++)
                    {
                        workSheet.Cells[startedRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Получаю данные компонентов для записи
                    List<ComponentDataForSpecification> componentsDataForSpecification = GetComponentDataForSpecification(projID);

                    // Получаю данные по изделиям
                    List<ArticleDataForSpecification> articleDataForSpecification = GetArticleDataForSpecification(projID);
                    // Под одинаковыми индексами у каждого списка находится информация по одному изделию
                    
                    // Записываю данные в документ
                    // Перебираю изделия и компоненты внутри друг друга т.к они сопоставимы
                    foreach (ArticleDataForSpecification dataForSpecification in articleDataForSpecification)
                    {
                        workSheet.Cells[currentRow, firstColumn, currentRow, lastColumn].Merge = true;
                        workSheet.Cells[currentRow, firstColumn].Value =
                            "Изделие: " + "\"" + dataForSpecification.articleDescription + "\" "+ "(" + dataForSpecification.articleName + ")";
                        workSheet.Cells[currentRow, firstColumn, currentRow, lastColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        currentRow += 1; // Записал строку, увеличил счетчик на единицу
                        foreach (ComponentDataForSpecification componentData in componentsDataForSpecification)
                        {
                            // Записываем только, если articleID изделия = articleID компонента
                            if (dataForSpecification.articleID == componentData.articleID)
                            {
                                // Настройка строки для записи и столбцов
                                for (int j = 1; j <= 10; j++)
                                {
                                    workSheet.Cells[currentRow, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    workSheet.Cells[currentRow, j].Style.Font.Name = "Times New Roman";
                                    workSheet.Cells[currentRow, j].Style.Font.Size = 10;
                                    workSheet.Cells[currentRow, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    workSheet.Cells[currentRow, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                // Запись по столбцам (от 1 до 10, A > J). Не все столбцы
                                workSheet.Cells[currentRow, 1].Value = counterOfComponents + 1;
                                workSheet.Cells[currentRow, 3].Value = componentData.Description1 +
                                    " " + componentData.PartNumber;
                                workSheet.Cells[currentRow, 4].Value = componentData.OrderNumber;
                                workSheet.Cells[currentRow, 5].Value = componentData.ManufacturerFullName;
                                workSheet.Cells[currentRow, 7].Value = componentData.QuantityUnit;
                                workSheet.Cells[currentRow, 8].Value = componentData.Count;
                                workSheet.Cells[currentRow, 9].Value = componentData.Description2;

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
                Exception createProjectSpecificationErr = new Exception("Ошибка в функции CreateProjectSpecification");
                throw createProjectSpecificationErr;
            }
        }

        // Функция формирования заявки на склад по проекту
        // И загрузки в перечень документов
        public void CreateProjectWarehouseRequest(string projID, string savePath)
        {
            try
            {
                //Ширина колонок от А до J
                const double widthCellA = 4;
                const double widthCellB = 12.8;
                const double widthCellC = 23;
                const double widthCellD = 15.57;
                const double widthCellE = 15.7;
                const double widthCellF = 11.5;
                const double widthCellG = 5.29;
                const double widthCellH = 5.57;
                const double widthCellI = 24;
                const double widthCellJ = 17;

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
                    for (int i = 1; i <= 10; i++)
                    {
                        workSheet.Cells[10, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    // Получаю все компоненты
                    List<ComponentData> nonFilteredComponents = GetNonFilteredComponents(projID);

                    // Фильтрую компоненты
                    List<ComponentData> filteredComponents = GetFilteredComponents(nonFilteredComponents);

                    // Получаю полные данные для заявки
                    List<DataForWarehouseRequest> dataForWarehouseRequests = GetDataForWarehouseRequest(filteredComponents);

                    // Заполняю
                    const int startedRow = 10;
                    int endRow = 11; // Конец строк с компонентами
                    for (int i = 0; i < dataForWarehouseRequests.Count; i++)
                    {
                        // Настройка строки для записи и столбцов
                        for (int j = 1; j <= 10; j++)
                        {
                            workSheet.Cells[startedRow + i + 1, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            workSheet.Cells[startedRow + i + 1, j].Style.Font.Name = "Times New Roman";
                            workSheet.Cells[startedRow + i + 1, j].Style.Font.Size = 10;
                            workSheet.Cells[startedRow + i + 1, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            workSheet.Cells[startedRow + i + 1, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        // Запись по столбцам (от 1 до 10, A > J)
                        workSheet.Cells[startedRow + i + 1, 1].Value = i + 1;
                        workSheet.Cells[startedRow + i + 1, 3].Value = dataForWarehouseRequests[i].Description1 +
                            " " + dataForWarehouseRequests[i].PartNumber;
                        workSheet.Cells[startedRow + i + 1, 4].Value = dataForWarehouseRequests[i].OrderNumber;
                        workSheet.Cells[startedRow + i + 1, 5].Value = dataForWarehouseRequests[i].ManufacturerFullName;
                        workSheet.Cells[startedRow + i + 1, 7].Value = dataForWarehouseRequests[i].QuantityUnit;
                        workSheet.Cells[startedRow + i + 1, 8].Value = dataForWarehouseRequests[i].Count;
                        workSheet.Cells[startedRow + i + 1, 9].Value = dataForWarehouseRequests[i].Description2;
                        endRow = startedRow + i + 1;
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
                Exception createProjectWarehouseRequestErr = new Exception("Ошибка в функции CreateProjectWarehouseRequest");
                throw createProjectWarehouseRequestErr;
            }      
        }

        // Функция получения списка компонентов для заявки на склад
        public List<ComponentData> GetNonFilteredComponents(string projID)
        {
            try
            {
                // Так как ID записан строкой в списке, то конвертируем
                int curProjectID = Convert.ToInt32(projID);
                List<ComponentData> nonFilteredComponents = new List<ComponentData>();
                using (DBContext DBCon = new DBContext())
                {
                    // Получил изделия по проекту
                    List<PArticle> pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == curProjectID).ToList();
                    // Перебираю проекты и заполняю компоненты
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Получил список всех компонентов в изделии
                        List<Component> components = DBCon.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToList();
                        // Записываю базовую информацию по компонентам
                        foreach (Component component in components)
                        {
                            ComponentData componentData = new ComponentData();
                            componentData.PartNumber = component.PartNumber;
                            componentData.Count = component.Count;
                            nonFilteredComponents.Add(componentData);
                        }
                    }
                }
                return nonFilteredComponents;
            }
            catch
            {
                Exception getNonFilteredComponentsErr = new Exception("Ошибка в функции GetNonFilteredComponents");
                throw getNonFilteredComponentsErr;
            }   
        }
        
        // Функция фильтрации списка компонентов для заявки на склад
        public List<ComponentData> GetFilteredComponents(List<ComponentData> nonFilteredComponents)
        {
            try
            {
                List<ComponentData> filteredComponents = new List<ComponentData>();
                // Получил отфильтрованные компоненты, но не посчитал их количество
                List<string> nonFilteredPartNumbers = new List<string>();
                foreach (ComponentData componentData in nonFilteredComponents)
                {
                    nonFilteredPartNumbers.Add(componentData.PartNumber);
                }
                string[] partNumbersArray = nonFilteredPartNumbers.Distinct().ToArray();
                // Считаю количество компонентов и формирую список
                for (int i = 0; i < partNumbersArray.Length; i++)
                {
                    int componentCount = 0; // Количество компонентов
                    ComponentData componentData = new ComponentData();
                    //Прохожу по всему списку неотфильтрованному и считаю количество компонентов
                    for (int j = 0; j < nonFilteredComponents.Count; j++)
                    {
                        // Если совпали названия, смотрим количество и добавляем
                        if (partNumbersArray[i] == nonFilteredComponents[j].PartNumber)
                        {
                            componentCount += nonFilteredComponents[j].Count;
                        }
                    }
                    componentData.PartNumber = partNumbersArray[i];
                    componentData.Count = componentCount;
                    filteredComponents.Add(componentData);
                }
                return filteredComponents;
            }
            catch
            {
                Exception getFilteredComponentsErr = new Exception("Ошибка в функции GetFilteredComponents");
                throw getFilteredComponentsErr;
            }   
        }

        // Функция получения данных для заявки на склад
        public List<DataForWarehouseRequest> GetDataForWarehouseRequest(List<ComponentData> filteredComponents)
        {
            try
            {
                List<DataForWarehouseRequest> dataForWarehouseRequests = new List<DataForWarehouseRequest>();
                for (int i = 0; i < filteredComponents.Count; i++)
                {
                    DataForWarehouseRequest dataForWarehouseRequest = new DataForWarehouseRequest();
                    string partNumber = filteredComponents[i].PartNumber;
                    dataForWarehouseRequest.Count = filteredComponents[i].Count;
                    dataForWarehouseRequest.PartNumber = filteredComponents[i].PartNumber;
                    using (DBContext DBCon = new DBContext())
                    {
                        dataForWarehouseRequest.Description1 = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description1).FirstOrDefault();
                        dataForWarehouseRequest.Description2 = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.Description2).FirstOrDefault();
                        dataForWarehouseRequest.OrderNumber = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.OrderNumber).FirstOrDefault();
                        dataForWarehouseRequest.ManufacturerFullName = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.ManufacturerFullName).FirstOrDefault();
                        int? quantityUnitID = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == partNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                        if (quantityUnitID == null)
                        {
                            dataForWarehouseRequest.QuantityUnit = "";
                        }
                        else
                        {
                            dataForWarehouseRequest.QuantityUnit = DBCon.QuantityUnits.Where(
                                o => o.QuantityUnitID == quantityUnitID).Select(
                                o1 => o1.Name).FirstOrDefault();
                        }
                    }
                    dataForWarehouseRequests.Add(dataForWarehouseRequest);
                }
                return dataForWarehouseRequests;
            }
            catch
            {
                Exception getDataForWarehouseRequestErr = new Exception("Ошибка в функции GetDataForWarehouseRequest");
                throw getDataForWarehouseRequestErr;
            }
        }

        // Функция получения данных для спецификации по проекту
        public List<ComponentDataForSpecification> GetComponentDataForSpecification(string projID)
        {
            try
            {
                List<ComponentDataForSpecification> dataForSpecification = new List<ComponentDataForSpecification>();
                int projectID = Convert.ToInt32(projID);
                using (DBContext DBCon = new DBContext())
                {
                    // Ищу изделия по проекту
                    PArticle[] pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == projectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        // Ищу компоненты по изделиям в проекте
                        Component[] components = DBCon.Components.Where(
                            o => o.PArticleID == pArticle.PArticleID).ToArray();
                        // Перебираю и записываю данные
                        foreach (Component component in components)
                        {
                            ComponentDataForSpecification componentData = new ComponentDataForSpecification();
                            string componentPartNumber = component.PartNumber;
                            componentData.articleID = component.PArticleID;
                            componentData.PartNumber = component.PartNumber;
                            componentData.Count = component.Count;
                            componentData.OrderNumber = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.OrderNumber).FirstOrDefault();
                            componentData.Description1 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description1).FirstOrDefault();
                            componentData.Description2 = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.Description2).FirstOrDefault();
                            componentData.ManufacturerFullName = DBCon.ComponentCatalogs.Where(
                                o => o.PartNumber == componentPartNumber).Select(
                                o1 => o1.ManufacturerFullName).FirstOrDefault();
                            int? quantityUnitID = DBCon.ComponentCatalogs.Where(
                            o => o.PartNumber == componentPartNumber).Select(
                            o1 => o1.QuantityUnitID).FirstOrDefault();
                            if (quantityUnitID == null)
                            {
                                componentData.QuantityUnit = "";
                            }
                            else
                            {
                                componentData.QuantityUnit = DBCon.QuantityUnits.Where(
                                    o => o.QuantityUnitID == quantityUnitID).Select(
                                    o1 => o1.Name).FirstOrDefault();
                            }
                            dataForSpecification.Add(componentData);
                        }
                    }
                }
                return dataForSpecification;
            }
            catch
            {
                Exception getComponentDataForSpecificationErr = new Exception("Ошибка в функции GetComponentDataForSpecification");
                throw getComponentDataForSpecificationErr;
            }
        }

        // Функция получения описания изделия для спецификации
        public List<ArticleDataForSpecification> GetArticleDataForSpecification(string projID)
        {
            try
            {
                List<ArticleDataForSpecification> articleDatasForSpecification = new List<ArticleDataForSpecification>();
                int currentProjectID = Convert.ToInt32(projID);
                using (DBContext DBCon = new DBContext())
                {
                    PArticle[] pArticles = DBCon.PArticles.Where(
                        o => o.ProjectID == currentProjectID).ToArray();
                    foreach (PArticle pArticle in pArticles)
                    {
                        ArticleDataForSpecification dataForSpec = new ArticleDataForSpecification();
                        string pArticleDecription = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Description).FirstOrDefault();
                        string pArticleName = DBCon.LocationDescriptions.Where(
                            o => o.LocationDescriptionID == pArticle.LocationDesriptionID).Select(
                            o1 => o1.Name).FirstOrDefault();
                        dataForSpec.articleName = pArticleName;
                        dataForSpec.articleDescription = pArticleDecription;
                        dataForSpec.articleID = pArticle.PArticleID;
                        articleDatasForSpecification.Add(dataForSpec);
                    }
                }
                return articleDatasForSpecification;
            }
            catch
            {
                Exception getArticleDataForSpecificationErr = new Exception("Ошибка в функции GetArticleDataForSpecification");
                throw getArticleDataForSpecificationErr;
            }   
        }
    }
}
