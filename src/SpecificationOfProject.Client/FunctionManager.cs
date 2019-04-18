using DataBaseLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecificationOfProject.Client
{
    class FunctionManager
    {
        // Функция получения свойств выбранного компонента
        public List<ComponentPropertiesInfo> GetPropertiesInfos(string componentName)
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
                if (componentPropertiesInfo.Value.Length < minStringLength)
                {
                    componentPropertiesInfo.Value = "";
                }
                componentPropertiesInfos.Add(componentPropertiesInfo);
                
                // Поставщик
                componentPropertiesInfo = new ComponentPropertiesInfo();
                componentPropertiesInfo.Property = "Поставщик";
                componentPropertiesInfo.Value = componentCatalog.SupplierFullName
                    + " (" + componentCatalog.SupplierSmallName + ")";
                if (componentPropertiesInfo.Value.Length < minStringLength)
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

    }
}
