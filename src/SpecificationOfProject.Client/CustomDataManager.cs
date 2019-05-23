using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecificationOfProject.Client
{
    // Описание свойств компонента
    public class ComponentPropertiesInfo
    {
        public string Property;
        public string Value;
        public bool isShown = true;
    }

    // Описание компонента для заявки на склад
    public class DataForWarehouseRequest
    {
        public string PartNumber;
        public string Description1;
        public string Description2;
        public string OrderNumber;
        public string ManufacturerFullName;
        public string QuantityUnit;
        public int Count;
    }

    // Класс для сбора первоначальной информации о компоненте
    public class ComponentData
    {
        public string PartNumber;
        public int Count;
    }

    // Описание компонентов для спецификации
    public class ComponentDataForSpecification
    {
        public int articleID;
        public string PartNumber;
        public string Description1;
        public string Description2;
        public string OrderNumber;
        public string ManufacturerFullName;
        public string QuantityUnit;
        public int Count;
    }

    // Описание изделий для спецификации
    public class ArticleDataForSpecification
    {
        public int articleID;
        public string articleName;
        public string articleDescription;
    }
}
