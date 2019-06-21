using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecificationOfProject.Client
{
    // Описание свойств компонента
    public class ComponentProperties
    {
        public string Property;
        public string Value;
        public bool isShown = true;
    }

    // Описание компонента для заявки на склад
    public class PropertyForWarehouseRequest
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
    public class ComponentInitialProperties
    {
        public string PartNumber;
        public int Count;
    }

    // Описание компонентов для спецификации
    public class ComponentPropertiesForSpecification
    {
        public int ArticleID;
        public string PartNumber;
        public string Description1;
        public string Description2;
        public string OrderNumber;
        public string ManufacturerFullName;
        public string QuantityUnit;
        public int Count;
    }

    // Описание изделий для спецификации
    public class ArticlePropertiesForSpecification
    {
        public int ArticleID;
        public string ArticleName;
        public string ArticleDescription;
    }
}
