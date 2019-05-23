using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DataBaseLibrary
{
    public class Proj // Проект
    {   
        [Required][Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ProjID { get; set; }
        [Required]
        public string Name { get; set; }
        public DateTime? Date { get; set; }
        public DateTime? Time { get; set; }
        public string Executor { get; set; } // Исполнитель

        // Ссылка на изделие
        public List<PArticle> Articles { get; set; }

        // Ссылка на документы к проекту
        public List<DocumentForProject> DocumentForProjects { get; set; }
    }

    public class PArticle // Изделие
    {
        [Required][ForeignKey("Proj")]
        public int ProjectID { get; set; }
        [Required][Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int PArticleID { get; set; }
        [Required][ForeignKey("LocationDescription")]
        public int LocationDesriptionID { get; set; }

        // Ссылка на проект
        public Proj Proj { get; set; }
        
        // Ссылка на описание изделий
        public LocationDescription LocationDescription { get; set; }

        // Ссылка на компонент
        public List<Component> Components { get; set; }
    }

    public class Component // Компонент
    {
        [Key]
        public int ComponentID { get; set; }
        [ForeignKey("PArticle")]
        public int PArticleID { get; set; }
        public int? SubArticleID { get; set; } // Код подизделия
        [Required]
        public string PartNumber { get; set; }
        [Required]
        public int Count { get; set; }

        // Ссылка на изделие
        public PArticle PArticle { get; set; }
    }

    public class ComponentCatalog // Справочник компонентов
    {
        [Required][Key, DatabaseGenerated(DatabaseGeneratedOption.None)]
        public string PartNumber { get; set; }
        public string TypeNumber { get; set; }
        public string OrderNumber { get; set; }
        public string ManufacturerSmallName { get; set; }
        public string ManufacturerFullName { get; set; }
        public string SupplierSmallName { get; set; }
        public string SupplierFullName { get; set; }
        public string Description1 { get; set; }
        public string Description2 { get; set; }
        public string TechnicalCharacteristics { get; set; }
        public string Note { get; set; }
        public string TerminalCrossSectionFrom { get; set; }
        public string TerminalCrossSectionTo { get; set; }
        public string ElectricalCurrent { get; set; }
        public string ElectricalSwitchingCapacity { get; set; }
        public string Voltage { get; set; }
        public string VoltageType { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public double Depth { get; set; }
        public double Weight { get; set; }
        [ForeignKey("MountingSite")]
        public int MountingSiteID { get; set; }
        public double MountingSpace { get; set; }
        public string PartGroup { get; set; }
        [ForeignKey("QuantityUnit")]
        public int? QuantityUnitID { get; set; }
        public double PackagingQuantity { get; set; }
        public string LastChange { get; set; }

        // Ссылка на единицы измерения
        public QuantityUnit QuantityUnit { get; set; }
        // Ссылка на монтажную поверхность
        public MountingSite MountingSite { get; set; }
        // Ссылка на прайсы
        public List<PriceList> PriceLists { get; set; }

    }

    public class Area // Область применения (Гл. групп., справочник)
    {
        [Required][Key]
        public int AreaID { get; set; }
        [Required]
        public int InternalValue { get; set; } // Для дефолтного значения
        [Required]
        public string Name { get; set; }
    }

    public class Section // Раздел (групп., справочник)
    {
        [Required][Key]
        public int SectionID { get; set; }
        [Required]
        public int InternalValue { get; set; } // Для дефолтного значения
        [Required]
        public string Name { get; set; }
    }

    public class Sphere // Сфера применения (подгрупп., справочник)
    {
        [Required][Key]
        public int SphereID { get; set; }
        [Required]
        public int InternalValue { get; set; } // Для дефолтного значения
        [Required]
        public string Name { get; set; }
    }

    public class QuantityUnit // Единицы измерения
    {
        [Required]
        [Key]
        public int QuantityUnitID { get; set; }
        [Required]
        public string Name { get; set; }

        // Ссылка на справочник компонентов
        public List<ComponentCatalog> ComponentCatalogs { get; set; }
    }

    public class PriceList // Прейскурант
    {
        [Required][Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int PriceListID { get; set; }
        [Required][ForeignKey("ComponentCatalog")]
        public string PartNumber { get; set; }
        public double SalesPrice1 { get; set; }
        public double SalesPrice2 { get; set; }
        public double PurchasePrice1 { get; set; }
        public double PurchasePrice2 { get; set; }
        public double PackagingPrice1 { get; set; }
        public double PackagingPrice2 { get; set; }
        public int PriceUnit { get; set; }
        public DateTime LastChange { get; set; }

        // Ссылка на справочник компонентов
        public ComponentCatalog ComponentCatalog { get; set; }
    }

    public class MountingSite // Монтажная поверхность
    {
        [Required][Key]
        public int MountingSiteID { get; set; }
        [Required]
        public int InternalValue { get; set; } // Для дефолтного значения
        [Required]
        public string Name { get; set; }

        // Ссылка на справочник компонентов
        public List<ComponentCatalog> ComponentCatalogs { get; set; }
    }

    public class LocationDescription // Структурные идентификаторы (справочник)
    {
        [Required][Key]
        public int LocationDescriptionID { get; set; }
        [Required]
        public string Name { get; set; }
        [Required]
        public string Description { get; set; }

        // Ссылка на изделия
        public List<PArticle> PArticles { get; set; }
    }

    public class DocumentForProject // Документ к проекту
    {
        [Required][Key]
        public int DocumentID { get; set; }
        [Required][ForeignKey("Proj")]
        public int ProjectID { get; set; }
        [Required]
        public string DocumentName { get; set; }
        [Required]
        public string DocumentType { get; set; }
        [Required]
        public string DocumentPath { get; set; }

        // Ссылка на проекты
        public Proj Proj { get; set; }
    }
}