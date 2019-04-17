namespace Eplan.EplAddin.SpecificationOfProjects
{

    // Модель данных для спецификации
    public class ComponentInfo
    {
        public string PartNumber;
        public int Count;
    }

    // Модель данных для считывания свойств компонента из EPLAN
    public class ComponentCatalogInfo
    {
        public string PartNumber;
        public string TypeNumber;
        public string OrderNumber;
        public string ManufacturerSmallName;
        public string ManufacturerFullName;
        public string SupplierSmallName;
        public string SupplierFullName;
        public string Description1;
        public string Description2;
        public string TechnicalCharacteristics;
        public string Note;
        public string TerminalCrossSectionFrom;
        public string TerminalCrossSectionTo;
        public string ElectricalCurrent;
        public string ElectricalSwitchingCapacity;
        public string Voltage;
        public string VoltageType;
        public double Height;
        public double Width;
        public double Depth;
        public double Weight;
        public int MountingSiteID;
        public double MountingSpace;
        public string PartGroup; // Преобразуем по шаблону "0-00-000"
        public string QuantityUnit;
        public double PackagingQuantity;
        public string LastChange;
        public double SalesPrice1;
        public double SalesPrice2;
        public double PurchasePrice1;
        public double PurchasePrice2;
        public double PackagingPrice1;
        public double PackagingPrice2;
        public int PriceUnit;
    }

    public class LocationInfo // Модель для выборки структурных обозначений
    {
        public string Name;
        public string Description;
    }
}
