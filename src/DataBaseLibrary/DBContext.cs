using System.Data.Entity;
using DataBaseLibrary.Migrations;

namespace DataBaseLibrary
{
    public class DataBaseInitializer : MigrateDatabaseToLatestVersion<DBContext, Configuration> { }

    public class DBContext : DbContext
    {
        public DBContext() : base(@"Data Source=WIN-CKJ0931HT91\SQLEXPRESS;" +
            "Initial Catalog=SavushkinProject;" +
            "Persist Security Info = False;" +
            "Integrated Security=False;" +
            "User ID=sa;" +
            "Password=root;" +
            "Trusted_Connection=false;")
        {
            Database.SetInitializer<DBContext>(new DataBaseInitializer());
        }

        public DbSet<PArticle> PArticles { get; set; } // Изделия
        public DbSet<Component> Components { get; set; } // Компоненты
        public DbSet<MountingSite> MountingSites { get; set; } // Монтажные поверхности
        public DbSet<PriceList> PriceLists { get; set; } // Прайсы
        public DbSet<Proj> Projs { get; set; } // Проекты
        public DbSet<QuantityUnit> QuantityUnits { get; set; } // Ед. измерения
        public DbSet<ComponentCatalog> ComponentCatalogs { get; set; } // Справочник компонентов
        public DbSet<Area> Areas { get; set; } // Область
        public DbSet<Section> Sections { get; set; } // Раздел
        public DbSet<Sphere> Spheres { get; set; } // Сфера 
        public DbSet<LocationDescription> LocationDescriptions { get; set; } // Структурные обозначения
    }
}
