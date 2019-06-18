using System.Data.Entity;
using System.Collections.Generic;


namespace DataBaseLibrary
{
    public class DataBaseInitializer : CreateDatabaseIfNotExists<DataBaseContext>
    {

        protected override void Seed(DataBaseContext context)
        {
            // Инициализация базы данных при её создании
            var dataBaseDefaultValue = new DataBaseDefaultValues();
            var mountingSites = dataBaseDefaultValue.GetMountingSites();
            var areas = dataBaseDefaultValue.GetAreas();
            var sections = dataBaseDefaultValue.GetSections();
            var spheres = dataBaseDefaultValue.GetSpheres();

            for (int item = 0; item < mountingSites.Count; item++)
            {
                context.MountingSites.Add(mountingSites[item]);
            }

            for (int item = 0; item < areas.Count; item++)
            {
                context.Areas.Add(areas[item]);
            }

            for (int item = 0; item < sections.Count; item++)
            {
                context.Sections.Add(sections[item]);
            }

            for (int item = 0; item < spheres.Count; item++)
            {
                context.Spheres.Add(spheres[item]);
            }

            context.SaveChanges();
            base.Seed(context);
        }

    }

    public class DataBaseContext : DbContext
    {
        public DataBaseContext() : base()
        {
            Database.SetInitializer<DataBaseContext>(new DataBaseInitializer());
        }

        public DataBaseContext(string nameOrConnectionString) : base(nameOrConnectionString)
        {
            Database.SetInitializer<DataBaseContext>(new DataBaseInitializer());
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
        public DbSet<DocumentForProject> DocumentForProjects { get; set; } // Документы для проекта
    }
}
