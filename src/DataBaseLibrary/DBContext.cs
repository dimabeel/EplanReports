using System.Data.Entity;
using System.Collections.Generic;


namespace DataBaseLibrary
{
    public class DataBaseInitializer : CreateDatabaseIfNotExists<DBContext>
    {

        protected override void Seed(DBContext context)
        {
            // Инициализация базы данных при её создании
            DBDefaultData DBDefaultData = new DBDefaultData();
            List<MountingSite> mountingSites = DBDefaultData.GetMountingSites();
            List<Area> areas = DBDefaultData.GetAreas();
            List<Section> sections = DBDefaultData.GetSections();
            List<Sphere> spheres = DBDefaultData.GetSpheres();

            foreach (MountingSite mountingSite in mountingSites)
            {
                context.MountingSites.Add(mountingSite);
            }

            for (int i = 0; i < areas.Count; i++)
            {
                context.Areas.Add(areas[i]);
            }

            for (int i = 0; i < sections.Count; i++)
            {
                context.Sections.Add(sections[i]);
            }

            for (int i = 0; i < spheres.Count; i++)
            {
                context.Spheres.Add(spheres[i]);
            }

            context.SaveChanges();
            base.Seed(context);
        }

    }

    public class DBContext : DbContext
    {
        public DBContext() : base()
        {
            Database.SetInitializer<DBContext>(new DataBaseInitializer());
        }

        public DBContext(string nameOrConnectionString) : base(nameOrConnectionString)
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
        public DbSet<DocumentForProject> DocumentForProjects { get; set; } // Документы для проекта
    }
}
