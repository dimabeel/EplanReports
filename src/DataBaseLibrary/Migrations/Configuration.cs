namespace DataBaseLibrary.Migrations
{
    using System.Collections.Generic;
    using System.Data.Entity.Migrations;
    using DataBaseLibrary;

    public sealed class Configuration : DbMigrationsConfiguration<DBContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
        }

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
}
