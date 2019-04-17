using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;

namespace Eplan.AplAddin.SpecificationListOfObjects
{
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref bool bLoadStart)
        {
            bLoadStart = true;
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }

        public bool OnInit()
        {
            return true;
        }

        public bool OnInitGui()
        {
            //Добавить меню для действия
            Menu MyMenu = new Menu();
            MyMenu.AddMainMenu("Add-in's", Menu.MainMenuName.eMainMenuUtilities,
                "Get project specification", "GetSpecification", "In progress", 1);
            return true;
        }

        public bool OnExit()
        {
            return true;
        }
    }
}

 //  - Методы OnRegister и OnUnregister вызываются по одному разу, при первом 
 //    подключении и удалении Add-Ina соответственно.
 //  - Метод OnExit вызывается при закрытии EPLANa.
 //  - Метод OnInit вызывается при загрузке EPLANa для инициализации Add-Ina.
 //  - Метод OnInitGui вызывается при загрузке EPLANa для инициализации 
 //    Add-Ina и пользовательского интерфейса.
