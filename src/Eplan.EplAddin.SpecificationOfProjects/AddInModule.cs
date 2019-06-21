using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;


namespace Eplan.EplAddin.SpecificationListOfObjects
{
    public class AddInModule : IEplAddIn, IEplAddInShadowCopy
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
            MyMenu.AddMainMenu("Надстройки", Menu.MainMenuName.eMainMenuUtilities,
                "Выгрузить спецификацию", "GetSpecification", "In progress", 1);
            return true;
        }

        public bool OnExit()
        {
            return true;
        }

        // До инициализации получаю путь к надстройке
        public void OnBeforeInit(string strOriginalAssemblyPath)
        {
            OriginalAssemblyPath = strOriginalAssemblyPath;
        }
        
        // static переменная для пути
        public static string OriginalAssemblyPath;
    }
}

 //  - Методы OnRegister и OnUnregister вызываются по одному разу, при первом 
 //    подключении и удалении Add-Ina соответственно.
 //  - Метод OnExit вызывается при закрытии EPLANa.
 //  - Метод OnInit вызывается при загрузке EPLANa для инициализации Add-Ina.
 //  - Метод OnInitGui вызывается при загрузке EPLANa для инициализации 
 //    Add-Ina и пользовательского интерфейса.
