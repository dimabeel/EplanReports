using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Eplan.EplAddin.SpecificationListOfObjects
{
    class ConfigIniFile
    {
        string filePath; // Имя файла и путь

        // Подключаем kernel32.dll и описываем его функцию WritePrivateProfileString
        [DllImport("kernel32")]
        static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

        // Еще раз подключаем kernel32.dll, а теперь описываем функцию GetPrivateProfileString
        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        // С помощью конструктора записываем путь до файла и его имя.
        public ConfigIniFile(string IniPath)
        {
            filePath = new FileInfo(IniPath).FullName.ToString();
        }

        //Читаем ini-файл и возвращаем значение указного ключа из заданной секции.
        public string ReadINI(string Section, string Key)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(Section, Key, "", RetVal, 255, filePath);
            return RetVal.ToString();
        }

        //Записываем в ini-файл. Запись происходит в выбранную секцию в выбранный ключ.
        public void Write(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, filePath);
        }

        //Удаляем ключ из выбранной секции.
        public void DeleteKey(string Key, string Section = null)
        {
            Write(Section, Key, null);
        }

        //Удаляем выбранную секцию
        public void DeleteSection(string Section = null)
        {
            Write(Section, null, null);
        }

        //Проверяем, есть ли такой ключ, в этой секции
        public bool KeyExists(string Key, string Section = null)
        {
            return ReadINI(Section, Key).Length > 0;
        }
    }
}
