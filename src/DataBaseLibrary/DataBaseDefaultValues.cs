﻿using System.Collections.Generic;

namespace DataBaseLibrary
{
    // Класс, содержащий дефолтные данные для БД
    public class DataBaseDefaultValues
    {
        // Дейолтные монтажные поверхности
        List<MountingSite> mountingSites = new List<MountingSite>
        {
            new MountingSite {InternalValue = 0, Name = "Не определено"},
            new MountingSite {InternalValue = 1, Name = "Монтажная плата"},
            new MountingSite {InternalValue = 2, Name = "Поворотная рама"},
            new MountingSite {InternalValue = 3, Name = "Дверь"},
            new MountingSite {InternalValue = 4, Name = "Боковая часть"},
            new MountingSite {InternalValue = 5, Name = "Крыша"},
        };

        // Дефолтные области применения
        List<Area> areas = new List<Area>
        {
            new Area {InternalValue = 0, Name = "Не определено"},
            new Area {InternalValue = 1, Name = "Электротехника"},
            new Area {InternalValue = 2, Name = "Fluid-техника"},
            new Area {InternalValue = 3, Name = "Механика"},
            new Area {InternalValue = 4, Name = "Технология производственных процессов"},
        };

        // Дефолтные разделы
        List<Section> sections = new List<Section>
        {
            new Section {InternalValue = 0, Name = "Не определено"},
            new Section {InternalValue = 1, Name = "Общее"},
            new Section {InternalValue = 2, Name = "Реле, контакторы"},
            new Section {InternalValue = 3, Name = "Клеммы"},
            new Section {InternalValue = 4, Name = "Штекеры"},
            new Section {InternalValue = 5, Name = "Преобразователи"},
            new Section {InternalValue = 6, Name = "Защитные устройства"},
            new Section {InternalValue = 7, Name = "Лампы, полупроводники"},
            new Section {InternalValue = 8, Name = "Сигнальные устройства"},
            new Section {InternalValue = 9, Name = "Двигатели"},
            new Section {InternalValue = 10, Name = "Измерительные приборы, контрольное оборудование"},
            new Section {InternalValue = 11, Name = "Резисторы"},
            new Section {InternalValue = 12, Name = "Переключатели, селекторы"},
            new Section {InternalValue = 13, Name = "Трансформаторы"},
            new Section {InternalValue = 14, Name = "Модуляторы"},
            new Section {InternalValue = 15, Name = "Механическое оборудование с электрическим приводом"},
            new Section {InternalValue = 16, Name = "Спец. функциональные элементы"},
            new Section {InternalValue = 17, Name = "Разное"},
            new Section {InternalValue = 18, Name = "Конденсаторы"},
            new Section {InternalValue = 19, Name = "Цифровые элементы"},
            new Section {InternalValue = 20, Name = "Генераторы, электрообеспечение"},
            new Section {InternalValue = 21, Name = "Индуктивности"},
            new Section {InternalValue = 22, Name = "Усилители, регуляторы"},
            new Section {InternalValue = 23, Name = "Сильноточные коммутационные устройства"},
            new Section {InternalValue = 24, Name = "Крышки, фильтры"},
            new Section {InternalValue = 25, Name = "Пути передачи"},
            new Section {InternalValue = 26, Name = "ПЛК"},
            new Section {InternalValue = 29, Name = "Кабели/соединения"},
            new Section {InternalValue = 30, Name = "Агрегаты и установки"},
            new Section {InternalValue = 32, Name = "Приводы"},
            new Section {InternalValue = 33, Name = "Приводные двигатели"},
            new Section {InternalValue = 35, Name = "Управляющий терминал"},
            new Section {InternalValue = 36, Name = "Муфты"},
            new Section {InternalValue = 37, Name = "Системы трубопроводов"},
            new Section {InternalValue = 38, Name = "Выводы устройства для измерения"},
            new Section {InternalValue = 39, Name = "Измерительные приборы"},
            new Section {InternalValue = 40, Name = "Насосы"},
            new Section {InternalValue = 41, Name = "Получатели сигнала"},
            new Section {InternalValue = 42, Name = "Спец. функциональные элементы"},
            new Section {InternalValue = 43, Name = "Накопитель"},
            new Section {InternalValue = 44, Name = "Клапаны"},
            new Section {InternalValue = 47, Name = "Принадлежности"},
            new Section {InternalValue = 48, Name = "Присоединительные платы"},
            new Section {InternalValue = 49, Name = "Корпус"},
            new Section {InternalValue = 50, Name = "Принадлежности корпуса для наружной установки"},
            new Section {InternalValue = 51, Name = "Принадлежности корпуса для внутренней установки"},
            new Section {InternalValue = 52, Name = "Запорные системы"},
            new Section {InternalValue = 53, Name = "Кабельные каналы"},
            new Section {InternalValue = 54, Name = "Сборные шины"},
            new Section {InternalValue = 55, Name = "Электрошкаф"},
            new Section {InternalValue = 56, Name = "Вытяжной колпак"},
            new Section {InternalValue = 57, Name = "Электролизер"},
            new Section {InternalValue = 58, Name = "Опора"},
            new Section {InternalValue = 59, Name = "Дымовая труба"},
            new Section {InternalValue = 61, Name = "Приводной двигатель"},
            new Section {InternalValue = 62, Name = "Запорная арматура"},
            new Section {InternalValue = 63, Name = "Трехходовая арматура"},
            new Section {InternalValue = 64, Name = "Обратная арматура"},
            new Section {InternalValue = 65, Name = "Резервуар"},
            new Section {InternalValue = 66, Name = "Вывод резервуара"},
            new Section {InternalValue = 67, Name = "Сепаратор"},
            new Section {InternalValue = 68, Name = "Фильтр"},
            new Section {InternalValue = 69, Name = "Сито"},
            new Section {InternalValue = 70, Name = "Конвейер"},
            new Section {InternalValue = 71, Name = "Подъемник"},
            new Section {InternalValue = 72, Name = "Транспортер"},
            new Section {InternalValue = 73, Name = "Оболочка резервуара"},
            new Section {InternalValue = 74, Name = "Змеевик резервуара"},
            new Section {InternalValue = 75, Name = "Паровой котел"},
            new Section {InternalValue = 76, Name = "Охладитель"},
            new Section {InternalValue = 77, Name = "Печь"},
            new Section {InternalValue = 78, Name = "Сушилка"},
            new Section {InternalValue = 79, Name = "Испаритель"},
            new Section {InternalValue = 80, Name = "Теплообменник"},
            new Section {InternalValue = 81, Name = "Меситель"},
            new Section {InternalValue = 82, Name = "Миксер"},
            new Section {InternalValue = 83, Name = "Мешалка"},
            new Section {InternalValue = 84, Name = "Компрессор"},
            new Section {InternalValue = 86, Name = "Вакуумный насос"},
            new Section {InternalValue = 87, Name = "Вентилятор"},
            new Section {InternalValue = 88, Name = "Компрессор"},
            new Section {InternalValue = 89, Name = "Часть трубопровода"},
            new Section {InternalValue = 90, Name = "Сепаратор"},
            new Section {InternalValue = 91, Name = "Центрифуга"},
            new Section {InternalValue = 92, Name = "Измер. устройство"},
            new Section {InternalValue = 93, Name = "Весы"},
            new Section {InternalValue = 94, Name = "Форм. машина"},
            new Section {InternalValue = 95, Name = "Просеиватель"},
            new Section {InternalValue = 96, Name = "Сортировщик"},
            new Section {InternalValue = 97, Name = "Измельчитель"},
            new Section {InternalValue = 98, Name = "Дозатор"},
        };

        // Дефолтные сферы применения
        List<Sphere> spheres = new List<Sphere>
        {
            new Sphere {InternalValue = 0, Name = "Не определено"},
            new Sphere {InternalValue = 1, Name = "Общее"},
            new Sphere {InternalValue = 2, Name = "Контакторы"},
            new Sphere {InternalValue = 3, Name = "Вспомогательный блок"},
            new Sphere {InternalValue = 4, Name = "Клемма"},
            new Sphere {InternalValue = 5, Name = "Несущая шина"},
            new Sphere {InternalValue = 6, Name = "Конечный угол"},
            new Sphere {InternalValue = 7, Name = "Клеммная табличка"},
            new Sphere {InternalValue = 8, Name = "Другие таблички"},
            new Sphere {InternalValue = 9, Name = "Табличка колодки"},
            new Sphere {InternalValue = 10, Name = "Поперечный соединитель"},
            new Sphere {InternalValue = 11, Name = "Шины"},
            new Sphere {InternalValue = 12, Name = "Крышка"},
            new Sphere {InternalValue = 13, Name = "Крышка"},
            new Sphere {InternalValue = 14, Name = "Перегородка"},
            new Sphere {InternalValue = 15, Name = "Корпус штекера"},
            new Sphere {InternalValue = 16, Name = "Штекерные принадлежности"},
            new Sphere {InternalValue = 17, Name = "Инструмент"},
            new Sphere {InternalValue = 18, Name = "Тестовые принадлежности"},
            new Sphere {InternalValue = 19, Name = "Изделие, источник (вид соединения и кабеля источника)"},
            new Sphere {InternalValue = 20, Name = "Изделие, цель (вид соединения и кабеля на цели)"},
            new Sphere {InternalValue = 47, Name = "Установки высокого давления"},
            new Sphere {InternalValue = 48, Name = "Резервуары высокого давления"},
            new Sphere {InternalValue = 49, Name = "Гидравлические баки"},
            new Sphere {InternalValue = 50, Name = "Комплектные агрегаты"},
            new Sphere {InternalValue = 51, Name = "Холодильные агрегаты"},
            new Sphere {InternalValue = 52, Name = "Агрегаты охлаждения электрошкафов"},
            new Sphere {InternalValue = 53, Name = "Винтовые компрессоры"},
            new Sphere {InternalValue = 54, Name = "Спиральные компрессоры"},
            new Sphere {InternalValue = 55, Name = "Теплообменник"},
            new Sphere {InternalValue = 56, Name = "Централизованные смазочные агрегаты"},
            new Sphere {InternalValue = 57, Name = "Присоединительные платы"},
            new Sphere {InternalValue = 58, Name = "Защитные крышки"},
            new Sphere {InternalValue = 59, Name = "Адаптерные пластины"},
            new Sphere {InternalValue = 60, Name = "Запорные шаровые клапаны"},
            new Sphere {InternalValue = 61, Name = "Фланцы"},
            new Sphere {InternalValue = 62, Name = "Шаровые клапаны"},
            new Sphere {InternalValue = 63, Name = "Переходные платы"},
            new Sphere {InternalValue = 64, Name = "Монтажные платы для последовательного присоединения"},
            new Sphere {InternalValue = 65, Name = "Управляющие блоки"},
            new Sphere {InternalValue = 66, Name = "Промежуточные платы"},
            new Sphere {InternalValue = 67, Name = "Гидромоторы"},
            new Sphere {InternalValue = 68, Name = "Приводы качания"},
            new Sphere {InternalValue = 69, Name = "Цилиндры"},
            new Sphere {InternalValue = 71, Name = "Фильтры системы вентиляции"},
            new Sphere {InternalValue = 72, Name = "Двойные фильтры с переключаемыми камерами"},
            new Sphere {InternalValue = 73, Name = "Фильтры сжатого воздуха"},
            new Sphere {InternalValue = 74, Name = "Одноступенчатые фильтры"},
            new Sphere {InternalValue = 75, Name = "Линейные фильтры"},
            new Sphere {InternalValue = 76, Name = "Обратный фильтр"},
            new Sphere {InternalValue = 77, Name = "Автоматические фильтры с обратной промывкой фильтрующих элементов"},
            new Sphere {InternalValue = 79, Name = "Насосные установки с муфтой"},
            new Sphere {InternalValue = 80, Name = "Измерительные муфты"},
            new Sphere {InternalValue = 81, Name = "Запорные муфты"},
            new Sphere {InternalValue = 82, Name = "Шланги"},
            new Sphere {InternalValue = 83, Name = "Разное"},
            new Sphere {InternalValue = 85, Name = "Манометры"},
            new Sphere {InternalValue = 86, Name = "Термометры"},
            new Sphere {InternalValue = 87, Name = "Компрессоры"},
            new Sphere {InternalValue = 88, Name = "Насосы с постоянной производительностью"},
            new Sphere {InternalValue = 89, Name = "Центробежные насосы"},
            new Sphere {InternalValue = 90, Name = "Винтовые насосы"},
            new Sphere {InternalValue = 91, Name = "Насосы с поворотными лопастями"},
            new Sphere {InternalValue = 92, Name = "Реле давления"},
            new Sphere {InternalValue = 93, Name = "Реле уровня"},
            new Sphere {InternalValue = 94, Name = "Манометры получателя сигнала"},
            new Sphere {InternalValue = 95, Name = "Бесконтактные переключатели"},
            new Sphere {InternalValue = 96, Name = "Расходомеры"},
            new Sphere {InternalValue = 97, Name = "Температурные датчики"},
            new Sphere {InternalValue = 99, Name = "Запорные резервуары"},
            new Sphere {InternalValue = 100, Name = "Аккумуляторы давления"},
            new Sphere {InternalValue = 101, Name = "Поршневые гидропневматические аккумуляторы"},
            new Sphere {InternalValue = 102, Name = "Аварийные резервуары"},
            new Sphere {InternalValue = 103, Name = "Блоки резервуаров"},
            new Sphere {InternalValue = 104, Name = "Гидромультипликаторы давления"},
            new Sphere {InternalValue = 105, Name = "Нагнетательные клапаны"},
            new Sphere {InternalValue = 106, Name = "Воздушные клапаны"},
            new Sphere {InternalValue = 107, Name = "Поршневые распределители"},
            new Sphere {InternalValue = 108, Name = "Пневматический шаровой клапан"},
            new Sphere {InternalValue = 109, Name = "Предохранительные клапаны манометра"},
            new Sphere {InternalValue = 110, Name = "Переключатели манометра"},
            new Sphere {InternalValue = 111, Name = "Прогрессивные распределители"},
            new Sphere {InternalValue = 112, Name = "Обратные клапаны"},
            new Sphere {InternalValue = 113, Name = "Запорные клапаны"},
            new Sphere {InternalValue = 114, Name = "Проточные клапаны"},
            new Sphere {InternalValue = 115, Name = "Делительные клапаны"},
            new Sphere {InternalValue = 116, Name = "Клапаны-термостаты"},
            new Sphere {InternalValue = 117, Name = "Ходовые коаксиальные клапаны"},
            new Sphere {InternalValue = 118, Name = "Ходовые клапаны"},
            new Sphere {InternalValue = 122, Name = "Кронштейны крепления насоса"},
            new Sphere {InternalValue = 124, Name = "Емкости"},
            new Sphere {InternalValue = 126, Name = "Резьбовые соединения"},
            new Sphere {InternalValue = 127, Name = "Водоотделители"},
            new Sphere {InternalValue = 128, Name = "Коробки для чертежей"},
            new Sphere {InternalValue = 130, Name = "Дверь"},
            new Sphere {InternalValue = 131, Name = "Задняя/боковая стенка"},
            new Sphere {InternalValue = 132, Name = "Крыша"},
            new Sphere {InternalValue = 133, Name = "Пол"},
            new Sphere {InternalValue = 134, Name = "Цоколь"},
            new Sphere {InternalValue = 135, Name = "Монтажные платы"},
            new Sphere {InternalValue = 136, Name = "Принадлежности"},
            new Sphere {InternalValue = 139, Name = "Монтажные шины кабеля"},
            new Sphere {InternalValue = 141, Name = "Направляющие шины"},
            new Sphere {InternalValue = 142, Name = "Профильные С-шины"},
            new Sphere {InternalValue = 146, Name = "Система"},
            new Sphere {InternalValue = 147, Name = "Кронштейн"},
            new Sphere {InternalValue = 148, Name = "Адаптер"},
            new Sphere {InternalValue = 150, Name = "Отдельная часть"},
            new Sphere {InternalValue = 151, Name = "Корпус"},
            new Sphere {InternalValue = 152, Name = "Карта аналогового входа ПЛК"},
            new Sphere {InternalValue = 153, Name = "Карта цифрового входа ПЛК"},
            new Sphere {InternalValue = 154, Name = "Карта аналогового выхода ПЛК"},
            new Sphere {InternalValue = 155, Name = "Карта цифрового выхода ПЛК"},
            new Sphere {InternalValue = 157, Name = "Бункер"},
            new Sphere {InternalValue = 158, Name = "Элеватор"},
            new Sphere {InternalValue = 159, Name = "Колонна"},
            new Sphere {InternalValue = 160, Name = "Реактор"},
            new Sphere {InternalValue = 161, Name = "Запорная арматура"},
            new Sphere {InternalValue = 162, Name = "Запорный клапан"},
            new Sphere {InternalValue = 163, Name = "Запорная заслонка"},
            new Sphere {InternalValue = 164, Name = "Обратная арматура"},
            new Sphere {InternalValue = 165, Name = "Обратный клапан"},
            new Sphere {InternalValue = 166, Name = "Обратная заслонка"},
            new Sphere {InternalValue = 167, Name = "Трехходовая арматура"},
            new Sphere {InternalValue = 168, Name = "Трехходовой клапан"},
            new Sphere {InternalValue = 169, Name = "Трехходовая заслонка"},
            new Sphere {InternalValue = 170, Name = "Дроссельный клапан"},
            new Sphere {InternalValue = 171, Name = "Заглушка"},
            new Sphere {InternalValue = 172, Name = "Воронка"},
            new Sphere {InternalValue = 173, Name = "Выпуск"},
            new Sphere {InternalValue = 174, Name = "Глушители"},
            new Sphere {InternalValue = 175, Name = "Смотровое стекло"},
            new Sphere {InternalValue = 176, Name = "Смес. сопло"},
            new Sphere {InternalValue = 177, Name = "Грязеуловитель"},
            new Sphere {InternalValue = 178, Name = "Конденсатоотводчик"},
            new Sphere {InternalValue = 179, Name = "Фланц. пара"},
            new Sphere {InternalValue = 180, Name = "Уменьшение"},
            new Sphere {InternalValue = 181, Name = "Муфты"},
            new Sphere {InternalValue = 182, Name = "Сифон"},
            new Sphere {InternalValue = 183, Name = "Пресс"},
        };

        // Геттеры
        public List<Sphere> GetSpheres()
        {
            return spheres;
        }

        public List<Section> GetSections()
        {
            return sections;
        }

        public List<Area> GetAreas()
        {
            return areas;
        }

        public List<MountingSite> GetMountingSites()
        {
            return mountingSites;
        }
    }
}
