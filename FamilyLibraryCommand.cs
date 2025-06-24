
using Autodesk.Revit.Attributes; // Атрибуты для команд Revit 
using Autodesk.Revit.DB; //  классы для работы с базой данных 
using Autodesk.Revit.DB.Structure; // Классы для работы сструктурными элементами
using Autodesk.Revit.UI; // Классы для работы с интерфейсом Revit
using Microsoft.WindowsAPICodePack.Dialogs; 
using System.IO; 
using System.Linq; 
using System.Collections.Generic; 
using System.Windows.Forms; 
using System; 
using View = Autodesk.Revit.DB.View; 
using Form = System.Windows.Forms.Form; // Создаем псевдоним для класса 
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog; 


namespace FamilyLibraryPresenter
{
 
    [Transaction(TransactionMode.Manual)]
    
    public class FamilyLibraryCommand : IExternalCommand
    {
        // ПОЛЯ КЛАССА
        private double koordinatX = 0; // Текущая координата X для размещения семейств
        private double koordinatY = 0; // Текущая координата Y для размещения семейств
        private double maxVisotaVStroke = 0; // Максимальная высота семейства для расчета смещения
        private const double otstup = 10.0; // Постоянное значение отступа между семействами
        private const int kolichStolbtsov = 5; // Постоянное значение количества столбцов
        private Document dokument; // Документ Revit, с которым мы работаем
        private UIDocument uiDokument; // Документ с привязкой к интерфейсу
        private View aktivnVid; // Активный вид, для размещаются семейства
        private Level uroven; // Уровень, на котором будут размещаться семейства
        private View3D vid3D; // 3D-вид, который будет использоваться для отображения
        private List<string> logOshibok = new List<string>(); // для записи ошибок и информации о размещении
        private Dictionary<string, List<(FamilySymbol Symbol, string Path)>> sgruppirovannyeSimvoly = new Dictionary<string, List<(FamilySymbol, string)>>(); // для хранения сгруппированных символов семейств

        // ГЛАВНЫЙ МЕТОД, ВЫПОЛНЯЮЩИЙСЯ ПРИ ЗАПУСКЕ 
        public Result Execute(ExternalCommandData dannyeKomandy, ref string soobshenie, ElementSet elementy)
        {
            // Получаем доступ к активному документу и
            uiDokument = dannyeKomandy.Application.ActiveUIDocument; // Получаем UI-документ
            dokument = uiDokument.Document; // Получаем сам документ

            // ИЩЕМ ПОДХОДЯЩИЙ 3D-ВИД
            vid3D = new FilteredElementCollector(dokument) // Создаем  в документе
                .OfClass(typeof(View3D)) // Указываем, что ищем элементы класса View3D
                .Cast<View3D>() // Преобразуем результат в коллекцию 
                .FirstOrDefault(v => !v.IsTemplate && v.Name.ToLower().Contains("3d")); // Находим первый 3D-вид, который не является шаблоном и имя которого содержит "3d"

            // ПРОВЕРЯЕМ, НАЙДЕН ЛИ 3D-ВИД
            if (vid3D == null) // Если 3D-вид не найден
            {
                RevitTaskDialog.Show("Ошибка", "3D-вид не найден"); // окно ошибка
                return Result.Failed; // завершаем
            }

            // УСТАНАВЛИВАЕМ НАЙДЕННЫЙ ВИД АКТИВНЫМ
            uiDokument.ActiveView = vid3D; // Устанавливаем 3D-вид как активный
            aktivnVid = vid3D; // Сохраняем его 
            
            // ИЩЕМ САМЫЙ НИЖНИЙ УРОВЕНЬ В ПРОЕКТЕ
            uroven = new FilteredElementCollector(dokument) // Создаем новый 
                .OfClass(typeof(Level)) // Ищем элементы класса уровнь
                .Cast<Level>() // Преобразуем результат в коллекцию уровней
                .OrderBy(l => l.Elevation) // Сортируем уровни по их высоте 
                .FirstOrDefault(); // Берем первый самый низкий уровень

            // СОЗДАЕМ И ПОКАЗЫВАЕМ ДИАЛОГ ВЫБОРА ПАПКИ
            var dialog = new CommonOpenFileDialog // Создаем диалоговое окно
            {
                IsFolderPicker = true, // это окно для выбора папки, а не файла
                Title = "Выберите папку с семействами" // заголовок окна
            };

            //  ВЫБРАЛ ЛИ ПОЛЬЗОВАТЕЛЬ ПАПКУ?
            if (dialog.ShowDialog() != CommonFileDialogResult.Ok) // Если пользователь закрыл окно или нажал Отмена
                return Result.Cancelled; // Завершаем команду

            // СОХРАНЯЕМ ПУТЬ К ВЫБРАННОЙ ПАПКЕ
            string putKPapke = dialog.FileName; 

            // ДИАЛОГ ВЫБОРА СПОСОБА ГРУППИРОВКИ
            var dialogGruppirovki = new GroupingDialog(); // Создаем окно
            if (dialogGruppirovki.ShowDialog() != DialogResult.OK) // Если  нажал "ОК"
                return Result.Cancelled; // Завершаем команду

            // РАБОТА С ФАЙЛАМИ И РАЗМЕЩЕНИЕМ
            try // для отлова возможных ошибок
            {
                string[] fayliSemeystv = Directory.GetFiles(putKPapke, "*.rfa", SearchOption.AllDirectories); // Получаем массив путей ко всем файлам .rfa в выбранной папке и ее подпапках
                LoadFamilies(fayliSemeystv, dialogGruppirovki.GroupByDirectory); // Вызываем метод для загрузки семейств
                PlaceFamilies(); // Вызываем метод для размещения семейств
            }
            catch (Exception isklyuchenie) // Если в блоке try произошла ошибка
            {
                RevitTaskDialog.Show("Ошибка", isklyuchenie.Message); // Показываем  сообщение об ошибке ОШИБКА 
                return Result.Failed; // Завершаем команду faile
            }

            // ЗАПИСЬ ЛОГА И ВЫВОД РЕЗУЛЬТАТА
            if (logOshibok.Count > 0) // Если в списке ошибок есть записи
            {
                string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "RevitFamilyPlacementErrors.log"); // Создаем путь для лог-файла на рабочем столе
                File.WriteAllLines(logPath, logOshibok); // Записываем все строки из списка ошибок в файл
                System.Diagnostics.Process.Start("notepad.exe", logPath); // Открываем созданный лог-файл в Блокноте
                RevitTaskDialog.Show("Готово", $"Размещено: {sgruppirovannyeSimvoly.Sum(g => g.Value.Count)} типоразмеров. Ошибки сохранены в: {logPath}"); // Показываем итоговое сообщение
            }
            else // Если ошибок не было
            {
                RevitTaskDialog.Show("Готово", $"Размещено: {sgruppirovannyeSimvoly.Sum(g => g.Value.Count)} типоразмеров. Ошибок нет."); // Показываем сообщение об успешном завершении без ошибок
            }

            return Result.Succeeded; // Возвращаем  "успешно"
        }

        //  ЗАГРУЗКАА СЕМЕЙСТВ ИЗ ФАЙЛОВ
        private void LoadFamilies(string[] faylySemeystv, bool gruppirovatPoPapke)
        {
            using (Transaction tranzaktsiya = new Transaction(dokument, "Загрузка семейств")) // Создаем транзакцию для внесения изменений в документ
            {
                tranzaktsiya.Start(); // Начинаем транзакцию

                foreach (var fayl in faylySemeystv) // Проходим по каждому файлу семейства
                {
                    if (!dokument.LoadFamily(fayl, out Family semeystvo)) // Пытаемся загрузить семейство из файла
                        continue; // Если загрузка не удалась, переходим к следующему файлу

                    // ОПРЕДЕЛЯЕМ КЛЮЧ ДЛЯ ГРУППИРОВКИ
                    string klyuchGruppy = gruppirovatPoPapke // Если выбран способ группировки по папке
                        ? Path.GetDirectoryName(fayl) // Ключом будет имя директории
                        : semeystvo.FamilyCategory?.Name ?? "Без категории"; // Иначе ключом будет имя категории семейства

                    // ДОБАВЛЯЕМ НОВУЮ ГРУППУ, ЕСЛИ ЕЕ НЕТ
                    if (!sgruppirovannyeSimvoly.ContainsKey(klyuchGruppy)) // Если в словаре еще нет такой группы
                        sgruppirovannyeSimvoly.Add(klyuchGruppy, new List<(FamilySymbol Symbol, string Path)>()); // Создаем новую запись с пустым списком

                    // ПОЛУЧАЕМ И ДОБАВЛЯЕМ ВСЕ ТИПОРАЗМЕРЫ (СИМВОЛЫ) СЕМЕЙСТВА В ГРУППУ
                    foreach (ElementId id in semeystvo.GetFamilySymbolIds()) // Получаем все ID типоразмеров в семействе
                    {
                        if (dokument.GetElement(id) is FamilySymbol simvol) // Получаем элемент по ID и проверяем, что это типоразмер
                        {
                            sgruppirovannyeSimvoly[klyuchGruppy].Add((simvol, fayl)); // Добавляем типоразмер и путь к файлу в соответствующую группу
                        }
                    }
                }

                tranzaktsiya.Commit(); // Подтверждаем все изменения, сделанные в транзакции
            }

            // СОРТИРУЕМ СЕМЕЙСТВА ВНУТРИ КАЖДОЙ ГРУППЫ
            foreach (var gruppa in sgruppirovannyeSimvoly) // Проходим по каждой группе в словаре
            {
                gruppa.Value.Sort((a, b) => string.Compare(a.Symbol.Family.Name, b.Symbol.Family.Name, StringComparison.Ordinal)); // Сортируем список типоразмеров по имени семейства
            }
        }

        // РАЗМЕЩЕНИЯ СЕМЕЙСТВ НА ВИДЕ
        private void PlaceFamilies()
        {
            using (Transaction tranzaktsiya = new Transaction(dokument, "Размещение семейств")) // Создаем новую транзакцию для размещения
            {
                tranzaktsiya.Start(); // Начинаем ее

                try // Оборачиваем в try-catch для отлова ошибок
                {
                    TextNoteType tipTekstovoyZametki = GetTextNoteType(); // Получаем стандартный тип текстовой заметки для заголовков

                    // ПРОХОДИМ ПО КАЖДОЙ ГРУППЕ (ОТСОРТИРОВАННОЙ ПО ИМЕНИ)
                    foreach (var gruppa in sgruppirovannyeSimvoly.OrderBy(g => g.Key))
                    {
                        if (!string.IsNullOrEmpty(gruppa.Key)) // Если у группы есть имя
                        {
                            CreateGroupHeader(gruppa.Key, tipTekstovoyZametki); // Создаем текстовый заголовок для группы
                            koordinatY += 4.0; // Смещаемся вниз по оси Y для размещения следующей строки
                            maxVisotaVStroke = 0; // Сбрасываем максимальную высоту для новой строки
                        }

                        // ПРОХОДИМ ПО КАЖДОМУ ТИПОРАЗМЕРУ В ГРУППЕ
                        foreach (var (simvol, put) in gruppa.Value)
                        {
                            try // Внутренний try-catch для отлова ошибок размещения одного семейства
                            {
                                if (!simvol.IsActive) // Если типоразмер не активен
                                    simvol.Activate(); // Активируем его

                                // ОПРЕДЕЛЯЕМ ГАБАРИТЫ СЕМЕЙСТВА
                                var gabarit = simvol.get_BoundingBox(null); // Получаем габаритный контейнер (BoundingBox)
                                double shirina = gabarit?.Max.X - gabarit?.Min.X ?? 5.0; // Вычисляем ширину, если габаритов нет - ставим 5.0
                                double vysota = gabarit?.Max.Y - gabarit?.Min.Y ?? 3.0; // Вычисляем высоту, если габаритов нет - ставим 3.0

                                // РАЗМЕЩАЕМ ЭКЗЕМПЛЯР
                                XYZ pozitsiya = new XYZ(koordinatX, koordinatY, 0); // Определяем координаты для размещения
                                var ekzemplyar = PlaceFamilyInstance(simvol, pozitsiya); // Вызываем метод для размещения экземпляра

                                // СОЗДАЕМ ПОДПИСЬ
                                CreateTextNote(ekzemplyar, simvol, tipTekstovoyZametki, pozitsiya, vysota); // Создаем текстовую подпись для семейства

                                // ЛОГИРУЕМ ДЕЙСТВИЕ
                                logOshibok.Add($"Размещено: {simvol.Family.Name} ({simvol.Name}) X={koordinatX} Y={koordinatY} bboxNull={gabarit==null}"); // Добавляем запись в лог

                                // ОБНОВЛЯЕМ КООРДИНАТЫ ДЛЯ СЛЕДУЮЩЕГО СЕМЕЙСТВА
                                koordinatX += shirina + otstup; // Смещаемся вправо на ширину семейства плюс отступ
                                if (koordinatX > kolichStolbtsov * (shirina + otstup)) // Если вышли за пределы количества столбцов
                                {
                                    koordinatX = 0; // Возвращаемся в начало строки (X=0)
                                    koordinatY += maxVisotaVStroke + otstup; // Смещаемся вниз на высоту самой большой ячейки в строке плюс отступ
                                    maxVisotaVStroke = 0; // Сбрасываем максимальную высоту для новой строки
                                }

                                // ОБНОВЛЯЕМ МАКСИМАЛЬНУЮ ВЫСОТУ
                                if (vysota > maxVisotaVStroke) // Если текущее семейство выше максимального в этой строке
                                    maxVisotaVStroke = vysota; // Обновляем максимальную высоту
                            }
                            catch (Exception isklyuchenie) // Если при размещении одного семейства произошла ошибка
                            {
                                logOshibok.Add($"{simvol.Family.Name}: {isklyuchenie.Message}"); // Записываем ошибку в лог
                            }
                        }
                    }

                    tranzaktsiya.Commit(); // Подтверждаем все размещения
                }
                catch (Exception isklyuchenie) // Если произошла глобальная ошибка в транзакции
                {
                    tranzaktsiya.RollBack(); // Откатываем все изменения
                    logOshibok.Add($"Ошибка транзакции: {isklyuchenie.Message}"); // Записываем ошибку в лог
                    throw; // Пробрасываем ошибку дальше
                }
            }
        }
        
        //  ПОЛУЧЕНИЯ СТАНДАРТНОГО ТИПА ТЕКСТА
        private TextNoteType GetTextNoteType()
        {
            return new FilteredElementCollector(dokument) // Создаем коллектор
                .OfClass(typeof(TextNoteType)) // Ищем типы текстовых заметок
                .Cast<TextNoteType>() // Преобразуем
                .FirstOrDefault(); // Возвращаем первый найденный
        }

        // СОЗДАНИЯ ЗАГОЛОВКА ГРУППЫ
        private void CreateGroupHeader(string imyaGruppy, TextNoteType tipTekstovoyZametki)
        {
            XYZ pozitZagolovka = new XYZ(koordinatX, koordinatY, 0); // Определяем позицию заголовка
            TextNote.Create(dokument, aktivnVid.Id, pozitZagolovka, imyaGruppy, tipTekstovoyZametki.Id); // Создаем текстовую заметку
        }

        // МЕТОД РАЗМЕЩЕНИЯ ОДНОГО  СЕМЕЙСТВА
        private FamilyInstance PlaceFamilyInstance(FamilySymbol simvol, XYZ pozitsiya)
        {
            if (uroven != null) // Если уровень найден
                return dokument.Create.NewFamilyInstance(pozitsiya, simvol, uroven, StructuralType.NonStructural); // Размещаем на уровне
            else // Если уровень не найден
                return dokument.Create.NewFamilyInstance(pozitsiya, simvol, StructuralType.NonStructural); // Размещаем без привязки к уровню
        }

        //  СОЗДАНИЯ ПОДПИСИ ДЛЯ СЕМЕЙСТВА
        private void CreateTextNote(FamilyInstance ekzemplyar, FamilySymbol simvol, TextNoteType tipTekstovoyZametki, XYZ pozitsiya, double vysota)
        {
            // ФОРМИРУЕМ ТЕКСТ ПОДПИСИ
            string imySemeystva = simvol.Family.Name; // Получаем имя семейства
            string imyTipa = simvol.Name; // Получаем имя типоразмера
            string textZametki = (imySemeystva.Length + imyTipa.Length + 2 > 30) // Если общая длина текста больше 30 символов
                ? $"{imySemeystva}:\n{imyTipa}" // Раздел/яем имя семейства и типоразмера переносом строки
                : $"{imySemeystva}: {imyTipa}"; // Иначе,  оставляем в одну строку
            
            // ОПРЕДЕЛЯЕМ ПОЗИЦИЮ ПОДПИСИ
            XYZ pozitZametki = new XYZ(pozitsiya.X, pozitsiya.Y - vysota / 2 - 1.0, 0); // Рассчитываем позицию под семейством

            // РАССЧИТЫВАЕМ РАЗМЕР ТЕКСТА
            double razmerText = Math.Max(0.0015, Math.Min(vysota / 3, 0.0035)); // Размер текста пропорционален высоте семейства (от 0.5 до 1.1 мм)
            TextNoteType tipZametki = GetOrCreateTextNoteType(razmerText); // Получаем или создаем нужный тип текста
            
            // СОЗДАЕМ ТЕКСТОВУЮ ЗАМЕТКУ
            TextNote.Create(dokument, aktivnVid.Id, pozitZametki, textZametki, tipZametki.Id); // Создаем подпись
        }

        // МЕТОД ПОЛУЧЕНИЯ ИЛИ СОЗДАНИЯ СПЕЦИАЛЬНОГО ТИПА ТЕКСТА
        private TextNoteType GetOrCreateTextNoteType(double razmerText)
        {
            // ИЩЕМ СУЩЕСТВУЮЩИЙ ТИП С НУЖНЫМИ ПАРАМЕТРАМИ
            var ygeSushestvuet = new FilteredElementCollector(dokument) // Создаем коллектор
                .OfClass(typeof(TextNoteType)) // Ищем типы текста
                .Cast<TextNoteType>() // Преобразуем
                .FirstOrDefault(t => // Ищем первый, который удовлетворяет условиям
                    Math.Abs(t.get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble() - razmerText) < 0.001 && // Сравниваем размер текста (с погрешностью)
                    t.get_Parameter(BuiltInParameter.TEXT_FONT).AsString() == "Arial" && // Проверяем, что шрифт - Arial
                    t.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).AsInteger() == 1 // Проверяем, что фон прозрачный (1)
                );
            if (ygeSushestvuet != null) // Если такой тип уже существует
                return ygeSushestvuet; // Возвращаем его
            
            // ЕСЛИ ТИП НЕ НАЙДЕН, СОЗДАЕМ НОВЫЙ
            var bazTip = new FilteredElementCollector(dokument) // Ищем любой существующий тип текста
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .First(); // Берем первый попавшийся как основу
            var newTip = bazTip.Duplicate($"Arial_Small_Transparent") as TextNoteType; // Дублируем его с новым именем
            newTip.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(razmerText); // Устанавливаем наш размер текста
            newTip.get_Parameter(BuiltInParameter.TEXT_FONT).Set("Arial"); // Устанавливаем шрифт Arial
            newTip.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1); // Устанавливаем прозрачный фон
            return newTip; // Возвращаем созданный тип
        }
    }

    // КЛАСС ДЛЯ ДИАЛОГОВОГО ОКНА ВЫБОРА ГРУППИРОВКИ
    public class GroupingDialog : Form // Наследуемся от стандартной формы Windows
    {
        public bool GroupByDirectory { get; private set; } = true; // Свойство, которое будет хранить выбор пользователя (по умолчанию - по папке)

        // КОНСТРУКТОР ОКНА
        public GroupingDialog()
        {
            // НАСТРОЙКИ ФОРМЫ
            this.Text = "Выберите способ группировки"; // Заголовок окна
            this.Size = new System.Drawing.Size(420, 300); // Размер окна
            this.FormBorderStyle = FormBorderStyle.FixedDialog; // Стиль рамки
            this.StartPosition = FormStartPosition.CenterScreen; // по центру экрана
            this.MaximizeBox = false; // Убираем кнопку "Развернуть"
            this.MinimizeBox = false; // Убираем кнопку "Свернуть"
            this.TopMost = true; // поверх всех остальных

            // ЭЛЕМЕНТЫ УПРАВЛЕНИЯ
            var metka = new Label // Текстовая метка с вопросом
            {
                Text = "Как группировать семейства?", // Текст
                Location = new System.Drawing.Point(20, 20), // Положение
                AutoSize = true // Автоматический размер
            };

            var knopkaPoPapke = new RadioButton // Кнопка выбора "По структуре папок"
            {
                Text = "По структуре папок", // Текст
                Checked = true, // Выбрана по умолчанию
                Location = new System.Drawing.Point(20, 50), // Положение
                AutoSize = true // Автоматический размер
            };

            var knopkaPoImeni = new RadioButton // Кнопка выбора "По имени семейства"
            {
                Text = "По имени семейства", // Текст
                Location = new System.Drawing.Point(20, 80), // Положение
                AutoSize = true // Автоматический размер
            };

            var knopkaOk = new Button // Кнопка "ОК"
            {
                Text = "OK", // Текст
                DialogResult = DialogResult.OK, // Результат диалога при нажатии
                Location = new System.Drawing.Point(100, 150), // Положение
                Size = new System.Drawing.Size(100, 30) // Размер
            };

            var knopkaOtmeny = new Button // Кнопка "Отмена"
            {
                Text = "Отмена", // Текст
                DialogResult = DialogResult.Cancel, // Результат диалога при нажатии
                Location = new System.Drawing.Point(220, 150), // Положение
                Size = new System.Drawing.Size(100, 30) // Размер
            };

            // НАЗНАЧАЕМ ОБРАБОТЧИКИ СОБЫТИЙ
            knopkaPoPapke.CheckedChanged += (s, e) => GroupByDirectory = knopkaPoPapke.Checked; // При изменении состояния кнопки обновляем свойство
            knopkaPoImeni.CheckedChanged += (s, e) => GroupByDirectory = !knopkaPoImeni.Checked; // При изменении состояния второй кнопки также обновляем свойство

            // УСТАНАВЛИВАЕМ КНОПКИ ПО УМОЛЧАНИЮ
            this.AcceptButton = knopkaOk; // Кнопка "ОК" будет нажиматься по Enter
            this.CancelButton = knopkaOtmeny; // Кнопка "Отмена" будет нажиматься по Escape

            // ДОБАВЛЯЕМ ЭЛЕМЕНТЫ НА ФОРМУ
            this.Controls.Add(metka); // Добавляем метку
            this.Controls.Add(knopkaPoPapke); // Добавляем первую радиокнопку
            this.Controls.Add(knopkaPoImeni); // Добавляем вторую радиокнопку
            this.Controls.Add(knopkaOk); // Добавляем кнопку ОК
            this.Controls.Add(knopkaOtmeny); // Добавляем кнопку Отмена
        }
    }
}
