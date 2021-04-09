using System;
using Serilog;
using Topshelf;
using System.IO;
using System.Timers;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Service
{   
    /// <summary>
    /// Создание и настройка службы Windows
    /// </summary>
    public class GPIXlsSvc : ServiceControl
    {
        // private readonly Mode TagsWriteMode = Mode.Off;
        private string WriteMode { get; }
        private Timer Timer { get; set; }

        /// <summary>
        /// Конструктор класса
        /// </summary>
        public GPIXlsSvc()
        {
            WriteMode = Program.IniFile["Service"]["WriteMode"].StringValue;
            this.Timer = new Timer() { AutoReset = true, Interval = TimeSpan.FromSeconds(double.Parse(Program.IniFile["Service"]["Interval"].StringValue)).TotalMilliseconds };
            this.Timer.Elapsed += TimerCallback;
        }

        /// <summary>
        /// Событие: Запуск службы
        /// </summary>
        /// <param name="hostControl">Ссылка на экземпляр объекта службы.</param>
        public bool Start(HostControl hostControl)
        {
            Log.Information("Попытка запуска службы...");
            this.Timer.Start();
            Log.Information("Служба успешно запущена!");

            return true;
        }

        /// <summary>
        /// Событие: Остановка службы
        /// </summary>
        /// <param name="hostControl">Ссылка на экземпляр объекта службы.</param>
        public bool Stop(HostControl hostControl)
        {
            Log.Information("Попытка остановки службы...");
            this.Timer.Stop();
            Log.Information("Служба успешно остановлена!");

            return true;
        }

        /// <summary>
        /// Событие: Пауза службы
        /// </summary>
        /// <param name="hostControl">Ссылка на экземпляр объекта службы.</param>
        public bool Pause(HostControl hostControl)
        {
            Log.Information("Paused");

            return true;
        }

        /// <summary>
        /// Событие: Возобновление службы
        /// </summary>
        /// <param name="hostControl">Ссылка на экземпляр объекта службы.</param>
        public bool Continue(HostControl hostControl)
        {
            Log.Information("Continued");

            return true;
        }

        /// <summary>
        /// Событие: Сработка таймера
        /// </summary>
        /// <param name="Source">.</param>
        /// <param name="E">.</param>
        private void TimerCallback(object Source, ElapsedEventArgs E)
        {
            Log.Information("Начало обработки Excel-файлов...");
            if (this.Timer.Enabled)
                this.Timer.Stop();
            Log.Information("Остановка таймера службы...");
            // _ = new List<string>();
            try
            {
                Log.Information("Генерация списка обрабатываемых файлов...");
                List<string> fileNames = Files(new DirectoryInfo(Program.IniFile["Service"]["Folder"].StringValue), "*" + Program.IniFile["Service"]["Filter"].StringValue + "*", true);
                if (fileNames.Count > 0)
                {
                    Log.Information("Количество файлов для обработки: {Count}", fileNames.Count.ToString());
                    foreach (string fileName in fileNames)
                    {
                        Log.Information("Текущий файл: " + fileName.ToString());
                        XlsData xlsData = new XlsData(fileName);
                        Log.Information("Загрузка данных из Excel-файла...");
                        int Tank = 0;
                        string Tag = string.Empty;
                        string Park = string.Empty;
                        if (xlsData.Oper.Count > 0)
                        {
                            Log.Information("Загружено строк данных: {Count}", xlsData.Oper.Count.ToString());
                            for (int count = 0; count < xlsData.Oper.Count; count++)
                            {
                                Park = xlsData.Park[count].Split(' ')[1];
                                Log.Information("Запись данных в PI Server, строка №{Count}", count.ToString());
                                if (!String.IsNullOrEmpty(xlsData.TimeStamp))
                                {
                                    Log.Information("Дата создания отчета: {Time}", xlsData.TimeStamp);
                                    if (!String.IsNullOrEmpty(Park))
                                    {
                                        Log.Information("Запись данных парка: Park-{Park}", Park);
                                        Tank = Convert.ToInt32(xlsData.Tank[count]);
                                        if (Tank != 0)
                                        {
                                            Log.Information("Запись данных резервуара: R-{Tank}", Tank.ToString());
                                            Tag = TagBuilder("Product", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Product[count]);
                                            Tag = TagBuilder("Full", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Full[count].ToString());
                                            Tag = TagBuilder("Naliv", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Naliv[count].ToString());
                                            Tag = TagBuilder("Oper", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Oper[count].ToString());
                                            Tag = TagBuilder("Dead", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Dead[count].ToString());
                                            Tag = TagBuilder("Direct", Park, Tank);
                                            Write(Tag, xlsData.TimeStamp, xlsData.Direct[count].ToString());
                                        }
                                    }
                                }
                                else
                                {
                                    Log.Error("Отсутствует дата создания отчета, дальнейшая обработка невозможна!");
                                    break;
                                }
                            }
                        }
                        Log.Information("Запись дополнительных значений...");
                        Tag = Program.IniFile["Tags"]["FreeVolumeSK"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.FreeVolumeOfSK));
                        Tag = Program.IniFile["Tags"]["FreeVolumeDGKL"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.FreeVolumeOfDGKL));
                        Tag = Program.IniFile["Tags"]["FreeVolumeDT"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.FreeVolumeOfDT));
                        Tag = Program.IniFile["Tags"]["RealesationSK"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.RealesationOfSK));
                        Tag = Program.IniFile["Tags"]["RealesationDGKL"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.RealesationOfDGKL));
                        Tag = Program.IniFile["Tags"]["RealesationDT"].StringValue;
                        Write(Tag, xlsData.TimeStamp, Convert.ToString(xlsData.RealesationOfDT));
                        Log.Information("Перенос файла: {Path} в папку 'Processed'", Path.GetFileName(fileName));
                        if (!Directory.Exists(Program.IniFile["Service"]["Folder"].StringValue + @"\Processed\"))
                            Directory.CreateDirectory(Program.IniFile["Service"]["Folder"].StringValue + @"\Processed\");
                        File.Move(fileName, Program.IniFile["Service"]["Folder"].StringValue + @"\Processed\" + Path.GetFileName(fileName));
                    }
                    Log.Information("Завершение обработки файлов...");
                }
                else
                {
                    Log.Information("Файлы для обработки отсутствуют!");
                }
            }
            catch (System.Exception Ex)
            {
                Log.Error(Ex, "Ошибка обработки файлов!");
            }
            Log.Information("Запуск таймера службы...");
            this.Timer.Start();
        }

        /// <summary>
        /// Генерация тэга по маске.
        /// </summary>
        /// <param name="Name">Имя тэга в конфигурационном файле.</param>
        /// <param name="Park">Номер парка.</param>
        /// <param name="Tank">Номер резервуара.</param>
        private string TagBuilder(string Name, string Park, int Tank)
        {
            string tag = Program.IniFile["Tags"][Name].StringValue;
            tag = Regex.Replace(tag, @"\x2D\x2A\x2E", "-" + Park + ".");
            tag = Regex.Replace(tag, @"R\x2A\x3A", "R" + Tank.ToString() + ":");
            return tag;
        }

        /// <summary>
        /// Запись значения в тэг на временную метку.
        /// </summary>
        /// <param name="Tag">Тэг.</param>
        /// <param name="Timestamp">Временная метка.</param>
        /// <param name="Value">Значение.</param>
        private void Write(string Tag, string Timestamp, string Value)
        {
            
            if (WriteMode == "On" || WriteMode == "on" || WriteMode == "ON")
            {
                Log.Information("Запись значения в тэг: {Tag}", Tag);
                try
                {
                    PISDK.PISDK PiSDK = new PISDK.PISDK();
                    PISDK.Server PiServer = PiSDK.Servers[Program.IniFile["Service"]["PIServer"].StringValue];
                    PISDK.PIPoint PiPoint = PiServer.PIPoints[Tag];
                    DateTime timeStamp = DateTime.Parse(Timestamp);
                    PISDK.PIValue PiValue = PiPoint.Data.Snapshot;
                    if (PiValue.Value.GetType().IsCOMObject)
                    {
                        switch (Value)
                        {
                            case "Конденсат":
                                Value = "СК";
                                break;
                            case "Дистиллят":
                                Value = "ДГКЛ";
                                break;
                            case "ТД":
                                Value = "Топливо дизельное";
                                break;
                        }
                    }
                    PiPoint.Data.UpdateValue(Value, timeStamp, PISDK.DataMergeConstants.dmReplaceDuplicates, null);
                    Log.Information("Записано значение: {Value}", Value);
                }
                catch (System.Exception Ex)
                {
                    Log.Error(Ex, "Ошибка записи значения: {Value} в тэг: {Tag}", Value, Tag);
                }
            }
            else
            {
                Log.Information("Запись отключена, в тэг: {Tag} значение {Value} не записано!", Tag, Value);
            }
        }

        /// <summary>
        /// Рекурсивный поиск файлов внутри базового каталога на основе маски (паттерна)
        /// </summary>
        /// <param name="Directory">Путь к базовой директории, с которой начинается поиск.</param>
        /// <param name="Pattern">Маска для фильтрации имен файлов.</param>
        /// <param name="Recursive">Поддержка рекурсивного вызова процедуры для обработки вложенных директорий.</param>
        private static List<string> Files(DirectoryInfo Directory, string Pattern, bool Recursive)
        {
            List<string> files = new List<string>();

            try
            {
                foreach (FileInfo file in Directory.GetFiles(Pattern))                                                                                // Циклический перебор файлов внутри директории
                {
                    files.Add(file.FullName);                                                                                                         // Добавление пути к файлу в массив, если он подходит по критериям (маске)
                }
                if (Recursive)                                                                                                                        // Проверка флага поддержки рекурсии, если включен, то
                {
                    foreach (DirectoryInfo SubDirectory in Directory.GetDirectories())                                                                // Циклический перебор подпапок внутри папки
                    {
                        Files(SubDirectory, Pattern, Recursive);                                                                                      // Рекурсивный вызов самой себя для добавления пути к файлу в массив
                    }
                }
            }
            catch (Exception Ex)                                                                                                                      // Обработка ошибок в случае их возникновения
            {
                Log.Error(Ex, "Ошибка чтения файлов из локальной папки: ");
            }

            return files;
        }
    }
}