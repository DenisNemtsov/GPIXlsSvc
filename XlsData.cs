using System;
using Serilog;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Service
{
    /// <summary>
    /// Загрузка и хранение данных полученных из Excel-файлов
    /// </summary>
    internal class XlsData
    {
        public string TimeStamp { get; set; }
        public double FreeVolumeOfSK { get; set; }
        public double FreeVolumeOfDGKL { get; set; }
        public double FreeVolumeOfDT { get; set; }
        public double RealesationOfSK { get; set; }
        public double RealesationOfDGKL { get; set; }
        public double RealesationOfDT { get; set; }
        public List<string> Product { get; set; }
        public List<string> Park { get; set; }
        public List<double> Tank { get; set; }
        public List<double> Full { get; set; }
        public List<double> Naliv { get; set; }
        public List<double> Oper { get; set; }
        public List<double> Dead { get; set; }
        public List<string> Direct { get; set; }
        
        /// <summary>
        /// Конструктор класса
        /// </summary>
        public XlsData()
        {

        }

        /// <summary>
        /// Конструктор класса
        /// </summary>
        public XlsData(string FileName)
        {
            Product = new List<string>();
            Park = new List<string>();
            Tank = new List<double>();
            Full = new List<double>();
            Naliv = new List<double>();
            Oper = new List<double>();
            Dead = new List<double>();
            Direct = new List<string>();
            if (FileName != "")
                LoadDataFromXlsFile(FileName);
        }

        /// <summary>
        /// Загрузка данных из Excel-файла
        /// </summary>
        /// <param name="FileName">Полный путь к Excel-файлу с данными.</param>
        private void LoadDataFromXlsFile(string FileName)
        {
            int rowCount = 1;                                                                                                                         // Переменная для хранения счетчика пройденных строк
            FileInfo xlsFile = new FileInfo(FileName);

            try
            {
                using (ExcelPackage xlsPackage = new ExcelPackage(xlsFile))
                {
                    ExcelWorksheet xlsWorksheet = xlsPackage.Workbook.Worksheets[0];
                    while (xlsWorksheet.Cells[rowCount, 1].Text != "1")                                                                               // Поиск строки на которой заканчивается "шапка" таблицы
                        rowCount++;
                    rowCount++;                                                                                                                       // Переход на следующую после "шапки" строку
                    int rowStart = rowCount;
                    while (xlsWorksheet.Cells[rowCount, 3].Value != null)                                                                             // Перебор всех непустых строк в 3-ем столбце
                    {
                        if (xlsWorksheet.Cells[rowCount, 3].Value is System.String)                                                                   // Поиск строки на которой заканчивается блок данных ("Итого:")
                        {
                            int rowEnd = rowCount;
                            for (int k = rowStart; k != rowEnd; k++)                                                                                  // Циклическое заполнение структуры текущим блоком данных
                            {
                                this.Product.Add((string)xlsWorksheet.Cells[rowStart, 1].Value);                                                      // Чтение данных из столбца "Наименование продукта" (1)
                                this.Park.Add((string)xlsWorksheet.Cells[k, 2].Value);                                                                // Чтение данных из столбца "Наименование парка" (2)    
                                this.Tank.Add((double)xlsWorksheet.Cells[k, 3].Value);                                                                // Чтение данных из столбца "Наименование резервуара" (3)
                                this.Full.Add((double)xlsWorksheet.Cells[k, 4].Value);                                                                // Чтение данных из столбца "Полный товарный объем" (4)
                                this.Naliv.Add((double)xlsWorksheet.Cells[k, 5].Value);                                                               // Чтение данных из столбца "Масса налитого продукта" (5)
                                this.Oper.Add((double)xlsWorksheet.Cells[k, 6].Value);                                                                // Чтение данных из столбца "Оперативный учет" (6)    
                                this.Dead.Add((double)xlsWorksheet.Cells[k, 7].Value);                                                                // Чтение данных из столбца "Технологические остатки" (7)
                                this.Direct.Add((string)xlsWorksheet.Cells[k, 8].Value);                                                              // Чтение данных из столбца "Состояние резервуара" (8)
                            }
                            switch ((string)xlsWorksheet.Cells[rowStart, 1].Value)                                                                    // Чтение значений "Реализация за прошедшие сутки (тн.)" и "Свободный объём для заполнения (тн.)"
                            {
                                case "Конденсат":
                                    this.FreeVolumeOfSK = (double)xlsWorksheet.Cells[rowEnd + 1, 5].Value;
                                    this.RealesationOfSK = (double)xlsWorksheet.Cells[rowEnd + 2, 5].Value;
                                break;
                                case "Дистиллят":
                                    this.FreeVolumeOfDGKL = (double)xlsWorksheet.Cells[rowEnd + 1, 5].Value;
                                    this.RealesationOfDGKL = (double)xlsWorksheet.Cells[rowEnd + 2, 5].Value;
                                break;
                                case "ТД":
                                    this.FreeVolumeOfDT = (double)xlsWorksheet.Cells[rowEnd + 1, 5].Value;
                                    this.RealesationOfDT = (double)xlsWorksheet.Cells[rowEnd + 2, 5].Value;
                                break;
                            }
                            rowCount = rowCount + 3;                                                                                                  // Переход к следующему блоку данных
                            rowStart = rowCount;                                                                                                      // Сохранение начала следующего блока данных
                        }
                        rowCount++;                                                                                                                   // Инкрементация счетчика обработанных строк
                    }
                    rowCount = 1;
                    string TimeStamp = (string)xlsWorksheet.Cells[3, 2].Value;                                                                        // Чтение строки из файла для получения временной метки
                    Regex Regex = new Regex(@"(\d{1,2}\.\d{1,2}\.\d{2,4})");                                                                          // Регулярное выражение для выделения даты
                    Match Match = Regex.Match(TimeStamp);
                    string Date = Match.Value;                                                                                                        // Сохранение даты в переменную                                                                                        
                    
                    switch (Program.IniFile["Service"]["SaveTime"].StringValue)                                                                       // Обработка настройки времени сохранения данных в тэги
                    {
                        case "Now()":                                                                                                                 // Запись данных с текущим временем
                            string Time = DateTime.Now.ToShortTimeString();
                            this.TimeStamp = Date + " " + Time;                                                                                       // Запись временной метки в массив данных
                            break;
                        case "GetFromFile()":                                                                                                         // Запись данных с временем взятым из шапки файла
                            Regex = new Regex(@"(\d{1,2}:\d{1,2})");                                                                                  // Регулярное выражение для выделения времени
                            Match = Regex.Match(TimeStamp);
                            Time = Match.Value;                                                                                                       // Сохранение времени в переменную
                            DateTime tempTime = DateTime.Parse(Time);                                                                                 // Преобразуем строковое значение времени в DateTime-переменную
                            tempTime += new TimeSpan(2, 0, 0);                                                                                        // Прибавляем 2 часа (По просьбам трудящихся)
                            break;
                        default:                                                                                                                      // Запись данных со временем указанным в файле настроек
                            this.TimeStamp = Date + " " + Program.IniFile["Service"]["SaveTime"].StringValue;                                         // Запись временной метки в массив данных
                            break;
                    }
                }
            }
            catch (Exception Ex)
            {
                Log.Error(Ex, "Ошибка загрузки данных из Excel-файла");
            }
        }
    }
}