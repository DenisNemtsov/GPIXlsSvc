using System;
using System.IO;
using Serilog;
using Topshelf;
using SharpConfig;
using System.Collections.Generic;

namespace Service
{
    class Program
    {
        public static Configuration IniFile;
        
        /// <summary>
        /// Точка входа в программу
        /// </summary>
        static void Main(string[] args)
        {
            var Result = HostFactory.Run(Congiguration =>
            {
                 Log.Logger = new LoggerConfiguration()
                    .WriteTo.Console()
                    .WriteTo.File(@"GPIXlsSvc.log", rollingInterval: RollingInterval.Day, outputTemplate: "{Timestamp:yyyy.MM.dd HH:mm:ss} [{Level:w3}] {Message:lj}{NewLine}{Exception}")
                    .CreateLogger();

                Congiguration.UseSerilog();

                if (File.Exists("GPIXlsSvc.ini"))
                {
                    IniFile = Configuration.LoadFromFile("GPIXlsSvc.ini");
                }
                else
                {
                    Log.Error("Файл настроек GPIXlsSvc.ini не найден,");
                    Log.Error("запуск приложения невозможен!");
                    Environment.Exit(-1);
                }

                Congiguration.Service<GPIXlsSvc>(Service =>
                {
                    Service.ConstructUsing(GPIXlsSvc => new GPIXlsSvc());
                    Service.WhenStarted((GPIXlsSvc, hostControl) => GPIXlsSvc.Start(hostControl));
                    Service.WhenStopped((GPIXlsSvc, hostControl) => GPIXlsSvc.Stop(hostControl));
                });
                
                Congiguration.RunAsLocalSystem();
                Congiguration.StartAutomatically();

                Congiguration.SetDescription("Служба передачи данных получаемых от УМТСиК");
                Congiguration.SetDisplayName("GPIXlsSvc");
                Congiguration.SetServiceName("GPIXlsSvc");
            });

            var exitCode = (int) Convert.ChangeType(Result, Result.GetTypeCode());
            Environment.ExitCode = exitCode;
        }
    }
}
