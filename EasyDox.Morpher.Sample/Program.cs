using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace EasyDox.Morpher.Sample
{
    class Program
    {
        static void Main()
        {
            // В данном примере для склонения по падежам используется веб-сервис morpher.ru,
            // но вы можете легко перейти на библиотеку Morpher.Russian.dll.
            // Она платная, заказать ее можно на странице http://morpher.ru/Buy.aspx
            var webService = new global::Morpher.WebService.V2.Client();

            // Если вам не нужны функции склонения, вызовите конструктор Engine без параметров.
            var engine = new Engine(new MorpherFunctionPack(webService, webService));

            var fieldValues = new Dictionary<string, string>
                {
                    {"Номер договора", "333"},
                    {"Город", "Москва"},
                    {"Дата договора", DateTime.Now.ToShortDateString ()},
                    {"Краткое наименование", "ООО «Тюльпан»"},
                    {"Полное наименование", "Общество с ограниченной ответственностью «Тюльпан»"},
                    {"Должность представителя", "Генеральный директор"},
                    {"ФИО представителя", "Иванов Петр Сергеевич"},
                };

            Console.Write("Генерируем договор... ");

            const string outputPath = "dogovor1.docx";

            var errors = engine.Merge("dogovor.docx", fieldValues, outputPath);

            Console.WriteLine("готово.");

            foreach (var error in errors)
            {
                Console.WriteLine(error.Accept(new ErrorToRussianString()));
            }

            // Показываем договор пользователю (если установлен Ворд).
            // Даже если в процессе генерации возникли ошибки (не найдено поле или функция),
            // документ все равно будет создан. Значения полей с ошибками не изменятся.
            Process.Start(outputPath);
        }
    }
}
