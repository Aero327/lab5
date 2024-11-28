// Разработать консольное приложение с дружественным интерфейсом с возможностью выбора
// заданий. Приложение должно выполнять следующие функции:
// 1. Чтение базы данных из excel файла.
// 2. Просмотр базы данных.
// 3. Удаление элементов (по ключу).
// 4. Корректировка элементов (по ключу).
// 5. Добавление элементов.
// 6. Реализация 4 запросов (формулировки запросов придумать самостоятельно и отразить в
// отчёте, можно использовать запрос, данный в примере):
// 1. 1 запрос с обращением к одной таблице
// 2. 1 запрос с обращением к двум таблицам
// 3. 2 запроса с обращением к трем таблицам
// 2 запроса должны возвращать перечень, 2 запроса одно значение.
// 7. Во время всего сеанса работы ведется полное протоколирование действий в текстовом
// файле (в начале сеанса запросить, будет ли это новый файл или дописывать в уже
// существующий).

using System.Text;

namespace lab5;

class Program
{
    public static void Main(string[] args)
    {
        const string path = "LR5-var13.xls";
        var sheetNames = Repository.GetSheetNames(path);
        
        Console.OutputEncoding = Encoding.UTF8;
        Console.InputEncoding = Encoding.UTF8;
        
        var input = "";
        int choice;

        while (input != "1" && input != "2")
        {
            Console.WriteLine("Настройки протоколирования.\n" +
                              "Записывать действия в новый текстовый файл или уже существующий? (1 - новый, 2 - существующий): ");
            
            input = Console.ReadLine();
        }
        
        var add = input == "2";
        var logger = new Log("log.txt", add);
        logger.Write("Инициализация логгера");
        
        var repository = new Repository(logger);
        repository.LoadData(path);
        logger.Write($"Данные из файла {path} успешно загружены");

        do
        {
            Console.WriteLine("""
            ===========

            Функционал приложения:
            1. Просмотр базы данных.
            2. Удаление элементов (по ключу).
            3. Корректировка элементов (по ключу).
            4. Добавление элементов.
            5. Найти количество начислений на сумму более 1.
            6. Определить держателя счёта с максимальным начислением в рублях за 25 декабря 2021 года.
            7. Найти держателя счета с минимальным количеством операций, совершенных после 27 декабря
             2021 года, и валюту, в которой было проведено больше всего его транзакций.
            8. Вывести полные данные об операции, где сумма операции больше среднего значения по всем
             операциям и меньше всех остальных
            0. Выход из приложения

            ===========
            """);

            do
            {
                Console.Write("Выберите действие: ");
                input = Console.ReadLine();
            } while (!(int.TryParse(input, out var n) && n >= 0 && n <= 8));
            logger.Write($"Пользователь выбрал действие \"{input}\"");
            
            switch (input)
            {
                case "1":
                    repository.PrintData();
                    break;
                
                case "2":
                    Console.WriteLine("Введите имя листа, в котором вы хотите удалить элемент: ");
                    var sheetName = Console.ReadLine();
                    logger.Write($"Получено имя листа: {sheetName}");

                    while (!sheetNames.Contains(sheetName))
                    {
                        Console.WriteLine("Неверное имя листа. Введите правильное имя листа: ");
                        logger.Write("Введено неверное имя листа. Запрошен повторный ввод");
                        sheetName = Console.ReadLine();
                    }

                    Console.WriteLine("Введите номер строки, которую хотите удалить: ");
                    input = Console.ReadLine();
                    logger.Write($"Получен номер строки: {input}");
                    
                    int id;
                    while (!int.TryParse(input, out id) || id < 0)
                    {
                        Console.WriteLine("Неверный номер. Введите правильный номер: ");
                        logger.Write("Введен неверный номер. Запрошен повторный ввод");
                        input = Console.ReadLine();
                    }

                    repository.DeleteRowById(path, sheetName!, id);
                    break;
                
                case "3":
                    try
                    {
                        Console.WriteLine("Введите имя листа, в котором вы хотите обновить значение: ");
                        sheetName = Console.ReadLine();
                        logger.Write($"Получено имя листа: {sheetName}");

                        while (!sheetNames.Contains(sheetName))
                        {
                            Console.WriteLine("Неверное имя листа. Введите правильное имя листа: ");
                            sheetName = Console.ReadLine();
                            logger.Write("Введено неверное имя листа. Запрошен повторный ввод");
                        }

                        Console.WriteLine("Введите номер строки, которую хотите обновить: ");
                        input = Console.ReadLine();
                        logger.Write($"Получен номер строки: {input}");
                        
                        while (!int.TryParse(input, out id) || id < 0)
                        {
                            Console.WriteLine("Неверный номер строки. Введите положительное целое число: ");
                            logger.Write("Введен неверный номер. Запрошен повторный ввод");
                            input = Console.ReadLine();
                        }

                        var cols = Repository.GetColumnNumber(path, sheetName!);
                        var values = new List<string> { input };
                        Console.WriteLine("Ввод новых данных.");
                        logger.Write("Ввод новых данных.");
                        for (var i = 1; i < cols; i++)
                        {
                            Console.WriteLine($"Введите значение для столбца {i + 1}: ");
                            input = Console.ReadLine();
                            values.Add(input);
                            logger.Write($"Введено значение \"{input}\"");
                        }
                        repository.UpdateRowById(path, sheetName!, id, values);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Ошибка при обновлении строки: {e.Message}");
                        logger.Write($"Ошибка при обновлении строки: {e.Message}");
                    }
                    break;
                
                case "4":
                    try
                    {
                        Console.WriteLine("Введите имя листа, в который хотите добавить данные: ");
                        sheetName = Console.ReadLine();
                        logger.Write($"Получено имя листа: {sheetName}");

                        while (!sheetNames.Contains(sheetName))
                        {
                            Console.WriteLine("Неверное имя листа. Введите правильное имя листа: ");
                            sheetName = Console.ReadLine();
                            logger.Write("Введено неверное имя листа. Запрошен повторный ввод");
                        }

                        var cols = Repository.GetColumnNumber(path, sheetName!);
                        var values = new List<string>();

                        Console.WriteLine("Ввод новых данных.");
                        logger.Write("Ввод новых данных.");
                        for (var i = 1; i < cols; i++)
                        {
                            Console.WriteLine($"Введите значение для столбца {i + 1}: ");
                            input = Console.ReadLine();
                            values.Add(input);
                            logger.Write($"Введено значение \"{input}\"");
                        }
                        repository.InsertRow(path, sheetName!, values);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Ошибка при добавлении строки: {e.Message}");
                        logger.Write($"Ошибка при добавлении строки: {e.Message}");
                    }
                    break;
                
                case "5":
                    var result1 = repository.Query1();
                    Console.WriteLine($"Результат: {result1}");
                    break;
                
                case "6":
                    var result2 = repository.Query2();
                    Console.WriteLine($"Результат: {result2}");
                    break;
                
                case "7":
                    var result3 = repository.Query3();
                    Console.WriteLine($"Результат: \n{string.Join('\t', result3)}");
                    break;
                
                case "8":
                    //
                    var result4 = repository.Query4();
                    Console.WriteLine($"Результат: \n{string.Join('\t', result4)}");
                    break;
            }
        } while (input != "0");
    }
}