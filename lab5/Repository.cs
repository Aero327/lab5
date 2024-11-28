using System.Runtime.CompilerServices;
using Aspose.Cells;

namespace lab5;

public class Repository
{
    private List<ExchangeRates> _exchangeRates;
    private List<Receipts> _receipts;
    private List<Accounts> _accounts;
    private Log _logger;

    public Repository(Log logger)
    {
        _exchangeRates = new List<ExchangeRates>();
        _receipts = new List<Receipts>();
        _accounts = new List<Accounts>();
        _logger = logger;
    }
    
    public static string[] GetSheetNames(string fileName)
    {
        var workbook = new Workbook(fileName);
        var sheetCount = workbook.Worksheets.Count;
        
        var sheetNames = new string[sheetCount];

        for (var i = 0; i < sheetCount; i++)
        {
            sheetNames[i] = workbook.Worksheets[i].Name;
        }

        return sheetNames;
    }

    public static int GetColumnNumber(string fileName, string sheetName)
    {
        var workbook = new Workbook(fileName);
        var worksheet = workbook.Worksheets[sheetName];
        return worksheet.Cells.MaxDataColumn + 1;
    }
    
    public void LoadData(string fileName)
    {
        try
        {
            var accountsSheet = new Workbook(fileName).Worksheets[0];
            _accounts = accountsSheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .Select(row => new Accounts(
                    int.Parse(row.GetCellOrNull(0).Value.ToString()),
                    row.GetCellOrNull(1).Value.ToString(),
                    DateTime.Parse(row.GetCellOrNull(2).Value.ToString())
                ))
                .ToList();
            _logger.Write("Получена таблица \"Счета\"");

            var exchangeRatesSheet = new Workbook(fileName).Worksheets[1];
            _exchangeRates = exchangeRatesSheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .Select(row => new ExchangeRates(
                    int.Parse(row.GetCellOrNull(0).Value.ToString()),
                    row.GetCellOrNull(1).Value.ToString(),
                    double.Parse(row.GetCellOrNull(2).Value.ToString()),
                    row.GetCellOrNull(3).Value.ToString()
                ))
                .ToList();
            _logger.Write("Получена таблица \"Курс валют\"");

            var receiptsSheet = new Workbook(fileName).Worksheets[2];
            _receipts = receiptsSheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .Select(row => new Receipts(
                    int.Parse(row.GetCellOrNull(0).Value.ToString()),
                    int.Parse(row.GetCellOrNull(1).Value.ToString()),
                    int.Parse(row.GetCellOrNull(2).Value.ToString()),
                    DateTime.Parse(row.GetCellOrNull(3).Value.ToString()),
                    double.Parse(row.GetCellOrNull(4).Value.ToString())
                ))
                .ToList();
            _logger.Write("Получена таблица \"Начисления\"");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка при чтении Excel файла {fileName}: {e.Message}");
            _logger.Write($"Ошибка при чтении Excel файла {fileName}: {e.Message}");
        }
    }

    public void PrintData()
    {
        Console.WriteLine("Вывод данных Excel файла");
        _logger.Write("Вывод данных Excel файла");
        
        Console.WriteLine("Счета:");
        _accounts.ForEach(account => Console.WriteLine(account.ToString()));
        _logger.Write("Выведена таблица \"Счета\"");

        Console.WriteLine("\nКурс валют:");
        _exchangeRates.ForEach(exchangeRate => Console.WriteLine(exchangeRate.ToString()));
        _logger.Write("Выведена таблица \"Курс валют\"");

        Console.WriteLine("\nПоступления:");
        _receipts.ForEach(receipt => Console.WriteLine(receipt.ToString()));
        _logger.Write("Выведена таблица \"Начисления\"");
    }

    public void DeleteRowById(string fileName, string sheetName, int id)
    {
        try
        {
            Console.WriteLine($"Удаление строки с номером {id} в листе {sheetName} в файле {fileName}");
            _logger.Write($"Удаление строки с номером {id} в листе {sheetName} в файле {fileName}");

            var workbook = new Workbook(fileName);
            var worksheet = workbook.Worksheets[sheetName];
            var rowToDelete = worksheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .FirstOrDefault(row => int.Parse(row.GetCellOrNull(0).StringValue) == id);
            _logger.Write("Найдена строка для удаления");

            if (rowToDelete != null) worksheet.Cells.DeleteRow(rowToDelete.Index);
            else
            {
                Console.WriteLine($"Строка с номером {id} в листе {sheetName} в файле {fileName} не найдена");
                _logger.Write($"Строка с номером {id} в листе {sheetName} в файле {fileName} не найдена");
                return;
            }

            workbook.Save(fileName);
            workbook.Dispose();
            worksheet.Dispose();

            Console.WriteLine($"Строка с номером {id} в листе {sheetName} в файле {fileName} успешно удалена");
            _logger.Write($"Строка с номером {id} в листе {sheetName} в файле {fileName} успешно удалена");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка при удалении строки с номером {id} в листе {sheetName} в файле {fileName}: {e.Message}");
            _logger.Write($"Ошибка при удалении строки с номером {id} в листе {sheetName} в файле {fileName}: {e.Message}");
        }
    }

    public void UpdateRowById(string fileName, string sheetName, int id, List<string> values)
    {
        try
        {
            Console.WriteLine($"Обновление строки с номером {id} в листе {sheetName} в файле {fileName}");
            _logger.Write($"Обновление строки с номером {id} в листе {sheetName} в файле {fileName}");
            
            var workbook = new Workbook(fileName);
            var worksheet = workbook.Worksheets[sheetName];
            var rowToUpdate = worksheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .FirstOrDefault(row => row.GetCellOrNull(0) != null 
                    ? int.Parse(row.GetCellOrNull(0).StringValue) == id 
                    : false);
            _logger.Write("Найдена строка для обновления");

            if (rowToUpdate != null)
            {
                for (var i = 0; i < values.Count; i++)
                {
                    rowToUpdate.GetCellOrNull(i).PutValue(values[i]);
                }
            }
            else
            {
                Console.WriteLine($"Строка с номером {id} в листе {sheetName} в файле {fileName} не найдена");
                _logger.Write($"Строка с номером {id} в листе {sheetName} в файле {fileName} не найдена");
                return;
            }

            workbook.Save(fileName);
            workbook.Dispose();
            worksheet.Dispose();

            Console.WriteLine($"Строка с номером {id} в листе {sheetName} в файле {fileName} успешно обновлена");
            _logger.Write($"Строка с номером {id} в листе {sheetName} в файле {fileName} успешно обновлена");
            LoadData(fileName);
        }
        catch (Exception e)
        {
            Console.WriteLine(
                $"Ошибка при обновлении строки с номером {id} в листе {sheetName} в файле {fileName}: {e.Message}");
            _logger.Write(
                $"Ошибка при обновлении строки с номером {id} в листе {sheetName} в файле {fileName}: {e.Message}");
        }
    }

    public void InsertRow(string fileName, string sheetName, List<string> values)
    {
        try
        {
            Console.WriteLine($"Добавление строки в лист {sheetName} в файле {fileName}");
            _logger.Write($"Добавление строки в лист {sheetName} в файле {fileName}");

            var workbook = new Workbook(fileName);
            var worksheet = workbook.Worksheets[sheetName];
            var rows = worksheet.Cells.Rows.Cast<Row>()
                .Skip(1)
                .ToList();
            _logger.Write("Получены столбцы таблицы");
            
            var nextId = rows.Max(row => row.GetCellOrNull(0) != null 
                ? int.Parse(row.GetCellOrNull(0).StringValue) 
                : 0) + 1;
            _logger.Write("Получен следующий ID");
            var nextRow = worksheet.Cells.MaxDataRow + 1;
            _logger.Write("Получена номер следующей строки");
            values.Insert(0, nextId.ToString());
            
            for (var i = 0; i < values.Count; i++)
            {
                worksheet.Cells[nextRow, i].PutValue(values[i]);
            }

            workbook.Save(fileName);
            workbook.Dispose();
            worksheet.Dispose();

            Console.WriteLine(
                $"Новая строка под номером {nextId} в листе {sheetName} в файле {fileName} успешно обновлена");
            _logger.Write(
                $"Новая строка под номером {nextId} в листе {sheetName} в файле {fileName} успешно обновлена");
            LoadData(fileName);
        }
        catch (Exception e)
        {
            Console.WriteLine(
                $"Ошибка при добавлении новой строки номер в листе {sheetName} в файле {fileName}: {e.Message}");
            _logger.Write($"Ошибка при добавлении новой строки номер в листе {sheetName} в файле {fileName}: {e.Message}");
        }
    }

    public int Query1()
    {
        // Найти количество начислений на сумму более 1.

        var receiptsCount = _receipts
            .Count(receipt => receipt.Total > 1);
        _logger.Write("Получен ответ для запроса 1");
        
        return receiptsCount;
    }

    public string Query2()
    {
        // Определить держателя счёта с максимальным начислением в рублях за 25 декабря 2021 года.
        
        var accountMaxFullName = from account in _accounts
            join receipt in _receipts on account.Id equals receipt.AccountId
            where receipt.Date == new DateTime(2021, 12, 25) 
            orderby receipt.Total descending 
            select account.FullName;
        _logger.Write("Получен ответ для запроса 2");

        return accountMaxFullName.First();
    }

    public List<string> Query3()
    {
        // Найти держателя счета с минимальным количеством операций, совершенных после 27 декабря
        // 2021 года, и валюту, в которой было проведено больше всего транзакций.
        
        // количество операций для каждого счета после 27 декабря 2021 года
        var operationsAfterDate = _receipts
            .Where(receipt => receipt.Date > new DateTime(2021, 12, 27))
            .GroupBy(receipt => receipt.AccountId)
            .Select(group => new
            {
                AccountId = group.Key,
                OperationCount = group.Count()
            })
            .OrderBy(group => group.OperationCount)
            .FirstOrDefault();
        _logger.Write("Найдено количество операций для каждого счета после 27 декабря 2021 года для запроса 3");

        // количество транзакций по каждой валюте
        var mostPopularCurrency = _receipts
            .Where(receipt => receipt.AccountId == operationsAfterDate.AccountId)
            .GroupBy(receipt => receipt.CurrencyId)
            .Select(group => new
            {
                CurrencyId = group.Key,
                TransactionCount = group.Count()
            })
            .OrderByDescending(group => group.TransactionCount)
            .FirstOrDefault();
        _logger.Write("Найдено количество транзакций по каждой валюте для запроса 3");

        // имя держателя счета с минимальным количеством операций
        var accountHolder = _accounts
            .Where(account => account.Id == operationsAfterDate.AccountId)
            .Select(account => account.FullName)
            .FirstOrDefault();
        _logger.Write("Найдено имя держателя счета с минимальным количеством операций для запроса 3");

        // название самой популярной валюты
        var currencyName = _exchangeRates
            .Where(exchangeRates => exchangeRates.Id == mostPopularCurrency.CurrencyId)
            .Select(exchangeRates => exchangeRates.FullName)
            .FirstOrDefault();
        _logger.Write("Найдено название самой популярной валюты для запроса 3");
        
        return new List<string> { accountHolder, operationsAfterDate.OperationCount.ToString(), currencyName };
    }

    public List<string> Query4()
    {
        // Вывести полные данные об операции, где сумма операции больше среднего значения по всем
        // операциям и меньше всех остальных
        
        var averageSum = _receipts.Sum(receipt => receipt.Total) / _receipts.Count;

        var result = (from receipts in _receipts
            join account in _accounts on receipts.AccountId equals account.Id
            join exchangeRates in _exchangeRates on receipts.CurrencyId equals exchangeRates.Id
            where receipts.Total > averageSum
            orderby receipts.Total descending
            select new {
                account.FullName,
                DepositOpeningDate = account.DepositOpeningDate.ToString("dd/MM/yyyy"),
                CurrencyName = exchangeRates.FullName,
                ReceiptDate = receipts.Date.ToString("dd/MM/yyyy"),
                ReceiptSum = receipts.Total.ToString()
            })
            .First();
        _logger.Write("Получены результирующие значения для запроса 4");

        return new List<string>
            {
                result.FullName, 
                result.DepositOpeningDate, 
                result.CurrencyName, 
                result.ReceiptDate, 
                result.ReceiptSum
                
            };
    }
}