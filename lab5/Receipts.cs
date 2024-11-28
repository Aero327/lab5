namespace lab5;

public class Receipts
{
    private int _id;
    private int _accountId;
    private int _currencyId;
    private DateTime _date;
    private double _total;

    public Receipts(int id, int accountId, int currencyId, DateTime date, double total)
    {
        _id = id;
        _accountId = accountId;
        _currencyId = currencyId;
        _date = date;
        _total = total;
    }

    public int Id
    {
        get => _id;
        set => _id = value;
    }

    public int AccountId
    {
        get => _accountId;
        set => _accountId = value;
    }

    public int CurrencyId
    {
        get => _currencyId;
        set => _currencyId = value;
    }

    public DateTime Date
    {
        get => _date;
        set => _date = value;
    }

    public double Total
    {
        get => _total;
        set => _total = value;
    }

    public override string ToString()
    {
        return $"id: {_id}, accountId: {_accountId}, currencyId: {_currencyId} date: {_date:dd-MM-yyyy}, total: {_total}";
    }
}