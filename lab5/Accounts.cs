namespace lab5;

public class Accounts
{
    private int _id;
    private string _fullName;
    private DateTime _depositOpeningDate;

    public Accounts(int id, string fullName, DateTime depositOpeningDate)
    {
        _id = id;
        _fullName = fullName;
        _depositOpeningDate = depositOpeningDate;
    }

    public DateTime DepositOpeningDate
    {
        get => _depositOpeningDate;
        set => _depositOpeningDate = value;
    }

    public string FullName
    {
        get => _fullName;
        set => _fullName = value ?? throw new ArgumentNullException(nameof(value));
    }

    public int Id
    {
        get => _id;
        set => _id = value;
    }

    public override string ToString()
    {
        return $"Id: {_id}, FullName: {_fullName}, DepositOpeningDate: {_depositOpeningDate:dd-MM-yyyy}";
    }
}