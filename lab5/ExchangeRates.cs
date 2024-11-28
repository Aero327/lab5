namespace lab5;

public class ExchangeRates
{
    private int _id;
    private string _countryCode;
    private double _rate;
    private string _fullName;

    public ExchangeRates(int id, string countryCode, double rate, string fullName)
    {
        _id = id;
        _countryCode = countryCode;
        _rate = rate;
        _fullName = fullName;
    }


    public string FullName
    {
        get => _fullName;
        set => _fullName = value ?? throw new ArgumentNullException(nameof(value));
    }

    public double Rate
    {
        get => _rate;
        set => _rate = value;
    }

    public string CountryCode
    {
        get => _countryCode;
        set => _countryCode = value ?? throw new ArgumentNullException(nameof(value));
    }

    public int Id
    {
        get => _id;
        set => _id = value;
    }

    public override string ToString()
    {
        return $"id: {_id}, countryCode: {_countryCode}, rate: {_rate}, fullName: {_fullName}";
    }
}