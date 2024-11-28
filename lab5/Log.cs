namespace lab5;

public class Log
{
    private string _logFileName;
    private bool _add;

    public Log(string logFileName, bool add)
    {
        _logFileName = logFileName;
        _add = add;

        if (!add)
        {
            File.WriteAllText(_logFileName, string.Empty);
        }
    }
    
    public void Write(string message)
    {
        if (_add)
        {
            File.AppendAllText(_logFileName, message + "\n");
        } else
        {
            File.WriteAllText(_logFileName, message + "\n");
            _add = true;
        }
    }
}