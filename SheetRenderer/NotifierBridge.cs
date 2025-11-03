using ExcelDna.Integration;

public sealed class NotifierBridge
{
    public void Info(string title, string message, int timeoutMs = 3000)
        => ExcelAsyncUtil.QueueAsMacro(() => Notifier.Info(title, message, timeoutMs));

    public void Warn(string title, string message, int timeoutMs = 3000)
        => ExcelAsyncUtil.QueueAsMacro(() => Notifier.Warn(title, message, timeoutMs));

    public void Error(string title, string message, int timeoutMs = 5000)
        => ExcelAsyncUtil.QueueAsMacro(() => Notifier.Error(title, message, timeoutMs));
}
