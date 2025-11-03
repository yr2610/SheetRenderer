using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

public static class Notifier
{
    private static NotifyIcon _icon;

    public static void Initialize()
    {
        if (_icon != null)
        {
            return;
        }

        _icon = new NotifyIcon();
        // アイコンは適当でOK（差し替え可能）
        _icon.Icon = SystemIcons.Application;
        _icon.Visible = true;
        _icon.Text = "SheetRenderer";

        // クリックでログファイルを開く（無ければ作成）
        _icon.BalloonTipClicked += (s, e) => FileLogger.OpenLog();
    }

    public static void Dispose()
    {
        if (_icon != null)
        {
            _icon.Visible = false;
            _icon.Dispose();
            _icon = null;
        }
    }

    public static void Info(string title, string message, int timeoutMs = 3000)
    {
        ShowBalloon(title, message, ToolTipIcon.Info, timeoutMs);
        FileLogger.Info($"{title}: {message}");
    }

    public static void Warn(string title, string message, int timeoutMs = 3000)
    {
        ShowBalloon(title, message, ToolTipIcon.Warning, timeoutMs);
        FileLogger.Warn($"{title}: {message}");
    }

    public static void Error(string title, string message, int timeoutMs = 5000)
    {
        ShowBalloon(title, message, ToolTipIcon.Error, timeoutMs);
        FileLogger.Error($"{title}: {message}");
    }

    private static void ShowBalloon(string title, string message, ToolTipIcon icon, int timeoutMs)
    {
        if (_icon == null)
        {
            Initialize();
        }
        // バルーンはだいたい 200〜250 文字で省略されます（要約にとどめる）
        _icon.BalloonTipTitle = title ?? "";
        _icon.BalloonTipText = message ?? "";
        _icon.BalloonTipIcon = icon;
        _icon.ShowBalloonTip(timeoutMs);
    }
}
