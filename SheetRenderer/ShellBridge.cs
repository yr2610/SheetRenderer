using System;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;

public sealed class ShellBridge
{
    // ExcelのUIスレッドID（AutoOpen等でInitializeOnExcelUiThreadを呼んでセット）
    private static int _uiThreadId = -1;

    /// <summary>
    /// 必ずExcelのUIスレッドで呼ぶ（AutoOpenなど）
    /// </summary>
    public static void InitializeOnExcelUiThread()
    {
        _uiThreadId = Thread.CurrentThread.ManagedThreadId;
    }

    /// <summary>
    /// WSH Shell.Popup 互換（timeoutは未対応）
    /// 戻り値: vbOK=1, vbCancel=2, vbAbort=3, vbRetry=4, vbIgnore=5, vbYes=6, vbNo=7
    /// </summary>
    public int Popup(string text, string title = "情報", int type = 0)
    {
        int ShowAndMap()
        {
            var icon = MapIcon(type);
            var buttons = MapButtons(type);
            var dr = MessageBox.Show(text ?? "", title ?? "情報", buttons, icon);
            return MapDialogResultToVb(dr);
        }

        bool isUiThread = Thread.CurrentThread.ManagedThreadId == _uiThreadId;

        // UIスレッドなら同期で即表示（絶対に待たない）
        if (isUiThread || _uiThreadId == -1)   // 初期化前は保守的に直呼び
        {
            return ShowAndMap();
        }

        // 非UIスレッドならUIに投げて完了待ち（この待ちはUIをブロックしない）
        int result = 2; // 既定: vbCancel
        using (var done = new ManualResetEventSlim(false))
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { result = ShowAndMap(); }
                finally { done.Set(); }
            });

            // 無限待ちは怖いのでタイムアウト保険（例: 60秒）
            if (!done.Wait(TimeSpan.FromSeconds(60)))
            {
                // 失敗時はvbCancel相当で返す or 例外を投げるなど運用方針に合わせて
                return 2;
            }
        }
        return result;
    }

    // --- WSH互換のビットマップ ---

    private static MessageBoxIcon MapIcon(int type)
    {
        switch (type & 0xF0)   // 上位4bit: 16, 32, 48, 64
        {
            case 0x10: return MessageBoxIcon.Error;        // 16 = Stop/Critical
            case 0x20: return MessageBoxIcon.Question;     // 32 = ?
            case 0x30: return MessageBoxIcon.Exclamation;  // 48 = !
            case 0x40: return MessageBoxIcon.Information;  // 64 = i
            default: return MessageBoxIcon.None;
        }
    }

    private static MessageBoxButtons MapButtons(int type)
    {
        switch (type & 0x0F)   // 下位4bit: 0..5
        {
            case 0x0: return MessageBoxButtons.OK;
            case 0x1: return MessageBoxButtons.OKCancel;
            case 0x2: return MessageBoxButtons.AbortRetryIgnore;
            case 0x3: return MessageBoxButtons.YesNoCancel;
            case 0x4: return MessageBoxButtons.YesNo;
            case 0x5: return MessageBoxButtons.RetryCancel;
            default: return MessageBoxButtons.OK;
        }
    }

    private static int MapDialogResultToVb(DialogResult r)
    {
        switch (r)
        {
            case DialogResult.OK: return 1; // vbOK
            case DialogResult.Cancel: return 2; // vbCancel
            case DialogResult.Abort: return 3; // vbAbort
            case DialogResult.Retry: return 4; // vbRetry
            case DialogResult.Ignore: return 5; // vbIgnore
            case DialogResult.Yes: return 6; // vbYes
            case DialogResult.No: return 7; // vbNo
            default: return 2;
        }
    }
}
