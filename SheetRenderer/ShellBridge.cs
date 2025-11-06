using System.Windows.Forms;
using ExcelDna.Integration;

public sealed class ShellBridge
{
    /// <summary>
    /// VBScript / WSH の Shell.Popup に相当する実装。
    /// https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/aa767739(v=vs.85)
    /// </summary>
    public int Popup(string text, string title = "情報", int type = 0)
    {
        int result = 2; // vbCancel (既定)
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var icon = MessageBoxIcon.None;
            var buttons = MessageBoxButtons.OK;

            // --- Icon 部分 ---
            int iconBits = type & 0xF0;
            switch (iconBits)
            {
                case 0x10: icon = MessageBoxIcon.Error; break;         // 16 = Stop
                case 0x20: icon = MessageBoxIcon.Question; break;      // 32 = ?
                case 0x30: icon = MessageBoxIcon.Exclamation; break;   // 48 = !
                case 0x40: icon = MessageBoxIcon.Information; break;   // 64 = i
            }

            // --- Button 部分 ---
            int buttonBits = type & 0x0F;
            switch (buttonBits)
            {
                case 0x0: buttons = MessageBoxButtons.OK; break;
                case 0x1: buttons = MessageBoxButtons.OKCancel; break;
                case 0x2: buttons = MessageBoxButtons.AbortRetryIgnore; break;
                case 0x3: buttons = MessageBoxButtons.YesNoCancel; break;
                case 0x4: buttons = MessageBoxButtons.YesNo; break;
                case 0x5: buttons = MessageBoxButtons.RetryCancel; break;
            }

            var r = MessageBox.Show(text ?? "", title ?? "情報", buttons, icon);

            // --- 戻り値をVB互換で返す ---
            switch (r)
            {
                case DialogResult.OK: result = 1; break;       // vbOK
                case DialogResult.Cancel: result = 2; break;   // vbCancel
                case DialogResult.Abort: result = 3; break;    // vbAbort
                case DialogResult.Retry: result = 4; break;    // vbRetry
                case DialogResult.Ignore: result = 5; break;   // vbIgnore
                case DialogResult.Yes: result = 6; break;      // vbYes
                case DialogResult.No: result = 7; break;       // vbNo
            }
        });

        return result;
    }
}
