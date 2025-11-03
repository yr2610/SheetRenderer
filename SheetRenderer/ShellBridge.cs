using System.Windows.Forms;
using ExcelDna.Integration;

public sealed class ShellBridge
{
    public void Popup(string text, string title = "情報", int type = 0)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var icon = MessageBoxIcon.None;
            if ((type & 0x10) != 0) icon = MessageBoxIcon.Warning;     // vbExclamation
            if ((type & 0x20) != 0) icon = MessageBoxIcon.Information; // vbInformation
            if ((type & 0x30) != 0) icon = MessageBoxIcon.Question;    // vbQuestion
            if ((type & 0x40) != 0) icon = MessageBoxIcon.Error;       // vbCritical

            var buttons = MessageBoxButtons.OK;
            if ((type & 0x2) != 0) buttons = MessageBoxButtons.OKCancel;
            else if ((type & 0x3) != 0) buttons = MessageBoxButtons.AbortRetryIgnore;
            else if ((type & 0x4) != 0) buttons = MessageBoxButtons.YesNoCancel;
            else if ((type & 0x5) != 0) buttons = MessageBoxButtons.YesNo;
            else if ((type & 0x6) != 0) buttons = MessageBoxButtons.RetryCancel;

            MessageBox.Show(text ?? "", title ?? "情報", buttons, icon);
        });
    }
}
