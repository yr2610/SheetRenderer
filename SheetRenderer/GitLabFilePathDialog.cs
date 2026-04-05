using System;
using System.Drawing;
using System.Windows.Forms;

internal static class GitLabFilePathDialog
{
    public static bool TryShow(string initialFilePath, out string filePath)
    {
        filePath = null;

        using (var form = new Form())
        using (var lblFilePath = new Label())
        using (var txtFilePath = new TextBox())
        using (var btnOk = new Button())
        using (var btnCancel = new Button())
        {
            form.Text = "取得するファイルを指定";
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ShowInTaskbar = false;
            form.ClientSize = new Size(700, 120);
            form.Font = new Font("Meiryo UI", 11f);

            lblFilePath.Text = "File Path";
            lblFilePath.Left = 14;
            lblFilePath.Top = 20;
            lblFilePath.Width = 120;
            lblFilePath.Height = 20;

            txtFilePath.AutoSize = false;
            txtFilePath.Height = 28;
            txtFilePath.Left = 140;
            txtFilePath.Top = 16;
            txtFilePath.Width = 540;
            txtFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFilePath.Text = initialFilePath ?? "";

            btnOk.Text = "OK";
            btnOk.Width = 90;
            btnOk.Height = 28;
            btnOk.Left = form.ClientSize.Width - btnOk.Width - 110;
            btnOk.Top = form.ClientSize.Height - btnOk.Height - 14;
            btnOk.DialogResult = DialogResult.OK;

            btnCancel.Text = "Cancel";
            btnCancel.Width = 90;
            btnCancel.Height = 28;
            btnCancel.Left = form.ClientSize.Width - btnCancel.Width - 14;
            btnCancel.Top = form.ClientSize.Height - btnCancel.Height - 14;
            btnCancel.DialogResult = DialogResult.Cancel;

            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            form.Controls.Add(lblFilePath);
            form.Controls.Add(txtFilePath);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            form.Shown += (s, e) =>
            {
                txtFilePath.Focus();
                txtFilePath.SelectAll();
            };

            txtFilePath.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    btnOk.PerformClick();
                }
            };

            string tempResult = null;

            btnOk.Click += (s, e) =>
            {
                string value = (txtFilePath.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(value))
                {
                    MessageBox.Show(form, "File Path を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtFilePath.Focus();
                    return;
                }

                tempResult = value;
                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            DialogResult dr = form.ShowDialog();
            if (dr == DialogResult.OK && !string.IsNullOrWhiteSpace(tempResult))
            {
                filePath = tempResult;
                return true;
            }

            return false;
        }
    }
}
