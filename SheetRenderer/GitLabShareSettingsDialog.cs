using System;
using System.Drawing;
using System.Windows.Forms;

internal static class GitLabShareSettingsDialog
{
    public static bool TryShow(
        GitLabShareInfo initialShare,
        GitLabLastInput pullDefaults,
        out GitLabShareInfo result)
    {
        result = null;

        if (initialShare == null)
        {
            initialShare = new GitLabShareInfo();
        }

        using (var form = new Form())
        using (var lblBaseUrl = new Label())
        using (var txtBaseUrl = new TextBox())
        using (var lblProjectId = new Label())
        using (var txtProjectId = new TextBox())
        using (var lblRefName = new Label())
        using (var txtRefName = new TextBox())
        using (var btnOk = new Button())
        using (var btnCancel = new Button())
        {
            form.Text = "共有先設定";
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ShowInTaskbar = false;
            form.ClientSize = new Size(700, 190);
            form.Font = new Font("Meiryo UI", 11f);

            txtBaseUrl.AutoSize = false;
            txtBaseUrl.Height = 28;
            txtProjectId.AutoSize = false;
            txtProjectId.Height = 28;
            txtRefName.AutoSize = false;
            txtRefName.Height = 28;

            form.Shown += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtBaseUrl.Text))
                {
                    txtBaseUrl.Focus();
                    txtBaseUrl.SelectAll();
                }
                else if (string.IsNullOrWhiteSpace(txtProjectId.Text))
                {
                    txtProjectId.Focus();
                    txtProjectId.SelectAll();
                }
                else
                {
                    txtRefName.Focus();
                    txtRefName.SelectAll();
                }
            };

            int leftLabel = 14;
            int leftInput = 160;
            int top = 16;
            int rowH = 32;
            int inputW = 520;
            int labelW = 140;

            ConfigureLabel(lblBaseUrl, "Base URL", leftLabel, top, labelW);
            ConfigureTextBox(txtBaseUrl, leftInput, top - 2, inputW);
            txtBaseUrl.Text = !string.IsNullOrWhiteSpace(initialShare.BaseUrl)
                ? initialShare.BaseUrl
                : (pullDefaults == null ? "" : (pullDefaults.BaseUrl ?? ""));

            top += rowH;
            ConfigureLabel(lblProjectId, "Project ID", leftLabel, top, labelW);
            ConfigureTextBox(txtProjectId, leftInput, top - 2, inputW);
            txtProjectId.Text = initialShare.ProjectId ?? "";

            top += rowH;
            ConfigureLabel(lblRefName, "Ref (branch/tag)", leftLabel, top, labelW);
            ConfigureTextBox(txtRefName, leftInput, top - 2, inputW);
            txtRefName.Text = !string.IsNullOrWhiteSpace(initialShare.RefName)
                ? initialShare.RefName
                : GetDefaultRefName(pullDefaults);

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

            form.Controls.Add(lblBaseUrl);
            form.Controls.Add(txtBaseUrl);
            form.Controls.Add(lblProjectId);
            form.Controls.Add(txtProjectId);
            form.Controls.Add(lblRefName);
            form.Controls.Add(txtRefName);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            txtBaseUrl.KeyDown += (s, e) => MoveNextOnEnter(e, txtProjectId);
            txtProjectId.KeyDown += (s, e) => MoveNextOnEnter(e, txtRefName);
            txtRefName.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    btnOk.PerformClick();
                }
            };

            GitLabShareInfo tempResult = null;

            btnOk.Click += (s, e) =>
            {
                string baseUrl = (txtBaseUrl.Text ?? "").Trim().TrimEnd('/');
                string projectId = (txtProjectId.Text ?? "").Trim();
                string refName = (txtRefName.Text ?? "").Trim();

                if (string.IsNullOrWhiteSpace(baseUrl))
                {
                    MessageBox.Show(form, "共有先の Base URL を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtBaseUrl.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(projectId))
                {
                    MessageBox.Show(form, "共有先の Project ID を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtProjectId.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(refName))
                {
                    refName = "main";
                }

                if (!baseUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                    !baseUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(form, "共有先の Base URL は http:// または https:// から始めてください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtBaseUrl.Focus();
                    return;
                }

                tempResult = new GitLabShareInfo
                {
                    BaseUrl = baseUrl,
                    ProjectId = projectId,
                    RefName = refName
                };

                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            DialogResult dr = form.ShowDialog();
            if (dr == DialogResult.OK && tempResult != null)
            {
                result = tempResult;
                return true;
            }

            return false;
        }
    }

    private static string GetDefaultRefName(GitLabLastInput pullDefaults)
    {
        if (pullDefaults == null || string.IsNullOrWhiteSpace(pullDefaults.RefName))
        {
            return "main";
        }

        return pullDefaults.RefName;
    }

    private static void ConfigureLabel(Label label, string text, int left, int top, int width)
    {
        label.Text = text;
        label.Left = left;
        label.Top = top + 4;
        label.Width = width;
        label.Height = 20;
    }

    private static void ConfigureTextBox(TextBox tb, int left, int top, int width)
    {
        tb.Left = left;
        tb.Top = top;
        tb.Width = width;
        tb.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
    }

    private static void MoveNextOnEnter(KeyEventArgs e, Control next)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            if (next != null)
            {
                next.Focus();
            }
        }
    }
}
