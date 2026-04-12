using System;
using System.Drawing;
using System.Windows.Forms;

internal static class GitLabRepoDialog
{
    public static bool TryShow(GitLabLastInput initial, out GitLabLastInput result)
    {
        result = null;

        if (initial == null)
        {
            initial = new GitLabLastInput();
        }

        using (var form = new Form())
        using (var lblBaseUrl = new Label())
        using (var txtBaseUrl = new TextBox())
        using (var lblProjectId = new Label())
        using (var txtProjectId = new TextBox())
        using (var lblRefName = new Label())
        using (var txtRefName = new TextBox())
        using (var chkPullEnabled = new CheckBox())
        using (var btnOk = new Button())
        using (var btnCancel = new Button())
        {
            form.Text = "取得元設定";
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ShowInTaskbar = false;
            form.ClientSize = new Size(700, 228);
            form.Font = new Font("Meiryo UI", 11f);

            txtBaseUrl.AutoSize = false;
            txtBaseUrl.Height = 28;
            txtProjectId.AutoSize = false;
            txtProjectId.Height = 28;
            txtRefName.AutoSize = false;
            txtRefName.Height = 28;

            int leftLabel = 14;
            int leftInput = 160;
            int top = 16;
            int rowH = 32;
            int inputW = 520;
            int labelW = 140;

            ConfigureLabel(lblBaseUrl, "Base URL", leftLabel, top, labelW);
            ConfigureTextBox(txtBaseUrl, leftInput, top - 2, inputW);
            txtBaseUrl.Text = initial.BaseUrl ?? "";

            top += rowH;
            ConfigureLabel(lblProjectId, "Project ID", leftLabel, top, labelW);
            ConfigureTextBox(txtProjectId, leftInput, top - 2, inputW);
            txtProjectId.Text = initial.ProjectId ?? "";

            top += rowH;
            ConfigureLabel(lblRefName, "Ref (branch/tag)", leftLabel, top, labelW);
            ConfigureTextBox(txtRefName, leftInput, top - 2, inputW);
            txtRefName.Text = string.IsNullOrWhiteSpace(initial.RefName) ? "main" : initial.RefName;

            top += rowH + 4;
            chkPullEnabled.Left = leftInput;
            chkPullEnabled.Top = top - 2;
            chkPullEnabled.Width = inputW;
            chkPullEnabled.Height = 24;
            chkPullEnabled.Text = "取得元から最新版を取得する";
            chkPullEnabled.Checked = initial.PullEnabled != false;

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
                else if (string.IsNullOrWhiteSpace(txtRefName.Text))
                {
                    txtRefName.Focus();
                    txtRefName.SelectAll();
                }
                else
                {
                    chkPullEnabled.Focus();
                }
            };

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
            form.Controls.Add(chkPullEnabled);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            txtBaseUrl.KeyDown += (s, e) => MoveNextOnEnter(e, txtProjectId);
            txtProjectId.KeyDown += (s, e) => MoveNextOnEnter(e, txtRefName);
            txtRefName.KeyDown += (s, e) => MoveNextOnEnter(e, chkPullEnabled);
            chkPullEnabled.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    btnOk.PerformClick();
                }
            };

            GitLabLastInput tempResult = null;

            btnOk.Click += (s, e) =>
            {
                string baseUrl = (txtBaseUrl.Text ?? "").Trim().TrimEnd('/');
                string projectId = (txtProjectId.Text ?? "").Trim();
                string refName = (txtRefName.Text ?? "").Trim();
                bool pullEnabled = chkPullEnabled.Checked;

                if (pullEnabled && string.IsNullOrWhiteSpace(baseUrl))
                {
                    MessageBox.Show(form, "Base URL を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtBaseUrl.Focus();
                    return;
                }

                if (pullEnabled && string.IsNullOrWhiteSpace(projectId))
                {
                    MessageBox.Show(form, "Project ID を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtProjectId.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(refName))
                {
                    refName = "main";
                }

                if (pullEnabled &&
                    !baseUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                    !baseUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(form, "Base URL は http:// または https:// から始めてください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtBaseUrl.Focus();
                    return;
                }

                tempResult = new GitLabLastInput
                {
                    BaseUrl = baseUrl,
                    ProjectId = projectId,
                    RefName = refName,
                    FilePath = initial.FilePath,
                    PullEnabled = pullEnabled
                };

                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            DialogResult dialogResult = form.ShowDialog();
            if (dialogResult == DialogResult.OK && tempResult != null)
            {
                result = tempResult;
                return true;
            }

            return false;
        }
    }

    private static void ConfigureLabel(Label label, string text, int left, int top, int width)
    {
        label.Text = text;
        label.Left = left;
        label.Top = top + 4;
        label.Width = width;
        label.Height = 20;
    }

    private static void ConfigureTextBox(TextBox textBox, int left, int top, int width)
    {
        textBox.Left = left;
        textBox.Top = top;
        textBox.Width = width;
        textBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
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
