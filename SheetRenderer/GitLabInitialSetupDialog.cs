using System;
using System.Drawing;
using System.Windows.Forms;

internal static class GitLabInitialSetupDialog
{
    public static bool TryShow(
        GitLabLastInput initialPull,
        GitLabShareInfo initialShare,
        out GitLabLastInput pullResult,
        out GitLabShareInfo shareResult)
    {
        pullResult = null;
        shareResult = null;

        if (initialPull == null)
        {
            initialPull = new GitLabLastInput();
        }

        if (initialShare == null)
        {
            initialShare = new GitLabShareInfo();
        }

        using (var form = new Form())
        using (var grpPull = new GroupBox())
        using (var grpShare = new GroupBox())
        using (var lblBaseUrl = new Label())
        using (var txtBaseUrl = new TextBox())
        using (var lblProjectId = new Label())
        using (var txtProjectId = new TextBox())
        using (var lblRefName = new Label())
        using (var txtRefName = new TextBox())
        using (var lblFilePath = new Label())
        using (var txtFilePath = new TextBox())
        using (var lblShareBaseUrl = new Label())
        using (var txtShareBaseUrl = new TextBox())
        using (var lblShareProjectId = new Label())
        using (var txtShareProjectId = new TextBox())
        using (var lblShareRefName = new Label())
        using (var txtShareRefName = new TextBox())
        using (var btnOk = new Button())
        using (var btnCancel = new Button())
        {
            form.Text = "GitLab 初期セットアップ";
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ShowInTaskbar = false;
            form.ClientSize = new Size(760, 430);
            form.Font = new Font("Meiryo UI", 11f);

            grpPull.Text = "取得元";
            grpPull.Left = 12;
            grpPull.Top = 12;
            grpPull.Width = form.ClientSize.Width - 24;
            grpPull.Height = 170;
            grpPull.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            grpShare.Text = "共有先";
            grpShare.Left = 12;
            grpShare.Top = grpPull.Bottom + 10;
            grpShare.Width = form.ClientSize.Width - 24;
            grpShare.Height = 132;
            grpShare.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            ConfigureTextBox(txtBaseUrl);
            ConfigureTextBox(txtProjectId);
            ConfigureTextBox(txtRefName);
            ConfigureTextBox(txtFilePath);
            ConfigureTextBox(txtShareBaseUrl);
            ConfigureTextBox(txtShareProjectId);
            ConfigureTextBox(txtShareRefName);

            int leftLabel = 18;
            int leftInput = 170;
            int inputWidth = grpPull.Width - leftInput - 22;
            int rowHeight = 32;
            int top = 28;
            int labelWidth = 140;

            ConfigureLabel(lblBaseUrl, "Base URL", leftLabel, top, labelWidth);
            ConfigureTextBoxBounds(txtBaseUrl, leftInput, top - 2, inputWidth);
            txtBaseUrl.Text = initialPull.BaseUrl ?? "";

            top += rowHeight;
            ConfigureLabel(lblProjectId, "Project ID", leftLabel, top, labelWidth);
            ConfigureTextBoxBounds(txtProjectId, leftInput, top - 2, inputWidth);
            txtProjectId.Text = initialPull.ProjectId ?? "";

            top += rowHeight;
            ConfigureLabel(lblRefName, "Ref (branch/tag)", leftLabel, top, labelWidth);
            ConfigureTextBoxBounds(txtRefName, leftInput, top - 2, inputWidth);
            txtRefName.Text = string.IsNullOrWhiteSpace(initialPull.RefName) ? "main" : initialPull.RefName;

            top += rowHeight;
            ConfigureLabel(lblFilePath, "File Path", leftLabel, top, labelWidth);
            ConfigureTextBoxBounds(txtFilePath, leftInput, top - 2, inputWidth);
            txtFilePath.Text = initialPull.FilePath ?? "";

            grpPull.Controls.Add(lblBaseUrl);
            grpPull.Controls.Add(txtBaseUrl);
            grpPull.Controls.Add(lblProjectId);
            grpPull.Controls.Add(txtProjectId);
            grpPull.Controls.Add(lblRefName);
            grpPull.Controls.Add(txtRefName);
            grpPull.Controls.Add(lblFilePath);
            grpPull.Controls.Add(txtFilePath);

            int shareInputWidth = grpShare.Width - leftInput - 22;
            int shareTop = 28;

            ConfigureLabel(lblShareBaseUrl, "Base URL", leftLabel, shareTop, labelWidth);
            ConfigureTextBoxBounds(txtShareBaseUrl, leftInput, shareTop - 2, shareInputWidth);
            txtShareBaseUrl.Text = !string.IsNullOrWhiteSpace(initialShare.BaseUrl)
                ? initialShare.BaseUrl
                : (initialPull.BaseUrl ?? "");

            shareTop += rowHeight;
            ConfigureLabel(lblShareProjectId, "Project ID", leftLabel, shareTop, labelWidth);
            ConfigureTextBoxBounds(txtShareProjectId, leftInput, shareTop - 2, shareInputWidth);
            txtShareProjectId.Text = initialShare.ProjectId ?? "";

            shareTop += rowHeight;
            ConfigureLabel(lblShareRefName, "Ref (branch/tag)", leftLabel, shareTop, labelWidth);
            ConfigureTextBoxBounds(txtShareRefName, leftInput, shareTop - 2, shareInputWidth);
            txtShareRefName.Text = !string.IsNullOrWhiteSpace(initialShare.RefName)
                ? initialShare.RefName
                : (string.IsNullOrWhiteSpace(initialPull.RefName) ? "main" : initialPull.RefName);

            grpShare.Controls.Add(lblShareBaseUrl);
            grpShare.Controls.Add(txtShareBaseUrl);
            grpShare.Controls.Add(lblShareProjectId);
            grpShare.Controls.Add(txtShareProjectId);
            grpShare.Controls.Add(lblShareRefName);
            grpShare.Controls.Add(txtShareRefName);

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

            form.Controls.Add(grpPull);
            form.Controls.Add(grpShare);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            form.Shown += (s, e) =>
            {
                txtFilePath.Focus();
                txtFilePath.SelectAll();
            };

            txtBaseUrl.KeyDown += (s, e) => MoveNextOnEnter(e, txtProjectId);
            txtProjectId.KeyDown += (s, e) => MoveNextOnEnter(e, txtRefName);
            txtRefName.KeyDown += (s, e) => MoveNextOnEnter(e, txtFilePath);
            txtFilePath.KeyDown += (s, e) => MoveNextOnEnter(e, txtShareBaseUrl);
            txtShareBaseUrl.KeyDown += (s, e) => MoveNextOnEnter(e, txtShareProjectId);
            txtShareProjectId.KeyDown += (s, e) => MoveNextOnEnter(e, txtShareRefName);
            txtShareRefName.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    btnOk.PerformClick();
                }
            };

            GitLabLastInput tempPullResult = null;
            GitLabShareInfo tempShareResult = null;

            btnOk.Click += (s, e) =>
            {
                string baseUrl = (txtBaseUrl.Text ?? "").Trim().TrimEnd('/');
                string projectId = (txtProjectId.Text ?? "").Trim();
                string refName = (txtRefName.Text ?? "").Trim();
                string filePath = (txtFilePath.Text ?? "").Trim();

                string shareBaseUrl = (txtShareBaseUrl.Text ?? "").Trim().TrimEnd('/');
                string shareProjectId = (txtShareProjectId.Text ?? "").Trim();
                string shareRefName = (txtShareRefName.Text ?? "").Trim();

                if (!ValidateBaseUrl(form, txtBaseUrl, baseUrl))
                {
                    return;
                }

                if (string.IsNullOrWhiteSpace(projectId))
                {
                    MessageBox.Show(form, "Project ID を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtProjectId.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(refName))
                {
                    refName = "main";
                }

                if (!ValidateBaseUrl(form, txtShareBaseUrl, shareBaseUrl, "共有先の Base URL を入力してください。"))
                {
                    return;
                }

                if (string.IsNullOrWhiteSpace(shareProjectId))
                {
                    MessageBox.Show(form, "共有先の Project ID を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtShareProjectId.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(shareRefName))
                {
                    shareRefName = "main";
                }

                tempPullResult = new GitLabLastInput
                {
                    BaseUrl = baseUrl,
                    ProjectId = projectId,
                    RefName = refName,
                    FilePath = filePath
                };

                tempShareResult = new GitLabShareInfo
                {
                    BaseUrl = shareBaseUrl,
                    ProjectId = shareProjectId,
                    RefName = shareRefName
                };

                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            DialogResult dr = form.ShowDialog();
            if (dr == DialogResult.OK && tempPullResult != null && tempShareResult != null)
            {
                pullResult = tempPullResult;
                shareResult = tempShareResult;
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

    private static void ConfigureTextBox(TextBox tb)
    {
        tb.AutoSize = false;
        tb.Height = 28;
        tb.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
    }

    private static void ConfigureTextBoxBounds(TextBox tb, int left, int top, int width)
    {
        tb.Left = left;
        tb.Top = top;
        tb.Width = width;
    }

    private static bool ValidateBaseUrl(Form form, TextBox textBox, string baseUrl, string emptyMessage = "Base URL を入力してください。")
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            MessageBox.Show(form, emptyMessage, "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            textBox.Focus();
            return false;
        }

        if (!baseUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
            !baseUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(form, "Base URL は http:// または https:// から始めてください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            textBox.Focus();
            return false;
        }

        return true;
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
