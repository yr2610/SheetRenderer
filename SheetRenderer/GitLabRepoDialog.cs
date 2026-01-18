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
        using (var lblFilePath = new Label())
        using (var txtFilePath = new TextBox())
        using (var btnOk = new Button())
        using (var btnCancel = new Button())
        {
            form.Text = "GitLab Repo Settings";
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ShowInTaskbar = false;
            form.ClientSize = new Size(700, 230);

            form.Font = new Font("Meiryo UI", 11f);

            // （必要なら）TextBox高さ
            txtBaseUrl.AutoSize = false;
            txtBaseUrl.Height = 28;
            txtProjectId.AutoSize = false;
            txtProjectId.Height = 28;
            txtRefName.AutoSize = false;
            txtRefName.Height = 28;
            txtFilePath.AutoSize = false;
            txtFilePath.Height = 28;

            // 初期フォーカス
            form.Shown += (s, e) =>
            {
                txtFilePath.Focus();
                txtFilePath.SelectAll();
            };

            int leftLabel = 14;
            int leftInput = 140;
            int top = 16;
            int rowH = 32;
            int inputW = 540;
            int labelW = 120;

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

            top += rowH;
            ConfigureLabel(lblFilePath, "File Path", leftLabel, top, labelW);
            ConfigureTextBox(txtFilePath, leftInput, top - 2, inputW);
            txtFilePath.Text = initial.FilePath ?? "";

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
            form.Controls.Add(lblFilePath);
            form.Controls.Add(txtFilePath);
            form.Controls.Add(btnOk);
            form.Controls.Add(btnCancel);

            // Enterキーで次の欄へ（最後はOK）
            txtBaseUrl.KeyDown += (s, e) => MoveNextOnEnter(e, txtProjectId);
            txtProjectId.KeyDown += (s, e) => MoveNextOnEnter(e, txtRefName);
            txtRefName.KeyDown += (s, e) => MoveNextOnEnter(e, txtFilePath);
            txtFilePath.KeyDown += (s, e) =>
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
                string baseUrl = (txtBaseUrl.Text ?? "").Trim();
                string projectId = (txtProjectId.Text ?? "").Trim();
                string refName = (txtRefName.Text ?? "").Trim();
                string filePath = (txtFilePath.Text ?? "").Trim();

                // 最小バリデーション：まずは「空じゃない」を強制（FilePathは運用によって空でもOKにできる）
                if (string.IsNullOrWhiteSpace(baseUrl))
                {
                    MessageBox.Show(form, "Base URL を入力してください。", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtBaseUrl.Focus();
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

                // BaseUrlは末尾スラッシュを落としておく（後段で TrimEnd('/') してもOKだが、ここで整える）
                baseUrl = baseUrl.TrimEnd('/');

                // 超ゆるいURLチェック（厳密にはしない。入力ミスの早期検出用）
                if (!baseUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
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
                    FilePath = filePath
                };

                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            // 親が無い場合でも中央に出したいので ShowDialog() でOK
            DialogResult dr = form.ShowDialog();
            if (dr == DialogResult.OK && tempResult != null)
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
