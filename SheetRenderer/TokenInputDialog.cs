using System;
using System.Drawing;
using System.Windows.Forms;

internal sealed class TokenInputDialog : Form
{
    private readonly TextBox _txtToken;
    private readonly CheckBox _chkRemember;

    public string TokenText
    {
        get
        {
            return (_txtToken.Text ?? "").Trim();
        }
    }

    public bool RememberToken
    {
        get
        {
            return _chkRemember.Checked;
        }
    }

    private TokenInputDialog(string baseUrl, string projectId)
    {
        Text = "GitLab Token";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(520, 210);

        Font = SystemFonts.MessageBoxFont;

        Label lblInfo = new Label();
        lblInfo.AutoSize = true;
        lblInfo.Location = new Point(12, 12);
        lblInfo.Text =
            "This project requires a GitLab access token.\r\n" +
            "Enter your token (masked).";

        Label lblBaseUrl = new Label();
        lblBaseUrl.AutoSize = true;
        lblBaseUrl.Location = new Point(12, 62);
        lblBaseUrl.Text = "BaseUrl:";

        TextBox txtBaseUrl = new TextBox();
        txtBaseUrl.Location = new Point(90, 58);
        txtBaseUrl.Size = new Size(410, 22);
        txtBaseUrl.ReadOnly = true;
        txtBaseUrl.Text = baseUrl;

        Label lblProjectId = new Label();
        lblProjectId.AutoSize = true;
        lblProjectId.Location = new Point(12, 90);
        lblProjectId.Text = "ProjectId:";

        TextBox txtProjectId = new TextBox();
        txtProjectId.Location = new Point(90, 86);
        txtProjectId.Size = new Size(410, 22);
        txtProjectId.ReadOnly = true;
        txtProjectId.Text = projectId;

        Label lblToken = new Label();
        lblToken.AutoSize = true;
        lblToken.Location = new Point(12, 120);
        lblToken.Text = "Token:";

        _txtToken = new TextBox();
        _txtToken.Location = new Point(90, 116);
        _txtToken.Size = new Size(410, 22);
        _txtToken.UseSystemPasswordChar = true;

        _chkRemember = new CheckBox();
        _chkRemember.AutoSize = true;
        _chkRemember.Location = new Point(90, 146);
        _chkRemember.Text = "Save this token on this PC (encrypted)";
        _chkRemember.Checked = true;

        Button btnOk = new Button();
        btnOk.Text = "OK";
        btnOk.Size = new Size(90, 28);
        btnOk.Location = new Point(320, 172);
        btnOk.DialogResult = DialogResult.OK;

        Button btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.Size = new Size(90, 28);
        btnCancel.Location = new Point(410, 172);
        btnCancel.DialogResult = DialogResult.Cancel;

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        Controls.Add(lblInfo);
        Controls.Add(lblBaseUrl);
        Controls.Add(txtBaseUrl);
        Controls.Add(lblProjectId);
        Controls.Add(txtProjectId);
        Controls.Add(lblToken);
        Controls.Add(_txtToken);
        Controls.Add(_chkRemember);
        Controls.Add(btnOk);
        Controls.Add(btnCancel);

        Shown += (s, e) =>
        {
            _txtToken.Focus();
        };

        btnOk.Click += (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(TokenText))
            {
                MessageBox.Show(this, "Token is empty.", "GitLab Token", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
        };
    }

    public static DialogResult ShowDialog(IWin32Window owner, string baseUrl, string projectId, out string token, out bool remember)
    {
        token = null;
        remember = false;

        using (TokenInputDialog dlg = new TokenInputDialog(baseUrl, projectId))
        {
            DialogResult result = dlg.ShowDialog(owner);
            if (result != DialogResult.OK)
            {
                return result;
            }

            token = dlg.TokenText;
            remember = dlg.RememberToken;
            return result;
        }
    }
}
