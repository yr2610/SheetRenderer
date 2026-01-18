using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

internal sealed class GitLabTokenManagerDialog : Form
{
    private readonly CheckedListBox _list;
    private readonly Button _btnDeleteSelected;
    private readonly Button _btnSelectAll;
    private readonly Button _btnClearAll;
    private readonly Button _btnClose;
    private readonly Label _lblInfo;

    private readonly TokenKeyInfo[] _items;

    public GitLabTokenManagerDialog(TokenKeyInfo[] items)
    {
        _items = items ?? new TokenKeyInfo[0];

        Text = "トークン管理";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Width = 700;
        Height = 420;

        _lblInfo = new Label();
        _lblInfo.AutoSize = false;
        _lblInfo.Left = 12;
        _lblInfo.Top = 12;
        _lblInfo.Width = ClientSize.Width - 24;
        _lblInfo.Height = 40;
        _lblInfo.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _lblInfo.Text =
            "保存済みのアクセストークン一覧です。\r\n" +
            "削除したい項目にチェックを付けて「選択削除」を押してください。";
        Controls.Add(_lblInfo);

        _list = new CheckedListBox();
        _list.Left = 12;
        _list.Top = 60;
        _list.Width = ClientSize.Width - 24;
        _list.Height = ClientSize.Height - 60 - 60;
        _list.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _list.CheckOnClick = true;
        _list.HorizontalScrollbar = true;
        _list.Font = new Font("Meiryo UI", 9f);

        foreach (var item in _items)
        {
            _list.Items.Add(item.DisplayText, false);
        }

        _list.SelectedIndexChanged += (s, e) => UpdateButtons();
        _list.ItemCheck += (s, e) =>
        {
            // ItemCheck はチェック状態が変わる前に呼ばれるので、少し遅延して更新
            BeginInvoke(new Action(UpdateButtons));
        };

        Controls.Add(_list);

        int bottom = ClientSize.Height - 46;

        _btnDeleteSelected = new Button();
        _btnDeleteSelected.Text = "選択削除";
        _btnDeleteSelected.Left = 12;
        _btnDeleteSelected.Top = bottom;
        _btnDeleteSelected.Width = 110;
        _btnDeleteSelected.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        _btnDeleteSelected.Click += (s, e) => DeleteSelected();
        Controls.Add(_btnDeleteSelected);

        _btnSelectAll = new Button();
        _btnSelectAll.Text = "全選択";
        _btnSelectAll.Left = 12 + 110 + 8;
        _btnSelectAll.Top = bottom;
        _btnSelectAll.Width = 90;
        _btnSelectAll.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        _btnSelectAll.Click += (s, e) => SetAllChecked(true);
        Controls.Add(_btnSelectAll);

        _btnClearAll = new Button();
        _btnClearAll.Text = "全解除";
        _btnClearAll.Left = 12 + 110 + 8 + 90 + 8;
        _btnClearAll.Top = bottom;
        _btnClearAll.Width = 90;
        _btnClearAll.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        _btnClearAll.Click += (s, e) => SetAllChecked(false);
        Controls.Add(_btnClearAll);

        _btnClose = new Button();
        _btnClose.Text = "閉じる";
        _btnClose.Left = ClientSize.Width - 12 - 90;
        _btnClose.Top = bottom;
        _btnClose.Width = 90;
        _btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        _btnClose.Click += (s, e) => Close();
        Controls.Add(_btnClose);

        UpdateButtons();
    }

    private void UpdateButtons()
    {
        bool hasItems = _list.Items.Count > 0;
        bool hasChecked = _list.CheckedIndices.Count > 0;

        _btnDeleteSelected.Enabled = hasChecked;
        _btnSelectAll.Enabled = hasItems;
        _btnClearAll.Enabled = hasItems;
    }

    private void SetAllChecked(bool check)
    {
        _list.BeginUpdate();
        try
        {
            for (int i = 0; i < _list.Items.Count; i++)
            {
                _list.SetItemChecked(i, check);
            }
        }
        finally
        {
            _list.EndUpdate();
        }
        UpdateButtons();
    }

    private void DeleteSelected()
    {
        if (_list.CheckedIndices.Count == 0)
        {
            return;
        }

        var targets = GetCheckedTokenKeys();
        if (targets.Count == 0)
        {
            return;
        }

        string msg =
            "選択したトークンを削除しますか？\r\n\r\n" +
            string.Join("\r\n", targets.Select(x => x.DisplayText).Take(20).ToArray()) +
            (targets.Count > 20 ? "\r\n..." : "") +
            "\r\n\r\n（次回同期時に再入力が必要になります）";

        var dr = MessageBox.Show(this, msg, "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        if (dr != DialogResult.OK)
        {
            return;
        }

        int removed = 0;

        foreach (var t in targets)
        {
            if (TokenStore.Delete(t.BaseUrl, t.ProjectId))
            {
                removed++;
            }
        }

        // 画面上からも削除（後ろから消す）
        var checkedIdx = _list.CheckedIndices.Cast<int>().OrderByDescending(x => x).ToArray();
        foreach (int idx in checkedIdx)
        {
            _list.Items.RemoveAt(idx);
        }

        MessageBox.Show(this, "削除しました: " + removed + " 件", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        UpdateButtons();
    }

    private List<TokenKeyInfo> GetCheckedTokenKeys()
    {
        var list = new List<TokenKeyInfo>();

        foreach (int idx in _list.CheckedIndices)
        {
            if (idx < 0 || idx >= _items.Length)
            {
                continue;
            }
            list.Add(_items[idx]);
        }

        return list;
    }

    public static void ShowDialogSafe(IWin32Window owner)
    {
        TokenKeyInfo[] items = TokenStore.GetAllTokenKeys();
        using (var dlg = new GitLabTokenManagerDialog(items))
        {
            dlg.ShowDialog(owner);
        }
    }
}
