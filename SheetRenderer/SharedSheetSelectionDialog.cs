using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

internal sealed class SharedSheetSelectionItem
{
    public bool Selected { get; set; }
    public string SheetName { get; set; }
    public string SheetId { get; set; }
    public string ActionLabel { get; set; }
    public string StatusDetail { get; set; }
    public bool HasConflict { get; set; }
    public string DiffText { get; set; }
    public SharedSheetDocument Document { get; set; }
}

internal sealed class SharedSheetSelectionDialog : Form
{
    private readonly DataGridView grid;
    private readonly Button btnOk;
    private readonly Button btnCancel;
    private readonly Label lblInfo;
    private readonly BindingSource bindingSource;
    private readonly List<SharedSheetSelectionItem> items;

    public SharedSheetSelectionDialog(
        IEnumerable<SharedSheetSelectionItem> items,
        string infoText = null,
        bool readOnlyMode = false,
        string okButtonText = "共有開始")
    {
        this.items = (items ?? Enumerable.Empty<SharedSheetSelectionItem>()).ToList();

        Text = "変更共有";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Width = 920;
        Height = 460;

        lblInfo = new Label();
        lblInfo.AutoSize = false;
        lblInfo.Left = 12;
        lblInfo.Top = 12;
        lblInfo.Width = ClientSize.Width - 24;
        lblInfo.Height = 40;
        lblInfo.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        lblInfo.Text = string.IsNullOrWhiteSpace(infoText)
            ? "共有するシートを選択してください。ダブルクリックで差分を確認できます。"
            : infoText;
        Controls.Add(lblInfo);

        grid = new DataGridView();
        grid.Left = 12;
        grid.Top = 60;
        grid.Width = ClientSize.Width - 24;
        grid.Height = ClientSize.Height - 120;
        grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersVisible = false;
        grid.MultiSelect = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AutoGenerateColumns = false;
        grid.EditMode = DataGridViewEditMode.EditOnEnter;
        grid.Font = new Font("Meiryo UI", 9f);

        grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = "Selected",
            Name = "Selected",
            HeaderText = "",
            Width = 40,
            ReadOnly = readOnlyMode
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "SheetName",
            Name = "SheetName",
            HeaderText = "シート名",
            Width = 220,
            ReadOnly = true
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "ActionLabel",
            Name = "ActionLabel",
            HeaderText = "状態",
            Width = 80,
            ReadOnly = true
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "StatusDetail",
            Name = "StatusDetail",
            HeaderText = "詳細",
            Width = 260,
            ReadOnly = true
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "SheetId",
            Name = "SheetId",
            HeaderText = "Sheet ID",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            ReadOnly = true
        });

        bindingSource = new BindingSource();
        bindingSource.DataSource = this.items;
        grid.DataSource = bindingSource;
        Controls.Add(grid);

        btnOk = new Button();
        btnOk.Text = okButtonText;
        btnOk.Left = ClientSize.Width - 196;
        btnOk.Top = ClientSize.Height - 46;
        btnOk.Width = 90;
        btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnOk.DialogResult = DialogResult.OK;
        Controls.Add(btnOk);

        btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.Left = ClientSize.Width - 98;
        btnCancel.Top = ClientSize.Height - 46;
        btnCancel.Width = 90;
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        btnCancel.DialogResult = DialogResult.Cancel;
        Controls.Add(btnCancel);

        if (readOnlyMode)
        {
            btnCancel.Visible = false;
            btnOk.Left = ClientSize.Width - btnOk.Width - 12;
        }

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        if (!readOnlyMode)
        {
            grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (grid.IsCurrentCellDirty)
                {
                    grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
        }

        grid.CellDoubleClick += (s, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= this.items.Count)
            {
                return;
            }

            SharedSheetSelectionItem item = this.items[e.RowIndex];
            if (item == null)
            {
                return;
            }

            ShowDiffDialog(this, item);
        };

        grid.DataBindingComplete += (s, e) =>
        {
            foreach (DataGridViewRow row in grid.Rows)
            {
                SharedSheetSelectionItem item = row.DataBoundItem as SharedSheetSelectionItem;
                if (item == null)
                {
                    continue;
                }

                row.DefaultCellStyle.BackColor = Color.White;
                if (item.HasConflict)
                {
                    row.DefaultCellStyle.BackColor = Color.MistyRose;
                }
                else if (string.Equals(item.ActionLabel, "マージ", StringComparison.Ordinal))
                {
                    row.DefaultCellStyle.BackColor = Color.LemonChiffon;
                }
                else if (string.Equals(item.ActionLabel, "新規", StringComparison.Ordinal))
                {
                    row.DefaultCellStyle.BackColor = Color.Honeydew;
                }
            }
        };

        Shown += (s, e) =>
        {
            if (grid.Rows.Count > 0)
            {
                grid.CurrentCell = grid.Rows[0].Cells[0];
            }
        };
    }

    private static void ShowDiffDialog(IWin32Window owner, SharedSheetSelectionItem item)
    {
        using (var form = new Form())
        using (var textBox = new TextBox())
        using (var closeButton = new Button())
        {
            form.Text = "差分確認: " + (item.SheetName ?? item.SheetId ?? "");
            form.StartPosition = FormStartPosition.CenterParent;
            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.MinimizeBox = false;
            form.MaximizeBox = true;
            form.ShowInTaskbar = false;
            form.Width = 980;
            form.Height = 620;
            form.Font = new Font("Meiryo UI", 9f);

            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Both;
            textBox.ReadOnly = true;
            textBox.WordWrap = false;
            textBox.Left = 12;
            textBox.Top = 12;
            textBox.Width = form.ClientSize.Width - 24;
            textBox.Height = form.ClientSize.Height - 56;
            textBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBox.Font = new Font("Meiryo UI", 9f);
            textBox.Text = string.IsNullOrWhiteSpace(item.DiffText)
                ? "差分はありません。"
                : item.DiffText;

            closeButton.Text = "閉じる";
            closeButton.Width = 90;
            closeButton.Height = 28;
            closeButton.Left = form.ClientSize.Width - closeButton.Width - 12;
            closeButton.Top = form.ClientSize.Height - closeButton.Height - 12;
            closeButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            closeButton.DialogResult = DialogResult.OK;

            form.Controls.Add(textBox);
            form.Controls.Add(closeButton);
            form.AcceptButton = closeButton;
            form.CancelButton = closeButton;

            form.ShowDialog(owner);
        }
    }

    public static void ShowDiff(
        IWin32Window owner,
        SharedSheetSelectionItem item)
    {
        if (item == null)
        {
            return;
        }

        ShowDiffDialog(owner, item);
    }

    public static bool TryShowRevertConfirmation(
        IWin32Window owner,
        SharedSheetSelectionItem item,
        string infoText = null,
        string okButtonText = "取り消す")
    {
        if (item == null)
        {
            return false;
        }

        using (var form = new Form())
        using (var infoLabel = new Label())
        using (var textBox = new TextBox())
        using (var okButton = new Button())
        using (var cancelButton = new Button())
        {
            form.Text = "変更の取り消し: " + (item.SheetName ?? item.SheetId ?? "");
            form.StartPosition = FormStartPosition.CenterParent;
            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.MinimizeBox = false;
            form.MaximizeBox = true;
            form.ShowInTaskbar = false;
            form.Width = 980;
            form.Height = 620;
            form.Font = new Font("Meiryo UI", 9f);

            infoLabel.AutoSize = false;
            infoLabel.Left = 12;
            infoLabel.Top = 12;
            infoLabel.Width = form.ClientSize.Width - 24;
            infoLabel.Height = 44;
            infoLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            infoLabel.Text = string.IsNullOrWhiteSpace(infoText)
                ? "このシートの未共有の変更を取り消します。元に戻せません。本当に取り消しますか？"
                : infoText;

            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Both;
            textBox.ReadOnly = true;
            textBox.WordWrap = false;
            textBox.Left = 12;
            textBox.Top = infoLabel.Bottom + 8;
            textBox.Width = form.ClientSize.Width - 24;
            textBox.Height = form.ClientSize.Height - 108;
            textBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBox.Font = new Font("Meiryo UI", 9f);
            textBox.Text = string.IsNullOrWhiteSpace(item.DiffText)
                ? "差分はありません。"
                : item.DiffText;

            okButton.Text = okButtonText;
            okButton.Width = 90;
            okButton.Height = 28;
            okButton.Left = form.ClientSize.Width - 192;
            okButton.Top = form.ClientSize.Height - okButton.Height - 12;
            okButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            okButton.DialogResult = DialogResult.OK;

            cancelButton.Text = "キャンセル";
            cancelButton.Width = 90;
            cancelButton.Height = 28;
            cancelButton.Left = form.ClientSize.Width - 96;
            cancelButton.Top = form.ClientSize.Height - cancelButton.Height - 12;
            cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            cancelButton.DialogResult = DialogResult.Cancel;

            form.Controls.Add(infoLabel);
            form.Controls.Add(textBox);
            form.Controls.Add(okButton);
            form.Controls.Add(cancelButton);
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            return form.ShowDialog(owner) == DialogResult.OK;
        }
    }

    public List<SharedSheetSelectionItem> GetSelectedItems()
    {
        return items.Where(x => x != null && x.Selected && x.Document != null).ToList();
    }

    public static void ShowConflictReview(
        IWin32Window owner,
        IEnumerable<SharedSheetSelectionItem> items)
    {
        using (var dialog = new SharedSheetSelectionDialog(
            items,
            "競合があるため変更共有できません。先に最新版取得を実行してください。ダブルクリックで差分を確認できます。",
            readOnlyMode: true,
            okButtonText: "閉じる"))
        {
            dialog.ShowDialog(owner);
        }
    }

    public static bool TryShow(
        IWin32Window owner,
        IEnumerable<SharedSheetSelectionItem> items,
        out List<SharedSheetSelectionItem> selectedItems)
    {
        selectedItems = null;

        using (var dialog = new SharedSheetSelectionDialog(items))
        {
            DialogResult result = dialog.ShowDialog(owner);
            if (result != DialogResult.OK)
            {
                return false;
            }

            selectedItems = dialog.GetSelectedItems();
            return true;
        }
    }
}
