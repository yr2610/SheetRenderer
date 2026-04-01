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
    public SharedSheetDocument Document { get; set; }
}

internal sealed class SharedSheetSelectionDialog : Form
{
    private readonly DataGridView _grid;
    private readonly Button _btnOk;
    private readonly Button _btnCancel;
    private readonly Label _lblInfo;
    private readonly BindingSource _bindingSource;

    private readonly List<SharedSheetSelectionItem> _items;

    public SharedSheetSelectionDialog(IEnumerable<SharedSheetSelectionItem> items)
    {
        _items = (items ?? Enumerable.Empty<SharedSheetSelectionItem>()).ToList();

        Text = "変更共有";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Width = 760;
        Height = 460;

        _lblInfo = new Label();
        _lblInfo.AutoSize = false;
        _lblInfo.Left = 12;
        _lblInfo.Top = 12;
        _lblInfo.Width = ClientSize.Width - 24;
        _lblInfo.Height = 40;
        _lblInfo.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _lblInfo.Text = "共有するシートを選択してください。";
        Controls.Add(_lblInfo);

        _grid = new DataGridView();
        _grid.Left = 12;
        _grid.Top = 60;
        _grid.Width = ClientSize.Width - 24;
        _grid.Height = ClientSize.Height - 120;
        _grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.RowHeadersVisible = false;
        _grid.MultiSelect = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.AutoGenerateColumns = false;
        _grid.EditMode = DataGridViewEditMode.EditOnEnter;
        _grid.Font = new Font("Meiryo UI", 9f);

        _grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = "Selected",
            Name = "Selected",
            HeaderText = "",
            Width = 40
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "SheetName",
            Name = "SheetName",
            HeaderText = "シート名",
            Width = 260,
            ReadOnly = true
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "ActionLabel",
            Name = "ActionLabel",
            HeaderText = "操作",
            Width = 80,
            ReadOnly = true
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = "SheetId",
            Name = "SheetId",
            HeaderText = "Sheet ID",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            ReadOnly = true
        });

        _bindingSource = new BindingSource();
        _bindingSource.DataSource = _items;
        _grid.DataSource = _bindingSource;
        Controls.Add(_grid);

        _btnOk = new Button();
        _btnOk.Text = "共有開始";
        _btnOk.Left = ClientSize.Width - 196;
        _btnOk.Top = ClientSize.Height - 46;
        _btnOk.Width = 90;
        _btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        _btnOk.DialogResult = DialogResult.OK;
        Controls.Add(_btnOk);

        _btnCancel = new Button();
        _btnCancel.Text = "Cancel";
        _btnCancel.Left = ClientSize.Width - 98;
        _btnCancel.Top = ClientSize.Height - 46;
        _btnCancel.Width = 90;
        _btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        _btnCancel.DialogResult = DialogResult.Cancel;
        Controls.Add(_btnCancel);

        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        _grid.CurrentCellDirtyStateChanged += (s, e) =>
        {
            if (_grid.IsCurrentCellDirty)
            {
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        };

        _grid.CellDoubleClick += (s, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= _items.Count)
            {
                return;
            }

            SharedSheetSelectionItem item = _items[e.RowIndex];
            if (item == null)
            {
                return;
            }

            MessageBox.Show(
                this,
                "差分表示は今後追加予定です。\n\n" + item.SheetName,
                "変更共有",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        };

        Shown += (s, e) =>
        {
            if (_grid.Rows.Count > 0)
            {
                _grid.CurrentCell = _grid.Rows[0].Cells[0];
            }
        };
    }

    public List<SharedSheetSelectionItem> GetSelectedItems()
    {
        return _items.Where(x => x != null && x.Selected).ToList();
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
