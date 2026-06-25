using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

internal sealed class SharedSheetConflictResolution
{
    public string SheetId { get; set; }
    public string SheetName { get; set; }
    public string RowId { get; set; }
    public string CellAddress { get; set; }
    public object BaseValue { get; set; }
    public object LocalValue { get; set; }
    public object RemoteValue { get; set; }
    public object ResolvedValue { get; set; }
    public string ResolutionSource { get; set; }
    public bool IsResolved { get; set; }
    internal Action<object> ApplyValue { get; set; }

    public string StatusText
    {
        get { return IsResolved ? "解決済み" : "未解決"; }
    }

    public string BaseText
    {
        get { return FormatValue(BaseValue); }
    }

    public string LocalText
    {
        get { return FormatValue(LocalValue); }
    }

    public string RemoteText
    {
        get { return FormatValue(RemoteValue); }
    }

    public string ResolvedText
    {
        get { return IsResolved ? FormatValue(ResolvedValue) : string.Empty; }
    }

    public void ApplyResolvedValue()
    {
        if (!IsResolved)
        {
            throw new InvalidOperationException("未解決の共有値競合があります。");
        }

        ApplyValue?.Invoke(ResolvedValue);
    }

    public static string FormatValue(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return "(空)";
        }

        if (value is DateTime dateTimeValue)
        {
            return dateTimeValue.ToString("yyyy-MM-dd HH:mm:ss");
        }

        return value.ToString();
    }
}

internal sealed class SharedSheetConflictResolutionDialog : Form
{
    private readonly BindingList<SharedSheetConflictResolution> conflicts;
    private readonly DataGridView grid;
    private readonly TextBox txtBase;
    private readonly TextBox txtLocal;
    private readonly TextBox txtRemote;
    private readonly TextBox txtResolved;
    private readonly Label lblRemaining;
    private readonly Button btnOk;

    private SharedSheetConflictResolutionDialog(IEnumerable<SharedSheetConflictResolution> conflicts)
    {
        this.conflicts = new BindingList<SharedSheetConflictResolution>(
            (conflicts ?? Enumerable.Empty<SharedSheetConflictResolution>()).ToList());

        Text = "共有値の競合解決";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Width = 1240;
        Height = 760;
        MinimumSize = new Size(920, 620);
        Font = new Font("Meiryo UI", 9f);

        var lblInfo = new Label
        {
            AutoSize = false,
            Left = 12,
            Top = 12,
            Width = ClientSize.Width - 24,
            Height = 38,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "ローカルと共有先の両方で変更されたセルがあります。各セルの採用値を決めてください。一覧の行をダブルクリックすると対象セルに移動します。キャンセルした場合、共有値は反映されません。"
        };
        Controls.Add(lblInfo);

        grid = CreateGrid();
        grid.Left = 12;
        grid.Top = 56;
        grid.Width = ClientSize.Width - 24;
        grid.Height = 300;
        grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        grid.DataSource = this.conflicts;
        Controls.Add(grid);

        var detailPanel = new TableLayoutPanel
        {
            Left = 12,
            Top = grid.Bottom + 10,
            Width = ClientSize.Width - 24,
            Height = ClientSize.Height - grid.Bottom - 74,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ColumnCount = 4,
            RowCount = 3,
            Padding = new Padding(0)
        };
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        detailPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24f));
        detailPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        detailPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));
        Controls.Add(detailPanel);

        detailPanel.Controls.Add(CreateHeaderLabel("ベース"), 0, 0);
        detailPanel.Controls.Add(CreateHeaderLabel("ローカル"), 1, 0);
        detailPanel.Controls.Add(CreateHeaderLabel("共有先"), 2, 0);
        detailPanel.Controls.Add(CreateHeaderLabel("採用値（自由入力可）"), 3, 0);

        txtBase = CreateValueTextBox(readOnly: true);
        txtLocal = CreateValueTextBox(readOnly: true);
        txtRemote = CreateValueTextBox(readOnly: true);
        txtResolved = CreateValueTextBox(readOnly: false);
        detailPanel.Controls.Add(txtBase, 0, 1);
        detailPanel.Controls.Add(txtLocal, 1, 1);
        detailPanel.Controls.Add(txtRemote, 2, 1);
        detailPanel.Controls.Add(txtResolved, 3, 1);

        detailPanel.Controls.Add(CreateResolveButton("ローカルを採用", () => ResolveSelected("ローカル", SelectedConflict?.LocalValue)), 1, 2);
        detailPanel.Controls.Add(CreateResolveButton("共有先を採用", () => ResolveSelected("共有先", SelectedConflict?.RemoteValue)), 2, 2);

        var customButtonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(2, 4, 0, 0)
        };
        customButtonPanel.Controls.Add(CreateResolveButton("入力値を採用", () => ResolveSelected("自由入力", txtResolved.Text), 112));
        customButtonPanel.Controls.Add(CreateResolveButton("未解決に戻す", ResetSelected, 112));
        detailPanel.Controls.Add(customButtonPanel, 3, 2);

        lblRemaining = new Label
        {
            AutoSize = false,
            Left = 12,
            Top = ClientSize.Height - 43,
            Width = 300,
            Height = 28,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        };
        Controls.Add(lblRemaining);

        btnOk = new Button
        {
            Text = "反映",
            Left = ClientSize.Width - 196,
            Top = ClientSize.Height - 46,
            Width = 90,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            DialogResult = DialogResult.OK
        };
        Controls.Add(btnOk);

        var btnCancel = new Button
        {
            Text = "キャンセル",
            Left = ClientSize.Width - 98,
            Top = ClientSize.Height - 46,
            Width = 90,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            DialogResult = DialogResult.Cancel
        };
        Controls.Add(btnCancel);

        AcceptButton = btnOk;
        CancelButton = btnCancel;

        grid.SelectionChanged += (s, e) => LoadSelectedConflict();
        grid.CellDoubleClick += Grid_CellDoubleClick;
        FormClosing += Dialog_FormClosing;
        Shown += (s, e) =>
        {
            if (grid.Rows.Count > 0)
            {
                grid.CurrentCell = grid.Rows[0].Cells[0];
            }
            LoadSelectedConflict();
            UpdateResolutionState();
        };
    }

    private SharedSheetConflictResolution SelectedConflict
    {
        get
        {
            return grid.CurrentRow == null
                ? null
                : grid.CurrentRow.DataBoundItem as SharedSheetConflictResolution;
        }
    }

    private static Label CreateHeaderLabel(string text)
    {
        return new Label
        {
            Text = text,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
    }

    private static TextBox CreateValueTextBox(bool readOnly)
    {
        return new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            ReadOnly = readOnly,
            Margin = new Padding(2)
        };
    }

    private static Button CreateResolveButton(string text, Action onClick, int width = 132)
    {
        var button = new Button
        {
            Text = text,
            Width = width,
            Height = 28,
            Margin = new Padding(2)
        };
        button.Click += (s, e) => onClick();
        return button;
    }

    private static DataGridView CreateGrid()
    {
        var result = new DataGridView
        {
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoGenerateColumns = false,
            EditMode = DataGridViewEditMode.EditProgrammatically,
            MultiSelect = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        result.Columns.Add(CreateTextColumn("StatusText", "状態", 72));
        result.Columns.Add(CreateTextColumn("SheetName", "シート名", 140));
        result.Columns.Add(CreateTextColumn("CellAddress", "セル", 70));
        result.Columns.Add(CreateTextColumn("RowId", "Row ID", 110));
        result.Columns.Add(CreateTextColumn("BaseText", "ベース", 175));
        result.Columns.Add(CreateTextColumn("LocalText", "ローカル", 175));
        result.Columns.Add(CreateTextColumn("RemoteText", "共有先", 175));
        result.Columns.Add(CreateTextColumn("ResolutionSource", "採用方法", 85));
        result.Columns.Add(CreateTextColumn("ResolvedText", "採用値", 175, DataGridViewAutoSizeColumnMode.Fill));

        result.CellFormatting += (s, e) =>
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            SharedSheetConflictResolution conflict = result.Rows[e.RowIndex].DataBoundItem as SharedSheetConflictResolution;
            if (conflict != null)
            {
                result.Rows[e.RowIndex].DefaultCellStyle.BackColor = conflict.IsResolved
                    ? Color.Honeydew
                    : Color.MistyRose;
            }
        };

        return result;
    }

    private static DataGridViewTextBoxColumn CreateTextColumn(
        string propertyName,
        string headerText,
        int width,
        DataGridViewAutoSizeColumnMode autoSizeMode = DataGridViewAutoSizeColumnMode.None)
    {
        return new DataGridViewTextBoxColumn
        {
            DataPropertyName = propertyName,
            Name = propertyName,
            HeaderText = headerText,
            Width = width,
            AutoSizeMode = autoSizeMode,
            ReadOnly = true
        };
    }

    private void LoadSelectedConflict()
    {
        SharedSheetConflictResolution conflict = SelectedConflict;
        txtBase.Text = conflict == null ? string.Empty : SharedSheetConflictResolution.FormatValue(conflict.BaseValue);
        txtLocal.Text = conflict == null ? string.Empty : SharedSheetConflictResolution.FormatValue(conflict.LocalValue);
        txtRemote.Text = conflict == null ? string.Empty : SharedSheetConflictResolution.FormatValue(conflict.RemoteValue);
        txtResolved.Text = conflict == null || !conflict.IsResolved
            ? string.Empty
            : ConvertResolvedValueToEditableText(conflict.ResolvedValue);
    }

    private static string ConvertResolvedValueToEditableText(object value)
    {
        return value == null || value == DBNull.Value ? string.Empty : value.ToString();
    }

    private void ResolveSelected(string source, object value)
    {
        SharedSheetConflictResolution conflict = SelectedConflict;
        if (conflict == null)
        {
            return;
        }

        conflict.ResolutionSource = source;
        conflict.ResolvedValue = value;
        conflict.IsResolved = true;
        RefreshConflict(conflict);
        SelectNextUnresolved();
    }

    private void ResetSelected()
    {
        SharedSheetConflictResolution conflict = SelectedConflict;
        if (conflict == null)
        {
            return;
        }

        conflict.ResolutionSource = null;
        conflict.ResolvedValue = null;
        conflict.IsResolved = false;
        RefreshConflict(conflict);
    }

    private void RefreshConflict(SharedSheetConflictResolution conflict)
    {
        int index = conflicts.IndexOf(conflict);
        if (index >= 0)
        {
            conflicts.ResetItem(index);
        }
        if (index >= 0)
        {
            grid.InvalidateRow(index);
        }
        LoadSelectedConflict();
        UpdateResolutionState();
    }

    private void SelectNextUnresolved()
    {
        UpdateResolutionState();
        if (SelectedConflict == null)
        {
            return;
        }

        int currentIndex = conflicts.IndexOf(SelectedConflict);
        for (int offset = 1; offset <= conflicts.Count; offset++)
        {
            int index = (currentIndex + offset) % conflicts.Count;
            if (!conflicts[index].IsResolved)
            {
                grid.CurrentCell = grid.Rows[index].Cells[0];
                return;
            }
        }
    }

    private void UpdateResolutionState()
    {
        int unresolvedCount = conflicts.Count(x => !x.IsResolved);
        lblRemaining.Text = "競合: " + conflicts.Count + " 件 / 未解決: " + unresolvedCount + " 件";
        btnOk.Enabled = conflicts.Count > 0 && unresolvedCount == 0;
    }

    private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0)
        {
            return;
        }

        var conflict = grid.Rows[e.RowIndex].DataBoundItem as SharedSheetConflictResolution;
        SelectConflictCell(conflict);
    }

    private void SelectConflictCell(SharedSheetConflictResolution conflict)
    {
        if (conflict == null)
        {
            return;
        }

        ExcelSelectionHelper.QueueSelectCell(
            conflict.SheetName,
            conflict.CellAddress,
            Text);
    }

    private void Dialog_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (DialogResult != DialogResult.OK)
        {
            return;
        }

        int unresolvedCount = conflicts.Count(x => !x.IsResolved);
        if (unresolvedCount <= 0)
        {
            return;
        }

        e.Cancel = true;
        MessageBox.Show(
            this,
            "未解決の競合が " + unresolvedCount + " 件あります。",
            Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
    }

    public static bool TryResolve(
        IWin32Window owner,
        IEnumerable<SharedSheetConflictResolution> conflicts)
    {
        using (var dialog = new SharedSheetConflictResolutionDialog(conflicts))
        {
            return dialog.ShowDialog(owner) == DialogResult.OK;
        }
    }
}
