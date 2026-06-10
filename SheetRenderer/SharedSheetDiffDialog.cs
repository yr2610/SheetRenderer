using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

internal sealed class SharedSheetDiffDialog : Form
{
    private readonly BindingList<SharedSheetDiffEntry> entries;
    private readonly DataGridView grid;
    private readonly TextBox txtBase;
    private readonly TextBox txtLocal;
    private readonly TextBox txtRemote;
    private readonly bool showRemote;

    private SharedSheetDiffDialog(SharedSheetSelectionItem item)
    {
        List<SharedSheetDiffEntry> sourceEntries = item == null || item.DiffEntries == null
            ? new List<SharedSheetDiffEntry>()
            : item.DiffEntries.Where(x => x != null).ToList();

        entries = new BindingList<SharedSheetDiffEntry>(sourceEntries);
        showRemote = entries.Any(x => x.HasRemoteValue);

        Text = "差分確認: " + (item?.SheetName ?? item?.SheetId ?? "");
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        ShowInTaskbar = false;
        Width = showRemote ? 1180 : 980;
        Height = 680;
        MinimumSize = new Size(showRemote ? 980 : 820, 540);
        Font = new Font("Meiryo UI", 9f);

        var lblInfo = new Label
        {
            AutoSize = false,
            Left = 12,
            Top = 12,
            Width = ClientSize.Width - 24,
            Height = 28,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = entries.Count == 0
                ? "差分はありません。"
                : "差分 " + entries.Count + " 件"
        };
        Controls.Add(lblInfo);

        grid = CreateGrid(showRemote);
        grid.Left = 12;
        grid.Top = 46;
        grid.Width = ClientSize.Width - 24;
        grid.Height = 300;
        grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        grid.DataSource = entries;
        Controls.Add(grid);

        var detailPanel = new TableLayoutPanel
        {
            Left = 12,
            Top = grid.Bottom + 10,
            Width = ClientSize.Width - 24,
            Height = ClientSize.Height - grid.Bottom - 74,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ColumnCount = showRemote ? 3 : 2,
            RowCount = 2,
            Padding = new Padding(0)
        };
        Controls.Add(detailPanel);

        float columnWidth = showRemote ? 33.333f : 50f;
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, columnWidth));
        detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, columnWidth));
        if (showRemote)
        {
            detailPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, columnWidth));
        }
        detailPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24f));
        detailPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        detailPanel.Controls.Add(CreateHeaderLabel("Base"), 0, 0);
        detailPanel.Controls.Add(CreateHeaderLabel("Local"), 1, 0);
        if (showRemote)
        {
            detailPanel.Controls.Add(CreateHeaderLabel("Remote"), 2, 0);
        }

        txtBase = CreateValueTextBox();
        txtLocal = CreateValueTextBox();
        detailPanel.Controls.Add(txtBase, 0, 1);
        detailPanel.Controls.Add(txtLocal, 1, 1);

        if (showRemote)
        {
            txtRemote = CreateValueTextBox();
            detailPanel.Controls.Add(txtRemote, 2, 1);
        }

        var closeButton = new Button
        {
            Text = "閉じる",
            Left = ClientSize.Width - 102,
            Top = ClientSize.Height - 46,
            Width = 90,
            Height = 28,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            DialogResult = DialogResult.OK
        };
        Controls.Add(closeButton);
        AcceptButton = closeButton;
        CancelButton = closeButton;

        grid.SelectionChanged += (s, e) => LoadSelectedEntry();
        grid.CellDoubleClick += Grid_CellDoubleClick;
        Shown += (s, e) =>
        {
            if (grid.Rows.Count > 0)
            {
                grid.CurrentCell = grid.Rows[0].Cells[0];
            }

            LoadSelectedEntry();
        };
    }

    private SharedSheetDiffEntry SelectedEntry
    {
        get
        {
            return grid.CurrentRow == null
                ? null
                : grid.CurrentRow.DataBoundItem as SharedSheetDiffEntry;
        }
    }

    private static DataGridView CreateGrid(bool showRemote)
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

        result.Columns.Add(CreateTextColumn("CellAddress", "セル", 80));
        result.Columns.Add(CreateTextColumn("StateLabel", "状態", 92));
        result.Columns.Add(CreateTextColumn("BaseText", "Base", 240));
        result.Columns.Add(CreateTextColumn("LocalText", "Local", 240, showRemote
            ? DataGridViewAutoSizeColumnMode.None
            : DataGridViewAutoSizeColumnMode.Fill));
        if (showRemote)
        {
            result.Columns.Add(CreateTextColumn("RemoteText", "Remote", 240, DataGridViewAutoSizeColumnMode.Fill));
        }

        result.CellFormatting += (s, e) =>
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            SharedSheetDiffEntry entry = result.Rows[e.RowIndex].DataBoundItem as SharedSheetDiffEntry;
            if (entry == null)
            {
                return;
            }

            result.Rows[e.RowIndex].DefaultCellStyle.BackColor = GetStateBackColor(entry.StateLabel);
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

    private static Color GetStateBackColor(string stateLabel)
    {
        if (string.Equals(stateLabel, "競合", StringComparison.Ordinal))
        {
            return Color.MistyRose;
        }

        if (string.Equals(stateLabel, "共有先変更", StringComparison.Ordinal))
        {
            return Color.Lavender;
        }

        if (string.Equals(stateLabel, "同一変更", StringComparison.Ordinal))
        {
            return Color.LemonChiffon;
        }

        return Color.White;
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

    private static TextBox CreateValueTextBox()
    {
        return new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Margin = new Padding(2)
        };
    }

    private void LoadSelectedEntry()
    {
        SharedSheetDiffEntry entry = SelectedEntry;
        txtBase.Text = entry == null ? string.Empty : entry.BaseText;
        txtLocal.Text = entry == null ? string.Empty : entry.LocalText;
        if (showRemote && txtRemote != null)
        {
            txtRemote.Text = entry == null ? string.Empty : entry.RemoteText;
        }
    }

    private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0)
        {
            return;
        }

        var entry = grid.Rows[e.RowIndex].DataBoundItem as SharedSheetDiffEntry;
        SelectDiffCell(entry);
    }

    private void SelectDiffCell(SharedSheetDiffEntry entry)
    {
        if (entry == null ||
            string.IsNullOrWhiteSpace(entry.CellAddress) ||
            string.Equals(entry.CellAddress, "?", StringComparison.Ordinal))
        {
            return;
        }

        ExcelSelectionHelper.QueueSelectCell(
            entry.SheetName,
            entry.CellAddress,
            Text);
    }

    public static void Show(
        IWin32Window owner,
        SharedSheetSelectionItem item)
    {
        using (var dialog = new SharedSheetDiffDialog(item))
        {
            dialog.ShowDialog(owner);
        }
    }
}
